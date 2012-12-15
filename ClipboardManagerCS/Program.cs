using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Drawing;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Reflection;

namespace ClipboardManagerCS
{
	static class Program
	{
		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main()
		{
			Application.EnableVisualStyles();
			Application.SetCompatibleTextRenderingDefault(false);
			Application.Run(new MainForm());
		}
	}

	class MainForm : Form
	{
		private const string releaseVersion = "1.0.1.0";

		private NotifyIcon notifyIcon;
		private ContextMenuStrip contextMenuStrip;
		private ToolStripMenuItem exitItem;
		private ToolStripMenuItem clearItem;
		private ToolStripMenuItem ignoreShiftDelItem;
		private List<ClipboardItem> listClipboardItem;
		private ClipboardItem currentClipboardItem;

		public MainForm()
		{
			listClipboardItem = new List<ClipboardItem>();
			contextMenuStrip = new ContextMenuStrip();

			clearItem = new ToolStripMenuItem();
			clearItem.Text = "Auto Clear Formatting";
			clearItem.Checked = true;
			clearItem.Click += (o, e) =>
			{
				clearItem.Checked = !clearItem.Checked;
				SetClipboard(currentClipboardItem);
			};

			ignoreShiftDelItem = new ToolStripMenuItem();
			ignoreShiftDelItem.Text = "Ignore Shift+Del Clipboard Changes";
			ignoreShiftDelItem.Checked = true;
			ignoreShiftDelItem.Click += (o, e) =>
			{
				ignoreShiftDelItem.Checked = !ignoreShiftDelItem.Checked;
			};

			exitItem = new ToolStripMenuItem();
			exitItem.Text = "Exit";
			exitItem.Click += (o, e) => { CloseApplication(); };

			notifyIcon = new NotifyIcon();
			notifyIcon.Text = "ClipboardManagerCS " + releaseVersion;
			notifyIcon.Icon = new Icon(Properties.Resources.Icon1, 40, 40);
			notifyIcon.ContextMenuStrip = contextMenuStrip;

			notifyIcon.Click += (s, a) =>
			{
				try
				{
					var args = (MouseEventArgs)a;
					contextMenuStrip.Items.Clear();

					foreach(var ci in listClipboardItem)
						contextMenuStrip.Items.Add(ci.MenuItem);

					contextMenuStrip.Items.Add("-");
					contextMenuStrip.Items.Add(clearItem);
					contextMenuStrip.Items.Add(ignoreShiftDelItem);
					contextMenuStrip.Items.Add(exitItem);

					MethodInfo mi = typeof(NotifyIcon).GetMethod("ShowContextMenu", BindingFlags.Instance | BindingFlags.NonPublic);
					mi.Invoke(notifyIcon, null);
				}
				catch(Exception ex)
				{
					Console.WriteLine(ex);
				}
			};
		}

		protected override void OnLoad(EventArgs e)
		{
			MainForm.keyListener = this;
			hookID = SetHook(hookCallback);
			notifyIcon.Visible = true;
			this.Visible = false;
			this.ShowInTaskbar = false;
			base.OnLoad(e);
		}

		private void CloseApplication()
		{
			notifyIcon.Visible = false;
			UnhookWindowsHookEx(hookID);
			this.Close();
		}

		private void AddClipboardItem()
		{
			currentClipboardItem = GetClipboard();
			listClipboardItem.Add(currentClipboardItem);

			//remove last item if list is greater than maximum size
			if(listClipboardItem.Count > 10)
				listClipboardItem.RemoveAt(0);

			AutoClearFormatting();
		}

		private void AutoClearFormatting()
		{
			if(!clearItem.Checked)
				return;

			var dataObject = Clipboard.GetDataObject();
			Type format = typeof(string);

			if(!dataObject.GetDataPresent(format))
				return;

			Clipboard.SetData(format.ToString(), dataObject.GetData(format));
		}

		private ClipboardItem GetClipboard()
		{
			var clipboardItem = new ClipboardItem();
			var dataObject = Clipboard.GetDataObject();
			var formats = dataObject.GetFormats();

			foreach(var format in formats)
			{
				var data = dataObject.GetData(format);
				clipboardItem.Objects.Add(format, data);
			}

			var menuItem = new ToolStripMenuItem();
			clipboardItem.MenuItem = menuItem;
			menuItem.Tag = clipboardItem;
			menuItem.Click += menuItemClickHandler;
			menuItem.Text = DateTime.Now.ToString();

			//mark the current clipboard as checked
			foreach(var ci in listClipboardItem)
				ci.MenuItem.Checked = false;

			menuItem.Checked = true;

			//determine clipboard content text
			var text = (string)dataObject.GetData("Text");

			if(text == null)
			{
				var data = (string[])dataObject.GetData("FileName");
				text = string.Join("\n", data);
			}

			menuItem.DropDown.Items.Add(text);

			return clipboardItem;
		}

		private void SetClipboard(ClipboardItem clipboardItem)
		{
			var dataObject = new DataObject();

			foreach(var pair in clipboardItem.Objects)
				dataObject.SetData(pair.Key, pair.Value);

			Clipboard.SetDataObject(dataObject, true, 5, 100);
			AutoClearFormatting();

			//uncheck all other items
			foreach(var ci in listClipboardItem)
				ci.MenuItem.Checked = false;

			//check current item
			clipboardItem.MenuItem.Checked = true;

			currentClipboardItem = clipboardItem;
		}

		private void menuItemClickHandler(Object sender, EventArgs args)
		{
			var menuItem = (ToolStripMenuItem)sender;
			SetClipboard((ClipboardItem)menuItem.Tag);
		}

		//get the item older than current item
		public void BackClipboardItem()
		{
			//if there is no older item then do nothing
			if(listClipboardItem.Count <= 1)
				return;

			var backItem = currentClipboardItem;

			//look from the oldest item in list
			for(int i = 1; i < listClipboardItem.Count; i++)
			{
				if(listClipboardItem[i] == currentClipboardItem)
					SetClipboard(listClipboardItem[i - 1]);
			}
		}

		//ignore clipboard changes
		public void IgnoreClipboardItem()
		{
			if(!ignoreShiftDelItem.Checked)
				return;

			//if there is no previous clipboard item then do nothing
			if(currentClipboardItem == null)
				return;

			SetClipboard(currentClipboardItem);
		}

		private const int WH_KEYBOARD_LL = 13;
		private const int WM_KEYDOWN = 0x0100;
		private const int WM_KEYUP = 0x0101;

		private static LowLevelKeyboardProc hookCallback = HookCallback;
		private static IntPtr hookID = IntPtr.Zero;
		private static bool[] keyPressedDown = new bool[256];
		private static MainForm keyListener;
		private static bool copyToggleDown = false;
		private static bool backToggleDown = false;
		private static bool ignoreToggleDown = false;

		private static IntPtr SetHook(LowLevelKeyboardProc proc)
		{
			using(Process curProcess = Process.GetCurrentProcess())
			using(ProcessModule curModule = curProcess.MainModule)
			{
				return SetWindowsHookEx(WH_KEYBOARD_LL, proc, GetModuleHandle(curModule.ModuleName), 0);
			}
		}

		private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

		public static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
		{
			int vkCode = Marshal.ReadInt32(lParam);

			if(nCode >= 0)
			{
				if(wParam == (IntPtr)WM_KEYDOWN)
					keyPressedDown[vkCode] = true;

				if(wParam == (IntPtr)WM_KEYUP)
					keyPressedDown[vkCode] = false;
			}

			if(!copyToggleDown && keyPressedDown[162] && keyPressedDown[67])
			{
				//ctrl+c down
				copyToggleDown = true;
			}
			else if(copyToggleDown && !keyPressedDown[162] && !keyPressedDown[67])
			{
				//ctrl+c up
				copyToggleDown = false;
				keyListener.AddClipboardItem();
			}

			if(!backToggleDown && keyPressedDown[162] && keyPressedDown[66])
			{
				//ctrl+b down
				backToggleDown = true;
			}
			else if(backToggleDown && !keyPressedDown[162] && !keyPressedDown[66])
			{
				//ctrl+b up
				backToggleDown = false;
				keyListener.BackClipboardItem();
			}

			if(!ignoreToggleDown && keyPressedDown[160] && keyPressedDown[46])
			{
				//shift+delete down
				ignoreToggleDown = true;
			}
			else if(ignoreToggleDown && !keyPressedDown[160] && !keyPressedDown[46])
			{
				//shift+delete up
				ignoreToggleDown = false;
				keyListener.IgnoreClipboardItem();
			}

			return CallNextHookEx(hookID, nCode, wParam, lParam);
		}

		[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

		[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		[return: MarshalAs(UnmanagedType.Bool)]
		private static extern bool UnhookWindowsHookEx(IntPtr hhk);

		[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

		[DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		private static extern IntPtr GetModuleHandle(string lpModuleName);
	}

	internal class ClipboardItem
	{
		public Dictionary<String, Object> Objects = new Dictionary<String, Object>();
		public ToolStripMenuItem MenuItem;
	}
}
