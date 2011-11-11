using System;
using System.Collections.Generic;
using System.Linq;
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
			Application.SetCompatibleTextRenderingDefault( false );
			Application.Run(new MainForm());
		}
	}

	class MainForm : Form
	{
		private NotifyIcon notifyIcon;
		private ContextMenuStrip contextMenuStrip;
		private ToolStripMenuItem exitItem;
		private ToolStripMenuItem settingItem;
		private List<ClipboardItem> listClipboardItem;
		
		public MainForm()
		{
			contextMenuStrip = new ContextMenuStrip();

			exitItem = new ToolStripMenuItem();
			exitItem.Text = "Exit";
			exitItem.Click += ( o, e ) => { CloseApplication(); };

			settingItem = new ToolStripMenuItem();
			settingItem.Text = "Setting";
			settingItem.Click += ( o, e ) => { AddClipboardItem(); };
			
			notifyIcon = new NotifyIcon();
			notifyIcon.Text = "Clipboard Manager CS";
			notifyIcon.Icon = new Icon(Properties.Resources.Icon1, 40, 40);
			notifyIcon.ContextMenuStrip = contextMenuStrip;
			
			notifyIcon.Click += ( s, a ) => {
				var args = ( MouseEventArgs )a;
				contextMenuStrip.Items.Clear();

				foreach( var ci in listClipboardItem )
					contextMenuStrip.Items.Add( ci.MenuItem );

				contextMenuStrip.Items.Add( "-" );
				contextMenuStrip.Items.Add( settingItem );
				contextMenuStrip.Items.Add( exitItem );

				MethodInfo mi = typeof( NotifyIcon ).GetMethod( "ShowContextMenu", BindingFlags.Instance | BindingFlags.NonPublic );
				mi.Invoke( notifyIcon, null );
			};

			listClipboardItem = new List<ClipboardItem>();
		}

		protected override void OnLoad( EventArgs e )
		{
			MainForm.keyListener = this;
			hookID = SetHook( hookCallback );
			notifyIcon.Visible = true;
			this.Visible = false;
			this.ShowInTaskbar = false;
			base.OnLoad( e );
		}

		private void CloseApplication()
		{
			notifyIcon.Visible = false;
			UnhookWindowsHookEx( hookID );
			this.Close();
		}

		private void AddClipboardItem()
		{
			listClipboardItem.Add( GetClipboard() );

			if( listClipboardItem.Count > 10 )
				listClipboardItem.RemoveAt( 0 );
		}

		private void menuItemClickHandler( Object sender, EventArgs args )
		{
			var menuItem = ( ToolStripMenuItem )sender;

			foreach( var ci in listClipboardItem )
				ci.MenuItem.Checked = false;

			menuItem.Checked = true;
			
			SetClipboard( ( ClipboardItem )menuItem.Tag );
		}

		private ClipboardItem GetClipboard()
		{
			var clipboardItem = new ClipboardItem();
			var dataObject = Clipboard.GetDataObject();
			var formats = dataObject.GetFormats();

			foreach( var format in formats )
			{
				var data = dataObject.GetData( format );
				clipboardItem.Objects.Add( format, data );
			}

			var menuItem = new ToolStripMenuItem();
			clipboardItem.MenuItem = menuItem;
			menuItem.Tag = clipboardItem;
			menuItem.Click += menuItemClickHandler;
			menuItem.Text = DateTime.Now.ToString();

			//mark the current clipboard as checked
			foreach( var ci in listClipboardItem )
				ci.MenuItem.Checked = false;

			menuItem.Checked = true;

			//determine mouse hover tooltip
			menuItem.AutoToolTip = true;
			menuItem.ToolTipText = (string)dataObject.GetData( "Text" );

			if( menuItem.ToolTipText == null )
			{
				var data = ( string[] )dataObject.GetData( "FileName" );
				menuItem.ToolTipText = string.Join( "\n", data );
			}

			return clipboardItem;
		}

		private void SetClipboard( ClipboardItem clipboardItem )
		{
			var dataObject = new DataObject();

			foreach( var pair in clipboardItem.Objects )
				dataObject.SetData( pair.Key, pair.Value );

			Clipboard.SetDataObject( dataObject, true, 5, 100 );
		}

		private const int WH_KEYBOARD_LL = 13;
		private const int WM_KEYDOWN = 0x0100;
		private const int WM_KEYUP = 0x0101;

		private static LowLevelKeyboardProc hookCallback = HookCallback;
		private static IntPtr hookID = IntPtr.Zero;
		private static bool[] keyPressedDown = new bool[256];
		private static MainForm keyListener;
		private static bool copyToggleDown = false;

		private static IntPtr SetHook( LowLevelKeyboardProc proc )
		{
			using( Process curProcess = Process.GetCurrentProcess() )
			using( ProcessModule curModule = curProcess.MainModule )
			{
				return SetWindowsHookEx( WH_KEYBOARD_LL, proc, GetModuleHandle( curModule.ModuleName ), 0 );
			}
		}

		private delegate IntPtr LowLevelKeyboardProc( int nCode, IntPtr wParam, IntPtr lParam );

		public static IntPtr HookCallback( int nCode, IntPtr wParam, IntPtr lParam )
		{
			int vkCode = Marshal.ReadInt32( lParam );

			if( nCode >= 0 )
			{
				if( wParam == ( IntPtr )WM_KEYDOWN )
					keyPressedDown[ vkCode ] = true;

				if( wParam == ( IntPtr )WM_KEYUP )
					keyPressedDown[ vkCode ] = false;
			}

			if( !copyToggleDown && keyPressedDown[ 162 ] && keyPressedDown[ 67 ] )
			{
				copyToggleDown = true;
			}
			else if( copyToggleDown && !keyPressedDown[ 162 ] && !keyPressedDown[ 67 ] )
			{
				copyToggleDown = false;
				keyListener.AddClipboardItem();
			}

			return CallNextHookEx( hookID, nCode, wParam, lParam );
		}

		[DllImport( "user32.dll", CharSet = CharSet.Auto, SetLastError = true )]
		private static extern IntPtr SetWindowsHookEx( int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId );

		[DllImport( "user32.dll", CharSet = CharSet.Auto, SetLastError = true )]
		[return: MarshalAs( UnmanagedType.Bool )]
		private static extern bool UnhookWindowsHookEx( IntPtr hhk );

		[DllImport( "user32.dll", CharSet = CharSet.Auto, SetLastError = true )]
		private static extern IntPtr CallNextHookEx( IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam );

		[DllImport( "kernel32.dll", CharSet = CharSet.Auto, SetLastError = true )]
		private static extern IntPtr GetModuleHandle( string lpModuleName );
	}

	internal class ClipboardItem
	{
		public Dictionary<String, Object> Objects = new Dictionary<String, Object>();
		public ToolStripMenuItem MenuItem;
	}
}
