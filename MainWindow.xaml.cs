using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Automation;
using Condition = System.Windows.Automation;

namespace PCBYWPF
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
	{
		WinEventDelegate eventDelegate = null;		
		private System.Windows.Forms.NotifyIcon notificationObject;
		private WindowState normalWindowState = WindowState.Normal;

		public MainWindow()
		{			
			InitializeComponent();

			HorizontalCenterAlign();

			CreateGlobalWindowHook();
			CreateNotification();
		}

		#region Methods

		/// <summary>
		/// Create global windows hook that monitors a change of active window
		/// </summary>
		public void CreateGlobalWindowHook()
		{
			eventDelegate = new WinEventDelegate(WinEventProc);
			IntPtr m_hhook = SetWinEventHook(EVENT_SYSTEM_FOREGROUND, EVENT_SYSTEM_FOREGROUND, IntPtr.Zero, eventDelegate, 0, 0, WINEVENT_OUTOFCONTEXT);
		}

		public void WinEventProc(IntPtr hWinEventHook, uint eventType, IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
		{
			ActiveWindowCheck();
		}

		/// <summary>
		/// Create Notification Object and Hide Application
		/// </summary>
		public void CreateNotification()
		{
			notificationObject = new System.Windows.Forms.NotifyIcon();
			notificationObject.BalloonTipText = "Desktop Icon Finder running in background";
			notificationObject.BalloonTipTitle = "Desktop Icon Finder";
			notificationObject.Text = "Desktop Icon Finder";
			notificationObject.Icon = new System.Drawing.Icon("theressaApp.ico");
			notificationObject.Click += new EventHandler(notificationObject_Click);

			//Start application minimised
			WindowState = WindowState.Minimized;
			Hide();
			CheckTrayIcon();
			if (notificationObject != null)
				notificationObject.ShowBalloonTip(3000);
		}

		/// <summary>
		/// Check current active window
		/// </summary>
		private void ActiveWindowCheck()
		{
			const int nChars = 256;
			IntPtr handle = GetForegroundWindow();//Get Active Foreground Window
			StringBuilder Buff = new StringBuilder(nChars);

			//Get Active Window Title
			if (GetWindowText(handle, Buff, 256) > 0)
			{
				//If Active Window is Program Manager
				if (Buff.ToString() == "Program Manager")
				{
					//Pop up desktop icon finder application
					Show();
					WindowState = normalWindowState;

					if (handle != IntPtr.Zero)
					{
						//Run UIAutomation time consuming task as a parallel task
						DoUIAutomation_Task();
					}
				}
			}
		}
		/// <summary>
		/// Parallel task to handle UIAutomation Job
		/// </summary>
		private async void DoUIAutomation_Task()
		{
			desktopIconList.Text = "Loading Icon List..";
			desktopIconCount.Text = "Loading Icon Count..";

			var result = await Task.Run(() =>
			{
				return GetDesktopIconListAndCount();
			});

			if(desktopIconList != null)
				desktopIconList.Text = result.Value;

			if (desktopIconCount != null)
				desktopIconCount.Text = result.Key.ToString();
		}

		/// <summary>
		/// Get Desktop Icon List And Count
		/// </summary>
		/// <returns></returns>
		public KeyValuePair<int, string> GetDesktopIconListAndCount()
		{
			string iconList = "";
			int count = 0;
			
			AutomationElement elementWindow = AutomationElement.RootElement;//Get the Root Element

			AutomationElement programElement = elementWindow.FindFirst(TreeScope.Children, new PropertyCondition
			(AutomationElement.NameProperty, "Program Manager"));//Get Program Manager

			AutomationElement desktopElement = programElement.FindFirst(TreeScope.Children, new PropertyCondition
			(AutomationElement.NameProperty, "Desktop"));//Get Desktop

			var desktopElementList = desktopElement.FindAll(TreeScope.Children, Condition.Condition.TrueCondition);
			if (desktopElementList != null && desktopElementList.Count > 0)
			{
				count = desktopElementList.Count;
				foreach (AutomationElement item in desktopElementList)
				{
					iconList += item.Current.Name + "\r\n";
				}
			}
			return new KeyValuePair<int, string>(count, iconList);
		}
		/// <summary>
		/// Horizontal Center Align
		/// </summary>
		private void HorizontalCenterAlign()
		{
			this.Left = (System.Windows.SystemParameters.PrimaryScreenWidth / 2) - (this.Width / 2);
			this.Top = (System.Windows.SystemParameters.PrimaryScreenHeight / 2) - (this.Height / 2);
		}
		#endregion

		#region Events		

		void OnClose(object sender, CancelEventArgs args)
		{
			notificationObject.Dispose();
			notificationObject = null;
		}		
		void OnStateChanged(object sender, EventArgs args)
		{
			if (WindowState == WindowState.Minimized)
			{
				Hide();
				if (notificationObject != null)
					notificationObject.ShowBalloonTip(2000);
			}
			else
				normalWindowState = WindowState;
		}
		void OnIsVisibleChanged(object sender, DependencyPropertyChangedEventArgs args)
		{
			CheckTrayIcon();
		}
		void notificationObject_Click(object sender, EventArgs e)
		{
			Show();
			WindowState = normalWindowState;
		}
		void CheckTrayIcon()
		{
			ShowTrayIcon(!IsVisible);
		}
		void ShowTrayIcon(bool show)
		{
			if (notificationObject != null)
				notificationObject.Visible = show;
		}
		#endregion

		#region Win32

		delegate void WinEventDelegate(IntPtr hWinEventHook, uint eventType, IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime);

		[DllImport("user32.dll")]
		static extern IntPtr SetWinEventHook(uint eventMin, uint eventMax, IntPtr hmodWinEventProc, WinEventDelegate lpfnWinEventProc, uint idProcess, uint idThread, uint dwFlags);

		private const uint WINEVENT_OUTOFCONTEXT = 0;
		private const uint EVENT_SYSTEM_FOREGROUND = 3;

		[DllImport("user32.dll")]
		static extern IntPtr GetForegroundWindow();

		[DllImport("user32.dll")]
		static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

		[DllImport("user32.dll")]
		private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
		

		#endregion
	}
}
