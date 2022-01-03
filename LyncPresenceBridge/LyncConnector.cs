
using System;
using System.Windows.Forms;

using Microsoft.Lync.Model;

using System.Management;
using System.Diagnostics;
using System.Drawing;
using Uctrl.Arduino;

namespace LyncPresenceBridge
{
    internal class LyncConnectorAppContext : ApplicationContext
    {
        private NotifyIcon trayIcon;
        private ContextMenuStrip trayIconContextMenu;

        private LyncClient lyncClient;

        private readonly Arduino arduino = new Arduino();

        private ManagementEventWatcher usbWatcher;

        private bool isLyncIntegratedMode = true;

        private static readonly Color ColorFree = Color.DarkGreen;
        private static readonly Color ColorBusy = Color.DarkRed;
        private static readonly Color ColorDnD = Color.Maroon;
        private static readonly Color ColorAway = Color.DarkOrange;
        private static readonly Color ColorOoO = Color.Fuchsia;
        private static readonly Color ColorOff = Color.Black;

        public LyncConnectorAppContext()
        {
            Application.ApplicationExit += OnApplicationExit;
            AppDomain.CurrentDomain.ProcessExit += OnApplicationExit;

            // Setup UI, NotifyIcon
            InitializeComponent();

            trayIcon.Visible = true;

            // Open Arduino serial port it's not 0
            if (Properties.Settings.Default.ArduinoSerialPort > 0)
            {
                if (! arduino.OpenPort("COM" + Properties.Settings.Default.ArduinoSerialPort.ToString()))
                {
                    trayIcon.ShowBalloonTip(1000, "Error", "Could not open and init serial port.", ToolTipIcon.Warning);
                }
            }

            // Setup Lync Client Connection
            GetLyncClient();

            // Watch for USB Changes, try to monitor arduino plugin/removal
            InitializeUsbWatcher();

        }

        private void InitializeComponent()
        {
            trayIcon = new NotifyIcon();

            //The icon is added to the project resources.
            trayIcon.Icon = Properties.Resources.blink_off;

            // TrayIconContextMenu
            trayIconContextMenu = new ContextMenuStrip();
            trayIconContextMenu.SuspendLayout();
            trayIconContextMenu.Name = "trayIconContextMenu";

            // Tray Context Menu items to set color
            this.trayIconContextMenu.Items.Add("Free", null, new EventHandler(FreeMenuItem_Click));
            this.trayIconContextMenu.Items.Add("Busy", null, new EventHandler(BusyMenuItem_Click));
            this.trayIconContextMenu.Items.Add("DnD", null, new EventHandler(DnDMenuItem_Click));
            this.trayIconContextMenu.Items.Add("Away", null, new EventHandler(AwayMenuItem_Click));
            this.trayIconContextMenu.Items.Add("OoO", null, new EventHandler(OoOMenuItem_Click));
            this.trayIconContextMenu.Items.Add("Off", null, new EventHandler(OffMenuItem_Click));

            // Separation Line
            this.trayIconContextMenu.Items.Add(new ToolStripSeparator());

            // About Form Line
            this.trayIconContextMenu.Items.Add("About", null, new EventHandler(aboutMenuItem_Click));

            // Separation Line
            this.trayIconContextMenu.Items.Add(new ToolStripSeparator());

            // CloseMenuItem
            this.trayIconContextMenu.Items.Add("Exit", null, new EventHandler(CloseMenuItem_Click));


            trayIconContextMenu.ResumeLayout(false);
            trayIcon.ContextMenuStrip = trayIconContextMenu;
        }

        private void GetLyncClient()
        {
            try
            {
                // try to get the running lync client and register for change events, if Client is not running then ClientNoFound Exception is thrown by lync api
                lyncClient = LyncClient.GetClient();
                lyncClient.StateChanged += lyncClient_StateChanged;
                
                if (lyncClient.State == ClientState.SignedIn)
                    lyncClient.Self.Contact.ContactInformationChanged += Contact_ContactInformationChanged;
    
                SetCurrentContactState();
            }
            catch (ClientNotFoundException)
            {
                Debug.WriteLine("Lync Client not started.");

                SetLyncIntegrationMode(false);

                trayIcon.ShowBalloonTip(1000, "Error", "Lync Client not started. Running in manual mode now. Please use the context menu to change your blink color.", ToolTipIcon.Warning);
            }
            catch (Exception e)
            {
                Debug.WriteLine(e.ToString());

                trayIcon.ShowBalloonTip(1000, "Error", "Something went wrong by getting your Lync status. Running in manual mode now. Please use the context menu to change your blink color.", ToolTipIcon.Warning);
                Debug.WriteLine(e.Message);
            }
        }

        private void SetLyncIntegrationMode(bool isLyncIntegrated)
        {
            isLyncIntegratedMode = isLyncIntegrated;
            if (isLyncIntegratedMode)
            {
            }
        }

        /// <summary>
        /// Read the current Availability Information from Lync/Skype for Business and set the color 
        /// </summary>
        private void SetCurrentContactState()
        {
            if (lyncClient.State == ClientState.SignedIn)
            {
                var currentAvailability = (ContactAvailability)lyncClient.Self.Contact.GetContactInformation(ContactInformationType.Availability);
                var currentCalendarState = (ContactCalendarState)lyncClient.Self.Contact.GetContactInformation(ContactInformationType.CurrentCalendarState);
                
                switch (currentAvailability)
                {
                    case ContactAvailability.None: // ???
                        arduino.SetLEDs(currentCalendarState == ContactCalendarState.OutOfOffice ? ColorOoO : ColorOff);
                        break;
                    
                    case ContactAvailability.Free: // Free
                    case ContactAvailability.FreeIdle: // Free and idle
                        arduino.SetLEDs(ColorFree);
                        break;

                    case ContactAvailability.Busy: // Busy
                    case ContactAvailability.BusyIdle: // Busy and idle
                        arduino.SetLEDs(ColorBusy);
                        break;

                    case ContactAvailability.DoNotDisturb: // Do not disturb
                        arduino.SetLEDs(ColorDnD);
                        break;

                    case ContactAvailability.TemporarilyAway: // Be right back
                    case ContactAvailability.Away:            // Inactive/away, off work, appear away
                        arduino.SetLEDs(currentCalendarState == ContactCalendarState.OutOfOffice ? ColorOoO : ColorAway);
                        break;

                    case ContactAvailability.Offline: // Offline
                        arduino.SetLEDs(currentCalendarState == ContactCalendarState.OutOfOffice ? ColorOoO : ColorOff);
                        break;

                    case ContactAvailability.Invalid:
                    default:
                        arduino.SetLEDs(ColorOff);
                        break;
                }
            }
        }

        void lyncClient_StateChanged(object sender, ClientStateChangedEventArgs e)
        {
            switch (e.NewState)
            {
                case ClientState.Initializing:
                    break;

                case ClientState.Invalid:
                    break;

                case ClientState.ShuttingDown:
                    break;

                case ClientState.SignedIn:
                    lyncClient.Self.Contact.ContactInformationChanged += Contact_ContactInformationChanged;
                    SetCurrentContactState();
                    break;

                case ClientState.SignedOut:
                    trayIcon.ShowBalloonTip(1000, "", "You signed out in Lync. Switching to manual mode.", ToolTipIcon.Info);
                    break;

                case ClientState.SigningIn:
                    break;

                case ClientState.SigningOut:
                    break;

                case ClientState.Uninitialized:
                    break;
            }
        }

        void Contact_ContactInformationChanged(object sender, ContactInformationChangedEventArgs e)
        {
            if (e.ChangedContactInformation.Contains(ContactInformationType.Availability))
            {
                SetCurrentContactState();
            }
        }

        private void OnApplicationExit(object sender, EventArgs e)
        {
            //Cleanup so that the icon will be removed when the application is closed
            trayIcon.Visible = false;

            // stop (USB) ManagementEventWatcher
            usbWatcher.Stop();
            usbWatcher.Dispose();

            if (arduino.Port.IsOpen)
            {
                arduino.SetLEDs(ColorOff);
                arduino.Dispose();
            }
                
        }
        
        private void TrayIcon_DoubleClick(object sender, EventArgs e)
        {
        }

        private static void CloseMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private static void aboutMenuItem_Click(object sender, EventArgs e)
        {
            var about = new AboutForm();
            about.ShowDialog();
        }

        private void OffMenuItem_Click(object sender, EventArgs e)
        {
            arduino.SetLEDs(ColorOff);
        }

        private void OoOMenuItem_Click(object sender, EventArgs e)
        {
            arduino.SetLEDs(ColorOoO);
        }

        private void AwayMenuItem_Click(object sender, EventArgs e)
        {
            arduino.SetLEDs(ColorAway);
        }

        private void BusyMenuItem_Click(object sender, EventArgs e)
        {
            arduino.SetLEDs(ColorBusy);
        }

        private void DnDMenuItem_Click(object sender, EventArgs e)
        {
            arduino.SetLEDs(ColorDnD);
        }

        private void FreeMenuItem_Click(object sender, EventArgs e)
        {
            arduino.SetLEDs(ColorFree);
        }

        // Watch for USB changes to detect arduino removal
        private void InitializeUsbWatcher()
        {
            usbWatcher = new ManagementEventWatcher();
            var query = new WqlEventQuery("SELECT * FROM Win32_DeviceChangeEvent");
            usbWatcher.EventArrived += new EventArrivedEventHandler(watcher_EventArrived);
            usbWatcher.Query = query;
            usbWatcher.Start();
        }

        private static void watcher_EventArrived(object sender, EventArrivedEventArgs e)
        {
        }
    }
}
