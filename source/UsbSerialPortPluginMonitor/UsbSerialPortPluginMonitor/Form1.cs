using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Management;

namespace UsbSerialPortPluginMonitor
{
    public partial class mainForm : Form
    {
        private bool mustClose = false;

        public mainForm()
        {
            InitializeComponent();

            //
            notifyIcon.Visible = true;
            notifyIcon.BalloonTipText = " ";
            notifyIcon.Icon = new Icon(this.Icon, 40, 40);
            notifyIcon.BalloonTipIcon = ToolTipIcon.Info;

            //
            SerialPortService.PortsChanged += (sender1, changedArgs) => { HandlePortChanges(changedArgs); };
        }


        // Changes in the serial ports on the pc
        //
        //
        void HandlePortChanges(PortsChangedArgs args)
        {
            string Text = "";
            string description = "";
            var Ports = COMPortInfo.Info.Get();

            // concat multiple changes
            foreach (string name in args.SerialPorts)
            {
                Console.WriteLine(args.EventType + " " + name);
                if (Text != "")
                {
                    Text += " \r\n ";
                }
                
                description = "";

                if (args.EventType == EventType.Insertion)
                {
                    foreach (COMPortInfo.Info comPort in Ports)
                    {
                        if (comPort.Name == name)
                        {
                            description = comPort.Description;
                        }
                    }
                }

                Text = args.EventType + " " + name + " " + description;
            }

            if (Text != "")
            {
                notifyIcon.BalloonTipTitle = Text;
                notifyIcon.ShowBalloonTip(1500);
            }

            if (Visible)
            {
                // update Ports_comboBox list
                log_listBox.Invoke((MethodInvoker)(() =>
                {
                    updateComportList();
                }
                ));
            }
        }


        // 
        //
        //
        private void notifyIcon_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            Console.WriteLine("notifyIcon_MouseDoubleClick");
            ShowWindow();
        }


        //
        //
        //
        private void closeButton_Click(object sender, EventArgs e)
        {
            Hide();
        }


        //
        //
        //
        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            mustClose = true;
            Close();
        }


        //
        //
        //
        private void showToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ShowWindow();
        }


        //
        //
        //
        private void ShowWindow()
        {
            if (Visible == false)
            {
                WindowState = FormWindowState.Normal;
                updateComportList();
                Show();
            }
        }


        // update list with friendly names of comports
        //
        //
        private void updateComportList()
        {
            lock (log_listBox)
            {                
                var sortedPorts = COMPortInfo.Info.Get().OrderBy(Name => Convert.ToUInt32(Name.Name.Replace("COM", string.Empty)));
                string Text;
                log_listBox.Items.Clear();
                foreach (COMPortInfo.Info comPort in sortedPorts)
                {
                    Text = string.Format("{0} – {1}", comPort.Name, comPort.Description);
                    Console.WriteLine(Text);
                    log_listBox.Items.Add(Text);
                }
            }
        }


        //
        //
        //
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            // closed -> not by user clicking close button [x] form
            if ((e.CloseReason == CloseReason.UserClosing) && (mustClose == true))
            {
                return;
            }

            // just hide the form
            e.Cancel = true;
            Hide();
        }


        //
        //
        //
        private void mainForm_Shown(object sender, EventArgs e)
        {
            // hide form at start of program
            Hide();
        }
    }
}
