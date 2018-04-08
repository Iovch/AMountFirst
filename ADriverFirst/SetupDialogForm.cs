using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using ASCOM.Utilities;
using ASCOM.AMount;
using System.IO.Ports;

namespace ASCOM.AMount
{
    [ComVisible(false)]  // Form not registered for COM!

    public partial class SetupDialogForm : Form
    {
        public SerialPort objSerial;

        public SetupDialogForm()
        {
            InitializeComponent();
            // Initialise current values of user settings from the ASCOM Profile
            InitUI();
        }

        private void cmdOK_Click(object sender, EventArgs e) // OK button event handler
        {
            // Place any validation constraint checks here
            // Update the state variables with results from the dialogue
            Telescope.comPort = (string)comboBoxComPort.SelectedItem;
            Telescope.tl.Enabled = chkTrace.Checked;
            if (objSerial != null) objSerial.Close();
        }

        private void cmdCancel_Click(object sender, EventArgs e) // Cancel button event handler
        {
            if (objSerial != null) objSerial.Close();
            Close();
        }

        private void BrowseToAscom(object sender, EventArgs e) // Click on ASCOM logo event handler
        {
            try
            {
                System.Diagnostics.Process.Start("http://ascom-standards.org/");
            }
            catch (System.ComponentModel.Win32Exception noBrowser)
            {
                if (noBrowser.ErrorCode == -2147467259)
                    MessageBox.Show(noBrowser.Message);
            }
            catch (System.Exception other)
            {
                MessageBox.Show(other.Message);
            }
        }

        private void InitUI()
        {
            chkTrace.Checked = Telescope.tl.Enabled;
            // set the list of com ports to those that are currently available
            comboBoxComPort.Items.Clear();
            comboBoxComPort.Items.AddRange(System.IO.Ports.SerialPort.GetPortNames());      // use System.IO because it's static
            // select the current port if possible
            if (comboBoxComPort.Items.Contains(Telescope.comPort))
            {
                comboBoxComPort.SelectedItem = Telescope.comPort;
            }
        }

        private void chkTrace_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void comboBoxComPort_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (objSerial != null) objSerial.Write("PN100");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (objSerial != null) objSerial.Write("PW100");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (objSerial != null) objSerial.Write("PE100");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (objSerial != null) objSerial.Write("PS100");
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Telescope.comPort = (string)comboBoxComPort.SelectedItem;
            objSerial = new SerialPort(Telescope.comPort);
            objSerial.Open();
            objSerial.DiscardInBuffer();
            objSerial.DiscardOutBuffer();
        }
    }
}