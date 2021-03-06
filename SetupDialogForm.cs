using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using ASCOM.Utilities;
using ASCOM.GotoStar;
using System.Threading;

namespace ASCOM.GotoStar
{
    [ComVisible(false)]					// Form not registered for COM!
    public partial class SetupDialogForm : Form
    {
        private GotoStarCommunication gotoStar;

        public SetupDialogForm()
        {
            InitializeComponent();
            // Initialise current values of user settings from the ASCOM Profile
            InitUI();
            gotoStar = null;
        }

        private void cmdOK_Click(object sender, EventArgs e) // OK button event handler
        {
            // Place any validation constraint checks here
            // Update the state variables with results from the dialogue
            Telescope.comPort = (string)comboBoxComPort.SelectedItem;
            Telescope.LogDebug = chkTrace.Checked;
        }

        private void cmdCancel_Click(object sender, EventArgs e) // Cancel button event handler
        {
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
            chkTrace.Checked = Telescope.LogDebug;
            // set the list of com ports to those that are currently available
            comboBoxComPort.Items.Clear();
            comboBoxComPort.Items.AddRange(System.IO.Ports.SerialPort.GetPortNames());      // use System.IO because it's static
            btnTestConnection.Enabled = false;
            // select the current port if possible
            if (comboBoxComPort.Items.Contains(Telescope.comPort))
            {
                comboBoxComPort.SelectedItem = Telescope.comPort;
            }
        }

        private void btnTestConnection_Click(object sender, EventArgs e)
        {
            if (gotoStar == null)
            {
                gotoStar = new GotoStarCommunication();
                
            }
            gotoStar.OpenPort(comboBoxComPort.SelectedItem.ToString());
            lblTestResult.Text = gotoStar.TestConnection();
            /*
             * informal test area
             */
            gotoStar.GetRightAscension(out double ra);
            gotoStar.ClosePort();
        }

        private void comboBoxComPort_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnTestConnection.Enabled = !string.IsNullOrEmpty(comboBoxComPort.SelectedItem.ToString());

        }

    }
}
/*
 * Acties:
 * - Nadenken over een asynchrone pulse guide
 * - Misschien ook het guide-commando uitvaardigen?
 * - Inventariseren wat belangrijk is
 */