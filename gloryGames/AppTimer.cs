using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace gloryGames
{
    public class AppTimer { 
        public static int timerStart(MetroFramework.Controls.MetroComboBox cmbDuration, System.Windows.Forms.Timer timer, Label lblonoff, Label lblDevice, Button btnStartDevice)
        {
            int totalSeconds = 0;
            int num = cmbDuration.SelectedIndex;
            switch (num)
            {
                case 0:

                    totalSeconds = 900;
                    timer.Enabled = true;
                    break;

                case 1:

                    totalSeconds = 1200;
                    timer.Enabled = true;
                    break;

                case 2:

                    totalSeconds = 1800;
                    timer.Enabled = true;
                    break;

                case 3:

                    totalSeconds = 3600;
                    timer.Enabled = true;
                    break;

                case 4:

                    totalSeconds = 5400;
                    timer.Enabled = true;
                    break;

                case 5:

                    totalSeconds = 7200;
                    timer.Enabled = true;
                    break;

                case 6:

                    totalSeconds = 9000;
                    timer.Enabled = true;
                    break;

                case 7:

                    totalSeconds = 10800;
                    timer.Enabled = true;
                    break;

                case 8:

                    totalSeconds = 12600;
                    timer.Enabled = true;
                    break;

                case 9:

                    totalSeconds = 14400;
                    timer.Enabled = true;
                    break;

                case 10:

                    totalSeconds = 16200;
                    timer.Enabled = true;
                    break;

                case 11:

                    totalSeconds = 18000;
                    timer.Enabled = true;
                    break;
            }

            lblonoff.Text = "on";
            lblDevice.ForeColor = Color.Green;
            btnStartDevice.Enabled = false;
            return totalSeconds;
        }

        public static int timerTick(Label lblDevice, int totalSeconds, Button btnStartDevice, Label lblonoff, NotifyIcon notifyIcon, System.Windows.Forms.Timer timer)
        {
            if (totalSeconds > 0)
            {
                totalSeconds--;

                TimeSpan t = TimeSpan.FromSeconds(totalSeconds);
                lblDevice.Text = String.Format("{0:D2}:{1:D2}:{2:D2}", t.Hours, t.Minutes, t.Seconds);
            }
            else
            {
                timer.Stop();
                timer.Enabled = false;

                TimeSpan t = TimeSpan.FromSeconds(totalSeconds);
                lblDevice.Text = String.Format("{0:D2}:{1:D2}:{2:D2}", t.Hours, t.Minutes, t.Seconds);
                btnStartDevice.Enabled = true;
                lblonoff.Text = "off";
                notifyIcon.ShowBalloonTip(1000, "Post 1", "Finished Playing", ToolTipIcon.Info);

                notifyIcon.Visible = true;
            }
            return totalSeconds;
        }


        public static void timerStop(System.Windows.Forms.Timer timer, int totalSeconds, Label lblDevice, Button btnStartDevice, Label lblonoff, NotifyIcon notifyIcon)
        {
            timer.Stop();
            totalSeconds = 0;
            TimeSpan t = TimeSpan.FromSeconds(totalSeconds);
            lblDevice.Text = String.Format("{0:D2}:{1:D2}:{2:D2}", t.Hours, t.Minutes, t.Seconds);
            btnStartDevice.Enabled = true;
            lblonoff.Text = "off";
            notifyIcon.ShowBalloonTip(1000, "Post 1", "Finished Playing", ToolTipIcon.Info);
            lblDevice.ForeColor = Color.Red;

        }


        public static void timerPause(Label lblonoff, System.Windows.Forms.Timer timer, DateTime stoptime, Label lblDevice, DateTime starttime)
        {
            if (lblonoff.Text == "on")
            {
                timer.Stop();
                stoptime = DateTime.Now;
                lblonoff.Text = "p";
                lblDevice.ForeColor = Color.RoyalBlue;
            }
            else if (lblonoff.Text == "p")
            {
                starttime += (DateTime.Now - stoptime);

                timer.Start();
                lblonoff.Text = "on";
                lblDevice.ForeColor = Color.Green;
            }
        }
    }
}
