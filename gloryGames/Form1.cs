
using System.Timers;

namespace gloryGames
{
    public partial class Form1 : MetroFramework.Forms.MetroForm
    {
        private int totalSeconds;
        private int totalSeconds2;
        private int totalSeconds3;
        private int totalSeconds4;
        public Form1()
        {
            InitializeComponent();
            this.StyleManager = metroStyleManager1;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            #region Device
            cmbDevice.Items.Add("PS4_1");
            cmbDevice.Items.Add("PS4_2");
            cmbDevice.Items.Add("PS4_3");
            cmbDevice.Items.Add("PS4_4");
            #endregion

            #region Game Type
            cmbType.Items.AddRange(new string[] { "Simple Game", "Match","Match/Extra Time/Penalties"});
            #endregion

            #region Duration
            cmbDuration.Items.AddRange(new string[] { "15min","20min","30min", "1H","1H30min","2H","2H30min","3H","3H30min","4H","4H30min","5H" });
            #endregion

            #region Game Status
            cmbStatus.Items.AddRange(new string[] { "Online", "Offline" });
            #endregion

            #region Timer Status
            lblonoff1.Text = "off";
            lblonoff2.Text = "off";
            lblonoff3.Text = "off";
            lblonoff4.Text = "off";
            #endregion

            notifyIcon1.Icon = SystemIcons.Information;
            notifyIcon1.Visible = true;

            #region GridView
            table.ColumnCount = 6;

            table.Columns[0].Name = "Date";
            table.Columns[1].Name = "Device Number";
            table.Columns[2].Name = "Device Type";
            table.Columns[3].Name = "Game Status";
            table.Columns[4].Name = "Duration";
            table.Columns[5].Name = "Amount Paid";


            table.Columns[4].DefaultCellStyle.Format = "hh:mm";
            table.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(238,239,249);
            table.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            table.DefaultCellStyle.SelectionBackColor = Color.FromArgb(255, 196, 37);
            table.BackgroundColor = Color.White;


            #endregion


        }

      

        #region Day/Night
        private void N_D_Click(object sender, EventArgs e)
        {
            if (metroStyleManager1.Theme == MetroFramework.MetroThemeStyle.Dark)
            {
                metroStyleManager1.Theme = MetroFramework.MetroThemeStyle.Light;

                N_D.BackgroundImage = Properties.Resources.Night_Day_Light;

                txtPay.BackColor = Color.White;
                txtPay.ForeColor = Color.Black;
                btnAdd.ForeColor = Color.Black;
                btnAdd.FlatAppearance.BorderColor = Color.Black;
                btnExport.ForeColor = Color.Black;
                btnExport.FlatAppearance.BorderColor = Color.Black;
                btndelete.ForeColor = Color.Black;
                btndelete.FlatAppearance.BorderColor = Color.Black;
                table.BackgroundColor = Color.White;
                pictureBox1.Image = Properties.Resources.Glorygames;
                pictureBox1.Refresh();

                #region Labels
                label1.ForeColor = Color.Black;
                label2.ForeColor = Color.Black;
                label3.ForeColor = Color.Black;
                label4.ForeColor = Color.Black;
                label5.ForeColor = Color.Black;

                label7.ForeColor = Color.Black;
                label8.ForeColor = Color.Black;
                label9.ForeColor = Color.Black;
                label10.ForeColor = Color.Black;

                lblDevice1.ForeColor = Color.Black;
                lblDevice2.ForeColor = Color.Black;
                lblDevice3.ForeColor = Color.Black;
                lblDevice4.ForeColor = Color.Black;
                #endregion


            }
            else
            {
                metroStyleManager1.Theme = MetroFramework.MetroThemeStyle.Dark;

                N_D.BackgroundImage = Properties.Resources.Night_Day_Dark;

                txtPay.BackColor = Color.FromArgb(29, 29, 29);
                txtPay.ForeColor = Color.White;
                btnAdd.ForeColor = Color.White;
                btnAdd.FlatAppearance.BorderColor = Color.White;
                btnExport.ForeColor = Color.White;
                btnExport.FlatAppearance.BorderColor = Color.White;
                btndelete.ForeColor = Color.White;
                btndelete.FlatAppearance.BorderColor = Color.White;
                table.BackgroundColor = Color.FromArgb(29, 29, 29);
                pictureBox1.Image = Properties.Resources.Glorygames_Light;
                pictureBox1.Refresh();


                #region Labels
                label1.ForeColor = Color.White;
                label2.ForeColor = Color.White;
                label3.ForeColor = Color.White;
                label4.ForeColor = Color.White;
                label5.ForeColor = Color.White;

                label7.ForeColor = Color.White;
                label8.ForeColor = Color.White;
                label9.ForeColor = Color.White;
                label10.ForeColor = Color.White;

                lblDevice1.ForeColor = Color.White;
                lblDevice2.ForeColor = Color.White;
                lblDevice3.ForeColor = Color.White;
                lblDevice4.ForeColor = Color.White;

                #endregion

            }


        }
        #endregion

        #region Add Button
        private void btnAdd_MouseEnter(object sender, EventArgs e)
        {
            if (metroStyleManager1.Theme == MetroFramework.MetroThemeStyle.Light)
            {
                btnAdd.BackColor = Color.FromArgb(255, 196, 37);
                btnAdd.ForeColor = Color.Black;
            }
            else
            {
                btnAdd.BackColor = Color.FromArgb(255, 196, 37);
                btnAdd.ForeColor = Color.Black;
                btnAdd.FlatAppearance.BorderColor = Color.Black;
            }
        }

        private void btnAdd_MouseLeave(object sender, EventArgs e)
        {
            if (metroStyleManager1.Theme == MetroFramework.MetroThemeStyle.Light)
            {
                btnAdd.BackColor = Color.Transparent;
                btnAdd.ForeColor = Color.Black;
            }
            else
            {
                btnAdd.BackColor = Color.Transparent;
                btnAdd.ForeColor = Color.White;
                btnAdd.FlatAppearance.BorderColor = Color.White;
            }
        }
        #endregion

        #region button Export
        private void btnExport_Click(object sender, EventArgs e)
        {
            // creating Excel Application  
            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
            // creating new WorkBook within Excel application  
            Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
            // creating new Excelsheet in workbook  
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
            // see the excel sheet behind the program  
            app.Visible = true;
            // get the reference of first sheet. By default its name is Sheet1.  
            // store its reference to worksheet  
            worksheet = workbook.Sheets["Sheet1"];
            worksheet = workbook.ActiveSheet;
            // changing the name of active sheet  
            worksheet.Name = "Exported from gridview";
            // storing header part in Excel  
            for (int i = 1; i < table.Columns.Count + 1; i++)
            {
                worksheet.Cells[1, i] = table.Columns[i - 1].HeaderText;
            }
            // storing Each row and column value to excel sheet  
            for (int i = 0; i < table.Rows.Count - 1; i++)
            {
                for (int j = 0; j < table.Columns.Count; j++)
                {
                    worksheet.Cells[i + 2, j + 1] = table.Rows[i].Cells[j].Value.ToString();
                }
            }
            // save the application  
            var path= Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            var b = path+"\\"+ DateTime.Now.ToString("yyyy-dd-M--HH-mm-ss") + ".xlsx";

            workbook.SaveAs(b, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            workbook.Close();
        }




        #endregion

        #region Export
        private void btnExport_MouseEnter(object sender, EventArgs e)
        {
            if (metroStyleManager1.Theme == MetroFramework.MetroThemeStyle.Light)
            {
                btnExport.BackColor = Color.FromArgb(255, 196, 37);
                btnExport.ForeColor = Color.Black;
            }
            else
            {
                btnExport.BackColor = Color.FromArgb(255, 196, 37);
                btnExport.ForeColor = Color.Black;
                btnExport.FlatAppearance.BorderColor = Color.Black;
            }
        }

        private void btnExport_MouseLeave(object sender, EventArgs e)
        {
            if (metroStyleManager1.Theme == MetroFramework.MetroThemeStyle.Light)
            {
                btnExport.BackColor = Color.Transparent;
                btnExport.ForeColor = Color.Black;
            }
            else
            {
                btnExport.BackColor = Color.Transparent;
                btnExport.ForeColor = Color.White;
                btnExport.FlatAppearance.BorderColor = Color.White;
            }
        }
        #endregion

        #region Delete Button
        private void btndelete_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in table.SelectedRows)
            {
                table.Rows.Remove(row);
            }
        }
        #endregion

        #region Delete
        private void btndelete_MouseEnter(object sender, EventArgs e)
        {
            if (metroStyleManager1.Theme == MetroFramework.MetroThemeStyle.Light)
            {
                btndelete.BackColor = Color.FromArgb(255, 196, 37);
                btndelete.ForeColor = Color.Black;
            }
            else
            {
                btndelete.BackColor = Color.FromArgb(255, 196, 37);
                btndelete.ForeColor = Color.Black;
                btndelete.FlatAppearance.BorderColor = Color.Black;
            }
        }

        private void btndelete_MouseLeave(object sender, EventArgs e)
        {
            if (metroStyleManager1.Theme == MetroFramework.MetroThemeStyle.Light)
            {
                btndelete.BackColor = Color.Transparent;
                btndelete.ForeColor = Color.Black;
            }
            else
            {
                btndelete.BackColor = Color.Transparent;
                btndelete.ForeColor = Color.White;
                btndelete.FlatAppearance.BorderColor = Color.White;
            }
        }
        #endregion


        

        #region Prerequisites
        DateTime starttime1;
        DateTime starttime2;
        DateTime starttime3;
        DateTime starttime4;
        DateTime stoptime1;
        DateTime stoptime2;
        DateTime stoptime3;
        DateTime stoptime4;
        
        #endregion

        

        #region Start1
        private void btnStartDevice1_Click(object sender, EventArgs e)
        {
            int num= cmbDuration.SelectedIndex;
            switch (num)
            {
                case 0:

                    totalSeconds = 900;
                    timer1.Enabled = true;
                    break;

                case 1:

                    totalSeconds = 1200;
                    timer1.Enabled = true;
                    break;

                case 2:

                    totalSeconds = 1800;
                    timer1.Enabled = true;
                    break;

                case 3:

                    totalSeconds = 3600;
                    timer1.Enabled = true;
                    break;

                case 4:

                    totalSeconds = 5400;
                    timer1.Enabled = true;
                    break;

                case 5:

                    totalSeconds = 7200;
                    timer1.Enabled = true;
                    break;

                case 6:

                    totalSeconds = 9000;
                    timer1.Enabled = true;
                    break;

                case 7:

                    totalSeconds = 10800;
                    timer1.Enabled = true;
                    break;

                case 8:

                    totalSeconds = 12600;
                    timer1.Enabled = true;
                    break;

                case 9:

                    totalSeconds = 14400;
                    timer1.Enabled = true;
                    break;

                case 10:

                    totalSeconds = 16200;
                    timer1.Enabled = true;
                    break;

                case 11:

                    totalSeconds = 18000;
                    timer1.Enabled = true;
                    break;
            }
            
            lblonoff1.Text = "on";
            lblDevice1.ForeColor = Color.Green;
            btnStartDevice1.Enabled = false;
        }

        #endregion

        #region Start2
        private void btnStartDevice2_Click(object sender, EventArgs e)
        {
            int num = cmbDuration.SelectedIndex;
            switch (num)
            {
                case 0:

                    totalSeconds2 = 900;
                    timer2.Enabled = true;
                    break;

                case 1:

                    totalSeconds2 = 1200;
                    timer2.Enabled = true;
                    break;

                case 2:

                    totalSeconds2 = 1800;
                    timer2.Enabled = true;
                    break;

                case 3:

                    totalSeconds2 = 3600;
                    timer2.Enabled = true;
                    break;

                case 4:

                    totalSeconds2 = 5400;
                    timer2.Enabled = true;
                    break;

                case 5:

                    totalSeconds2 = 7200;
                    timer2.Enabled = true;
                    break;

                case 6:

                    totalSeconds2 = 9000;
                    timer2.Enabled = true;
                    break;

                case 7:

                    totalSeconds2 = 10800;
                    timer2.Enabled = true;
                    break;

                case 8:

                    totalSeconds2 = 12600;
                    timer2.Enabled = true;
                    break;

                case 9:

                    totalSeconds2 = 14400;
                    timer2.Enabled = true;
                    break;

                case 10:

                    totalSeconds2 = 16200;
                    timer2.Enabled = true;
                    break;

                case 11:

                    totalSeconds2 = 18000;
                    timer2.Enabled = true;
                    break;
            }
            
            lblonoff2.Text = "on";
            lblDevice2.ForeColor = Color.Green;
            btnStartDevice2.Enabled = false;
        }

        #endregion

        #region Start3
        private void btnStartDevice3_Click(object sender, EventArgs e)
        {

            int num = cmbDuration.SelectedIndex;
            switch (num)
            {
                case 0:

                    totalSeconds3 = 900;
                    timer3.Enabled = true;
                    break;

                case 1:

                    totalSeconds3 = 1200;
                    timer3.Enabled = true;
                    break;

                case 2:

                    totalSeconds3 = 1800;
                    timer3.Enabled = true;
                    break;

                case 3:

                    totalSeconds3 = 3600;
                    timer3.Enabled = true;
                    break;

                case 4:

                    totalSeconds3 = 5400;
                    timer3.Enabled = true;
                    break;

                case 5:

                    totalSeconds3 = 7200;
                    timer3.Enabled = true;
                    break;

                case 6:

                    totalSeconds3 = 9000;
                    timer3.Enabled = true;
                    break;

                case 7:

                    totalSeconds3 = 10800;
                    timer3.Enabled = true;
                    break;

                case 8:

                    totalSeconds3 = 12600;
                    timer3.Enabled = true;
                    break;

                case 9:

                    totalSeconds3 = 14400;
                    timer3.Enabled = true;
                    break;

                case 10:

                    totalSeconds3 = 16200;
                    timer3.Enabled = true;
                    break;

                case 11:

                    totalSeconds3 = 18000;
                    timer3.Enabled = true;
                    break;
            }
            lblonoff3.Text = "on";
            btnStartDevice3.Enabled = false;
            lblDevice3.ForeColor = Color.Green;
        }

        #endregion

        #region Start4
        private void btnStartDevice4_Click(object sender, EventArgs e)
        {
            int num = cmbDuration.SelectedIndex;
            switch (num)
            {
                case 0:

                    totalSeconds4 = 900;
                    timer4.Enabled = true;
                    break;

                case 1:

                    totalSeconds4 = 1200;
                    timer4.Enabled = true;
                    break;

                case 2:

                    totalSeconds4 = 1800;
                    timer4.Enabled = true;
                    break;

                case 3:

                    totalSeconds4 = 3600;
                    timer4.Enabled = true;
                    break;

                case 4:

                    totalSeconds4 = 5400;
                    timer4.Enabled = true;
                    break;

                case 5:

                    totalSeconds4 = 7200;
                    timer4.Enabled = true;
                    break;

                case 6:

                    totalSeconds4 = 9000;
                    timer4.Enabled = true;
                    break;

                case 7:

                    totalSeconds4 = 10800;
                    timer4.Enabled = true;
                    break;

                case 8:

                    totalSeconds4 = 12600;
                    timer4.Enabled = true;
                    break;

                case 9:

                    totalSeconds4 = 14400;
                    timer4.Enabled = true;
                    break;

                case 10:

                    totalSeconds4 = 16200;
                    timer4.Enabled = true;
                    break;

                case 11:

                    totalSeconds4 = 18000;
                    timer4.Enabled = true;
                    break;
            }
            lblonoff4.Text = "on";
            btnStartDevice4.Enabled = false;
            lblDevice4.ForeColor = Color.Green;
        }
        #endregion

        #region Pause
        private void btnPauseDevice1_Click(object sender, EventArgs e)
        {
            if (lblonoff1.Text == "on")
            {
                timer1.Stop();
                stoptime1 = DateTime.Now;
                lblonoff1.Text = "p";
                lblDevice1.ForeColor = Color.RoyalBlue;
            }
            else if (lblonoff1.Text == "p")
            {
                starttime1 += (DateTime.Now - stoptime1);

                timer1.Start();
                lblonoff1.Text = "on";
                lblDevice1.ForeColor = Color.Green;
            }

        }

        private void btnPauseDevice2_Click(object sender, EventArgs e)
        {
            if (lblonoff2.Text == "on")
            {
                timer2.Stop();
                stoptime2 = DateTime.Now;
                lblonoff2.Text = "p";
                lblDevice2.ForeColor = Color.RoyalBlue;
            }
            else if (lblonoff2.Text == "p")
            {
                starttime2 += (DateTime.Now - stoptime2);

                timer2.Start();
                lblonoff2.Text = "on";
                lblDevice2.ForeColor = Color.Green;
            }
        }

        private void btnPauseDevice3_Click(object sender, EventArgs e)
        {
            if (lblonoff3.Text == "on")
            {
                timer3.Stop();
                stoptime3 = DateTime.Now;
                lblonoff3.Text = "p";
                lblDevice3.ForeColor = Color.RoyalBlue;
            }
            else if (lblonoff3.Text == "p")
            {
                starttime3 += (DateTime.Now - stoptime3);

                timer3.Start();
                lblonoff3.Text = "on";
                lblDevice3.ForeColor = Color.Green;
            }
        }

        private void btnPauseDevice4_Click(object sender, EventArgs e)
        {
            if (lblonoff4.Text == "on")
            {
                timer4.Stop();
                stoptime4 = DateTime.Now;
                lblonoff4.Text = "p";
                lblDevice4.ForeColor = Color.RoyalBlue;
            }
            else if (lblonoff4.Text == "p")
            {
                starttime4 += (DateTime.Now - stoptime4);

                timer4.Start();
                lblonoff4.Text = "on";
                lblDevice4.ForeColor = Color.Green;
            }
        }
        #endregion

        #region Stop
        private void btnStopDevice1_Click(object sender, EventArgs e)
        {
            timer1.Stop();
            totalSeconds = 0;
            TimeSpan t = TimeSpan.FromSeconds(totalSeconds);
            lblDevice1.Text = String.Format("{0:D2}:{1:D2}:{2:D2}", t.Hours, t.Minutes, t.Seconds);
            btnStartDevice1.Enabled = true;
            lblonoff1.Text = "off";
            notifyIcon1.ShowBalloonTip(1000, "Post 1", "Finished Playing",ToolTipIcon.Info);
            lblDevice1.ForeColor = Color.Red;
                
        }

        private void btnStopDevice2_Click(object sender, EventArgs e)
        {
            timer2.Stop();
            totalSeconds2 = 0;
            TimeSpan t = TimeSpan.FromSeconds(totalSeconds2);
            lblDevice2.Text = String.Format("{0:D2}:{1:D2}:{2:D2}", t.Hours, t.Minutes, t.Seconds);
            btnStartDevice2.Enabled = true;
            lblonoff2.Text = "off";
            notifyIcon1.ShowBalloonTip(1000, "Post 2", "Finished Playing", ToolTipIcon.Info);
            lblDevice2.ForeColor = Color.Red;

        }

        private void btnStopDevice3_Click(object sender, EventArgs e)
        {
            timer3.Stop();
            totalSeconds3 = 0;
            TimeSpan t = TimeSpan.FromSeconds(totalSeconds3);
            lblDevice3.Text = String.Format("{0:D2}:{1:D2}:{2:D2}", t.Hours, t.Minutes, t.Seconds);
            btnStartDevice3.Enabled = true;
            lblonoff3.Text = "off";
            notifyIcon1.ShowBalloonTip(1000, "Post 3", "Finished Playing", ToolTipIcon.Info);
            lblDevice3.ForeColor = Color.Red;

        }

        private void btnStopDevice4_Click(object sender, EventArgs e)
        {
            timer4.Stop();
            totalSeconds4 = 0;
            TimeSpan t = TimeSpan.FromSeconds(totalSeconds4);
            lblDevice4.Text = String.Format("{0:D2}:{1:D2}:{2:D2}", t.Hours, t.Minutes, t.Seconds);
            btnStartDevice4.Enabled = true;
            lblonoff4.Text = "off";
            notifyIcon1.ShowBalloonTip(1000, "Post 4", "Finished Playing", ToolTipIcon.Info);
            lblDevice4.ForeColor = Color.Red;

        }
        #endregion

        #region Ticks
        private void timer1_Tick(object sender, EventArgs e)
        {
            if (totalSeconds > 0)
            {
                totalSeconds--;

                TimeSpan t = TimeSpan.FromSeconds(totalSeconds);
                lblDevice1.Text = String.Format("{0:D2}:{1:D2}:{2:D2}", t.Hours, t.Minutes, t.Seconds);
            }
            else
            {
                timer1.Stop();
                timer1.Enabled = false;

                TimeSpan t = TimeSpan.FromSeconds(totalSeconds);
                lblDevice1.Text = String.Format("{0:D2}:{1:D2}:{2:D2}", t.Hours, t.Minutes, t.Seconds);
                btnStartDevice1.Enabled = true;
                lblonoff1.Text = "off";
                notifyIcon1.ShowBalloonTip(1000, "Post 1", "Finished Playing", ToolTipIcon.Info);

                notifyIcon1.Visible = true;
            }
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            if (totalSeconds2 > 0)
            {
                totalSeconds2--;

                TimeSpan t = TimeSpan.FromSeconds(totalSeconds2);
                lblDevice2.Text = String.Format("{0:D2}:{1:D2}:{2:D2}", t.Hours, t.Minutes, t.Seconds);
            }
            else
            {
                timer2.Stop();
                timer2.Enabled = false;

                TimeSpan t = TimeSpan.FromSeconds(totalSeconds2);
                lblDevice2.Text = String.Format("{0:D2}:{1:D2}:{2:D2}", t.Hours, t.Minutes, t.Seconds);
                btnStartDevice2.Enabled = true;
                lblonoff2.Text = "off";
                notifyIcon1.ShowBalloonTip(1000, "Post 2", "Finished Playing", ToolTipIcon.Info);

                notifyIcon1.Visible = true;
            }
        }

        private void timer3_Tick(object sender, EventArgs e)
        {
            if (totalSeconds3 > 0)
            {
                totalSeconds3--;

                TimeSpan t = TimeSpan.FromSeconds(totalSeconds3);
                lblDevice3.Text = String.Format("{0:D2}:{1:D2}:{2:D2}", t.Hours, t.Minutes, t.Seconds);
            }
            else
            {
                timer3.Stop();
                timer3.Enabled = false;

                TimeSpan t = TimeSpan.FromSeconds(totalSeconds3);
                lblDevice3.Text = String.Format("{0:D2}:{1:D2}:{2:D2}", t.Hours, t.Minutes, t.Seconds);
                btnStartDevice3.Enabled = true;
                lblonoff3.Text = "off";
                notifyIcon1.ShowBalloonTip(1000, "Post 3", "Finished Playing", ToolTipIcon.Info);

                notifyIcon1.Visible = true;
            }
        }

        private void timer4_Tick(object sender, EventArgs e)
        {
            if (totalSeconds4 > 0)
            {
                totalSeconds4--;

                TimeSpan t = TimeSpan.FromSeconds(totalSeconds4);
                lblDevice4.Text = String.Format("{0:D2}:{1:D2}:{2:D2}", t.Hours, t.Minutes, t.Seconds);
            }
            else
            {
                timer4.Stop();
                timer4.Enabled = false;

                TimeSpan t = TimeSpan.FromSeconds(totalSeconds4);
                lblDevice4.Text = String.Format("{0:D2}:{1:D2}:{2:D2}", t.Hours, t.Minutes, t.Seconds);
                btnStartDevice4.Enabled = true;
                lblonoff4.Text = "off";
                notifyIcon1.ShowBalloonTip(1000, "Post 4", "Finished Playing", ToolTipIcon.Info);

                notifyIcon1.Visible = true;
            }
        }
        #endregion

        #region Add Action
        private void btnAdd_Click(object sender, EventArgs e)
        {
            try
            {

                DataGridViewRow row = new DataGridViewRow();
                row.CreateCells(table);

                row.Cells[0].Value = DateTime.Now.ToString();
                if (cmbDevice.SelectedItem == null)
                {
                    throw new Exception();
                }

                else
                {
                    if (cmbType.SelectedItem == null)
                    {
                        throw new Exception();
                    }
                    else
                    {
                        if (cmbStatus.SelectedItem == null)
                        {
                            throw new Exception();
                        }
                        else
                        {
                            if (cmbDuration.SelectedItem == null)
                            {
                                throw new Exception();
                            }
                        }
                    }
                }


                row.Cells[1].Value = cmbDevice.SelectedItem.ToString();

                row.Cells[2].Value = cmbType.SelectedItem.ToString();

                row.Cells[3].Value = cmbStatus.SelectedItem.ToString();


                switch (cmbDuration.SelectedItem.ToString())
                {
                    case "15min":

                        TimeSpan ts = new TimeSpan(0, 15, 0);
                        row.Cells[4].Value = string.Format("{0:hh\\:mm}", ts);
                        break;

                    case "20min":

                        ts = new TimeSpan(0, 20, 0);
                        row.Cells[4].Value = string.Format("{0:hh\\:mm}", ts);
                        break;

                    case "30min":

                        ts = new TimeSpan(0, 30, 0);
                        row.Cells[4].Value = string.Format("{0:hh\\:mm}", ts);
                        break;

                    case "1H":

                        ts = new TimeSpan(1, 0, 0);
                        row.Cells[4].Value = string.Format("{0:hh\\:mm}", ts);
                        break;

                    case "1H30min":

                        ts = new TimeSpan(1, 30, 0);
                        row.Cells[4].Value = string.Format("{0:hh\\:mm}", ts);
                        break;

                    case "2H":

                        ts = new TimeSpan(2, 0, 0);
                        row.Cells[4].Value = string.Format("{0:hh\\:mm}", ts);
                        break;

                    case "2H30min":

                        ts = new TimeSpan(2, 30, 0);
                        row.Cells[4].Value = string.Format("{0:hh\\:mm}", ts);
                        break;

                    case "3H":

                        ts = new TimeSpan(3, 0, 0);
                        row.Cells[4].Value = string.Format("{0:hh\\:mm}", ts);
                        break;

                    case "3H30min":

                        ts = new TimeSpan(3, 30, 0);
                        row.Cells[4].Value = string.Format("{0:hh\\:mm}", ts);
                        break;

                    case "4H":

                        ts = new TimeSpan(4, 0, 0);
                        row.Cells[4].Value = string.Format("{0:hh\\:mm}", ts);
                        break;

                    case "4H30min":

                        ts = new TimeSpan(4, 30, 0);
                        row.Cells[4].Value = string.Format("{0:hh\\:mm}", ts);
                        break;

                    case "5H":

                        ts = new TimeSpan(5, 0, 0);
                        row.Cells[4].Value = string.Format("{0:hh\\:mm}", ts);
                        break;
                }

                row.Cells[5].Value = txtPay.Text;
                table.Rows.Add(row);
            }
            catch (Exception)
            {
                MessageBox.Show("Fill fields", "Warning");

            }

        }

        #endregion



    }
}