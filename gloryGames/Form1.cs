
using System.Timers;

namespace gloryGames
{
    public partial class timer_Tick : MetroFramework.Forms.MetroForm
    {
        private int totalSeconds1;
        private int totalSeconds2;
        private int totalSeconds3;
        private int totalSeconds4;
        public timer_Tick()
        {
            InitializeComponent();
            this.StyleManager = metroStyleManager1;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            #region Device
            cmbDevice.Items.AddRange(new string[] { "Console 1", "Console 2", "Console 3", "Console 4" });
            #endregion

            #region Game Type
            cmbType.Items.AddRange(new string[] { "Simple Game", "Match","Match/Extra Time/Penalties"});
            #endregion

            #region Duration
            cmbDuration.Items.AddRange(new string[] { "15min","20min","30min", "1 Hour","1H30min","2 Hours","2H30min", "3 Hours", "3H30min", "4 Hours", "4H30min", "5 Hours" });
            #endregion

            #region Game Status
            cmbStatus.Items.AddRange(new string[] { "Online", "Offline" });
            #endregion

            #region Timer Status

            var labels = new List<Label> { lblonoff1, lblonoff2, lblonoff3, lblonoff4 };
            foreach (var label in labels)
            {
                label.Text = "off";
            }
            #endregion

            notifyIcon1.Icon = SystemIcons.Information;
            notifyIcon1.Visible = true;

            #region GridView
            table.ColumnCount = 6;
            String[] tabelHeader = { "Date", "Device Number","Device Type", "Game Status","Duration","Amount Paid"};

            for(int i=0; i < 6; i++)
            {
                table.Columns[i].Name = tabelHeader[i];
            }


            table.Columns[4].DefaultCellStyle.Format = "hh:mm";
            table.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(238,239,249);
            table.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            table.DefaultCellStyle.SelectionBackColor = Color.FromArgb(255, 196, 37);
            table.BackgroundColor = Color.White;


            #endregion


        }

      

        #region Light Mode/Dark Mode
        private void L_D_Click(object sender, EventArgs e)
        {
            var labels = new List<Label> { label1, label2, label3, label4, label5, label7, label8, label9, label10, lblDevice1, lblDevice2, lblDevice3, lblDevice4 };
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

                foreach (var label in labels)
                {
                    label.ForeColor = Color.Black;
                }
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
                foreach ( var label in labels )
                {
                    label.ForeColor = Color.White;
                }
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

        #region Export Button
        private void btnExport_Click(object sender, EventArgs e)
        {
            // creating Excel Application  
            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
            // creating new WorkBook within Excel application  
            Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
            // creating new Excelsheet in workbook  
            Microsoft.Office.Interop.Excel._Worksheet? worksheet = null;
            // see the excel sheet behind the program  
            app.Visible = false;
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
        DateTime starttime1, starttime2, starttime3, starttime4;
       
        DateTime stoptime1, stoptime2, stoptime3, stoptime4;
        #endregion

        

        #region Start1
        private void btnStartDevice1_Click(object sender, EventArgs e)
        {
            totalSeconds1 = AppTimer.timerStart(cmbDuration, timer1, lblonoff1, lblDevice1, btnStartDevice1);
        }

        #endregion



        #region Start2
        private void btnStartDevice2_Click(object sender, EventArgs e)
        {
            totalSeconds2 = AppTimer.timerStart(cmbDuration, timer2, lblonoff2, lblDevice2, btnStartDevice2);
        }

        #endregion



        #region Start3
        private void btnStartDevice3_Click(object sender, EventArgs e)
        {
            totalSeconds3 = AppTimer.timerStart(cmbDuration, timer3, lblonoff3, lblDevice3, btnStartDevice3);
        }

        #endregion



        #region Start4
        private void btnStartDevice4_Click(object sender, EventArgs e)
        {
            totalSeconds4 = AppTimer.timerStart(cmbDuration, timer4, lblonoff4, lblDevice4, btnStartDevice4);
        }
        #endregion



        #region Pause
        private void btnPauseDevice1_Click(object sender, EventArgs e)
        {
            AppTimer.timerPause(lblonoff1, timer1, stoptime1, lblDevice1, starttime1);
        }

        private void btnPauseDevice2_Click(object sender, EventArgs e)
        {
            AppTimer.timerPause(lblonoff2, timer2, stoptime2, lblDevice2, starttime2);
        }

        private void btnPauseDevice3_Click(object sender, EventArgs e)
        {
            AppTimer.timerPause(lblonoff3, timer3, stoptime3, lblDevice3, starttime3);
        }

        private void btnPauseDevice4_Click(object sender, EventArgs e)
        {
            AppTimer.timerPause(lblonoff4, timer4, stoptime4, lblDevice4, starttime4);
        }
        #endregion

        #region Stop
        private void btnStopDevice1_Click(object sender, EventArgs e)
        {
            AppTimer.timerStop(timer1, totalSeconds1, lblDevice1, btnStartDevice1, lblonoff1, notifyIcon1);
        }

        private void btnStopDevice2_Click(object sender, EventArgs e)
        {
            AppTimer.timerStop(timer2, totalSeconds2, lblDevice2, btnStartDevice2, lblonoff2, notifyIcon1);
        }

        private void btnStopDevice3_Click(object sender, EventArgs e)
        {
            AppTimer.timerStop(timer3, totalSeconds3, lblDevice3, btnStartDevice3, lblonoff3, notifyIcon1);
        }

        private void btnStopDevice4_Click(object sender, EventArgs e)
        {
            AppTimer.timerStop(timer4, totalSeconds4, lblDevice4, btnStartDevice4, lblonoff4, notifyIcon1);
        }
        #endregion

        #region Ticks
        private void timer1_Tick(object sender, EventArgs e)
        {
            totalSeconds1 = AppTimer.timerTick(lblDevice1, totalSeconds1, btnStartDevice1, lblonoff1, notifyIcon1, timer1);
        }
        

        private void timer2_Tick(object sender, EventArgs e)
        {
            totalSeconds2 = AppTimer.timerTick(lblDevice2, totalSeconds2, btnStartDevice2, lblonoff2, notifyIcon1, timer2);
        }


        private void timer3_Tick(object sender, EventArgs e)
        {
            totalSeconds3 = AppTimer.timerTick(lblDevice3, totalSeconds3, btnStartDevice3, lblonoff3, notifyIcon1, timer3);
        }


        private void timer4_Tick(object sender, EventArgs e)
        {
            totalSeconds4 = AppTimer.timerTick(lblDevice4, totalSeconds4, btnStartDevice4, lblonoff4, notifyIcon1, timer4);
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

                    case "1 Hour":

                        ts = new TimeSpan(1, 0, 0);
                        row.Cells[4].Value = string.Format("{0:hh\\:mm}", ts);
                        break;

                    case "1H30min":

                        ts = new TimeSpan(1, 30, 0);
                        row.Cells[4].Value = string.Format("{0:hh\\:mm}", ts);
                        break;

                    case "2 Hours":

                        ts = new TimeSpan(2, 0, 0);
                        row.Cells[4].Value = string.Format("{0:hh\\:mm}", ts);
                        break;

                    case "2H30min":

                        ts = new TimeSpan(2, 30, 0);
                        row.Cells[4].Value = string.Format("{0:hh\\:mm}", ts);
                        break;

                    case "3 Hours":

                        ts = new TimeSpan(3, 0, 0);
                        row.Cells[4].Value = string.Format("{0:hh\\:mm}", ts);
                        break;

                    case "3H30min":

                        ts = new TimeSpan(3, 30, 0);
                        row.Cells[4].Value = string.Format("{0:hh\\:mm}", ts);
                        break;

                    case "4 Hours":

                        ts = new TimeSpan(4, 0, 0);
                        row.Cells[4].Value = string.Format("{0:hh\\:mm}", ts);
                        break;

                    case "4H30min":

                        ts = new TimeSpan(4, 30, 0);
                        row.Cells[4].Value = string.Format("{0:hh\\:mm}", ts);
                        break;

                    case "5 Hours":

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