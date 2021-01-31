using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
//MySQL Database Libraries
using MySql.Data;
using MySql.Data.Common;
using MySql.Data.Types;
using MySql.Data.MySqlClient;
//MS-Access-OLEDB Database Libraries
using System.Data.OleDb;

namespace DarkDemo
{

    public partial class Form1 : Form
    {
        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;
        [DllImportAttribute("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [DllImportAttribute("user32.dll")]
        public static extern bool ReleaseCapture();

        public Form1()
        {
            InitializeComponent();
            date_label.Text = DateTimeOffset.Now.DateTime.ToLongDateString(); //Shows current date on label on the top
        }

        private void exit_button_Click(object sender, EventArgs e) //Exit Application Button
        {
            Application.Exit();
        }

        private void button1_Click(object sender, EventArgs e) //Checks if all input fields has values
        {
            string message = "0の場合でも全部記入してください";
            string title = "エラー";
            
            if (String.IsNullOrEmpty(urikake_input.Text))
            {
                MessageBox.Show(message, title);
            }
            else if (String.IsNullOrEmpty(credit_input.Text))
            {
                MessageBox.Show(message, title);
            }
            else if (String.IsNullOrEmpty(shanai_input.Text))
            {
                MessageBox.Show(message, title);
            }
            else if (String.IsNullOrEmpty(genkin_input.Text))
            {
                MessageBox.Show(message, title);
            }

            else
            {
                //Converts all input fields to to Integer and sums them into a currency value
                souuriage_label.Text = (Convert.ToInt32(urikake_input.Text) + Convert.ToInt32(credit_input.Text) + Convert.ToInt32(shanai_input.Text) + Convert.ToInt32(genkin_input.Text)).ToString("C");
            }

        }

        private void rejishime_button_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open("\\\\192.168.10.100\\Public\\rejishime.xlsx", 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            string currentdate = DateTime.Now.ToString("MM-dd"); 
            string currentday = DateTime.Now.ToString("dd");
            string currentmonth = DateTime.Now.ToString("MM");
            string currentyear = DateTime.Now.ToString("yyyy");
            string currentdate2 = DateTime.Now.ToString("yyyy-MM-dd"); //date format for database input


            string message = "０の場合でも全部記入してください";
            string title = "エラー";
            //Checks if everything's been entered
            if (String.IsNullOrEmpty(mansatsu_input.Text))
            {
                MessageBox.Show(message, title);
            }
            else if (String.IsNullOrEmpty(gosenensatsu_input.Text))
            {
                MessageBox.Show(message, title);
            }
            else if (String.IsNullOrEmpty(nisenensatsu_input.Text))
            {
                MessageBox.Show(message, title);
            }
            else if (String.IsNullOrEmpty(senensatsu_input.Text))
            {
                MessageBox.Show(message, title);
            }
            else if (String.IsNullOrEmpty(gohyakuendama_input.Text))
            {
                MessageBox.Show(message, title);
            }
            else if (String.IsNullOrEmpty(hyakuendama_input.Text))
            {
                MessageBox.Show(message, title);
            }
            else if (String.IsNullOrEmpty(gojyuendama_input.Text))
            {
                MessageBox.Show(message, title);
            }
            else if (String.IsNullOrEmpty(jyuendama_input.Text))
            {
                MessageBox.Show(message, title);
            }
            else if (String.IsNullOrEmpty(goendama_input.Text))
            {
                MessageBox.Show(message, title);
            }
            else if (String.IsNullOrEmpty(ichiendama_input.Text))
            {
                MessageBox.Show(message, title);
            }
            else if (String.IsNullOrEmpty(urikake_input.Text))
            {
                MessageBox.Show(message, title);
            }
            else if (String.IsNullOrEmpty(credit_input.Text))
            {
                MessageBox.Show(message, title);
            }
            else if (String.IsNullOrEmpty(shanai_input.Text))
            {
                MessageBox.Show(message, title);
            }
            else if (String.IsNullOrEmpty(genkin_input.Text))
            {
                MessageBox.Show(message, title);
            }
            else
            {
                xlWorkSheet.Cells[1, 14] = currentday; //Inserts current date into excel
                xlWorkSheet.Cells[1, 12] = currentmonth;
                xlWorkSheet.Cells[1, 10] = currentyear;

                //Uriage into Excel
                xlWorkSheet.Cells[12, 11] = urikake_input.Text; //Urikake1
                xlWorkSheet.Cells[25, 11] = urikake_input.Text; //Urikake2
                xlWorkSheet.Cells[14, 11] = credit_input.Text; //Credit1
                xlWorkSheet.Cells[27, 11] = credit_input.Text; //Credit2
                xlWorkSheet.Cells[18, 11] = shanai_input.Text; //Shanai1
                xlWorkSheet.Cells[31, 11] = shanai_input.Text; //Shanai2
                xlWorkSheet.Cells[23, 11] = genkin_input.Text; //Genkin

                //Number of bills into Excel
                xlWorkSheet.Cells[6, 4] = mansatsu_input.Text; //Mansatsu
                xlWorkSheet.Cells[7, 4] = gosenensatsu_input.Text; //Gosensatsu
                xlWorkSheet.Cells[9, 4] = nisenensatsu_input.Text; //Nisensatsu
                xlWorkSheet.Cells[10, 4] = senensatsu_input.Text; //Sensatsu
                xlWorkSheet.Cells[11, 4] = gohyakuendama_input.Text; //Gohyakuendama
                xlWorkSheet.Cells[13, 4] = hyakuendama_input.Text; //Hyakuendama
                xlWorkSheet.Cells[14, 4] = gojyuendama_input.Text; //Gojyuendama
                xlWorkSheet.Cells[15, 4] = jyuendama_input.Text; //Jyuendama
                xlWorkSheet.Cells[17, 4] = goendama_input.Text; //Goendama
                xlWorkSheet.Cells[18, 4] = ichiendama_input.Text; //Ichiendama

                //shuunyuuinshikakunin into Excel
                xlWorkSheet.Cells[28, 1] = zenjitsumaisu_input.Text; //Yesterdays number
                xlWorkSheet.Cells[28, 3] = ukeiremaisu_input.Text; //Todays recieved number
                xlWorkSheet.Cells[28, 4] = shiharaimaisu_input.Text; //Paid number

                //Determines the path to the Desktop location and saves
                string path = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                xlWorkBook.SaveAs(path + "\\" + currentdate + ".xlsx");
                xlWorkBook.Close();
                xlApp.Quit();

                //database record
                System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection();
                conn.ConnectionString = "Provider = Microsoft.ACE.OLEDB.12.0;Data Source=\\\\192.168.10.100\\Public\\レジデータ\\rejishime.accdb";
                
                try
                {

                    conn.Open();
                    string my_querry = "INSERT INTO souuriage_table(Uriage_date,Genkin,Urikake,Credit,Shanai,1manen,5senen,2senen,1senen,500en,100en,50en,10en,5en,1en,yesterday_maisu,ukeire_maisu,used_maisu)VALUES('" + currentdate2 + "','" + genkin_input.Text + "','" + urikake_input.Text + "','" + credit_input.Text + "','" + shanai_input.Text + "','" + mansatsu_input.Text + "','" + gosenensatsu_input.Text + "','" + nisenensatsu_input.Text + "','" + senensatsu_input.Text + "','" + gohyakuendama_input.Text + "','" + hyakuendama_input.Text + "','" + gojyuendama_input.Text + "','" + jyuendama_input.Text + "','" + goendama_input.Text + "','" + ichiendama_input.Text + "','" + zenjitsumaisu_input.Text + "','" + ukeiremaisu_input.Text + "','" + shiharaimaisu_input.Text + "')";

                    OleDbCommand cmd = new OleDbCommand(my_querry, conn);
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("更新終了しました", "成功");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("失敗：" + ex.Message);
                }
                finally
                {
                    conn.Close();
                }

                //Changes the button text to "finished" when pressed
                if (rejishime_button.Text == "レジ締め")
                {
                    rejishime_button.Text = "完了";
                }
                else
                {
                    rejishime_button.Text = "レジ締め";
                }
            }
        }

        private void Form1_MouseDown(object sender, MouseEventArgs e) //Make window draggable
        {
            ReleaseCapture();
            SendMessage(this.Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
        }

        private void draggable_window(object sender, MouseEventArgs e) //Make window draggable
        {
            ReleaseCapture();
            SendMessage(this.Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
        }

        private void Enter_tab(object sender, KeyEventArgs e) //Makes enter work as tab
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
                SelectNextControl(genkin_input, true, true, true, true);
            }
        }

        private void Enter_tab1(object sender, KeyEventArgs e) //Makes enter work as tab
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
                SelectNextControl(urikake_input, true, true, true, true);
            }
        }

        private void Enter_tab2(object sender, KeyEventArgs e) //Makes enter work as tab
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
                SelectNextControl(credit_input, true, true, true, true);
            }
        }

        private void Enter_tab3(object sender, KeyEventArgs e) //Makes enter work as tab
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
                SelectNextControl(shanai_input, true, true, true, true);
            }
        }

        private void Enter_tab4(object sender, KeyEventArgs e) //Makes enter work as tab
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
                SelectNextControl(mansatsu_input, true, true, true, true);
            }
        }

        private void Enter_tab5(object sender, KeyEventArgs e) //Makes enter work as tab
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
                SelectNextControl(gosenensatsu_input, true, true, true, true);
            }
        }

        private void Enter_tab6(object sender, KeyEventArgs e) //Makes enter work as tab
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
                SelectNextControl(nisenensatsu_input, true, true, true, true);
            }
        }

        private void Enter_tab7(object sender, KeyEventArgs e) //Makes enter work as tab
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
                SelectNextControl(senensatsu_input, true, true, true, true);
            }
        }

        private void Enter_tab8(object sender, KeyEventArgs e) //Makes enter work as tab
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
                SelectNextControl(gohyakuendama_input, true, true, true, true);
            }
        }

        private void Enter_tab9(object sender, KeyEventArgs e) //Makes enter work as tab
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
                SelectNextControl(hyakuendama_input, true, true, true, true);
            }
        }

        private void Enter_tab10(object sender, KeyEventArgs e) //Makes enter work as tab
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
                SelectNextControl(gojyuendama_input, true, true, true, true);
            }
        }

        private void Enter_tab11(object sender, KeyEventArgs e) //Makes enter work as tab
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
                SelectNextControl(jyuendama_input, true, true, true, true);
            }
        }

        private void Enter_tab12(object sender, KeyEventArgs e) //Makes enter work as tab
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
                SelectNextControl(goendama_input, true, true, true, true);
            }
        }

        private void Enter_tab13(object sender, KeyEventArgs e) //Makes enter work as tab
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
                SelectNextControl(ichiendama_input, true, true, true, true);
            }
        }

        private void Enter_tab14(object sender, KeyEventArgs e) //Makes enter work as tab
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
                SelectNextControl(zenjitsumaisu_input, true, true, true, true);
            }
        }

        private void Enter_tab15(object sender, KeyEventArgs e) //Makes enter work as tab
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
                SelectNextControl(ukeiremaisu_input, true, true, true, true);
            }
        }

        private void Enter_tab16(object sender, KeyEventArgs e) //Makes enter work as tab
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
                SelectNextControl(shiharaimaisu_input, true, true, true, true);
            }
        }

        private void topExitbutton_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void minimizeButton_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }
    }
}
