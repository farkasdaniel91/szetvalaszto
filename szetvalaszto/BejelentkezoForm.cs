using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Configuration;
namespace szetvalaszto
{
    public partial class BejelentkezoForm : Form
    {
        public static List<Par> Parok;
        public static string Hely = ConfigurationManager.AppSettings["Hely"];
        public static string ParokXlsx = ConfigurationManager.AppSettings["Parokxlsx"];
        public static bool isAdmin = Convert.ToBoolean(ConfigurationManager.AppSettings["isAdmin"]);
        public BejelentkezoForm()
        {
            InitializeComponent();
            this.LoadParok();
        }

        private void LoadParok()
        {
            BejelentkezoForm.Parok = new List<Par>();

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(Hely + ParokXlsx, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            Microsoft.Office.Interop.Excel.Range range = xlWorkSheet.UsedRange;

            string par = string.Empty;
            int evfolyam = 0;
            int rCnt = 0;
            int cCnt = 0;

            for (rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
            {
                for (cCnt = 1; cCnt <= 2; cCnt++)
                {
                    switch (cCnt)
                    {
                        case 1:
                            par = (range.Cells[rCnt, cCnt] as Microsoft.Office.Interop.Excel.Range).Value2.ToString();
                            break;
                        case 2:
                            evfolyam = Convert.ToInt32((range.Cells[rCnt, cCnt] as Microsoft.Office.Interop.Excel.Range).Value2);
                            break;
                    }
                }
                BejelentkezoForm.Parok.Add(new Par(evfolyam, par));
            }

            xlWorkBook.Close(true, null, null);
            xlApp.Quit();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if (isAdmin)
            {
                this.button2.Visible = true;
            }
            else
            {
                this.button2.Visible = false;
            }

            this.comboBox1.Items.Add("");
            this.comboBox1.Items.AddRange(BejelentkezoForm.Parok.Select(x => x.par).ToArray());
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (this.comboBox1.SelectedItem.ToString() == string.Empty)
            {
                return;
            }

            BejelentkezoForm.Parok.Remove(BejelentkezoForm.Parok.Where(x => x.par == this.comboBox1.SelectedItem.ToString()).First());
            PreferenciaMakerForm asd = new PreferenciaMakerForm(this.comboBox1.SelectedItem.ToString(), BejelentkezoForm.Parok);
            this.Hide();
            asd.ShowDialog();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SzetvalasztoHelper.CalculateEredmenyz();
        }
    }
}
