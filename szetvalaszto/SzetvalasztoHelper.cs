using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace szetvalaszto
{
    public static class SzetvalasztoHelper
    {

        #region Konstansok

        public static string PrefXlsx = ConfigurationManager.AppSettings["Prefxlsx"];
        public static string EredmenyXlsx = ConfigurationManager.AppSettings["Eredmenyxlsx"];
        public static int TaborokSzama = Convert.ToInt32(ConfigurationManager.AppSettings["TaborokSzama"]);
        public static int elsotaborletszama = Convert.ToInt32(ConfigurationManager.AppSettings["elsotaborletszama"]);
        public static int masodiktaborletszama = Convert.ToInt32(ConfigurationManager.AppSettings["masodiktaborletszama"]);
        public static int harmadiktaborletszama = Convert.ToInt32(ConfigurationManager.AppSettings["harmadiktaborletszama"]);
        public static int negyediktaborletszama = Convert.ToInt32(ConfigurationManager.AppSettings["negyediktaborletszama"]);

        #endregion

        public static List<Preferencia> preferenciak;
        public static List<Preferencia> Kotesek;
        public static List<Tabor> Taborok;

        public static void ExportPreferenciak(List<Preferencia> prefz)
        {
            Microsoft.Office.Interop.Excel.Workbook mWorkBook;
            Microsoft.Office.Interop.Excel.Sheets mWorkSheets;
            Microsoft.Office.Interop.Excel.Worksheet mWSheet1;
            Microsoft.Office.Interop.Excel.Application oXL;

            string path = BejelentkezoForm.Hely + SzetvalasztoHelper.PrefXlsx;
            oXL = new Microsoft.Office.Interop.Excel.Application();
            oXL.Visible = false;
            oXL.DisplayAlerts = false;
            mWorkBook = oXL.Workbooks.Open(path, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //Get all the sheets in the workbook
            mWorkSheets = mWorkBook.Worksheets;
            //Get the allready exists sheet
            mWSheet1 = (Microsoft.Office.Interop.Excel.Worksheet)mWorkSheets.get_Item(1);
            Microsoft.Office.Interop.Excel.Range range = mWSheet1.UsedRange;

            int colCount = range.Columns.Count;
            int rowCount = range.Rows.Count;

            if (rowCount <= 1)
            {
                mWSheet1.Cells[1, 1] = "Választó";
                mWSheet1.Cells[1, 2] = "Választott";
                mWSheet1.Cells[1, 3] = "PreferenciaPont";
            }

            for (int index = 1; index < prefz.Count + 1; index++)
            {
                mWSheet1.Cells[rowCount + index, 1] = prefz[index - 1].valaszto;
                mWSheet1.Cells[rowCount + index, 2] = prefz[index - 1].valasztott;
                mWSheet1.Cells[rowCount + index, 3] = prefz[index - 1].prefpont;
            }
            mWorkBook.SaveAs(path, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault,
            Missing.Value, Missing.Value, Missing.Value, Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive,
            Missing.Value, Missing.Value, Missing.Value,
            Missing.Value, Missing.Value);
            mWorkBook.Close(Missing.Value, Missing.Value, Missing.Value);
            mWSheet1 = null;
            mWorkBook = null;
            oXL.Quit();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        #region Calculation
        public static void CalculateEredmenyz()
        {
            readprefz();

            LoadKotesek();

            InitTaborok();
            
            //CreateResult();
            
        }

        private static void InitTaborok()
        {
            int[] taborokletszama = new int[4];
            taborokletszama[0] = elsotaborletszama;
            taborokletszama[1] = masodiktaborletszama;
            taborokletszama[2] = harmadiktaborletszama;
            taborokletszama[3] = negyediktaborletszama;

            Taborok = new List<Tabor>(TaborokSzama);
            for (int i = 0; i < TaborokSzama; i++)
            {
                Taborok.Add(new Tabor(taborokletszama[i]));
            }
        }

        /// <summary>
        /// Ez annyit csinál hogy veszi a preferencia sorokat és kötések formájában összesíti őket
        /// Pl:
        /// preferencia: a köt b-hez 5 ponttal
        /// preferencia: b köt a-hoz 5 ponttal
        /// implájing
        /// Kötés: a köt b-hez 10 ponttal
        /// </summary>
        private static void LoadKotesek()
        {
            Kotesek = new List<Preferencia>();
            foreach (Preferencia pref in preferenciak)
            {
                int prefpontCounter = pref.prefpont;
                List<Preferencia> prflst = Kotesek.Where(x => x.valasztott == pref.valaszto).ToList().Where(x => pref.valasztott == x.valaszto).ToList();
                if (prflst.Count > 0)
                {
                    continue;
                }

                foreach (var prf in preferenciak.Where(x => x.valasztott == pref.valaszto).ToList().Where(x => pref.valasztott == x.valaszto).ToList())
                {
                    prefpontCounter += prf.prefpont;
                }
                Kotesek.Add(new Preferencia(pref.valaszto, pref.valasztott, prefpontCounter, pref.kaszt));
            }
            Kotesek = Kotesek.OrderByDescending(x => x.prefpont).ToList();
        }

        private static void CreateResult()
        {
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlApp.Visible = true;

            Workbook wb = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            Worksheet ws = (Worksheet)wb.Worksheets[1];
        }

        public static void readprefz()
        {
            SzetvalasztoHelper.preferenciak = new List<Preferencia>();

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(BejelentkezoForm.Hely + PrefXlsx, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            Microsoft.Office.Interop.Excel.Range range = xlWorkSheet.UsedRange;

            string valaszto = string.Empty;
            string valasztott = string.Empty;
            int prefpont = 0;
            int rCnt = 0;
            int cCnt = 0;

            for (rCnt = 2; rCnt <= range.Rows.Count; rCnt++)
            {
                for (cCnt = 1; cCnt <= 3; cCnt++)
                {
                    switch (cCnt)
                    {
                        case 1:
                            valaszto = (range.Cells[rCnt, cCnt] as Microsoft.Office.Interop.Excel.Range).Value2.ToString();
                            break;
                        case 2:
                            valasztott = (range.Cells[rCnt, cCnt] as Microsoft.Office.Interop.Excel.Range).Value2.ToString();
                            break;
                        case 3:
                            prefpont = Convert.ToInt32((range.Cells[rCnt, cCnt] as Microsoft.Office.Interop.Excel.Range).Value2);                            break;
                    }
                }

                int evfolyam = BejelentkezoForm.Parok.Where(x => x.par == valaszto).Select(x => x.kaszt).First();
                SzetvalasztoHelper.preferenciak.Add(new Preferencia(valaszto, valasztott, prefpont, evfolyam));
            }

            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            SzetvalasztoHelper.preferenciak = SzetvalasztoHelper.preferenciak.OrderByDescending(x => x.prefpont).ThenByDescending(x => x.kaszt).ToList<Preferencia>();
        }

        #endregion

    }
}
