using Microsoft.Build.Framework;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace EspritLogger
{
    //TODO make this private
    public class ExcelFileHandler
    {

        private Excel.Application m_XLApp;
        private Excel._Workbook m_WorkBook;
        private Excel._Worksheet m_Sheet;

        private int m_tolerance;
        private string m_logName;

        private Dictionary<string, int> m_columnsMap;
        private Dictionary<string, ProjectData> m_currProjectWarnings;

        private ExcelFileHandler() { }

        public ExcelFileHandler(Dictionary<string, ProjectData> projectWarnings, string fileName) 
        {
            m_XLApp = new Excel.Application();
            m_XLApp.Visible = false;
            m_XLApp.DisplayAlerts = false;

            m_currProjectWarnings = projectWarnings;
            m_tolerance = Constants.DELTA_TOLERANCE;

            m_logName = fileName;
        }

        public bool UpdateLog(bool buildSucseeded) 
        {
            if (!OpenWorkBook())
            {
                if (!CreateWorkBook())
                {
                    return false;
                }
            }
            AddNewDataRow(buildSucseeded);
            return true;
        }

        /**
         * Updates exsiting regression statistics with new data
        **/
        private void AddNewDataRow(bool buildSucseeded) 
        {
            int row = m_Sheet.UsedRange.Rows.Count + 1;
            int column = 1;
            string buildTime = DateTime.Now.ToString();
            
            //int totalCount = 0;
            if (buildSucseeded)
            {
                m_Sheet.Cells[row, column] = buildTime;
                foreach (KeyValuePair<string, ProjectData> entry in m_currProjectWarnings)
                {
                    if (!m_columnsMap.ContainsKey(entry.Key))
                    {
                        AddNewProjectColumn(entry.Key);
                    }

                    column = m_columnsMap[entry.Key];
                    m_Sheet.Cells[row, column] = entry.Value.GetWarningCount();
                    ApplyColor(entry.Value.GetWarningCount(), column, row);
                    //totalCount += entry.Value;
                }
                //m_Sheet.Cells[row, column + 1] = totalCount;
            }
            else 
            {
                m_Sheet.Cells[row, column] = buildTime + "(Build failed)";
                int startColumn = column;
                int endColumn = m_Sheet.UsedRange.Columns.Count;

                Excel.Range range = (Excel.Range)m_Sheet.Range[m_Sheet.Cells[row, column], 
                    m_Sheet.Cells[row, endColumn]];

                range.Interior.Color = ColorTranslator.FromHtml(Constants.COLOR_RED);
            }
        }

        /**
         * Adds project to the spreadsheet as a new collumn
        **/
        private void AddNewProjectColumn(string projectName) 
        {
            int columnId = m_Sheet.UsedRange.Columns.Count + 1;
            m_Sheet.Cells[1, columnId] = projectName;
            m_columnsMap.Add(projectName, columnId);
        }
        /**
         * Results of warnings are colored red means regression, green the opposite
        **/
        private void ApplyColor(int currValue, int column, int row) 
        {
            if (row > 2)
            {
                int lastValue = GetIntFromCell(row - 1, column);
                int delta = (currValue - lastValue);
                if (delta > m_tolerance)
                {
                    ((Excel.Range)m_Sheet.Cells[row, column]).Interior.Color =
                        ColorTranslator.FromHtml(Constants.COLOR_RED);
                }
                else if (delta < 0)
                {
                    ((Excel.Range)m_Sheet.Cells[row, column]).Interior.Color =
                        ColorTranslator.FromHtml(Constants.COLOR_GREEN);
                }
                else if (delta == 0) 
                {
                    ((Excel.Range)m_Sheet.Cells[row, column]).Interior.Color =
                        ColorTranslator.FromHtml(Constants.COLOR_YELLOW);
                } 
            }
            else 
            {
                //TODO Change base row color
            }
        }
        /**
         * if no log file exists this method creates new spreadsheet
        **/
        private bool CreateWorkBook() 
        {
            m_WorkBook = (Excel._Workbook)(m_XLApp.Workbooks.Add(Missing.Value));
            m_Sheet = (Excel._Worksheet) m_WorkBook.ActiveSheet;
            if (m_WorkBook == null || m_Sheet == null)
            { 
                return false;
            }

            m_Sheet.Cells[1, 1] = "Date";
            m_columnsMap = new Dictionary<string, int>();
            foreach (KeyValuePair<string, ProjectData> entry in m_currProjectWarnings)
            {
                AddNewProjectColumn(entry.Key);
            }
            //addColumn("Total");

            ApplyStyle();

            return true;
        }
        /**
         * Here goes all style formating of excel spreadsheet
        **/
        private void ApplyStyle()
        {
            Excel.Range dateColumn = m_Sheet.get_Range("A:A", System.Type.Missing);
            dateColumn.EntireColumn.ColumnWidth = Constants.DATE_COLUMN_WIDTH;

            Excel.Range projectCollumns = m_Sheet.get_Range("B:AA", System.Type.Missing);
            projectCollumns.EntireColumn.ColumnWidth = Constants.PROJECTS_COLUMN_WIDTH;
        }
        /**
         * if previous logs exist we open it and update it
         **/
        public bool OpenWorkBook()
        {

            if (!File.Exists(m_logName)) 
            {
                return false;
            }

            try
            {
                m_WorkBook = (Excel._Workbook)(m_XLApp.Workbooks.Open(m_logName));
                m_Sheet = (Excel._Worksheet)m_WorkBook.ActiveSheet;
            }
            catch (Exception e)
            {
                throw new LoggerException("Cannot open log file: " + e.Message);
            }

            if (m_Sheet == null)
            {
                return false;
            }

            m_columnsMap = new Dictionary<string, int>();
            
            string projectName;
            int lastRow = m_Sheet.UsedRange.Rows.Count;
            for (int column = 2; column < m_Sheet.UsedRange.Columns.Count + 1; column++)
            {
                projectName = GetStringFromCell(1, column);
                m_columnsMap.Add(projectName, column);
            }
            return true;
        }
        /**
         * helper function to parce values in excel cells
         **/
        private string GetStringFromCell(int row, int column) 
        {
            string res = "";
            try 
	        {	        
		        Excel.Range objRange = (Excel.Range)m_Sheet.Cells[1, column];
                res = objRange.get_Value(Missing.Value).ToString();
	        }
	        catch (Exception e)
	        {
		        throw new LoggerException("Cannot read cell value as string type: " + e.Message);
	        }
            return res;
        }

        private int GetIntFromCell(int row, int column) 
        {
            int res = 0;
            try 
            {
                Excel.Range objRange = (Excel.Range) m_Sheet.Cells[row, column];
                res = Convert.ToInt32(objRange.Value);
            }
            catch(Exception e)
            {
                throw new LoggerException("Cannot read cell value as integral type: " + e.Message);
            }
            return res;
        }

        public bool SaveLogFileAs() 
        {
            try
            {
                m_WorkBook.SaveAs(m_logName);
            }
            catch (Exception e)
            {
                throw new LoggerException("Cannot save log file: " + e.Message);
            }
            this.Close();
            return true;
        }

        private void Close()
        {
            // Cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.FinalReleaseComObject(m_Sheet);

            m_WorkBook.Close(Type.Missing, Type.Missing, Type.Missing);
            Marshal.FinalReleaseComObject(m_WorkBook);

            m_XLApp.Quit();
            Marshal.FinalReleaseComObject(m_XLApp);
        }

        internal void UpdateAsFailure()
        {
            throw new NotImplementedException();
        }
    }
}
