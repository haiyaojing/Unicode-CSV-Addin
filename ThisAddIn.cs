/*
    Copyright 2011 Jaimon Mathew www.jaimon.co.uk

   Licensed under the Apache License, Version 2.0 (the "License");
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.     
 
*/

using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using System.Text;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace UnicodeCSVAddin
{
    public partial class ThisAddIn
    {
        private Excel.Application _app;
        private List<string> _unicodeFiles; //a list of opened Unicode CSV files. We populate this list on WorkBookOpen event to avoid checking for CSV files on every Save event.
        private bool _sFlag = false;

        private HashSet<string> _openFiles;

        //Unicode file byte order marks.
        private const string UTF_16BE_BOM = "FEFF";
        private const string UTF_16LE_BOM = "FFFE";
        private const string UTF_8_BOM = "EFBBBF";
        public Ribbon1 ribbon;

        private int? _lastScrollRow;
        private int? _lastScrollColumn;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            _app = this.Application;
            _unicodeFiles = new List<string>();
            _openFiles = new HashSet<string>();

            _lastScrollRow = null;
            _lastScrollColumn = null;

            _app.WorkbookOpen += app_WorkbookOpen;
            _app.WorkbookBeforeClose += app_WorkbookBeforeClose;
            _app.WorkbookBeforeSave += app_WorkbookBeforeSave;

        }


        void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            _app = null;
            _unicodeFiles = null;
            _lastScrollRow = null;
            _lastScrollColumn = null;
        }

        void app_WorkbookBeforeSave(Excel.Workbook Wb, bool SaveAsUI, ref bool Cancel)
        {
     
            //Override Save behaviour for Unicode CSV files.
            if (!SaveAsUI && !_sFlag && _unicodeFiles.Contains(Wb.FullName))
            {
                Cancel = true;
                SaveAsUnicodeCSV(false, false);
            }
            _sFlag = false;
        }

        //This is required to show our custom Ribbon
        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon1();
        }

        void app_WorkbookBeforeClose(Excel.Workbook wb, ref bool cancel)
        {
            if (wb.FullName.ToLower().EndsWith(".csv"))
            {
                var name = wb.FullName;
                if (_app.ActiveWorkbook.Saved == false && _openFiles.Contains(wb.FullName))
                {
                    var ret = MessageBox.Show($"是否保存对\"{Path.GetFileName(wb.FullName)}\"", "UnicodeCSVAddin", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
                    if (ret == DialogResult.Cancel)
                    {
                        cancel = true;
                        return;
                    }

                    _app.ActiveWorkbook.Saved = true;
                    string status = string.Empty;

                    if (ret == DialogResult.Yes)
                    {
                        status = "Saved";
                        SaveAsUnicodeCSV(true, false, false);
                    }
                    else if (ret == DialogResult.No)
                    {
                        status = "Closed";
                    }
                    
                    _app.StatusBar = status;
                }
                _openFiles.Remove(name);
                _unicodeFiles.Remove(name);
            }
        }

        void app_WorkbookOpen(Excel.Workbook wb)
        {
            //Check to see if the opened document is a Unicode CSV files, so we can override Excel's Save method
            if (wb.FullName.ToLower().EndsWith(".csv"))
            {
                if (_openFiles.Contains(wb.FullName))
                {
                    if (isFileUnicode(wb.FullName))
                    {
                        if (!_unicodeFiles.Contains(wb.FullName))
                        {
                            _unicodeFiles.Add(wb.FullName);
                        }
                        _app.StatusBar = wb.Name + " has been opened as a Unicode CSV file";

                        // 冻结前两行
                        Excel.Worksheet worksheet = _app.ActiveSheet; //当前活动sheet
                        Excel.Range usedRange = worksheet?.UsedRange;   //获取使用的格子二维数组
                        var rowCount = usedRange?.Rows?.Count;
                        _app.ActiveWindow.FreezePanes = false;
                        if (rowCount != null && rowCount >= 2)
                        {
                            _app.ActiveWindow.SplitRow = 2;
                            _app.ActiveWindow.FreezePanes = true;
                        }

                        var columnCount = usedRange?.Columns?.Count;
                        if (columnCount != null && columnCount >= 2)
                        {
                            _app.ActiveWindow.SplitColumn = 2;
                            _app.ActiveWindow.FreezePanes = true;
                        }

                        if (_lastScrollRow != null)
                        {
                            _app.ActiveWindow.ScrollRow = _lastScrollRow.Value;
                            _lastScrollRow = null;
                        }

                        if (_lastScrollColumn != null)
                        {
                            _app.ActiveWindow.ScrollColumn = _lastScrollColumn.Value;
                            _lastScrollColumn = null;
                        }

                        _app.ActiveWorkbook.Saved = true;
                        //                        this.ribbon?.AddRecorders(Path.GetDirectoryName(wb.FullName), Path.GetFileNameWithoutExtension(wb.FullName));
                    }
                }
                else
                {
                    if (!isFileUnicode(wb.FullName))
                    {
                        MessageBox.Show("请关闭表格并且将格式改为UTF8-BOM", "警告");
                        return;
                    }

                    var fileName = wb.FullName;
                    _app.ActiveWorkbook.Close();
                    SafeOpenFile(fileName);
                }
            }
            else
            {
                _app.StatusBar = "Ready";
            }
        }

        /// <summary>
        /// This method check whether Excel is in Cell Editing mode or not
        /// There are few ways to check this (eg. check to see if a standard menu item is disabled etc.)
        /// I know in cell editing mode app.DisplayAlerts throws an Exception, so here I'm relying on that behaviour
        /// </summary>
        /// <returns>true if Excel is in cell editing mode</returns>
        private bool isInCellEditingMode()
        {
            bool flag = false;
            try
            {
                _app.DisplayAlerts = false; //This will throw an Exception if Excel is in Cell Editing Mode
            }
            catch (Exception)
            {
                flag = true;
            }
            return flag;
        }

        public void OpenFileDialog()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "csv文件(*.csv)|*.csv";
            openFileDialog.RestoreDirectory = true;
            openFileDialog.FilterIndex = 1;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                openFile(openFileDialog.FileName);
            }
        }

        string CSVReadLine(StreamReader sr)
        {
            StringBuilder lineStrBuilder = new StringBuilder();
            int quotationMarksCount = 0;  // 引号计数
            int content = sr.Read();
            while (content != -1)
            {
                char c = Convert.ToChar(content);
                //遇到换行符 切引号数两位偶数 则认为是正式的换行符
                //这里不用\r判断 是因为特殊情况下流里面剩下一个\n不好处理
                if (c == '\n' && quotationMarksCount % 2 == 0)
                {
                    if (lineStrBuilder[lineStrBuilder.Length - 1] == '\r')  //正式换行为\r\n 要去掉前一个读入的\r
                        lineStrBuilder.Remove(lineStrBuilder.Length - 1, 1);
                    return lineStrBuilder.ToString();
                }
                //读到引号计数
                if (c == '"')
                    quotationMarksCount++;

                lineStrBuilder.Append(c);
                content = sr.Read();
            }
            //如果字符串构造器中有值 则返回对应字符串
            if (lineStrBuilder.Length > 0)
                return lineStrBuilder.ToString();
            //构造器中无内容 则认为已经读完
            return null;
        }

        void formatCSV(string fileName)
        {
            List<string> tmp = new List<string>();
            using (var fs = new FileStream(fileName, FileMode.Open))
            {
                using (var sr = new StreamReader(fs, Encoding.Default))
                {
                    while (true)
                    {
                        var line = CSVReadLine(sr);
                        if (string.IsNullOrEmpty(line))
                        {
                            break;
                        }

                        line = line.Replace("\r\n", "\\r\\n");
                        line = line.Replace("\r", "\\r").Replace("\n", "\\n");
                        tmp.Add(line);
                    }
                }
            }

            using (var fs = new FileStream(fileName, FileMode.Create))
            {
                using (var sw = new StreamWriter(fs, Encoding.UTF8))
                {
                    foreach (var line in tmp)
                    {
                        sw.WriteLine(line);
                    }
                }
            }
        }

        public void SafeOpenFile(string fileName)
        {
            formatCSV(fileName);
            openFile(fileName);
        }

        internal void OpenFileDialogEx()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "csv文件(*.csv)|*.csv";
            openFileDialog.RestoreDirectory = true;
            openFileDialog.FilterIndex = 1;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                var filename = openFileDialog.FileName;
                SafeOpenFile(filename);
            }
        }

        private void openFile(string filePath)
        {
            try
            {
                string[] list;
                using (var fs = new FileStream(filePath, FileMode.Open))
                {
                    using (var sr = new StreamReader(fs, Encoding.Default))
                    {
                        var line = sr.ReadLine();
                        if (string.IsNullOrEmpty(line))
                        {
                            MessageBox.Show($"{filePath}第一行文本为空");
                            return;
                        }

                        list = line.Split(new[] { ',' });
                    }
                }

                int columns = list.Length;
                var fieldInfo = new int[columns, 2];
                for (int i = 0; i < columns; i++)
                {
                    fieldInfo[i, 0] = i;
                    fieldInfo[i, 1] = 2;
                }

                _openFiles.Add(filePath);

                Globals.ThisAddIn.Application.Workbooks.OpenText(filePath, 65001, Comma: true,
                    DataType: Excel.XlTextParsingType.xlDelimited, TextQualifier: Excel.XlTextQualifier.xlTextQualifierDoubleQuote, FieldInfo: fieldInfo, Tab: false);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                throw;
            }
        }


        /// <summary>
        /// This will create a temporary file in Unicode text (*.txt) format, overwrite the current loaded file by replaing all tabs with a comma and reload the file.
        /// </summary>
        /// <param name="force">To force save the current file as a Unicode CSV.
        /// When called from the Ribbon items Save/SaveAs, <i>force</i> will be true
        /// If this parameter is true and the file name extention is not .csv, then a SaveAs dialog will be displayed to choose a .csv file</param>
        /// <param name="newFile">To show a SaveAs dialog box to select a new file name
        /// This will be set to true when called from the Ribbon item SaveAs</param>
        public void SaveAsUnicodeCSV(bool force, bool newFile, bool needReOpen = true)
        {
//            _app.ActiveWindow.FreezePanes = false;
            
            _app.StatusBar = "";
            bool currDispAlert = _app.DisplayAlerts;
            bool flag = true;
            int i;
            string filename = _app.ActiveWorkbook.FullName;
            GC.Collect();
            if (force) //then make sure a csv file name is selected.
            {
                if (newFile || !filename.ToLower().EndsWith(".csv"))
                {
                    Office.FileDialog d = _app.get_FileDialog(Office.MsoFileDialogType.msoFileDialogSaveAs);
                    i = _app.ActiveWorkbook.Name.LastIndexOf(".");
                    if (i >= 0)
                    {
                        d.InitialFileName = _app.ActiveWorkbook.Name.Substring(0, i);
                    }
                    else
                    {
                        d.InitialFileName = _app.ActiveWorkbook.Name;
                    }
                    d.AllowMultiSelect = false;
                    Office.FileDialogFilters f = d.Filters;
                    for (i = 1; i <= f.Count; i++)
                    {
                        if ("*.csv".Equals(f.Item(i).Extensions))
                        {
                            d.FilterIndex = i;
                            break;
                        }
                    }
                    if (d.Show() == 0) //User cancelled the dialog
                    {
                        flag = false;
                    }
                    else
                    {
                        filename = d.SelectedItems.Item(1);
                    }
                }
                if (flag && !filename.ToLower().EndsWith(".csv"))
                {
                    MessageBox.Show("Please select a CSV file name first");
                    flag = false;
                }
            }

            if (flag && filename.ToLower().EndsWith(".csv") && (force || _unicodeFiles.Contains(filename)))
            {
                if (isInCellEditingMode())
                {
                    MessageBox.Show("Please finish editing before saving");
                }
                else
                {
                    try
                    {
                        //Getting current selection to restore the current cell selection
//                        Excel.Range rng = app.ActiveCell;
//                        int row = rng.Row;
//                        int col = rng.Column;

                        string tempFile = System.IO.Path.GetTempFileName();
                        var orignalTempFileName = tempFile;
                        var tmpFileName = Path.GetFileNameWithoutExtension(tempFile);
                        tempFile = tempFile.Replace(tmpFileName,
                            Path.GetFileNameWithoutExtension(filename) + "_" + tmpFileName);

                        try
                        {
                            _sFlag = true; //This is to prevent this method getting called again from app_WorkbookBeforeSave event caused by the next SaveAs call
                            using (var sw = new StreamWriter(tempFile, false, Encoding.UTF8))
                            {
                                Excel.Worksheet worksheet = _app.ActiveSheet; //当前活动sheet
                                Excel.Range usedRange = worksheet.UsedRange;   //获取使用的格子二维数组
                                _lastScrollRow = _app.ActiveWindow.ScrollRow;
                                _lastScrollColumn = _app.ActiveWindow.ScrollColumn;
                                if (usedRange == null)
                                {
                                    MessageBox.Show("不可预知的异常");
                                    return;
                                }

                                var rowCount = usedRange.Rows.Count;
                                var columnCount = usedRange.Columns.Count;

                                var arr = new object[rowCount, columnCount];
                                arr = usedRange.Value;

                                for (int j = 1; j <= rowCount; j++)
                                {
                                    for (int jj = 1; jj <= columnCount; jj++)
                                    {
                                        if (jj != 1)
                                        {
                                            sw.Write(",");
                                        }

                                        var o = arr[j, jj];
                                        var value = o != null ? o.ToString() : string.Empty;
                                        if (value.Contains("\""))
                                            value = value.Replace("\"", "\"\"");
                                        if (value.Contains(",") || value.Contains("\"") || value.Contains("\n"))
                                            value = $"\"{value}\"";
                                        
                                        sw.Write(value);
                                        
                                    }

                                    sw.WriteLine();
                                }
                            }
                            _openFiles.Remove(filename);
                            _app.ActiveWorkbook.Close();
                            GC.Collect();
                            if (new FileInfo(tempFile).Length <= (1024 * 1024)) //If its less than 1MB, load the whole data to memory for character replacement
                            {
                                File.WriteAllText(filename, File.ReadAllText(tempFile, Encoding.Default).Replace("\t", ","), Encoding.UTF8);
                            }
                            else //otherwise read chunks for data (in 10KB chunks) into memory
                            {
                                using (StreamReader sr = new StreamReader(tempFile, Encoding.Default))
                                using (StreamWriter sw = new StreamWriter(filename, false, Encoding.UTF8))
                                {
                                    char[] buffer = new char[10 * 1024]; //10KB Chunks
                                    while (!sr.EndOfStream)
                                    {
                                        int cnt = sr.ReadBlock(buffer, 0, buffer.Length);
                                        for (i = 0; i < cnt; i++)
                                        {
                                            if (buffer[i] == '\t')
                                            {
                                                buffer[i] = ',';
                                            }
                                        }
                                        sw.Write(buffer, 0, cnt);
                                    }
                                }
                            }
                        }
                        finally
                        {
                            File.Delete(orignalTempFileName);
                        }
                        GC.Collect();

                        if (!needReOpen) return;

                        openFile(filename);
                        _app.StatusBar = "File has been saved as a Unicode CSV";
                        if (!_unicodeFiles.Contains(filename))
                        {
                            _unicodeFiles.Add(filename);
                        }
                        _app.ActiveWorkbook.Saved = true;
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show("Error occured while trying to save this file as Unicode CSV: " + e.Message);
                    }
                    finally
                    {
                        _app.DisplayAlerts = currDispAlert;
                    }
                }
            }
        }

        /// <summary>
        /// This method will try and read the first few bytes to see if it contains a Unicode BOM
        /// </summary>
        /// <param name="filename">File to check for including full path</param>
        /// <returns>true if its a Unicode file</returns>
        private bool isFileUnicode(string filename)
        {
            bool ret = false;
            try
            {
                byte[] buff = new byte[3];
                using (FileStream fs = new FileStream(filename, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    fs.Read(buff, 0, 3);
                }

                string hx = "";
                foreach (byte letter in buff)
                {
                    hx += string.Format("{0:X2}", Convert.ToInt32(letter));
                    //Checking to see the first bytes matches with any of the defined Unicode BOM
                    //We only check for UTF8 and UTF16 here.
                    ret = UTF_8_BOM.Equals(hx);
                    if (ret)
                    {
                        break;
                    }
                }
            }
            catch (IOException)
            {
                //ignore any exception
            }
            return ret;
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += ThisAddIn_Startup;
            this.Shutdown += ThisAddIn_Shutdown;
        }
        
        #endregion
    }
}
