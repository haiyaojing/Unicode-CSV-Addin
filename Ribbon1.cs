﻿/*
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
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

namespace UnicodeCSVAddin
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon1()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("UnicodeCSVAddin.Ribbon1.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, select the Ribbon XML item in Solution Explorer and then press F1

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
            Globals.ThisAddIn.ribbon = this;
        }

        public void OpenFile(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.OpenFileDialog();
        }

        public void OpenFileEx(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.OpenFileDialogEx();
        }

        public void SaveButtonAction(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.SaveAsUnicodeCSV(true, false);
        }

        public void SaveAsButtonAction(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.SaveAsUnicodeCSV(true, true);
        }

        private List<string> _internalValidList = new List<string>();
        private List<string> _internalValidPathList = new List<string>();
        private int _internalSelectIndex = 0;
        public int getItemCount(Office.IRibbonControl control)
        {
            return _internalValidList.Count;
        }

        public string getItemLabel(Office.IRibbonControl control, int index)
        {
            return index >= 0 && index < _internalValidList.Count ? _internalValidList[index] : string.Empty;
        }

        public int getSelectedItemIndex(Office.IRibbonControl control)
        {
            return _internalSelectIndex;
        }
        public void onValueChanged(Office.IRibbonControl control, object s, int index)
        {
            _internalSelectIndex = index;

            if (index >= 0 && index < _internalValidPathList.Count)
            {
                Globals.ThisAddIn.SafeOpenFile(_internalValidPathList[index]);
                return;
            }

            MessageBox.Show("无此表");
        }

        #endregion

        #region Helpers

        public void AddRecorders(string folder, string orginal)
        {
            _internalValidList.Clear();
            _internalValidPathList.Clear();
            _internalSelectIndex = 0;
            if (Directory.Exists(folder))
            {
                if (folder.Replace("\\", "/").Contains("Table/CSV"))
                {
                    var files = Directory.GetFiles(folder);
                    for(int i = 0; i < files.Length; i++)
                    {
                        var file = files[i];
                        var fileName = Path.GetFileNameWithoutExtension(file);
                        _internalValidList.Add(fileName);
                        _internalValidPathList.Add(file);

                        if (orginal.Equals(fileName))
                        {
                            _internalSelectIndex = i;
                        }
                    }
                }
            }

            this.ribbon.InvalidateControl("dropdown1");
        }

        public Bitmap GetImage(Office.IRibbonControl control)
        {
            return Properties.Resources.Boli;
        }

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
