﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace WeekEndingTabs
{
    public partial class Sheet5
    {
        private void Sheet5_Startup(object sender, System.EventArgs e)
        {
        }

        private void Sheet5_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(Sheet5_Startup);
            this.Shutdown += new System.EventHandler(Sheet5_Shutdown);
        }

        #endregion

    }
}
