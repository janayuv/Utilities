using System;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Core;

namespace utilities
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            // Additional startup logic if needed
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            // Cleanup logic if needed
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return Globals.Factory.GetRibbonFactory().CreateRibbonManager(new[] { new UtilitiesRibbon() });
        }

        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }
    }
}
