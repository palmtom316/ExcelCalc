using System;
using Microsoft.Office.Tools.Ribbon;

namespace ExcelSmartCostAddIn
{
    public partial class Ribbon1
    {
        private async void OnRecalcSheet(IRibbonControl control)
        {
            try
            {
                await QuantityRecalc.RecalculateActiveSheetAsync();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message, "ExcelSmartCost");
            }
        }

        private void OnImportLibrary(IRibbonControl control)
        {
            System.Windows.Forms.MessageBox.Show("TODO: Import dialog (boq/quota/enterprise/components).", "ExcelSmartCost");
        }

        private void OnManageComponents(IRibbonControl control)
        {
            System.Windows.Forms.MessageBox.Show("TODO: Components manager.", "ExcelSmartCost");
        }
    }
}
