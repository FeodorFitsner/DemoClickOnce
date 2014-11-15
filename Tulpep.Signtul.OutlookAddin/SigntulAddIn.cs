
namespace Tulpep.Signtul.OutlookAddin
{
    public partial class SigntulAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //ConfigureSignature.ConfigureSignatures(this.Application.Session.Accounts.Cast<Account>().Select(x => x.SmtpAddress));
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new RibbonMailtab();
        }


        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
