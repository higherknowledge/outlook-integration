using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;

namespace HigherKnowledge_addin
{
    public partial class ThisAddIn
    {
        static string myMail;

        public static string User
        {
            get
            {
                return myMail;
            }
        }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
           try
           {
                Outlook.Recipient rec = Application.Session.CurrentUser;
                string type = rec.AddressEntry.Type;
                if (type.Equals("EX"))
                    myMail = rec.AddressEntry.GetExchangeUser().PrimarySmtpAddress;
                else
                    myMail = rec.AddressEntry.Address;
            }
            catch(Exception)
            {
                MessageBox.Show("Could not fetch the user...");
            }
        }

        void ItemsOpen(ref bool Item)
        {
            //return false;
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon1();
        }



        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see http://go.microsoft.com/fwlink/?LinkId=506785
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
