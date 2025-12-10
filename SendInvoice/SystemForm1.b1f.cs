using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;

namespace SendInvoice
{
    [FormAttribute("0", "SystemForm1.b1f")]
    class SystemForm1 : SystemFormBase
    {
        public SystemForm1()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            if (Program.sendinvoiceekrani)
            {
                this.UIAPIRawForm.Items.Item("1").Click();
            }
        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
        }
    }
}
