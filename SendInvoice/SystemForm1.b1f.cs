using System;
using SAPbouiCOM.Framework;

namespace SendInvoice
{
    [FormAttribute("0", "SystemForm1.b1f")]
    class SystemForm1 : SystemFormBase
    {
        // Filter listesi: B1 dilinden bagimsiz olabilmek icin TR + EN anahtar kelimeler.
        // Yeni dil/keyword eklemek icin sadece bu listeye ekle.
        private static readonly string[] AutoOkTitleKeywords = new[]
        {
            "send",
            "gonder",
            "gönder",
            "onay",
            "emin",
            "confirm",
            "fatura"
        };

        public SystemForm1()
        {
        }

        public override void OnInitializeComponent()
        {
            // sendinvoiceekrani = true degilse, kullanicinin kendi acmis oldugu
            // bir sistem mesajidir, dokunmuyoruz.
            if (!Program.sendinvoiceekrani) return;

            try
            {
                string title = "";
                try { title = this.UIAPIRawForm.Title ?? ""; } catch { }

                string btn1Caption = "";
                try
                {
                    var btn = this.UIAPIRawForm.Items.Item("1").Specific as SAPbouiCOM.Button;
                    if (btn != null) btn1Caption = btn.Caption ?? "";
                }
                catch { }

                Program.Log("SystemForm acildi - Title=\"" + title + "\" Btn1=\"" + btn1Caption + "\"");

                //if (ShouldAutoOk(title))
                //{
                Program.Log("SystemForm auto-OK basiliyor");
                this.UIAPIRawForm.Items.Item("1").Click();
                //}
                //else
                //{
                //    // Beklenmeyen bir dialog gelmis: TIKLAMA. Kullanici gorsun, log'da kalsin.
                //    Program.Log("SystemForm auto-OK SKIP (title eslesmedi)");
                //}
            }
            catch (Exception ex)
            {
                Program.Log("SystemForm OnInitializeComponent error: " + ex.Message);
            }
        }

        private static bool ShouldAutoOk(string title)
        {
            if (string.IsNullOrWhiteSpace(title)) return false;
            string lower = title.ToLowerInvariant();
            foreach (string kw in AutoOkTitleKeywords)
            {
                if (lower.Contains(kw)) return true;
            }
            return false;
        }

        public override void OnInitializeFormEvents()
        {
        }
    }
}
