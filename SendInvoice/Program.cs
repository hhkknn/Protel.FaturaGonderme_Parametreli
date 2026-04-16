using SAPbobsCOM;
using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SendInvoice
{
    internal class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        private static void Main(string[] args)
        {
            try
            {
                Application oApp = null;
                if (args.Length < 1)
                {
                    oApp = new Application();
                }
                else
                {
                    oApp = new Application(args[0]);
                }
                Menu MyMenu = new Menu();
                MyMenu.AddMenuItems();
                oApp.RegisterMenuEventHandler(MyMenu.SBO_Application_MenuEvent);
                Application.SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
                oCompanyObject = ((SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany());
                oApp.AfterInitialized += oApp_AfterInitialized;
                oApp.Run();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        public static System.Timers.Timer timer = null;

        private static void oApp_AfterInitialized(object sender, EventArgs e)
        {
            timer = new System.Timers.Timer();
            timer.Interval = 3000 * 60;
            //timer.Interval = 9000;
            timer.Enabled = true;
            timer.Elapsed += timer_Elapsed;
            timer.Start();
        }

        public static bool sendinvoiceekrani = false;
        public static bool faturagondermeislemicalisiyor = false;

        private static void timer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            try
            {
                SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                SAPbobsCOM.Recordset oRS_1 = (SAPbobsCOM.Recordset)oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                SAPbobsCOM.Recordset oRS_2 = (SAPbobsCOM.Recordset)oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                SAPbobsCOM.Recordset oRS_3 = (SAPbobsCOM.Recordset)oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                SAPbobsCOM.Recordset oRS_4 = (SAPbobsCOM.Recordset)oCompanyObject.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                if (faturagondermeislemicalisiyor)
                {
                    return;
                }

                faturagondermeislemicalisiyor = true;

                
                string sql = "Select ISNULL(\"U_BasTar\",'1900-01-01') as BasTar,ISNULL(\"U_BitTar\",'1900-01-01') as BitTar,* from \"@DON_OTOPARAM\" ";
                oRS_4.DoQuery(sql);

                string kulkodu = "";
                string fatserisi = "";
                string arsivserisi = "";
                string bastar = "";
                string bittar = "";

                kulkodu = oRS_4.Fields.Item("U_KulKod").Value.ToString();

                if (kulkodu == "")
                {
                    faturagondermeislemicalisiyor = false;
                    return;
                }

                fatserisi = oRS_4.Fields.Item("U_FatSeri").Value.ToString();
                arsivserisi = oRS_4.Fields.Item("U_ArsivSeri").Value.ToString();
                bastar = oRS_4.Fields.Item("BasTar").Value.ToString();
                bittar = oRS_4.Fields.Item("BitTar").Value.ToString();

                List<string> kulKodlari = kulkodu
                    .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(x => x.Trim())
                    .Where(x => !string.IsNullOrEmpty(x))
                    .ToList();

                foreach (string kulKod in kulKodlari)
                {
                    try
                    {
                        // Her kullanıcı için formu YENİDEN aç
                        SAPbouiCOM.Framework.Application.SBO_Application.ActivateMenuItem("DonusumOne.SendInvoice");

                        // Form referanslarını her iterasyonda YENİDEN al
                        SAPbouiCOM.Form oform = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;
                        SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oform.Items.Item("Item_4").Specific;
                        SAPbouiCOM.ComboBox oComboBelgeTipi = (SAPbouiCOM.ComboBox)oform.Items.Item("Item_40").Specific;
                        SAPbouiCOM.ComboBox oComboSube = (SAPbouiCOM.ComboBox)oform.Items.Item("Item_47").Specific;

                        
                        if (bastar != "" && !bastar.Contains("1900"))
                        {
                            if (DateTime.TryParse(bastar, out var dtBas))
                                ((SAPbouiCOM.EditText)oform.Items.Item("Item_1").Specific).Value = dtBas.ToString("yyyyMMdd");
                        }

                        if (bittar != "" && !bittar.Contains("1900"))
                        {
                            if (DateTime.TryParse(bittar, out var dtBit))
                                ((SAPbouiCOM.EditText)oform.Items.Item("Item_3").Specific).Value = dtBit.ToString("yyyyMMdd");
                        }

                        ((SAPbouiCOM.EditText)oform.Items.Item("Item_101").Specific).Value = kulKod;

                        // isSend her iterasyonda sıfırlanmalı
                        bool isSend = false;

                        for (int h = 1; h <= 2; h++)
                        {
                            if (h == 1)
                            {
                                oComboBelgeTipi.Select("F", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                sendinvoiceekrani = true;
                                oform.Items.Item("Item_5").Click(); // listele butonu
                                sendinvoiceekrani = false;
                            }
                            else if (h == 2)
                            {
                                oComboBelgeTipi.Select("A", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                sendinvoiceekrani = true;
                                oform.Items.Item("Item_5").Click(); // listele butonu
                                sendinvoiceekrani = false;
                            }

                            if (oMatrix.RowCount == 0)
                            {
                                continue;
                            }

                            for (int i = 1; i <= oMatrix.RowCount; i++) // Faturaları gönderir.
                            {
                                string cari = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_0").Cells.Item(i).Specific).Value.ToString();
                                string tabloadi = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_19").Cells.Item(i).Specific).Value;
                                string belgeno = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Col_2").Cells.Item(i).Specific).Value;

                                sql = "Select \"U_AIF_DOC_OrdType\" from [" + tabloadi + "] where \"DocEntry\" = '" + belgeno + "'";
                                oRS_3.DoQuery(sql);

                                if (oComboBelgeTipi.Value.Trim() == "F")
                                {
                                    ((SAPbouiCOM.ComboBox)oMatrix.Columns.Item("Col_3").Cells.Item(i).Specific).Select(fatserisi, SAPbouiCOM.BoSearchKey.psk_ByDescription);
                                }
                                else if (oComboBelgeTipi.Value.Trim() == "A")
                                {
                                    ((SAPbouiCOM.ComboBox)oMatrix.Columns.Item("Col_3").Cells.Item(i).Specific).Select(arsivserisi, SAPbouiCOM.BoSearchKey.psk_ByDescription);
                                }

                                ((SAPbouiCOM.CheckBox)oMatrix.Columns.Item("#").Cells.Item(i).Specific).Caption = "Y"; // seçme işlemi
                                isSend = true;
                            }

                            sendinvoiceekrani = true;
                            if (isSend) // Matrix içerisinde işaretlenmiş bir satır yoksa gönder tuşuna basılmıyor.
                                oform.Items.Item("Item_6").Click(); // fatura gönder butonu.

                            isSend = false;
                            sendinvoiceekrani = false;
                        }

                        // Her kullanıcı işlemi bittikten sonra formu kapat
                        try
                        {
                            oform.Close();
                        }
                        catch (Exception exClose)
                        {
                            // Form kapatma hatası yoksayılır, bir sonraki kullanıcıya geçilir
                        }
                    }
                    catch (Exception exKulKod)
                    {
                        // Tek bir kullanıcı kodu hata verse bile diğerleri çalışmaya devam eder
                        sendinvoiceekrani = false;
                    }
                }

                faturagondermeislemicalisiyor = false;
            }
            catch (Exception ex)
            {
                faturagondermeislemicalisiyor = false;
            }
        }

        private static void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    System.Windows.Forms.Application.Exit();
                    break;

                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                    break;

                case SAPbouiCOM.BoAppEventTypes.aet_FontChanged:
                    break;

                case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                    break;

                case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    break;

                default:
                    break;
            }
        }

        public class MagazaBilgileri
        {
            public string cariKodu { get; set; }
            public string seriNo { get; set; }
            public string tip { get; set; }
        }

        public static SAPbobsCOM.Company oCompanyObject { get; set; }
    }
}
