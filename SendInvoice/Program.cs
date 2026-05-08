using SAPbobsCOM;
using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;

namespace SendInvoice
{
    internal class Program
    {
        // ============ DonusumOne.SendInvoice formundaki kontrol ID'leri ============
        private const string ITEM_BAS_TARIH = "Item_1";
        private const string ITEM_BIT_TARIH = "Item_3";
        private const string ITEM_MATRIX = "Item_4";
        private const string ITEM_LISTELE_BTN = "Item_5";
        private const string ITEM_GONDER_BTN = "Item_6";
        private const string ITEM_BELGE_TIPI_CB = "Item_40";
        private const string ITEM_KULKOD = "Item_101";

        private const string COL_SERI = "Col_3";
        private const string COL_CHECK = "#";

        private const string DOC_FATURA = "F";
        private const string DOC_ARSIV = "A";

        private const string MENU_SEND_INVOICE = "DonusumOne.SendInvoice";

        // ============ Zamanlama ============
        private const int DEFAULT_INTERVAL_MS = 3 * 60 * 1000; // 3 dk
        private const int MIN_INTERVAL_SEC = 30;            // alt sinir
        private const int FORM_READY_TIMEOUT_MS = 5000;
        private const int FORM_READY_POLL_MS = 100;

        // ============ State ============
        public static SAPbobsCOM.Company oCompanyObject { get; set; }
        public static bool sendinvoiceekrani = false;

        private static System.Windows.Forms.Form _uiSync;
        private static System.Timers.Timer _timer;
        private static int _busy; // 0=bos, 1=calisiyor (Interlocked ile)

        // ============ Logger ============
        private static readonly object _logLock = new object();
        private static string _logPath;

        public static void Log(string msg)
        {
            try
            {
                lock (_logLock)
                {
                    if (_logPath == null)
                    {
                        _logPath = Path.Combine(
                            AppDomain.CurrentDomain.BaseDirectory,
                            "SendInvoice.log");
                    }
                    File.AppendAllText(_logPath,
                        string.Format("{0:yyyy-MM-dd HH:mm:ss.fff}  [TID:{1}]  {2}{3}",
                            DateTime.Now,
                            Thread.CurrentThread.ManagedThreadId,
                            msg,
                            Environment.NewLine));
                }
            }
            catch
            {
                // Loglama bile patlasa akisi durdurma
            }
        }

        // ============ Main ============
        [STAThread]
        private static void Main(string[] args)
        {
            try
            {
                Log("=== SendInvoice baslatiliyor ===");

                Application oApp = args.Length < 1
                    ? new Application()
                    : new Application(args[0]);

                Menu MyMenu = new Menu();
                MyMenu.AddMenuItems();
                oApp.RegisterMenuEventHandler(MyMenu.SBO_Application_MenuEvent);
                Application.SBO_Application.AppEvent += SBO_Application_AppEvent;

                oCompanyObject = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();
                Log("DI Company baglantisi: " + (oCompanyObject != null && oCompanyObject.Connected));

                oApp.AfterInitialized += oApp_AfterInitialized;
                oApp.Run();
            }
            catch (Exception ex)
            {
                Log("Main exception: " + ex);
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        private static void oApp_AfterInitialized(object sender, EventArgs e)
        {
            try
            {
                // KRITIK: System.Timers.Timer ThreadPool'da calisir.
                // SAP UI API cagrilari STA/UI thread'inde olmak zorunda.
                // Bu yuzden timer'in gercek isini hidden bir WinForms form
                // uzerinden BeginInvoke ile UI thread'ine marshal ediyoruz.
                _uiSync = new System.Windows.Forms.Form
                {
                    ShowInTaskbar = false,
                    FormBorderStyle = System.Windows.Forms.FormBorderStyle.None,
                    WindowState = System.Windows.Forms.FormWindowState.Minimized,
                    Opacity = 0,
                    Visible = false
                };
                var _ = _uiSync.Handle; // handle'i bu thread'de zorla yarat
                Log("UI sync form hazir, UI TID=" + Thread.CurrentThread.ManagedThreadId);

                int intervalMs = DEFAULT_INTERVAL_MS;//ReadIntervalFromParams();
                Log("Timer interval (ms): " + intervalMs);

                _timer = new System.Timers.Timer(intervalMs);
                _timer.AutoReset = false; // re-entrancy'yi bastan engelle
                _timer.Elapsed += OnTimerElapsed;
                _timer.Start();
            }
            catch (Exception ex)
            {
                Log("AfterInitialized error: " + ex);
            }
        }

        private static int ReadIntervalFromParams()
        {
            // U_Interval kolonu varsa oku, yoksa default kullan.
            // Boylece projeyi 'Parametreli' isminin hakkini verecek sekilde
            // dakika ayari @DON_OTOPARAM tablosundan kontrol edilebilir.
            SAPbobsCOM.Recordset oRS = null;
            try
            {
                if (oCompanyObject == null || !oCompanyObject.Connected)
                    return DEFAULT_INTERVAL_MS;

                oRS = (SAPbobsCOM.Recordset)oCompanyObject.GetBusinessObject(BoObjectTypes.BoRecordset);
                oRS.DoQuery("Select TOP 1 \"U_Interval\" from \"@DON_OTOPARAM\"");
                if (oRS.RecordCount > 0)
                {
                    object v = oRS.Fields.Item(0).Value;
                    if (v != null)
                    {
                        int sec;
                        if (int.TryParse(v.ToString(), out sec) && sec >= MIN_INTERVAL_SEC)
                            return sec * 1000;
                    }
                }
            }
            catch (Exception ex)
            {
                // Kolon yoksa sorgu hata verir; default'a dus.
                Log("Interval okunamadi (default kullanilacak): " + ex.Message);
            }
            finally
            {
                ReleaseComObject(oRS);
            }
            return DEFAULT_INTERVAL_MS;
        }

        // ============ Timer ============
        private static void OnTimerElapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            if (Interlocked.Exchange(ref _busy, 1) == 1)
            {
                Log("Onceki tick hala calisiyor, atliyorum");
                return;
            }

            try
            {
                if (_uiSync != null && _uiSync.IsHandleCreated && !_uiSync.IsDisposed)
                {
                    // Asil isi UI thread'inde calistir
                    _uiSync.BeginInvoke(new Action(RunJobOnUiThread));
                }
                else
                {
                    Log("UI sync form hazir degil, tick atlandi");
                    Interlocked.Exchange(ref _busy, 0);
                    SafeRestartTimer();
                }
            }
            catch (Exception ex)
            {
                Log("OnTimerElapsed marshal exception: " + ex);
                Interlocked.Exchange(ref _busy, 0);
                SafeRestartTimer();
            }
        }

        private static void RunJobOnUiThread()
        {
            try
            {
                DoInvoiceJob();
            }
            catch (Exception ex)
            {
                Log("RunJobOnUiThread exception: " + ex);
            }
            finally
            {
                Interlocked.Exchange(ref _busy, 0);
                SafeRestartTimer();
            }
        }

        private static void SafeRestartTimer()
        {
            try
            {
                if (_timer != null) _timer.Start();
            }
            catch (Exception ex)
            {
                Log("Timer restart hatasi: " + ex.Message);
            }
        }

        // ============ Asil Is ============
        private static void DoInvoiceJob()
        {
            Log("--- Tick basladi (UI TID=" + Thread.CurrentThread.ManagedThreadId + ") ---");

            if (oCompanyObject == null || !oCompanyObject.Connected)
            {
                Log("DI Company bagli degil, atliyorum");
                return;
            }

            SAPbobsCOM.Recordset oRSParam = null;

            try
            {
                oRSParam = (SAPbobsCOM.Recordset)oCompanyObject.GetBusinessObject(BoObjectTypes.BoRecordset);

                string sql = "Select ISNULL(\"U_BasTar\",'1900-01-01') as BasTar," +
                             "ISNULL(\"U_BitTar\",'1900-01-01') as BitTar,* from \"@DON_OTOPARAM\"";
                oRSParam.DoQuery(sql);

                if (oRSParam.RecordCount == 0)
                {
                    Log("@DON_OTOPARAM bos, atliyorum");
                    return;
                }

                string kulkodu = SafeFieldString(oRSParam, "U_KulKod");
                if (string.IsNullOrWhiteSpace(kulkodu))
                {
                    Log("U_KulKod bos, atliyorum");
                    return;
                }

                string fatserisi = SafeFieldString(oRSParam, "U_FatSeri");
                string arsivserisi = SafeFieldString(oRSParam, "U_ArsivSeri");
                string bastar = SafeFieldString(oRSParam, "BasTar");
                string bittar = SafeFieldString(oRSParam, "BitTar");

                List<string> kulKodlari = kulkodu
                    .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(x => x.Trim())
                    .Where(x => !string.IsNullOrEmpty(x))
                    .ToList();

                Log("Toplam " + kulKodlari.Count + " kulkod islenecek: " + string.Join(",", kulKodlari));

                foreach (string kulKod in kulKodlari)
                {
                    try
                    {
                        ProcessKulKod(kulKod, fatserisi, arsivserisi, bastar, bittar);
                    }
                    catch (Exception exKulKod)
                    {
                        Log("KulKod=" + kulKod + " HATA: " + exKulKod);
                        sendinvoiceekrani = false;
                    }
                }

                Log("--- Tick bitti ---");
            }
            catch (Exception ex)
            {
                Log("DoInvoiceJob outer exception: " + ex);
                sendinvoiceekrani = false;
            }
            finally
            {
                ReleaseComObject(oRSParam);
            }
        }

        private static string SafeFieldString(SAPbobsCOM.Recordset rs, string fieldName)
        {
            try
            {
                object v = rs.Fields.Item(fieldName).Value;
                return v == null ? "" : v.ToString();
            }
            catch
            {
                return "";
            }
        }

        private static void ProcessKulKod(string kulKod, string fatserisi, string arsivserisi,
                                          string bastar, string bittar)
        {
            Log("KulKod=" + kulKod + " basliyor");

            SAPbouiCOM.Form oform = null;
            SAPbouiCOM.Matrix oMatrix = null;
            SAPbouiCOM.ComboBox oComboBelgeTipi = null;

            try
            {
                Application.SBO_Application.ActivateMenuItem(MENU_SEND_INVOICE);
                Log("KulKod=" + kulKod + " menu aktive edildi");

                oform = WaitForActiveForm();
                if (oform == null)
                {
                    Log("KulKod=" + kulKod + " form acilmadi (timeout), atliyorum");
                    return;
                }

                oMatrix = (SAPbouiCOM.Matrix)oform.Items.Item(ITEM_MATRIX).Specific;
                oComboBelgeTipi = (SAPbouiCOM.ComboBox)oform.Items.Item(ITEM_BELGE_TIPI_CB).Specific;

                SetDateField(oform, ITEM_BAS_TARIH, bastar);
                SetDateField(oform, ITEM_BIT_TARIH, bittar);
                SetEditText(oform, ITEM_KULKOD, kulKod);

                ProcessBelgeTipi(oform, oMatrix, oComboBelgeTipi, DOC_FATURA, fatserisi, kulKod);
                ProcessBelgeTipi(oform, oMatrix, oComboBelgeTipi, DOC_ARSIV, arsivserisi, kulKod);

                try
                {
                    oform.Close();
                }
                catch (Exception exClose)
                {
                    Log("KulKod=" + kulKod + " form close hatasi: " + exClose.Message);
                }
            }
            finally
            {
                ReleaseComObject(oComboBelgeTipi);
                ReleaseComObject(oMatrix);
                ReleaseComObject(oform);
            }
        }

        private static void SetDateField(SAPbouiCOM.Form oform, string itemId, string raw)
        {
            DateTime dt;
            if (!TryParseDate(raw, out dt)) return;

            SAPbouiCOM.EditText et = null;
            try
            {
                et = (SAPbouiCOM.EditText)oform.Items.Item(itemId).Specific;
                et.Value = dt.ToString("yyyyMMdd");
            }
            finally
            {
                ReleaseComObject(et);
            }
        }

        private static void SetEditText(SAPbouiCOM.Form oform, string itemId, string value)
        {
            SAPbouiCOM.EditText et = null;
            try
            {
                et = (SAPbouiCOM.EditText)oform.Items.Item(itemId).Specific;
                et.Value = value ?? "";
            }
            finally
            {
                ReleaseComObject(et);
            }
        }

        private static void ProcessBelgeTipi(SAPbouiCOM.Form oform, SAPbouiCOM.Matrix oMatrix,
                                             SAPbouiCOM.ComboBox oComboBelgeTipi,
                                             string belgeTipi, string seri, string kulKod)
        {
            Log("KulKod=" + kulKod + " BelgeTipi=" + belgeTipi + " basliyor");

            oComboBelgeTipi.Select(belgeTipi, SAPbouiCOM.BoSearchKey.psk_ByValue);

            sendinvoiceekrani = true;
            try
            {
                Log("KulKod=" + kulKod + " BelgeTipi=" + belgeTipi + " Listele ONCE");
                oform.Items.Item(ITEM_LISTELE_BTN).Click();
                Log("KulKod=" + kulKod + " BelgeTipi=" + belgeTipi + " Listele SONRA");
            }
            finally
            {
                sendinvoiceekrani = false;
            }

            int rowCount = oMatrix.RowCount;
            Log("KulKod=" + kulKod + " BelgeTipi=" + belgeTipi + " " + rowCount + " satir bulundu");
            if (rowCount == 0) return;

            for (int i = 1; i <= rowCount; i++)
            {
                SAPbouiCOM.ComboBox cbSeri = null;
                SAPbouiCOM.CheckBox ckSec = null;
                try
                {
                    cbSeri = (SAPbouiCOM.ComboBox)oMatrix.Columns.Item(COL_SERI).Cells.Item(i).Specific;
                    cbSeri.Select(seri, SAPbouiCOM.BoSearchKey.psk_ByDescription);

                    ckSec = (SAPbouiCOM.CheckBox)oMatrix.Columns.Item(COL_CHECK).Cells.Item(i).Specific;
                    ckSec.Caption = "Y";
                }
                catch (Exception exRow)
                {
                    Log("KulKod=" + kulKod + " BelgeTipi=" + belgeTipi +
                        " row=" + i + " HATA: " + exRow.Message);
                }
                finally
                {
                    ReleaseComObject(cbSeri);
                    ReleaseComObject(ckSec);
                }
            }

            sendinvoiceekrani = true;
            try
            {
                Log("KulKod=" + kulKod + " BelgeTipi=" + belgeTipi + " Gonder ONCE");
                oform.Items.Item(ITEM_GONDER_BTN).Click();
                Log("KulKod=" + kulKod + " BelgeTipi=" + belgeTipi + " Gonder SONRA");
            }
            finally
            {
                sendinvoiceekrani = false;
            }
        }

        private static SAPbouiCOM.Form WaitForActiveForm()
        {
            int waited = 0;
            while (waited < FORM_READY_TIMEOUT_MS)
            {
                try
                {
                    SAPbouiCOM.Form f = Application.SBO_Application.Forms.ActiveForm;
                    if (f != null)
                    {
                        // Beklenen item'lar yuklenmis mi? (Item_4 = matrix)
                        var probe = f.Items.Item(ITEM_MATRIX);
                        if (probe != null) return f;
                    }
                }
                catch
                {
                    // Form henuz hazir degil
                }
                Thread.Sleep(FORM_READY_POLL_MS);
                waited += FORM_READY_POLL_MS;
            }
            return null;
        }

        private static bool TryParseDate(string s, out DateTime dt)
        {
            dt = DateTime.MinValue;
            if (string.IsNullOrWhiteSpace(s)) return false;
            if (!DateTime.TryParse(s, out dt)) return false;
            return dt.Year > 1900;
        }

        // ============ COM Cleanup ============
        public static void ReleaseComObject(object obj)
        {
            try
            {
                if (obj != null && System.Runtime.InteropServices.Marshal.IsComObject(obj))
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
            }
            catch { }
        }

        // ============ App Event ============
        private static void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    Log("aet_ShutDown alindi, kapatiliyor");
                    try
                    {
                        Interlocked.Exchange(ref _busy, 1); // kimse yeni is baslatamasin
                        if (_timer != null)
                        {
                            _timer.Stop();
                            _timer.Dispose();
                        }
                    }
                    catch { }
                    try
                    {
                        if (_uiSync != null)
                        {
                            _uiSync.Close();
                            _uiSync.Dispose();
                        }
                    }
                    catch { }
                    Environment.Exit(0);
                    break;

                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                case SAPbouiCOM.BoAppEventTypes.aet_FontChanged:
                case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    break;

                default:
                    break;
            }
        }
    }
}
