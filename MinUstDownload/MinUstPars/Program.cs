using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace MinUstPars
{
    class Program
    {
        static int a = -1;
        static string param = "";
        static string curl = @"http://unro.minjust.ru/NKOReports.aspx";
        static Filter obj_f;

        static void Main(string[] args)
        {
            if (args.Length != 0)
            {
                param = args[0];
            }

            if (param.Length > 1)
            {
                if (isFile(param))
                {
                    obj_f = LoadXml(typeof(Filter), param) as Filter;
                    if (obj_f != null)
                        runBrowserThread(curl);
                }
                else
                {
                    Console.WriteLine(@"Неверный параметр запуска. Путь должен быть указан в формате c:\path\file.xml");
                    Console.WriteLine("Нажмите любую клавишу.");
                    Console.ReadKey();
                }
            }
            else
            {
                Console.WriteLine("Проверьте параметры запуска, не указан файл настроек.");
                Console.WriteLine("Нажмите любую клавишу.");
                Console.ReadKey();
            }
        }

        private static void runCheckAndCreatePdf(string url)
        {
            try
            {
                var th = new Thread(() =>
                {
                    var br = new WebBrowser();
                    //br.DocumentCompleted += browser_DocumentCompleted;
                    br.ScriptErrorsSuppressed = true;
                    br.Navigate(url);
                    Application.Run();
                });
                th.SetApartmentState(ApartmentState.STA);
                th.Start();
            }
            catch //(Exception ex)
            {
             //   Console.WriteLine(">>>Ошибка:" + ex.ToString());
            }
        }

        private static void runBrowserThread(string url)
        {
            try
            {
                var th = new Thread(() =>
                {
                    Console.WriteLine(">>>Запуск процесса парсинга");
                    var br = new WebBrowser();
                    br.DocumentCompleted += browser_DocumentCompleted;
                    br.ScriptErrorsSuppressed = true;
                    br.Navigate(url);
                    Application.Run();
                });
                th.SetApartmentState(ApartmentState.STA);
                th.Start();
            }
            catch (Exception ex)
            {
                Console.WriteLine(">>>Ошибка:" + ex.ToString());
            }
            finally
            {
                Console.WriteLine(">>>Парсинг");
            }
        }

        private static bool checkStopWord(string stopWord, string inputStr)
        {
            Regex rg = new Regex(stopWord, RegexOptions.IgnoreCase);
            if (rg.IsMatch(inputStr))
            {
                return true;
            }
            else
            {
                return false;
            }

        }
        public static bool SaveXml(object obj, string filename)
        {
            bool result = false;
            using (StreamWriter writer =
                new StreamWriter(filename, false))
                {
                    try
                    {
                        XmlSerializerNamespaces ns =
                            new XmlSerializerNamespaces();
                        ns.Add("", "");
                        XmlSerializer serializer =
                            new XmlSerializer(obj.GetType());
                        serializer.Serialize(writer, obj, ns);
                        result = true;
                    }
                    catch (Exception e)
                    {
                        Debug.WriteLine(e.Message);
                    }
                    finally
                    {
                        writer.Close();
                    }
                }
            return result;
        }

        public static bool isFile(string path)
        {
            try
            {
                FileAttributes attr = File.GetAttributes(path);
                if (attr.HasFlag(FileAttributes.Directory))
                    return false;
                else 
                {
                    if (File.Exists(path))
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
            }
            catch 
            {
                return false;
            }
        }

        public static object LoadXml(Type type, string filename)
        {
            object result = null;
            using (StreamReader reader = new StreamReader(filename))
            {
                try
                {
                    XmlSerializer serializer = new XmlSerializer(type);
                    result = serializer.Deserialize(reader);
                }
                catch (Exception e)
                {
                    Debug.WriteLine(e.Message);
                }
                finally
                {
                    reader.Close();
                }
            }
            return result;
        }

        private static bool loadFile(string url, string pathToFile)
        {
            try
            {
                WebClient webClient = new WebClient();
                webClient.DownloadDataAsync(new Uri(url), pathToFile);
                return true;
            }
            catch 
            {
                return false;
            }

            
        }

        private static string getStrTypeNKO(int startindex, string text, out string vid, out string ogrn)
        {
            string s = "";
            ogrn = "";
            string patt = "pdg_item_odd\" valign='Top'>";
            int sidx = text.IndexOf(patt, startindex) + patt.Length;
            for (int i = 0; i < 3; i++)
            {
                int edx = text.IndexOf(patt, sidx) + patt.Length;
                sidx = edx + 1;

                if (i == 0)
                {
                    try
                    {
                        int ln = text.IndexOf("</td>", sidx);
                        ogrn = text.Substring(sidx - 1, ln - sidx + 1);
                    }
                    catch
                    {
                        ogrn = "err";
                    }
                }

                if (i == 1)
                {
                    try
                    {
                        int ln = text.IndexOf("</td>", sidx);
                        s = text.Substring(sidx - 1, ln - sidx + 1);
                    }
                    catch
                    {
                        s = "err";
                    }
                }
            }

            try
            {
                int ln = text.IndexOf("</td>", sidx);
                vid = text.Substring(sidx - 1, ln - sidx + 1);
            }
            catch
            {
                vid = "err";
            }

            
            return s;
        }

        private static int checkInIndex(string indexReport, List<ReportFile> lrf)
        {
            try
            {
               int value = int.Parse(lrf.First(item => item.indexFile == indexReport).indexFile);
               return value;
            }
            catch
            {
                return -1;
            }
        }

        private static void addToLog(string line)
        {
            StreamWriter sw = new StreamWriter("log.txt", true);
            sw.WriteLine(line);
            sw.Close();
        }

        private static void browser_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            var webBrowser1 = sender as WebBrowser;
         
            if (webBrowser1.ReadyState != WebBrowserReadyState.Complete)
            {

                a++;
                Debug.WriteLine("complete" + a);
                Debug.WriteLine(e.Url.AbsolutePath + (sender as WebBrowser).Url.AbsolutePath);
            }

            if (a == 0)
            {

                string typeRep = (obj_f.typeReport == "*") ? "" : "'; document.getElementById('filter_nko_form').value = '" + obj_f.typeReport;
                HtmlDocument doc = webBrowser1.Document;
                HtmlElement head = doc.GetElementsByTagName("head")[0];
                HtmlElement s = doc.CreateElement("script");
                s.SetAttribute("text", "function setFilter() {  document.getElementById('filter_region').value = '" + obj_f.region + 
                    "'; document.getElementById('filter_dt_from').value = '" + obj_f.year + 
                    "'; document.getElementById('filter_dt_to').value = '" + obj_f.year +
                    typeRep + 
                    "';}");
                head.AppendChild(s);
                webBrowser1.Document.InvokeScript("setFilter");

                // webBrowser1.Navigate("javascript: alarm()");

                //a++;
            }
            else if (a == 1)
            {
                //webBrowser1.Navigate("javascript: Refresh()");
                webBrowser1.Document.InvokeScript("Refresh");
                // a++;
            }
            else if (a == 2)
            {
                Debug.WriteLine(a + "set page");
                string text = webBrowser1.DocumentText;
                int st = text.IndexOf("pdg_pos") + 7;
                st = text.IndexOf("<b>", st) + 19;

                string records = text.Substring(st, text.IndexOf("</b>", st) - st);

                records = records.Replace("&nbsp;", "");
                //Debug.WriteLine("javascript: __doPostBack('pdg','page_size:" + records + "')");
                webBrowser1.Navigate("javascript: __doPostBack('pdg','page_size:" + "65535" + "')");
                //a++;
            }
            else if(a == 5)
            {
                addToLog("[" + DateTime.Now.ToString() + "] Начало загрузки отчетов");
                Encoding enc = Encoding.GetEncoding("windows-1251");

                Stream stream = webBrowser1.DocumentStream;
                StreamReader sr = new StreamReader(stream, enc);
                string text1 = sr.ReadToEnd();
                stream.Close();

                //string text1 = webBrowser1.DocumentText;
                text1 = text1.Replace("pdg_item_left_even", "pdg_item_name_nko");
                text1 = text1.Replace("pdg_item_left_odd", "pdg_item_name_nko");

                text1 = text1.Replace("pdg_item_even", "pdg_item_odd");

                Regex patt_name = new Regex(@"pdg_item_name_nko\"" valign=\'Top\'\>(?<val>.*)\</td>");
                Regex pattern = new Regex(@"ShowPdfReport\('(?<val>.*)', event\)");

                Index idx = new Index();
                //idx.valuePdf = obj_f.pathPdf;

                List<ReportFile> lrf = new List<ReportFile>();

                #region  /////////////////// Загрузка индекса
                Index ixf = new Index();

                try
                {
                    ixf = (Index)LoadXml(typeof(Index), obj_f.pathIndex);
                    lrf = ixf.indexList;
                }
                catch { }
                ///////////////////
                #endregion

                //Console.WriteLine(text1);

                foreach (Match m in patt_name.Matches(text1))
                    if (m.Success)
                    {
                        Application.DoEvents();
                        string nameNKO = m.Groups["val"].Value;

                        string vid = "";
                        string ogrn = "";
                        string typeNKO = getStrTypeNKO(m.Index + m.Length, text1, out vid, out ogrn);

                        //Console.WriteLine("Вид:" + ogrn);

                        int ln = text1.IndexOf("pdg_info", m.Index);
                        string nt = text1.Substring(m.Index, ln - m.Index);
                        
                        string idexReport = pattern.Match(nt).Groups["val"].Value;

                        #region//////////фильтрация///////////
                        bool next = true;
                        foreach (string regx in obj_f.stopWords)
                        {
                            next = checkStopWord(regx, nameNKO);
                            if (next == true)
                                break;
                            else
                                next = false;
                        }

                        if (obj_f.typeReport == "*" || vid == obj_f.typeReport)
                        {
                            next = false;
                        }
                        else
                        {
                            next = true;
                        }
                        ////////////////////////////////////
                        #endregion

                        if (next == false)
                        {
                            ReportFile rf = new ReportFile();
                            rf.formNKO = typeNKO;
                            rf.nameNKO = nameNKO;
                            rf.indexFile = idexReport;
                            rf.urlPdf = @"http://unro.minjust.ru/Reports/" + idexReport + ".pdf";
                            rf.ogrn = ogrn;
                            rf.vid = vid;
                            rf.period = obj_f.year;


                           // string urlCreatePdf = @"http://unro.minjust.ru/NKOReports.aspx?mode=show_pdf_report&id=" + idexReport;

                            if (checkInIndex(idexReport, lrf) < 0)
                            {
                                //lrf.Add(rf);
                                Console.Write("|");
                                try
                                {
                                    //runCheckAndCreatePdf(urlCreatePdf);
                                    webBrowser1.Document.InvokeScript("ShowPdfReport('" + idexReport + "', event)");
                                    //МинЮст не хранит пдф файлы, а генерирует их. Поэтому урл по умолчанию не будет доступен
                                    Thread.Sleep(1000); //обязательная пауза, на сайте минюст 3 секунды, у нас 1

                                    using (var cli = new WebClient())
                                    {
                                        //cli.Headers.Add(HttpRequestHeader.Accept, @"application/x-ms-application, image/jpeg, application/xaml+xml, image/gif, image/pjpeg, application/x-ms-xbap, application/x-shockwave-flash, application/vnd.ms-excel, application/vnd.ms-powerpoint, application/msword, */*");
                                        cli.Headers.Add(HttpRequestHeader.UserAgent, "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/47.0.2526.106 Safari/537.36");
                                        
                                        string directory = obj_f.pathPdf + obj_f.year;
                                        if (!Directory.Exists(directory))
                                        {
                                            Directory.CreateDirectory(directory);
                                        }

                                        cli.DownloadFile(new Uri(rf.urlPdf), directory + @"\" + idexReport + ".pdf");
                                        rf.pathPdf = directory + @"\" + idexReport + ".pdf";
                                        lrf.Add(rf);
                                        addToLog("[" + DateTime.Now.ToString() + "] Загружен отчет №" + idexReport.ToString() + " [" + rf.urlPdf + "]");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    addToLog("["+DateTime.Now.ToString() + "] Ошибка загрузки:" + ex.Message + " [" + rf.urlPdf + "]");
                                }
                            }
                           
                        }
                    }

                idx.indexList = lrf;
                
                SaveXml(idx, obj_f.pathIndex);

                int cnt = (ixf != null) ? ixf.indexList.Count - lrf.Count : lrf.Count;
                addToLog("[" + DateTime.Now.ToString() + "] Сохранен и сформирован индекс. Новых записей: " + cnt.ToString());
                
                a++;
                Console.WriteLine("");
                Console.WriteLine("Процесс завершен. Завершение работы приложения...");
                Thread.Sleep(3000); 
                Application.ExitThread();
                Application.Exit();
        }
        }
    }
}
