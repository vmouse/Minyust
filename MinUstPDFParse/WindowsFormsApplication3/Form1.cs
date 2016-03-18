using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PdfToText;
using System.Text.RegularExpressions;
using System.Diagnostics;
using System.IO;
using System.Xml.Serialization;

namespace MinUstPdf
{
    public partial class Form1 : Form
    {
        public List<string> files = new List<string>();
        public string csv_path = "";
        
        public string _xml_index_path = "";
        public string _min_summ = "";
        public string _viewReport = "OH0002";

        public Index indexFile = new Index();
        public List<ReportFile> listReports = new List<ReportFile>();

        private bool autostart = false;

        public Form1()
        {
            string[] args = Environment.GetCommandLineArgs();
            InitializeComponent();
            if (args.Length > 1)
            {
                _xml_index_path = args[1];
                csv_path = args[2];
                _min_summ = args[3];
                _viewReport = (args.Length>5)? args[4]: "OH0002";
                autostart = true;
                //Text = _xml_index_path + " " + _min_summ + " " + _viewReport;
            }
        }

        public string GetWithIn(string str)
        {
            string rez = "";

            Regex pattern =
                new Regex(@"\((?<val>.*?)\)",
                    RegexOptions.Compiled |
                    RegexOptions.Singleline);

            foreach (Match m in pattern.Matches(str))
                if (m.Success)
                    //меж скобок ( )  
                    //rez.Add(m.Groups["val"].Value);
                    rez += m.Groups["val"].Value;

            return rez;
        }

        public double getSumm(string str)
        {
            double rez = 0;
            str = str.Replace("Tm_/F1", "$");

            string[] rt = str.Split('$');

            str = "";

            foreach (string t in rt)
            {
                Regex pat =
                new Regex(@"\((?<val>.*?)\)",
                    RegexOptions.Compiled |
                    RegexOptions.Singleline);
                str += "(";
                foreach (Match m in pat.Matches(t))
                    if (m.Success)
                    {
                        str += m.Groups["val"].Value;
                    }
                str += ")Tj" + '\n';

                
            }
            Debug.WriteLine(str);

            Regex pattern =
                new Regex(@"\((?<val>.*?)\)",
                    RegexOptions.Compiled |
                    RegexOptions.Singleline);

            foreach (Match m in pattern.Matches(str))
                if (m.Success)
                //меж скобок ( )  
                //rez.Add(m.Groups["val"].Value);
                {
                    double d = isDouble(m.Groups["val"].Value);
                    Debug.WriteLine(d + m.Groups["val"].Value);
                    if (d > 1)
                    {
                        
                        rez += d;//strToDouble(m.Groups["val"].Value);
                    }
                }

            return rez;
        }

        private ReportFile findInIndex(string pathPdf, List<ReportFile> lrf)
        {
            try
            {
                ReportFile value = lrf.First(item => item.pathPdf == pathPdf);
                return value;
            }
            catch
            {
                return null;
            }
        }

        private double strToDouble(string arr)
        {
            //Debug.WriteLine(arr);
            double d = 0;
            try
            {
                d = Double.Parse(arr.Replace(" ", "").Replace('.', ','));
            }
            catch
            {
            }
            return d;
        }

        public double isDouble(string s)
        {
            Regex pat = new Regex(@"\d+(?:[.,]\d+)?");
            double ret = 0;

            if (pat.IsMatch(s))
            {
                try
                {
                    ret = Double.Parse(Regex.Match(s, @"\d+(?:\.\d+)?").ToString());
                }
                catch
                {
                    ret = 0;
                }

            }

            return ret;
            /*
            s = System.Text.RegularExpressions.Regex.Match(s, @"\d+(?:\.\d+)?").ToString(); //выбираем только цифры

            Debug.WriteLine(s);

            /*Regex pattern =
                new Regex(@"^([0-9]{1,99})([.,][0-9]{1,3})?$",
                    RegexOptions.Compiled |
                    RegexOptions.Singleline);

            
            return pattern.IsMatch(s);*/
         
        }

        public string clearRegX(string s)
        {
            string tr = s;
            while (tr.IndexOf("(Страница:)Tj") > 0)
            {
                int six = s.IndexOf("(Страница:)Tj");
                int sev = s.IndexOf("(Форма:)Tj", six + 1);
                int sed = s.IndexOf("Q", sev + 1);

                string for_del = s.Substring(six, sed - six);

                tr = tr.Replace(for_del, "");
            }

            return tr;
        }

        private string getDolz(string s)
        {
            string ret = "";
            int i = s.IndexOf(',') + 1;
            if (i >= 1)
            {
                ret = s.Substring(i, s.Length - i);
            }
            else
            {
                if (getWord(s) > 1)
                {
                    ret = s.Substring(getWord(s), s.Length - getWord(s));
                }
            }

            return ret;
        }

        private int getWord(string s)
        {
            int r = 0;
            
            int a = s.ToLower().IndexOf("генер");
            r = a;
            if (r < 1)
            {
                int e = s.ToLower().IndexOf("испол");
                r = e;

                if (r < 1)
                {
                    int c = s.ToLower().IndexOf("главн");
                    r = c;

                    if (r < 1)
                    {
                        int b = s.ToLower().IndexOf("дирек");
                        r = b;

                        if (r < 1)
                        {
                            int d = s.ToLower().IndexOf("бухга");
                            r = d;
                        }

                    }
                }
            }

            return r;
        }

        private void addToFile(List<string> s)
        {
            TextWriter sw = new StreamWriter(csv_path, false, Encoding.UTF8);
            foreach (string st in s)
            {
                sw.WriteLine(st);
            }
            sw.Close();
            pb.Value = 0;

            if (!autostart)
            {
                MessageBox.Show("Парсинг завершен!");
            }else
                Application.Exit();

            
         
        }

        private string getDir(string s)
        {
            string ret = "";
            int i = s.IndexOf(',');
            if (i >= 1)
            {
                ret = s.Substring(0, i);
            }
            else
            {
                if (getWord(s) > 1)
                {
                    ret = s.Substring(0, getWord(s));
                }
            }

            return ret;
        }

        private string FormatCSV(string param, char type, bool end = false) // t - text, d - date, n - number
        {
            
            switch (type) {
                case 'x': param = "\"" + param.Replace("\"", "\"\"") + "\"";
                    break;
            }
            return param + ((end)?"":";");
        }

        private void startParse()
        {
            if (pathCSV.Text.Length > 1 && pathPDF.Text.Length > 1 && files.Count > 0)
            {

                pb.Value = 0;
                pb.Maximum = files.Count;
                pb.Step = 1;
                ReportFile rf = null;
//                richTextBox1.Clear();

                //csv:  form; year; org; adres; inn; kpp; direktor + dolzh; buh; date; summ1; summ2;

                List<string> fcsv = new List<string>();
                if (listReports.Count < 0)
                {
                    fcsv.Add("Вид Отчета; Год отчета; Организация; Адрес; Moscow; ОГРН; Дата ЕГРЮЛ; ИНН; КПП; Сумма раздела 1; Сумма раздела 2; Директор; Дата отчета; Бухгалтер; Имя файла; Дата создания файла");
                }
                else
                {
                    fcsv.Add("Вид Отчета; Год отчета; Организация; Адрес; Moscow; ОГРН; Дата ЕГРЮЛ; ИНН; КПП; Сумма раздела 1; Сумма раздела 2; Директор; Дата отчета; Бухгалтер; Имя файла; Дата создания файла; НаименованиеНКО; Учетный номер; ОГРН; Форма; Вид отчета; Период;");
                }

                bool _add = false;

                PDFParser pf = new PDFParser();

                foreach (string filename in files)
                {
                    if (listReports.Count > 0)
                    {
                        rf = findInIndex(filename, listReports);
                    }

                    string csv_line = "";
                    _add = false;
                    pb.Value += 1;
                    try
                    {
                        int tot = 0;
                        string s = pf.ExtractText(filename, out tot);
                        

                        s = s.Replace('\n', '_');

                        Debug.WriteLine(">>>>>>>>>>>>" + tot);

                        int ifrm = s.IndexOf("Форма:)Tj") + 9;
                        int efrm = s.IndexOf("Q", ifrm);

                        string frm = s.Substring(ifrm, efrm - ifrm);
                        frm = GetWithIn(frm);

                        csv_line += FormatCSV(frm, 't');

//                        richTextBox1.AppendText("Форма документа: " + frm + '\n');

                        if (frm.Equals(viewReport.Text))
                        {
                            _add = true;
                            int za = s.IndexOf("за ");

                            string year = s.Substring(za + 3, 4);
//                            richTextBox1.AppendText("Год документа: " + year + '\n');

                            csv_line += FormatCSV(year, 'n');


                            int eid = s.IndexOf(@"(\(полное", za) - 9;

                            string naminc = s.Substring(za, eid - za);

//                            richTextBox1.AppendText("Организация: " + GetWithIn(naminc) + '\n');
                            csv_line += FormatCSV(GetWithIn(naminc),'t');

                            int adid = s.IndexOf(@"организации\))Tj", eid) + 16;
                            int eadid = s.IndexOf(@"(\(адрес", adid) - 8;

                            string addr = s.Substring(adid, eadid - adid);

//                            richTextBox1.AppendText("Адрес: " + GetWithIn(addr) + '\n');

                            addr = GetWithIn(addr);
                            csv_line += FormatCSV(addr, 't');

                            if (addr.IndexOf("осква") > 0)
                            {
                                csv_line += FormatCSV("1",'n');
                            }
                            else
                            {
                                csv_line += FormatCSV("0",'n');
                            }

                            int oid = s.IndexOf("(ОГРН:)Tj", eadid) + 8;
                            int eoid = s.IndexOf("F1 11 Tf", oid);

                            string textOgrn = s.Substring(oid, eoid - oid);

                            string ogr = GetWithIn(textOgrn);

                            csv_line += FormatCSV(ogr, 't');

//                            richTextBox1.AppendText("ОГРН: " + ogr + '\n');

                            int i_egrl = s.IndexOf("(ЕГРЮЛ)Tj", eoid) + 9;
                            int e_egrl = s.IndexOf("Q_q", i_egrl);

                            string s_d_egrul = s.Substring(i_egrl, e_egrl - i_egrl);

//                            richTextBox1.AppendText("ДАТА ЕГРЮЛ: " + GetWithIn(s_d_egrul) + '\n');

                            csv_line += FormatCSV(GetWithIn(s_d_egrul),'d');

                            int inn = s.IndexOf("(ИНН/КПП:)Tj", eoid) + 12;
                            int einn = s.IndexOf("F1 11 Tf", inn);

                            string sinn = s.Substring(inn, einn - inn);
//                            richTextBox1.AppendText("ИНН: " + GetWithIn(sinn) + '\n');

                            csv_line += FormatCSV(GetWithIn(sinn), 't');

                            int ekpp = s.IndexOf("F1 11 Tf", einn + 1);
                            string kpp = s.Substring(einn, ekpp - einn).Replace("/", "");

//                            richTextBox1.AppendText("КПП: " + GetWithIn(kpp) + '\n');

                            csv_line += FormatCSV(GetWithIn(kpp),'t');

                            double sumall = 0;

                            //1.1
                            int sone = s.IndexOf("(бюджета, бюджетов субъектов Российской Федерации, бюджетов)Tj");
                            int stwo = s.IndexOf("Q_q_2 J_0 G_Q_q_Q_q", sone);

                            string rt = s.Substring(sone, stwo - sone);
                            rt = clearRegX(rt);
//                            richTextBox1.AppendText("Сумма 1.1: " + getSumm(rt).ToString() + '\n');

                            sumall += getSumm(rt);

                            //1.2
                            stwo = s.IndexOf("расходования целевых денежных средств, полученных от российских)Tj", stwo);
                            int stri = s.IndexOf("Q_q_2 J_0 G_Q_q_Q_q", stwo);
                            string wt = s.Substring(stwo + 7, stri - stwo - 7);
                            wt = clearRegX(wt);
//                            richTextBox1.AppendText("Сумма 1.2: " + getSumm(wt).ToString() + '\n');

                            sumall += getSumm(wt);
                            //1.3
                            stri = s.IndexOf("(международных и иностранных организаций, иностранных граждан и лиц без)Tj", stri);
                            int otwo = s.IndexOf("Q_q_2 J_0 G_Q_q_Q_q", stri);
                            string tt = s.Substring(stri + 7, otwo - stri - 7);
                            tt = clearRegX(tt);
//                            richTextBox1.AppendText("Сумма 1.3: " + getSumm(tt).ToString() + '\n');

                            sumall += getSumm(tt);

//                            richTextBox1.AppendText("Общая сумма по разделу 1: " + sumall.ToString() + '\n');

                            csv_line += FormatCSV(sumall.ToString(), 'n');

                            //////////Раздел 2////////////

                            int i2s = s.IndexOf("(продажи товаров, выполнения работ, оказания услуг)Tj", otwo);
                            int ei2s = s.IndexOf("Q_q_2 J_0 G_Q_q_Q_q", i2s);
                            string s2s = s.Substring(i2s, ei2s - i2s);
                            s2s = clearRegX(s2s);

//                            richTextBox1.AppendText("Общая сумма по разделу 2: " + getSumm(s2s).ToString() + '\n');

                            sumall += getSumm(s2s);

                            int _min = int.Parse(minSumm.Text);
                            if (sumall < _min && _min > 0)
                            {
                                _add = false;
                            }

                            csv_line += FormatCSV(getSumm(s2s).ToString(), 'n');

                            //ФИО Руководителей

                            int ifio = s.IndexOf("(Лицо, имеющее право без доверенности действовать от имени некоммерческой организации:)Tj", ei2s) + 2;
                            int efio = s.IndexOf(@"(\(фамилия,", ifio);
                            string sfio = s.Substring(ifio, efio - ifio);

                            sfio = clearRegX(sfio);
                            sfio = GetWithIn(sfio);

//                            richTextBox1.AppendText(getDir(sfio) + " == " + getDolz(sfio) + '\n');

                            csv_line += FormatCSV(sfio, 't');

                            //дата отчета
                            int edat = s.IndexOf(@"(\(дата\", efio + 2) - 2;
                            string sdat = s.Substring(efio + 4, edat - efio);

                            sdat = GetWithIn(sdat);

//                            richTextBox1.AppendText(sdat + '\n');
                            csv_line += FormatCSV(sdat, 'd');

                            //Бухгалтер

                            int ibuh = s.IndexOf("(Лицо, ответственное за ведение бухгалтерского учета:)Tj", ifio) + 2;
                            int ebih = s.IndexOf(@"(\(фамилия", ibuh);

                            string sbuh = s.Substring(ibuh, ebih - ibuh);

                            sbuh = clearRegX(sbuh);
                            sbuh = GetWithIn(sbuh);

                            csv_line += FormatCSV(sbuh, 't');

//                            richTextBox1.AppendText(getDir(sbuh) + " == " + getDolz(sbuh));

                            csv_line += FormatCSV(filename, 't');

                            csv_line += FormatCSV(File.GetCreationTime(filename).ToString("dd.MM.yyyy"), 'd'); // filedate

                            if (rf != null)
                            {
                                csv_line += FormatCSV(rf.nameNKO, 't');
                                csv_line += FormatCSV(rf.indexFile, 't');
                                csv_line += FormatCSV(rf.ogrn, 't');
                                csv_line += FormatCSV(rf.formNKO, 't');
                                csv_line += FormatCSV(rf.vid, 't');
                                csv_line += FormatCSV(rf.period, 'n', true);
                            }
                        }
                        if (_add)
                        {
                            fcsv.Add(csv_line); 
                        }
                    }
                    catch { }

                }

                addToFile(fcsv);

            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            startParse();
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            
        }

        public static bool isText(string s)
        {
            foreach (char c in s)
            {
                if (!Char.IsLetter(c) && !Char.IsSymbol(c))
                    return false;
            }
            return true;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if (autostart)
            {
                pathCSV.Text = csv_path;
                pathPDF.Text = _xml_index_path;
                minSumm.Text = _min_summ;
                viewReport.Text = _viewReport;
                System.Threading.Thread.Sleep(2000);
                startParse();
            }
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                pathCSV.Text = saveFileDialog1.FileName;
                csv_path = saveFileDialog1.FileName; 
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (files.Count > 0)
            {
                files.Clear();
            }


            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string dir_name = Path.GetDirectoryName(openFileDialog1.FileName);
                pathPDF.Text = dir_name;  
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

        private void button4_Click(object sender, EventArgs e)
        {
            Text = Path.GetExtension(@"D:\123\000.pdf");
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void pathPDF_TextChanged(object sender, EventArgs e)
        {
            if (Path.GetExtension(pathPDF.Text) == ".xml")
            {
                indexFile = (Index)LoadXml(typeof(Index), pathPDF.Text);
                listReports = indexFile.indexList;

                foreach (ReportFile rf in listReports)
                {
                    files.Add(rf.pathPdf);
                }
            }
            else
            {
                foreach (string filename in Directory.GetFiles(pathPDF.Text, "*.pdf", SearchOption.AllDirectories))
                {
                    if (files.Count < 4)
                    {
                        files.Add(filename);
                        Debug.WriteLine(filename);
                    }
                }
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            
        }
    }
}
