using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;

namespace MinUstPdf
{
    [XmlRoot("indexFileReports")]
    public class Index
    {
        /*[XmlElement("pathToPdfFiles")]
        public string valuePdf;*/
        [XmlArray("reportFiles")]
        [XmlArrayItem("item")]
        public List<ReportFile> indexList;
    }

    public class ReportFile
    {
        public string indexFile;
        public string nameNKO;
        public string formNKO;
        public string urlPdf;
        public string pathPdf;
        public string period;
        public string ogrn;
        public string vid;
    }
}
