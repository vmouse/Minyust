using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace MinUstPdf
{
    [XmlRoot("Minjust_Filter")]
    public class Filter
    {
        [XmlElement("pathToPdfFiles")]
        public string pathPdf;
        [XmlElement("pathToIndexFiles")]
        public string pathIndex;
        [XmlElement("valueOrgFormNKO")]
        public string orgForm;
        [XmlElement("valueRegion")]
        public string region;
        [XmlElement("valueViewReport")]
        public string typeReport;
        [XmlElement("year")]
        public string year;
        [XmlArray("stopWords")]
        [XmlArrayItem("wordItem")]
        public List<string> stopWords;
    }
}
