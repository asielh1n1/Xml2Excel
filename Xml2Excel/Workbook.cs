using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Serialization;

namespace Xml2Excel
{
    [XmlRoot("workbook")]
    public class Workbook
    {
        [XmlAttribute("author")]
        public string author { get; set; }
        [XmlAttribute("title")]
        public string title { get; set; }
        [XmlAttribute("subject")]
        public string subject { get; set; }
        [XmlAttribute("category")]
        public string category { get; set; }
        [XmlAttribute("keywords")]
        public string keywords { get; set; }
        [XmlAttribute("comments")]
        public string comments { get; set; }
        [XmlAttribute("status")]
        public string status { get; set; }
        [XmlAttribute("company")]
        public string company { get; set; }
        [XmlAttribute("manager")]
        public string manager { get; set; }
        [XmlArrayItem("worksheet", typeof(Worksheet))]
        [XmlArray("worksheets")]
        public List<Worksheet> Worksheets { get; set; }
    }
}
