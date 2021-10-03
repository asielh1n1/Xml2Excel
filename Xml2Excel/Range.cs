using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Serialization;

namespace Xml2Excel
{
    public class Range
    {
        [XmlAttribute("cell1")]
        public string cell1 { get; set; }
        [XmlAttribute("cell2")]
        public string cell2 { get; set; }
        [XmlAttribute("merge")]
        public bool merge { get; set; }
        [XmlAttribute("clear")]
        public bool clear { get; set; }
    }
}
