using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Serialization;

namespace Xml2Excel
{
    public class Column
    {
        [XmlAttribute("number")]
        public int number { get; set; }
        [XmlAttribute("width")]
        public int width { get; set; }
        [XmlAttribute("style")]
        public string style { get; set; }
        [XmlAttribute("adjustToContents")]
        public bool adjustToContents { get; set; }
    }
}
