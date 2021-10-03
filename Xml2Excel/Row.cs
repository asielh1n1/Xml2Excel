using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Serialization;

namespace Xml2Excel
{
    public class Row
    {
        [XmlAttribute("number")]
        public int number { get; set; }
        [XmlAttribute("height")]
        public int height { get; set; }
        [XmlAttribute("style")]
        public string style { get; set; }
        [XmlAttribute("adjustToContents")]
        public bool adjustToContents { get; set; }
    }
}
