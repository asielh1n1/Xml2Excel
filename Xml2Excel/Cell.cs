using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Serialization;

namespace Xml2Excel
{
    public class Cell
    {
        [XmlAttribute("row")]
        public int row { get; set; }
        [XmlAttribute("column")]
        public int column { get; set; }   
        [XmlText()]
        public string value { get; set; }
        [XmlAttribute("image")]
        public string image { get; set; }

        [XmlAttribute("imageScale")]
        public double imageScale { get; set; }
        [XmlAttribute("style")]
        public string style { get; set; }
        [XmlAttribute("formula")]
        public string formula { get; set; }
        [XmlAttribute("link")]
        public string link { get; set; }
        [XmlAttribute("extLink")]
        public string extLink { get; set; }
        [XmlAttribute("numberFormat")]
        public string numberFormat { get; set; }

        [XmlAttribute("formatId")]
        public int formatId { get; set; }



    }
}
