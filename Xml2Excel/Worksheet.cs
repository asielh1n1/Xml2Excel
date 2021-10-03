using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Serialization;

namespace Xml2Excel
{
    public class Worksheet
    {
        [XmlAttribute("name")]
        public string name { get; set; }
        [XmlAttribute("tabColor")]
        public string tabColor{ get;set; }
        [XmlArrayItem("cell", typeof(Cell))]
        [XmlArray("cells")]
        public List<Cell> Cells { get; set; }

        [XmlArrayItem("range", typeof(Range))]
        [XmlArray("ranges")]
        public List<Range> Ranges { get; set; }

        [XmlArrayItem("row", typeof(Row))]
        [XmlArray("rows")]
        public List<Row> Rows { get; set; }

        [XmlArrayItem("column", typeof(Column))]
        [XmlArray("columns")]
        public List<Column> Columns { get; set; }

        [XmlAttribute("rowHeight")]
        public double rowHeight { get; set; }
        [XmlAttribute("password")]
        public string password { get; set; }
    }
}
