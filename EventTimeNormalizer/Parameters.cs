using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EventTimeNormalizer
{
    public class Parameters
    {
        [CommandLine.Option('s', "If set skips the first excel row", DefaultValue = false, Required = false)]
        public bool UseFirstRowHeader { get; set; }

        [CommandLine.Option('i', "Input excel file path", Required = true)]
        public string InputExcelFile { get; set; }

        [CommandLine.Option('o', "Output excel file path", Required = true)]
        public string OutputExcelFile { get; set; }

        [CommandLine.Option('n', "Value cell column", DefaultValue = 2, Required = true)]
        public int ValueColumn { get; set; }

        [CommandLine.Option('d', "Date time cell column", DefaultValue = 0, Required = true)]
        public int DataColumn { get; set; }

        [CommandLine.Option('v', "Verbose", DefaultValue = false, Required = false)]
        public bool Verbose { get; set; }

        [CommandLine.Option('k', "Comma separated numbers composing the primary key (ie 1,2,3 for a three column composite key)", DefaultValue = "1", Required = true)]
        public string KeyCellsCSV { get; set; }
    }
}
