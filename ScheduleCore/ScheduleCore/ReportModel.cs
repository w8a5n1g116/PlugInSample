using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace PlugInSample.Model
{
    public class ReportModel
    {
        public ReportModel(string unit)
        {
            Unit = unit;
        }

        public string Name { get; set; }
        public double Plan { get; set; }
        public double Real { get; set; }
        public double Rate { get; set; }
        public string Unit { get; set; }
    }
}