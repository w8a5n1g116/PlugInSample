using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PlanSource.Model
{
    public class MoModel
    {
        public string MOId { get; set; }
        public string MOName { get; set; }
        public string ProductId { get; set; }
        public string ProductName { get; set; }
        public string PlannedDateFrom { get; set; }
        public string PlannedDateTo { get; set; }
        public string MOStatus { get; set; }
        public string SapDeliveryDate { get; set; }
    }
}
