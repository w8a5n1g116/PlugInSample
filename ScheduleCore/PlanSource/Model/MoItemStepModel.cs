using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Braincase.GanttChart;

namespace PlanSource.Model
{
    public class MoItemStepModel : Braincase.GanttChart.Task
    {
        public MoItemStepModel()
        {

        }

        public MoItemStepModel(ProjectManager manager)
                : base()
        {
            Manager = manager;
        }

        private ProjectManager Manager { get; set; }

        public string MOId { get; set; }
        public string MOItemTaskStepId { get; set; }
        public string MOName { get; set; }
        public string ProductId { get; set; }
        public string ProductName { get; set; }
        public string Specification { get; set; }
        public string Venture { get; set; }
        public string PlannedDateFrom { get; set; }
        public string PlannedDateTo { get; set; }
        public string TaskStatus { get; set; }
        public string SapDeliveryDate { get; set; }

        public new TimeSpan Start { get { return base.Start; } set { Manager.SetStart(this, value); } }
        public new TimeSpan End { get { return base.End; } set { Manager.SetEnd(this, value); } }
        public new TimeSpan Duration { get { return base.Duration; } set { Manager.SetDuration(this, value); } }
        public new float Complete { get { return base.Complete; } set { Manager.SetComplete(this, value); } }
    }
}
