//------------------------------------------------------------------------------
// <auto-generated>
//     此代码已从模板生成。
//
//     手动更改此文件可能导致应用程序出现意外的行为。
//     如果重新生成代码，将覆盖对此文件的手动更改。
// </auto-generated>
//------------------------------------------------------------------------------

namespace PlanSource.Model
{
    using System;
    using System.Collections.Generic;
    
    public partial class MOItemTask
    {
        public string MOItemTaskId { get; set; }
        public string MOItemTaskName { get; set; }
        public string Sequence { get; set; }
        public string MOId { get; set; }
        public string MOItemId { get; set; }
        public string ProductId { get; set; }
        public string WorkflowId { get; set; }
        public string WorkflowStepId { get; set; }
        public string BOMId { get; set; }
        public string WorkcenterId { get; set; }
        public string VentureId { get; set; }
        public string EmployeeNo { get; set; }
        public string ShiftId { get; set; }
        public Nullable<int> TaskPlannedQty { get; set; }
        public string TaskStatusId { get; set; }
        public string TaskStatus { get; set; }
        public string TaskPriorityLevel { get; set; }
        public string ParentMOItemTaskId { get; set; }
        public string SendQCReportId { get; set; }
        public string QCResult { get; set; }
        public Nullable<System.DateTime> PlannedDateFrom { get; set; }
        public Nullable<System.DateTime> PlannedDateTo { get; set; }
        public Nullable<System.DateTime> ExecuteDateFrom { get; set; }
        public Nullable<System.DateTime> ExecuteDateTo { get; set; }
        public string TaskComment { get; set; }
        public Nullable<System.DateTime> CreateDate { get; set; }
        public string WorkFlowStepName { get; set; }
        public string SpecificationID { get; set; }
        public Nullable<System.DateTime> ReleaseDate { get; set; }
    }
}