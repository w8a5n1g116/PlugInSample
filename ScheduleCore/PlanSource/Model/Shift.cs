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
    
    public partial class Shift
    {
        public string ShiftId { get; set; }
        public string ShiftName { get; set; }
        public string ShiftShortName { get; set; }
        public string ShiftDescription { get; set; }
        public string SiteId { get; set; }
        public string DivisionId { get; set; }
        public string TimeFrom { get; set; }
        public string TimeTo { get; set; }
        public string FactoryId { get; set; }
        public Nullable<System.DateTime> ModifyDate { get; set; }
        public Nullable<System.DateTime> CreateDate { get; set; }
        public string WorkcenterId { get; set; }
        public Nullable<decimal> TimeRange { get; set; }
    }
}