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
    
    public partial class SysUser
    {
        public string UserId { get; set; }
        public string UserName { get; set; }
        public string EmployeeNo { get; set; }
        public string UserDescription { get; set; }
        public string UserSex { get; set; }
        public byte[] UserPassword { get; set; }
        public string Division { get; set; }
        public string DivisionId { get; set; }
        public string duties { get; set; }
        public string JobsId { get; set; }
        public string JobName { get; set; }
        public Nullable<System.DateTime> ChangePasswordDate { get; set; }
        public Nullable<bool> IsPasswordLongLive { get; set; }
        public Nullable<bool> IsChangePasswordWhenLogin { get; set; }
        public Nullable<bool> IsActivity { get; set; }
        public string MailAddress { get; set; }
        public Nullable<bool> IsReceiveEmail { get; set; }
        public string Phone { get; set; }
        public Nullable<bool> Protected { get; set; }
        public Nullable<System.DateTime> PasswordErrorDate { get; set; }
        public Nullable<bool> IsInquireOnlineUser { get; set; }
        public Nullable<bool> IsRunSCADADesigner { get; set; }
        public string FactoryId { get; set; }
        public string SiteId { get; set; }
        public string ProductId { get; set; }
        public Nullable<System.DateTime> CheckStartDate { get; set; }
        public Nullable<System.DateTime> CheckEndDate { get; set; }
        public Nullable<bool> NoCheckEndDate { get; set; }
        public string ShiftId { get; set; }
        public string GroupsId { get; set; }
        public Nullable<bool> IsOnline { get; set; }
        public string FtpPhoto { get; set; }
        public string LatestActiveResource { get; set; }
        public string LatestActivePlugin { get; set; }
        public Nullable<System.DateTime> LatestActiveDate { get; set; }
        public string WorkshopId { get; set; }
        public string UserStatus { get; set; }
        public string AutoRunPluginCommand { get; set; }
        public Nullable<bool> IsAccount { get; set; }
        public Nullable<bool> IsVendorUser { get; set; }
        public Nullable<bool> IsOMFUser { get; set; }
        public Nullable<bool> IsAdmin { get; set; }
        public string WorkcenterId { get; set; }
        public string VentureId { get; set; }
        public Nullable<System.DateTime> ModifyDate { get; set; }
        public string CreateUserId { get; set; }
        public Nullable<System.DateTime> CreateDate { get; set; }
        public string LogonLimitedPC { get; set; }
    }
}
