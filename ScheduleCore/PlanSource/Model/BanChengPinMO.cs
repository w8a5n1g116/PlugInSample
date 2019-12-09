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
    
    public partial class BanChengPinMO
    {
        public string MOId { get; set; }
        public string MOName { get; set; }
        public string MODescription { get; set; }
        public string MOTypeId { get; set; }
        public string MOClassId { get; set; }
        public string MOStatusId { get; set; }
        public string MOStatus { get; set; }
        public string ProductId { get; set; }
        public string BOMId { get; set; }
        public string SalaryCalcTypeId { get; set; }
        public Nullable<System.DateTime> PlannedDateFrom { get; set; }
        public Nullable<System.DateTime> PlannedDateTo { get; set; }
        public Nullable<System.DateTime> ExecuteDateFrom { get; set; }
        public Nullable<System.DateTime> ExecuteDateTo { get; set; }
        public Nullable<decimal> MOQtyRequired { get; set; }
        public Nullable<decimal> MOQtyDone { get; set; }
        public Nullable<decimal> MOQtyCurrent { get; set; }
        public string BOMName { get; set; }
        public string WorkflowId { get; set; }
        public string PickStatus { get; set; }
        public string PriorityLevel { get; set; }
        public string SOId { get; set; }
        public string SOName { get; set; }
        public string CustomerId { get; set; }
        public string ParentMOId { get; set; }
        public string FactoryId { get; set; }
        public string BatchCode { get; set; }
        public Nullable<bool> IsPickingListUpload { get; set; }
        public Nullable<int> MinPack { get; set; }
        public string ApproveUserId { get; set; }
        public Nullable<System.DateTime> ApproveDate { get; set; }
        public Nullable<System.DateTime> ReleaseDate { get; set; }
        public Nullable<System.DateTime> StartWorkDate { get; set; }
        public Nullable<System.DateTime> CloseDate { get; set; }
        public string DeleteFlag { get; set; }
        public byte[] ERPDate { get; set; }
        public string ERPId { get; set; }
        public string ERPNo { get; set; }
        public string ItemSequence { get; set; }
        public Nullable<System.DateTime> ModifyDate { get; set; }
        public string CreateUserId { get; set; }
        public Nullable<System.DateTime> CreateDate { get; set; }
        public string SAP_AUART { get; set; }
        public string SAP_BUKRS { get; set; }
        public Nullable<System.DateTime> SAP_AEDAT { get; set; }
        public string SAP_WERKS { get; set; }
        public string SAP_GSBER { get; set; }
        public string SAP_GMEIN { get; set; }
        public string SAP_VERID { get; set; }
        public string SAP_PHAS0 { get; set; }
        public string SAP_PHAS1 { get; set; }
        public string SAP_PHAS2 { get; set; }
        public string SAP_PHAS3 { get; set; }
        public Nullable<double> SAP_PSMNG { get; set; }
        public string SAP_ADD3 { get; set; }
        public string SAP_ADD4 { get; set; }
        public string SAP_BESKZ { get; set; }
        public string CreateUserName { get; set; }
        public string SAP_Z0001 { get; set; }
        public string SAP_Z0002 { get; set; }
        public string SAP_Z0003 { get; set; }
        public string SAP_Z0006 { get; set; }
        public string SAP_Z0007 { get; set; }
        public string SAP_Z0008 { get; set; }
        public string SAP_Z0009 { get; set; }
        public string SAP_Z010 { get; set; }
        public string SAP_VDATU { get; set; }
        public string SAP_CHARG { get; set; }
        public Nullable<System.DateTime> SendMaterialDate { get; set; }
        public string SAP_Remarks { get; set; }
        public string SOSequence { get; set; }
    }
}
