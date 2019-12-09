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
    
    public partial class Product
    {
        public string ProductId { get; set; }
        public string ProductRootId { get; set; }
        public string ProductRevision { get; set; }
        public string ProductStd { get; set; }
        public string ProductModel { get; set; }
        public string ProductShortName { get; set; }
        public string ProductDescription { get; set; }
        public string ProductDescription_1 { get; set; }
        public string ProductDescription_2 { get; set; }
        public string ManufacturerInfor { get; set; }
        public string ProductTypeId { get; set; }
        public string ProductClassId { get; set; }
        public string PurchaseUOMId { get; set; }
        public string InventoryUOMId { get; set; }
        public string UsageUOMId { get; set; }
        public Nullable<bool> IsActivity { get; set; }
        public Nullable<bool> IsDefaultProduct { get; set; }
        public Nullable<bool> IsHumidity { get; set; }
        public string IsRoHS { get; set; }
        public string BOMId { get; set; }
        public string ProductBarcodeId { get; set; }
        public string CustomerId { get; set; }
        public string CustomerPN { get; set; }
        public string BoxLabelId { get; set; }
        public string WorkflowId { get; set; }
        public Nullable<decimal> StandardCost { get; set; }
        public Nullable<decimal> StandardCapacity { get; set; }
        public Nullable<decimal> CurrentCost { get; set; }
        public Nullable<decimal> Weight { get; set; }
        public Nullable<decimal> NetWeight { get; set; }
        public Nullable<bool> IsFIFO { get; set; }
        public Nullable<bool> IsExceedPicking { get; set; }
        public Nullable<int> SafeStorageTime { get; set; }
        public Nullable<decimal> SafetyInventory { get; set; }
        public Nullable<decimal> MaxiInventory { get; set; }
        public string DeleteFlag { get; set; }
        public Nullable<System.DateTime> ModifyDate { get; set; }
        public Nullable<double> StandCapacity { get; set; }
        public Nullable<bool> IsUrgency { get; set; }
        public string ProjectCode { get; set; }
        public Nullable<bool> IsSendQCByProductionDate { get; set; }
        public Nullable<bool> IsReceiveByBox { get; set; }
        public string ProductInspecTypeId { get; set; }
        public Nullable<int> StandardQtyOfBox { get; set; }
        public string MachineModel { get; set; }
        public string MachineNo { get; set; }
        public Nullable<int> PurchaseCycle { get; set; }
        public Nullable<bool> IsLowValueConsumable { get; set; }
        public string SNRuleId { get; set; }
        public string FactoryId { get; set; }
        public Nullable<int> BOMLevel { get; set; }
        public Nullable<decimal> LabourCost { get; set; }
        public Nullable<System.DateTime> CreateDate { get; set; }
        public string SAP_WERKS { get; set; }
        public string SAP_LVORM { get; set; }
        public string SAP_BESKZ { get; set; }
        public string ERPID { get; set; }
        public string ERPNo { get; set; }
        public byte[] ERPDate { get; set; }
        public string SAP_ZEINR { get; set; }
        public Nullable<bool> IsRM { get; set; }
        public Nullable<int> Warranty { get; set; }
        public Nullable<int> ServiceLife { get; set; }
    }
}
