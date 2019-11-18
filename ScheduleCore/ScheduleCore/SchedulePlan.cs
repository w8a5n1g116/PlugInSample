using PlanSource.Model;
using ScheduleCore.Model;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Braincase.GanttChart;
using PlugInSample.Model;
using System.Threading;

namespace ScheduleCore
{
    public class SchedulePlan
    {
        /// <summary>
        /// 数据库连接字符串
        /// </summary>
        public static string connectString = "metadata=res://*/Model.PlanModel.csdl|res://*/Model.PlanModel.ssdl|res://*/Model.PlanModel.msl;provider=System.Data.SqlClient;provider connection string=\"data source=192.168.0.54;initial catalog=OrBitMOM_Kocel;persist security info=True;user id=wangjunjie;password=123456;MultipleActiveResultSets=True;App=EntityFramework\"";
        //static string TestconnectString = "metadata=res://*/TestModel.TestPlanModel.csdl|res://*/TestModel.TestPlanModel.ssdl|res://*/TestModel.TestPlanModel.msl;provider=System.Data.SqlClient;provider connection string=\"data source=192.168.41.58;initial catalog=OrBitMOM_Test;persist security info=True;user id=wangjunjie;password=123456;MultipleActiveResultSets=True;App=EntityFramework\"";

        public static OrBitMOM_KocelEntities1 orbitEntity = new OrBitMOM_KocelEntities1(connectString);
        //OrBitMOM_TestEntities orbitEntity = new OrBitMOM_TestEntities(TestconnectString);


        public static OrBitMOM_KocelEntities1 GetEntity()
        {
            if (orbitEntity != null)
                return orbitEntity;
            else
            {
                orbitEntity = new OrBitMOM_KocelEntities1(connectString);
                return orbitEntity;
            }
                
        }

        /// <summary>
        /// 经营体工序队列
        /// </summary>
        private List<StepModel> stepModelList = new List<StepModel>();

        //private List<string> MoIdListLocal = new List<string>();
        public List<ProductSpecification> psList = new List<ProductSpecification>();

        public void getPSList()
        {
            psList = orbitEntity.ProductSpecification.ToList();
        }

        public string FactoryID;

        public SchedulePlan()
        {

        }

        public SchedulePlan(string userId)
        {
            if (string.IsNullOrEmpty(userId))
                userId = "SUR1000003KX";

            try
            {
                this.FactoryID = GetUserFactory(userId);
            }
            catch (Exception ex)
            {
                throw new Exception("用户ID:" + userId+"未分配工厂ID");
            }

            Thread thread = new Thread(getPSList);
            thread.Start();

            GetStepModelList();
        }

        /// <summary>
        /// 构建经营体工序队列
        /// </summary>
        public void GetStepModelList()
        {
            List<string> specificationList = GetSpecificationList();

            foreach (string spe in specificationList)
            {
                StepModel stepModel = new StepModel();
                stepModel.SpecificationId = spe;

                List<VentureModel> ventureModelList = new List<VentureModel>();
                List<string> VentureIdList = GetVentureBySpecificationId(spe);

                foreach (string ventureId in VentureIdList)
                {
                    VentureModel ventureModel = new VentureModel(spe,
                        ventureId,
                        GetVentureCapById(spe, ventureId),
                        orbitEntity.MOItemTask.Where(p => p.TaskStatus != "已登记" && p.TaskStatus != "已报工" && p.TaskStatus != "强制关闭" && p.SpecificationID == spe && p.PlannedDateFrom != null && p.VentureId == ventureId).OrderBy(p => p.PlannedDateFrom).ToList());
                    ventureModelList.Add(ventureModel);
                }

                stepModel.VentureModelList = ventureModelList;

                stepModelList.Add(stepModel);
            }
        }
        
        /// <summary>
        /// 获取从sap下达成品的生产任务单
        /// </summary>
        /// <returns></returns>
        public List<MoModel> GetMoList(int pageNumber)
        {
            int pageSize = 20;

            List<MoModel> MoModelList = new List<MoModel>();
            List<ChanchengPinMO> chanchengpinMoList = orbitEntity.ChanchengPinMO.Where(p => p.FactoryId == this.FactoryID).OrderBy(p => p.SAP_VDATU).Skip(pageNumber * pageSize).Take(pageSize).ToList();
            //List<MO> MoAllList = orbitEntity.MO.Where(p => p.MOStatus == "已审批" ).ToList();//&& p.FactoryId == "FAC1000000SF"

            List<MO> MoList = CovertMOList(chanchengpinMoList);//new List<MO>();

            //foreach (var mo in MoAllList)
            //{
            //    if(IsFinishedProduct(mo))
            //    {
            //        MoList.Add(mo);
            //    }
            //}

            foreach(var mo in MoList)
            {
                MoModel moModel = new MoModel();
                moModel.MOId = mo.MOId;
                moModel.MOName = mo.MOName;
                moModel.MOStatus = mo.MOStatus;
                moModel.PlannedDateFrom = mo.PlannedDateFrom != null ?((DateTime)mo.PlannedDateFrom).ToString("yyyy-MM-dd HH:mm:ss") : "";
                moModel.PlannedDateTo = mo.PlannedDateTo != null ? ((DateTime)mo.PlannedDateTo).ToString("yyyy-MM-dd HH:mm:ss") : "";
                if(mo.SAP_VDATU != null)
                    moModel.SapDeliveryDate = mo.SAP_VDATU;
                else
                    moModel.SapDeliveryDate = "";

                Product pro = orbitEntity.Product.Where(p => p.ProductId == mo.ProductId).FirstOrDefault();
                if(pro != null)
                {
                    moModel.ProductName = pro.ProductDescription;

                    ProductRoot pr = orbitEntity.ProductRoot.Where(p => p.ProductRootId == pro.ProductRootId).FirstOrDefault();
                    if(pr != null)
                    {
                        moModel.ProductId = pr.ProductName;
                    }
                    
                }

                MoModelList.Add(moModel);
            }

            return MoModelList.OrderBy(p => p.SapDeliveryDate).ToList();
        }

        /// <summary>
        /// 根据MOid获取工序计划
        /// </summary>
        /// <param name="MoId"></param>
        /// <returns></returns>
        public List<MoItemStepModel> GetMoItemStepModelList(string MoId)
        {
            List<MoItemStepModel> mismList = new List<MoItemStepModel>();

            List<MOItemTask> stepList = new List<MOItemTask>();

            List<MO> moList = GetMoListById(new List<string>() { MoId });

            foreach(var mo in moList)
            {
                stepList.AddRange(GetStepByMoId(MoId));
            }

            foreach (var mits in stepList)
            {
                MoItemStepModel mism = new MoItemStepModel();

                mism.MOId = mits.MOId;
                mism.MOItemTaskStepId = mits.MOItemTaskId;
                mism.PlannedDateFrom = mits.PlannedDateFrom != null ? ((DateTime)mits.PlannedDateFrom).ToString("yyyy-MM-dd HH:mm:ss") : "";
                mism.PlannedDateTo = mits.PlannedDateTo != null ? ((DateTime)mits.PlannedDateTo).ToString("yyyy-MM-dd HH:mm:ss") : "";
                MO mo = orbitEntity.MO.Where(p => p.MOId == mits.MOId).FirstOrDefault();
                if (mo.SAP_VDATU != null)
                    mism.SapDeliveryDate = mo.SAP_VDATU;
                else
                    mism.SapDeliveryDate = "";
                mism.MOName = mo.MOName;

                Specification sp = orbitEntity.Specification.Where(p => p.SpecificationId == mits.SpecificationID).FirstOrDefault();
                if(sp != null)
                {
                    SpecificationRoot sprt = orbitEntity.SpecificationRoot.Where(p => p.SpecificationRootId == sp.SpecificationRootId).FirstOrDefault();
                    if (sprt != null)
                        mism.Specification = sprt.SpecificationName;
                    else
                        mism.Specification = "";
                }

                Venture vt = orbitEntity.Venture.Where(p => p.VentureId == mits.VentureId).FirstOrDefault();
                if(vt != null)
                {
                    mism.Venture = vt.VentureName;
                }
                else
                    mism.Venture = "";

                Product pro = orbitEntity.Product.Where(p => p.ProductId == mism.ProductId).FirstOrDefault();
                if (pro != null)
                {
                    mism.ProductName = pro.ProductDescription;

                    ProductRoot pr = orbitEntity.ProductRoot.Where(p => p.ProductRootId == pro.ProductRootId).FirstOrDefault();
                    if (pr != null)
                    {
                        mism.ProductId = pr.ProductName;
                    }

                }

                mismList.Add(mism);
            }

            return mismList.OrderBy(p => p.SapDeliveryDate).ToList();
        }

        /// <summary>
        /// 获取已经被排产过的计划
        /// </summary>
        /// <returns></returns>
        public List<MoItemStepModel> GetScheduledMoItemStepModelList(int pageNumber)
        {
            int pageSize = 20;
            List<MoItemStepModel> moStepModelList = new List<MoItemStepModel>();

            List<MOItemTask> stepList = new List<MOItemTask>();

            foreach (var stepModel in stepModelList)
            {
                foreach(var ventureModel in stepModel.VentureModelList)
                {
                    stepList.AddRange(ventureModel.StepList);                   
                }
            }


            stepList = stepList.OrderByDescending(p => p.PlannedDateFrom).Skip(pageSize * pageNumber).Take(pageSize).ToList();


            List<MoItemStepModel> mismList = new List<MoItemStepModel>();

            foreach (var mits in stepList)
            {
                MoItemStepModel mism = new MoItemStepModel();

                mism.MOId = mits.MOId;
                mism.MOItemTaskStepId = mits.MOItemTaskId;
                mism.PlannedDateFrom = mits.PlannedDateFrom != null ? ((DateTime)mits.PlannedDateFrom).ToString("yyyy-MM-dd HH:mm:ss") : "";
                mism.PlannedDateTo = mits.PlannedDateTo != null ? ((DateTime)mits.PlannedDateTo).ToString("yyyy-MM-dd HH:mm:ss") : "";
                MO mo = orbitEntity.MO.Where(p => p.MOId == mits.MOId).FirstOrDefault();
                if (mo.SAP_VDATU != null)
                    mism.SapDeliveryDate = mo.SAP_VDATU;
                else
                    mism.SapDeliveryDate = "";
                mism.MOName = mo.MOName;

                Specification sp = orbitEntity.Specification.Where(p => p.SpecificationId == mits.SpecificationID).FirstOrDefault();
                if (sp != null)
                {
                    SpecificationRoot sprt = orbitEntity.SpecificationRoot.Where(p => p.SpecificationRootId == sp.SpecificationRootId).FirstOrDefault();
                    if (sprt != null)
                        mism.Specification = sprt.SpecificationName;
                    else
                        mism.Specification = "";
                }

                Venture vt = orbitEntity.Venture.Where(p => p.VentureId == mits.VentureId).FirstOrDefault();
                if (vt != null)
                {
                    mism.Venture = vt.VentureName;
                }
                else
                    mism.Venture = "";

                Product pro = orbitEntity.Product.Where(p => p.ProductId == mits.ProductId).FirstOrDefault();
                if (pro != null)
                {
                    mism.ProductName = pro.ProductDescription;

                    ProductRoot pr = orbitEntity.ProductRoot.Where(p => p.ProductRootId == pro.ProductRootId).FirstOrDefault();
                    if (pr != null)
                    {
                        mism.ProductId = pr.ProductName;
                    }

                }

                mism.TaskStatus = mits.TaskStatus;

                mismList.Add(mism);
            }

            moStepModelList.AddRange(mismList);

            return moStepModelList.OrderBy(p => p.SapDeliveryDate).ToList();
        }

        /// <summary>
        /// 通过Moid列表进行排产
        /// </summary>
        /// <param name="MoIdList"></param>
        public List<MoItemStepModel> DoSchedule(List<string> MoIdList)
        {
            List<MOItemTask> moitemtaskList = new List<MOItemTask>();

            //排产过程
            List<MO> moList = GetMoListById(MoIdList);

            foreach (var mo in moList)
            {
                try
                {
                    List<DateTime?> FinalStartTimeList = new List<DateTime?>();

                    List<MO> halfMoList = GetBomMoList(mo);
                    foreach(var halfMo in halfMoList)
                    {
                        try
                        {
                            List<MOItemTask> halfStepList = GetStepByMoId(halfMo.MOId);
                            FinalStartTimeList.Add(ScheduleProcess(halfStepList,null));

                            moitemtaskList.AddRange(halfStepList);
                        }
                        catch (Exception e)
                        {
                            throw new Exception("子单据编号为:" + halfMo.MOName + "的工单排产出现问题\n" + e.ToString());
                        }
                    }
                    List<MOItemTask> stepList = GetStepByMoId(mo.MOId);
                    ScheduleProcess(stepList, FinalStartTimeList.Max());

                    moitemtaskList.AddRange(stepList);
                } 
                catch(Exception e)
                {
                    throw new Exception("单据编号为:" + mo.MOName + "的工单排产出现问题\n" + e.ToString());
                }
            }

            List<MoItemStepModel> mismList = new List<MoItemStepModel>();

            foreach (var mits in moitemtaskList)
            {
                MoItemStepModel mism = new MoItemStepModel();

                mism.MOId = mits.MOId;
                mism.MOItemTaskStepId = mits.MOItemTaskId;
                mism.PlannedDateFrom = mits.PlannedDateFrom != null ? ((DateTime)mits.PlannedDateFrom).ToString("yyyy-MM-dd HH:mm:ss") : "";
                mism.PlannedDateTo = mits.PlannedDateTo != null ? ((DateTime)mits.PlannedDateTo).ToString("yyyy-MM-dd HH:mm:ss") : "";
                MO mo = orbitEntity.MO.Where(p => p.MOId == mits.MOId).FirstOrDefault();
                if (mo.SAP_VDATU != null)
                    mism.SapDeliveryDate = mo.SAP_VDATU;
                else
                    mism.SapDeliveryDate = "";
                mism.MOName = mo.MOName;

                Specification sp = orbitEntity.Specification.Where(p => p.SpecificationId == mits.SpecificationID).FirstOrDefault();
                if (sp != null)
                {
                    SpecificationRoot sprt = orbitEntity.SpecificationRoot.Where(p => p.SpecificationRootId == sp.SpecificationRootId).FirstOrDefault();
                    if (sprt != null)
                        mism.Specification = sprt.SpecificationName;
                    else
                        mism.Specification = "";
                }

                Venture vt = orbitEntity.Venture.Where(p => p.VentureId == mits.VentureId).FirstOrDefault();
                if (vt != null)
                {
                    mism.Venture = vt.VentureName;
                }
                else
                    mism.Venture = "";

                Product pro = orbitEntity.Product.Where(p => p.ProductId == mits.ProductId).FirstOrDefault();
                if (pro != null)
                {
                    mism.ProductName = pro.ProductDescription;

                    ProductRoot pr = orbitEntity.ProductRoot.Where(p => p.ProductRootId == pro.ProductRootId).FirstOrDefault();
                    if (pr != null)
                    {
                        mism.ProductId = pr.ProductName;
                    }

                }

                mism.TaskStatus = mits.TaskStatus;

                mismList.Add(mism);
            }

            return mismList;

        }

        /// <summary>
        /// 排产过程
        /// </summary>
        /// <param name="stepList"></param>
        public DateTime? ScheduleProcess(List<MOItemTask> stepList,DateTime? FinalStartTime)
        {
            bool isFuji = false;
            if (this.FactoryID == "FAC1000000SF")
                isFuji = true;

            if (!stepList.Any())
                throw new Exception("工单无工序计划，请检查!");

            //部件工单的结束时间
            DateTime? PartLastTime = null;

            string WorkflowId = stepList.FirstOrDefault().WorkflowId;

            List<WorkflowStepExpandDefault> wfsedList = GetSpecificationOrderByMoId(WorkflowId);

            foreach (var wfsed in wfsedList)
            {
                MOItemTask step = stepList.Where(p => p.SpecificationID == wfsed.SpecificationId).FirstOrDefault();

                if (step == null)
                    throw new Exception("工序计划不全，请检查；" + wfsed.SubWorkflowStepName + "无工序计划！");

                if (step.TaskStatus != "已登记")
                    continue;

                //获取产品工序标准工时
                int standardWorkingTime = GetStandardWorkingTimeByProductId_SpecificationId(step.ProductId, step.SpecificationID);

                StepModel stepModel = stepModelList.Where(p => p.SpecificationId == wfsed.SpecificationId).FirstOrDefault();

                //寻找上一个工序
                WorkflowStepExpandDefault previousWfsed = FindPrevioudWorkflowStepExpandDefault(wfsed);

                //如果不存在则是起始工序
                if (previousWfsed == null)
                {
                    if(FinalStartTime == null)
                        PartLastTime = stepModel.AddStep(step, standardWorkingTime, DateTime.Now,true, isFuji);
                    else
                        PartLastTime = stepModel.AddStep(step, standardWorkingTime, FinalStartTime,true, isFuji);
                }
                else
                {
                    StepModel previoudStepModel = stepModelList.Where(p => p.SpecificationId == previousWfsed.SpecificationId).FirstOrDefault();

                    DateTime? lastTime = previoudStepModel.FindStepEndTime(step.MOId);

                    if(lastTime>DateTime.Now)
                        PartLastTime = stepModel.AddStep(step, standardWorkingTime, lastTime,false, isFuji);
                    else
                        PartLastTime = stepModel.AddStep(step, standardWorkingTime, DateTime.Now, false, isFuji);
                }

            }

            return PartLastTime;
        }

        /// <summary>
        /// 调整计划
        /// </summary>
        /// <param name="mitsId"></param>
        /// <param name="startTime"></param>
        //public void AdjustPlan(string mitsId, DateTime startTime)
        //{
        //    MOItemTaskStep step = orbitEntity.MOItemTaskStep.Where(p => p.MOItemTaskStepId == mitsId).FirstOrDefault();

        //    MOItemTask mit = orbitEntity.MOItemTask.Where(p => p.MOItemTaskId == step.MOItemTaskId).FirstOrDefault();

        //    //全部待排工序
        //    List<MOItemTaskStep> mitsList = new List<MOItemTaskStep>();

        //    if (mit.ParentMOItemTaskId != null)
        //    {
        //        List<WorkflowStepExpandDefault> wfsedList = GetSpecificationOrderByMoId(mit.WorkflowId);
        //        //当前工序
        //        WorkflowStepExpandDefault Wfsed = wfsedList.Where(p => p.SpecificationId == step.SpecificationId).FirstOrDefault();

        //        StepModel stepModel = stepModelList.Where(p => p.SpecificationId == Wfsed.SpecificationId).FirstOrDefault();
        //        List<MOItemTaskStep> list = stepModel.RemoveMOItemTaskStepById(mitsId);

        //        while (FindNextWorkflowStepExpandDefault(Wfsed) != null)
        //        {
        //            stepModel = stepModelList.Where(p => p.SpecificationId == Wfsed.SpecificationId).FirstOrDefault();
        //            list = stepModel.RemoveMOItemTaskStepById(mitsId);

        //            if (list != null)
        //                mitsList.AddRange(list);

        //            Wfsed = FindNextWorkflowStepExpandDefault(Wfsed);
        //        }
        //    }
        //    else
        //    {

        //    }
        //}

        /// <summary>
        /// 通过工作流id获取其所需的工序 按正序排列
        /// </summary>
        /// <returns></returns>
        List<WorkflowStepExpandDefault> GetSpecificationOrderByMoId(string WorkflowId)
        {
            List<WorkflowStepExpandDefault> wfsedList = new List<WorkflowStepExpandDefault>();

            var wfsedListTemp = orbitEntity.WorkflowStepExpandDefault.Where(p => p.WorkflowId == WorkflowId).OrderBy(p => p.RowNumber).ToList();

            foreach(var wfsed in wfsedListTemp)
            {
                Specification sp = orbitEntity.Specification.Where(p => p.SpecificationId == wfsed.SpecificationId ).FirstOrDefault();

                if(sp.WorkflowTypeId == "URC100000A7J" && sp.IsMetadataFlowStep == false)
                {
                    wfsedList.Add(wfsed);
                }
            }

            return wfsedList;
        }

        /// <summary>
        /// 判断当前时间是否是工作时间
        /// </summary>
        /// <param name="startTime"></param>
        /// <returns></returns>
        bool IsWorkTime(DateTime startTime)
        {
            return orbitEntity.WorkCenterCalendar.Where(p => DbFunctions.TruncateTime(p.NoWorkBeginDate) >= startTime || DbFunctions.TruncateTime(p.NoWorkEndDate) <= startTime).Any();
        }

        /// <summary>
        /// 返回工序计划里所有工序的distinct
        /// </summary>
        /// <returns></returns>
        List<string> GetSpecificationList()
        {
            List<string> specificationList = new List<string>();

            //Factory factory = orbitEntity.Factory.Where(p => p.FactoryId == this.FactoryID).First();

            //List<MO> moList = orbitEntity.MO.Where(p => p.FactoryId == this.FactoryID && p.MOStatus != "已强制关闭" && p.MOStatus != "完工").ToList();

            //List<MOItemTask> moitemList = new List<MOItemTask>();
            //foreach(var mo in moList)
            //{
            //    var list = orbitEntity.MOItemTask.Where(p => p.MOId == mo.MOId).ToList();
            //    moitemList.AddRange(list);
            //}

            //specificationList = moitemList.Where(p => p.TaskStatus != "已报工" && p.TaskStatus != "强制关闭").Select(p => p.SpecificationID).Distinct().ToList();

            specificationList = orbitEntity.Specification.Where(p => p.FactoryId == this.FactoryID && p.WorkflowTypeId == "URC100000A7J").Select(p => p.SpecificationId).ToList();
            //
            return specificationList;
        }

        /// <summary>
        /// 根据id获取MO
        /// </summary>
        /// <param name="MoIdList"></param>
        /// <returns></returns>
        List<MO> GetMoListById(List<string> MoIdList)
        {
            List<MO> moList = new List<MO>();

            foreach(var id in MoIdList)
            {
                MO mo = orbitEntity.MO.Where(p => p.MOId == id).FirstOrDefault();
                if (mo != null)
                    moList.Add(mo);
            }

            return moList.OrderBy(p => p.PlannedDateTo).ToList();
        }

        /// <summary>
        /// 根据Moid获取MoItemTaskList
        /// </summary>
        /// <param name="moId"></param>
        /// <returns></returns>
        List<MOItemTask> GetMOItemTaskListById(string moId)
        {
            List<MOItemTask> moItemTaskList = new List<MOItemTask>();

            moItemTaskList = orbitEntity.MOItemTask.Where(p => p.MOId == moId).ToList();

            return moItemTaskList;
        }

        /// <summary>
        /// 根据MoItemTaskid获取StepList 按工作流顺序排序
        /// </summary>
        /// <param name="MoId"></param>
        /// <returns></returns>
        List<MOItemTask> GetStepByMoItemTaskId(string mitId)
        {
            List<MOItemTask> stepList = new List<MOItemTask>();

            MOItemTask mit = orbitEntity.MOItemTask.Where(p => p.MOItemTaskId == mitId).FirstOrDefault();

            List<WorkflowStepExpandDefault> wfsedList = GetSpecificationOrderByMoId(mit.WorkflowId);

            foreach (var wfsed in wfsedList)
            {

                StepModel stepModel = stepModelList.Where(p => p.SpecificationId == wfsed.SpecificationId).FirstOrDefault();

                VentureModel ventureModel = stepModel.VentureModelList.Where(p => p.StepList.Where(a => a.MOItemTaskId == mitId).Any()).FirstOrDefault();

                if(ventureModel != null)
                {
                    MOItemTask mits = ventureModel.StepList.Where(a => a.MOItemTaskId == mitId).FirstOrDefault();

                    stepList.Add(mits);
                }
            }

            return stepList;
        }

        /// <summary>
        /// 根据Moid获取StepList 按工作流顺序排序
        /// </summary>
        /// <param name="MoId"></param>
        /// <returns></returns>
        public List<MOItemTask> GetStepByMoId(string moId)
        {
            List<MOItemTask> stepList = new List<MOItemTask>();

            MO mo = orbitEntity.MO.Where(p => p.MOId == moId).FirstOrDefault();

            if(mo.WorkflowId == null)
            {
                Product product = orbitEntity.Product.Where(p => p.ProductId == mo.ProductId).FirstOrDefault();
                mo.WorkflowId = product.WorkflowId;
            }

            List<WorkflowStepExpandDefault> wfsedList = GetSpecificationOrderByMoId(mo.WorkflowId);

            foreach (var wfsed in wfsedList)
            {

                //StepModel stepModel = stepModelList.Where(p => p.SpecificationId == wfsed.SpecificationId).FirstOrDefault();

                //VentureModel ventureModel = stepModel.VentureModelList.Where(p => p.StepList != null && p.StepList.Where(a => a.MOId == moId).Any()).FirstOrDefault();

                //if (ventureModel != null)
                //{
                //    MOItemTask mits = ventureModel.StepList.Where(a => a.MOId == moId).FirstOrDefault();

                //    if(mits == null)
                //    {
                //        mits = orbitEntity.MOItemTask.Where(p => p.VentureId == ventureModel.VentureId && p.SpecificationId == stepModel.SpecificationId && p.MOId == mo.MOId).FirstOrDefault();
                //    }
                //    mits = orbitEntity.MOItemTask.Where(p => p.VentureId == ventureModel.VentureId && p.SpecificationId == stepModel.SpecificationId && p.MOId == mo.MOId).FirstOrDefault();

                //    stepList.Add(mits);
                //}

                MOItemTask mits = orbitEntity.MOItemTask.Where(p => p.SpecificationID == wfsed.SpecificationId && p.MOId == mo.MOId).FirstOrDefault();
                if(mits != null)
                    stepList.Add(mits);
            }

            return stepList;
        }

        /// <summary>
        /// 根据产品代码和工序id获取工序标准工时
        /// </summary>
        /// <param name="ProductId"></param>
        /// <param name="SpecificationId"></param>
        /// <returns></returns>
        int GetStandardWorkingTimeByProductId_SpecificationId(string ProductId,string SpecificationId)
        {
            ProductSpecification ps = psList.Where(p => p.ProductId == ProductId && p.SpecificationId == SpecificationId).FirstOrDefault();

            if (ps != null && ps.StandardWorkingTime != null)
                return (int)ps.StandardWorkingTime;
            else
                return 0;
        }

        /// <summary>
        /// 根据工序ID获取工序对应的经营体ID
        /// </summary>
        /// <param name="SpecificationId"></param>
        /// <returns></returns>
        List<string> GetVentureBySpecificationId(string SpecificationId)
        {
            List<string> VentureIdList = new List<string>();

            VentureIdList = orbitEntity.Specification_Venture.Where(p => p.SpecificationId == SpecificationId && p.ResourcesAmount != 0 && p.DayCapacity != 0).Select(p => p.VentureId).ToList();

            return VentureIdList;
        }

        /// <summary>
        /// 寻找当前工序的上一个工序
        /// </summary>
        /// <param name="wfsed"></param>
        /// <returns>如果存在上一个工序则返回，否则返回null</returns>
        WorkflowStepExpandDefault FindPrevioudWorkflowStepExpandDefault(WorkflowStepExpandDefault wfsed)
        {
            WorkflowStepExpandDefault previousWfsed = null;

            int index = Convert.ToInt32(wfsed.RowNumber);

            int previousIndex = index - 1;

            string previousIndexString = previousIndex.ToString("00");

            previousWfsed = orbitEntity.WorkflowStepExpandDefault.Where(p => p.RowNumber == previousIndexString && p.WorkflowId == wfsed.WorkflowId).FirstOrDefault();

            if (previousWfsed != null)
            {
                Specification sp = orbitEntity.Specification.Where(p => p.SpecificationId == previousWfsed.SpecificationId).FirstOrDefault();

                if (sp.WorkflowTypeId != "URC100000A7J")
                {
                    previousWfsed = null;
                }

                if (previousWfsed != null)
                    return previousWfsed;
            }

            while (previousWfsed == null && previousIndex > 0)
            {
                previousIndex = previousIndex - 1;
                previousIndexString = previousIndex.ToString("00");
                previousWfsed = orbitEntity.WorkflowStepExpandDefault.Where(p => p.RowNumber == previousIndexString && p.WorkflowId == wfsed.WorkflowId).FirstOrDefault();

                if(previousWfsed != null)
                {
                    Specification sp = orbitEntity.Specification.Where(p => p.SpecificationId == previousWfsed.SpecificationId).FirstOrDefault();

                    if (sp.WorkflowTypeId != "URC100000A7J")
                    {
                        previousWfsed = null;
                    }

                    if (previousWfsed != null)
                        return previousWfsed;
                }
                
            }

            if (previousWfsed != null)
                return previousWfsed;
            else
                return null;
        }

        /// <summary>
        /// 寻找当前工序的下一个工序
        /// </summary>
        /// <param name="wfsed"></param>
        /// <returns>如果存在下一个工序则返回，否则返回null</returns>
        WorkflowStepExpandDefault FindNextWorkflowStepExpandDefault(WorkflowStepExpandDefault wfsed)
        {
            WorkflowStepExpandDefault nextWfsed = null;

            int index = Convert.ToInt32(wfsed.RowNumber);

            int previousIndex = index + 1;

            string previousIndexString = previousIndex.ToString("00");

            nextWfsed = orbitEntity.WorkflowStepExpandDefault.Where(p => p.RowNumber == previousIndexString).FirstOrDefault();

            if (nextWfsed != null)
                return nextWfsed;
            else
                return null;
        }

        /// <summary>
        /// 保存工序时间调整到数据库
        /// </summary>
        public void SaveChange(List<string> MoIdList)
        {
            List<MOItemTask> mitsList = new List<MOItemTask>();

            foreach (var stepModel in stepModelList)
            {

                foreach(var ventureModel in stepModel.VentureModelList)
                {

                    mitsList.AddRange(ventureModel.StepList);
                    //List<MOItemTask> mitsList = ventureModel.StepList;

                    //foreach (var mits in mitsList)
                    //{

                    //    var mits_in_database = orbitEntity.MOItemTask.Where(p => p.MOItemTaskId == mits.MOItemTaskId && p.MOId == mits.MOId).FirstOrDefault();
                    //    if (mits_in_database != null)
                    //    {
                    //        mits_in_database.PlannedDateFrom = mits.PlannedDateFrom;
                    //        mits_in_database.PlannedDateTo = mits.PlannedDateTo;

                    //        mits_in_database.VentureId = ventureModel.VentureId;

                    //        if(mits_in_database.TaskStatus == "已登记")
                    //        {
                    //            mits_in_database.TaskStatus = "已排产";
                    //            mits_in_database.TaskStatusId = "01";
                    //        }

                    //        orbitEntity.SaveChanges();
                    //    }



                    //}


                }
            }

            var list = mitsList.Where(p => p.TaskStatus == "被排产").ToList();

            //var db_list = new List<MOItemTask>();

            foreach (var mits in list)
            {

                var mits_in_database = orbitEntity.MOItemTask.Where(p => p.MOItemTaskId == mits.MOItemTaskId && p.MOId == mits.MOId).FirstOrDefault();
                if (mits_in_database != null)
                {
                    mits_in_database.PlannedDateFrom = mits.PlannedDateFrom;
                    mits_in_database.PlannedDateTo = mits.PlannedDateTo;

                    mits_in_database.VentureId = mits.VentureId;

                    if (mits_in_database.TaskStatus == "被排产")
                    {
                        mits_in_database.TaskStatus = "已排产";
                        mits_in_database.TaskStatusId = "01";
                    }

                    
                }
            }

            orbitEntity.SaveChanges();

            MO mo = orbitEntity.MO.Where(p => p.MOId == MoIdList.FirstOrDefault()).FirstOrDefault();

            List<MO> halfMoList = GetBomMoList(mo);

            foreach(var hm in halfMoList)
            {
                MoIdList.Add(hm.MOId);
            }

            foreach (var moid in MoIdList)
            {
                var mo_in_database = orbitEntity.MO.Where(p => p.MOId == moid).FirstOrDefault();
                if (mo_in_database != null)
                {
                    mo_in_database.MOStatus = "已排产";
                    mo_in_database.MOStatusId = "11";

                    List<MOItemTask> moitemtaskList = GetStepByMoId(mo_in_database.MOId);

                    mo_in_database.PlannedDateFrom = moitemtaskList.First().PlannedDateFrom;
                    mo_in_database.PlannedDateTo = moitemtaskList.Last().PlannedDateTo;


                    orbitEntity.SaveChanges();
                }
            }
        }

        /// <summary>
        /// 判断订单是否是产成品订单
        /// </summary>
        /// <param name="mo"></param>
        /// <returns></returns>
        public bool IsFinishedProduct(MO mo)
        {
            Product pro = orbitEntity.Product.Where(p => p.ProductId == mo.ProductId).FirstOrDefault();
            if (pro != null)
            {
                ProductRoot pr = orbitEntity.ProductRoot.Where(p => p.ProductRootId == pro.ProductRootId).FirstOrDefault();
                if (pr != null)
                {
                    ProductFamily pf = orbitEntity.ProductFamily.Where(p => p.ProductFamilyId == pr.ProductFamilyId).FirstOrDefault();
                    if (pf != null)
                    {
                        if (pf.ProductFamilyName.Contains("产成品"))
                            return true;
                        else
                            return false;
                    }
                    else
                        return false;
                }
                else
                    return false;

            }
            else
                return false;
        }

        /// <summary>
        /// 根据成品MO找到半成品MO
        /// </summary>
        /// <param name="mo"></param>
        /// <returns></returns>
        public List<MO> GetBomMoList(MO mo)
        {
            //List<MO> MoAllList = orbitEntity.MO.Where(p => p.MOStatus == "已审批").ToList();

            List<BanChengPinMO> banchengpingMOList = orbitEntity.BanChengPinMO.Where(p => p.FactoryId == this.FactoryID).ToList();

            List<MO> MoList = CovertMOList(banchengpingMOList);//new List<MO>();

            //foreach (var m in MoAllList)
            //{
            //    if (!IsFinishedProduct(m))
            //    {
            //        MoList.Add(m);
            //    }
            //}

            List<MO> BomMoList = new List<MO>();

            Product product = orbitEntity.Product.Where(p => p.ProductId == mo.ProductId).FirstOrDefault();
            //BOM bom = orbitEntity.BOM.Where(p => p.BOMId == product.BOMId).FirstOrDefault();
            List<BOMItem> bomItemList = orbitEntity.BOMItem.Where(p => p.BOMId == product.BOMId).ToList();

            foreach (var bomitem in bomItemList)
            {
                for(int i = 0; i < bomitem.QtyRequired;i++)
                {
                    MO halfMo = MoList.Where(p => p.ProductId == bomitem.ProductId).FirstOrDefault();
                    if(halfMo != null)
                        BomMoList.Add(halfMo);
                    MoList.Remove(halfMo);
                }
            }

            return BomMoList;
        }

        public List<MO> CovertMOList(List<ChanchengPinMO> ccpMoList)
        {
            List<MO> moList = new List<MO>();
            
            foreach(var ccpMo in ccpMoList)
            {
                MO mo = new MO();
                //mo.ApproveDate = ccpMo.ApproveDate;
                //mo.ApproveUserId = ccpMo.ApproveUserId;
                //mo.BatchCode = ccpMo.BatchCode;
                mo.BOMId = ccpMo.BOMId;
                mo.BOMName = ccpMo.BOMName;
                //mo.CloseDate = ccpMo.CloseDate;
                //mo.CreateDate = ccpMo.CreateDate;
                //mo.CreateUserId = ccpMo.CreateUserId;
                //mo.CreateUserName = ccpMo.CreateUserName;
                //mo.CustomerId = ccpMo.CustomerId;
                //mo.DeleteFlag = ccpMo.DeleteFlag;
                //mo.ERPDate = ccpMo.ERPDate;
                //mo.ERPId = ccpMo.ERPId;
                //mo.ERPNo = ccpMo.ERPNo;
                mo.ExecuteDateFrom = ccpMo.ExecuteDateFrom;
                mo.ExecuteDateTo = ccpMo.ExecuteDateTo;
                mo.FactoryId = ccpMo.FactoryId;
                //mo.IsPickingListUpload = ccpMo.IsPickingListUpload;
                //mo.ItemSequence = ccpMo.ItemSequence;
                //mo.MinPack = ccpMo.MinPack;
                //mo.MOClassId = ccpMo.MOClassId;
                //mo.MODescription = ccpMo.MODescription;
                //mo.ModifyDate = ccpMo.ModifyDate;
                mo.MOId = ccpMo.MOId;
                mo.MOName = ccpMo.MOName;
                //mo.MOQtyCurrent = ccpMo.MOQtyCurrent;
                //mo.MOQtyDone = ccpMo.MOQtyDone;
                //mo.MOQtyRequired = ccpMo.MOQtyRequired;
                mo.MOStatus = ccpMo.MOStatus;
                mo.MOTypeId = ccpMo.MOTypeId;
                mo.ParentMOId = ccpMo.ParentMOId;
                //mo.PickStatus = ccpMo.PickStatus;
                mo.PlannedDateFrom = ccpMo.PlannedDateFrom;
                mo.PlannedDateTo = ccpMo.PlannedDateTo;
                //mo.PriorityLevel = ccpMo.PriorityLevel;
                mo.ProductId = ccpMo.ProductId;
                //mo.ReleaseDate = ccpMo.ReleaseDate;
                //mo.SalaryCalcTypeId = ccpMo.SalaryCalcTypeId;
                //mo.SAP_ADD3 = ccpMo.SAP_ADD3;
                //mo.SAP_ADD4 = ccpMo.SAP_ADD4;
                //mo.SAP_AEDAT = ccpMo.SAP_AEDAT;
                //mo.SAP_AUART = ccpMo.SAP_AUART;
                //mo.SAP_BESKZ = ccpMo.SAP_BESKZ;
                //mo.SAP_BUKRS = ccpMo.SAP_BUKRS;
                //mo.SAP_CHARG = ccpMo.SAP_CHARG;
                //mo.SAP_GMEIN = ccpMo.SAP_GSBER;
                //mo.SAP_PHAS0 = ccpMo.SAP_PHAS0;
                //mo.SAP_PHAS1 = ccpMo.SAP_PHAS1;
                //mo.SAP_PHAS2 = ccpMo.SAP_PHAS2;
                //mo.SAP_PHAS3 = ccpMo.SAP_PHAS3;
                //mo.SAP_PSMNG = ccpMo.SAP_PSMNG;
                mo.SAP_VDATU = ccpMo.SAP_VDATU;
                //mo.SAP_VERID = ccpMo.SAP_VERID;
                //mo.SAP_WERKS = ccpMo.SAP_WERKS;
                //mo.SAP_Z0001 = ccpMo.SAP_Z0001;
                //mo.SAP_Z0002 = ccpMo.SAP_Z0002;
                //mo.SAP_Z0003 = ccpMo.SAP_Z0003;
                //mo.SAP_Z0006 = ccpMo.SAP_Z0006;
                //mo.SAP_Z0007 = ccpMo.SAP_Z0007;
                //mo.SAP_Z0008 = ccpMo.SAP_Z0008;
                //mo.SAP_Z0009 = ccpMo.SAP_Z0009;
                //mo.SAP_Z010 = ccpMo.SAP_Z010;
                //mo.SendMaterialDate = ccpMo.SendMaterialDate;
                //mo.SOId = ccpMo.SOId;
                //mo.SOName = ccpMo.SOName;
                //mo.StartWorkDate = ccpMo.StartWorkDate;
                mo.WorkflowId = ccpMo.WorkflowId;

                moList.Add(mo);
            }

            return moList;
        }

        public List<MO> CovertMOList(List<BanChengPinMO> ccpMoList)
        {
            List<MO> moList = new List<MO>();

            foreach (var ccpMo in ccpMoList)
            {
                MO mo = new MO();
                //mo.ApproveDate = ccpMo.ApproveDate;
                //mo.ApproveUserId = ccpMo.ApproveUserId;
                //mo.BatchCode = ccpMo.BatchCode;
                mo.BOMId = ccpMo.BOMId;
                mo.BOMName = ccpMo.BOMName;
                //mo.CloseDate = ccpMo.CloseDate;
                //mo.CreateDate = ccpMo.CreateDate;
                //mo.CreateUserId = ccpMo.CreateUserId;
                //mo.CreateUserName = ccpMo.CreateUserName;
                //mo.CustomerId = ccpMo.CustomerId;
                //mo.DeleteFlag = ccpMo.DeleteFlag;
                //mo.ERPDate = ccpMo.ERPDate;
                //mo.ERPId = ccpMo.ERPId;
                //mo.ERPNo = ccpMo.ERPNo;
                mo.ExecuteDateFrom = ccpMo.ExecuteDateFrom;
                mo.ExecuteDateTo = ccpMo.ExecuteDateTo;
                mo.FactoryId = ccpMo.FactoryId;
                //mo.IsPickingListUpload = ccpMo.IsPickingListUpload;
                //mo.ItemSequence = ccpMo.ItemSequence;
                //mo.MinPack = ccpMo.MinPack;
                //mo.MOClassId = ccpMo.MOClassId;
                //mo.MODescription = ccpMo.MODescription;
                //mo.ModifyDate = ccpMo.ModifyDate;
                mo.MOId = ccpMo.MOId;
                mo.MOName = ccpMo.MOName;
                //mo.MOQtyCurrent = ccpMo.MOQtyCurrent;
                //mo.MOQtyDone = ccpMo.MOQtyDone;
                //mo.MOQtyRequired = ccpMo.MOQtyRequired;
                mo.MOStatus = ccpMo.MOStatusId;
                mo.MOTypeId = ccpMo.MOTypeId;
                mo.ParentMOId = ccpMo.ParentMOId;
                //mo.PickStatus = ccpMo.PickStatus;
                mo.PlannedDateFrom = ccpMo.PlannedDateFrom;
                mo.PlannedDateTo = ccpMo.PlannedDateTo;
                //mo.PriorityLevel = ccpMo.PriorityLevel;
                mo.ProductId = ccpMo.ProductId;
                //mo.ReleaseDate = ccpMo.ReleaseDate;
                //mo.SalaryCalcTypeId = ccpMo.SalaryCalcTypeId;
                //mo.SAP_ADD3 = ccpMo.SAP_ADD3;
                //mo.SAP_ADD4 = ccpMo.SAP_ADD4;
                //mo.SAP_AEDAT = ccpMo.SAP_AEDAT;
                //mo.SAP_AUART = ccpMo.SAP_AUART;
                //mo.SAP_BESKZ = ccpMo.SAP_BESKZ;
                //mo.SAP_BUKRS = ccpMo.SAP_BUKRS;
                //mo.SAP_CHARG = ccpMo.SAP_CHARG;
                //mo.SAP_GMEIN = ccpMo.SAP_GSBER;
                //mo.SAP_PHAS0 = ccpMo.SAP_PHAS0;
                //mo.SAP_PHAS1 = ccpMo.SAP_PHAS1;
                //mo.SAP_PHAS2 = ccpMo.SAP_PHAS2;
                //mo.SAP_PHAS3 = ccpMo.SAP_PHAS3;
                //mo.SAP_PSMNG = ccpMo.SAP_PSMNG;
                mo.SAP_VDATU = ccpMo.SAP_VDATU;
                //mo.SAP_VERID = ccpMo.SAP_VERID;
                //mo.SAP_WERKS = ccpMo.SAP_WERKS;
                //mo.SAP_Z0001 = ccpMo.SAP_Z0001;
                //mo.SAP_Z0002 = ccpMo.SAP_Z0002;
                //mo.SAP_Z0003 = ccpMo.SAP_Z0003;
                //mo.SAP_Z0006 = ccpMo.SAP_Z0006;
                //mo.SAP_Z0007 = ccpMo.SAP_Z0007;
                //mo.SAP_Z0008 = ccpMo.SAP_Z0008;
                //mo.SAP_Z0009 = ccpMo.SAP_Z0009;
                //mo.SAP_Z010 = ccpMo.SAP_Z010;
                //mo.SendMaterialDate = ccpMo.SendMaterialDate;
                //mo.SOId = ccpMo.SOId;
                //mo.SOName = ccpMo.SOName;
                //mo.StartWorkDate = ccpMo.StartWorkDate;
                mo.WorkflowId = ccpMo.WorkflowId;

                moList.Add(mo);
            }

            return moList;
        }

        /// <summary>
        /// 根据UserId获取用户所在的工厂ID
        /// </summary>
        /// <param name="userID"></param>
        /// <returns></returns>
        public string GetUserFactory(string userID)
        {
            SysUser user = orbitEntity.SysUser.Where(p => p.UserId == userID).FirstOrDefault();
            return user.FactoryId;
        }

        /// <summary>
        /// 根据经营体id获取该经营体下甘特图数据
        /// </summary>
        /// <param name="ventureId"></param>
        /// <param name="manager"></param>
        /// <returns></returns>
        public List<MoItemStepModel> GetMoItemStopModelByVentureId(string ventureId, ProjectManager manager)
        {
            List<MoItemStepModel> moItemStepModelList = new List<MoItemStepModel>();

            if (stepModelList.Where(p => p.VentureModelList.Where(q => q.VentureId == ventureId && q.StepList.Any()).Any()).Any())
            {
                List<MOItemTask> stepList = stepModelList.Where(p => p.VentureModelList.Where(q => q.VentureId == ventureId).Any()).FirstOrDefault().VentureModelList.Where(q => q.VentureId == ventureId).FirstOrDefault().StepList.OrderBy( p => p.PlannedDateFrom).ToList();

                DateTime startTime = (DateTime)stepList.FirstOrDefault().PlannedDateFrom;

                foreach (var mits in stepList)
                {
                    MoItemStepModel mism = new MoItemStepModel(manager);

                    mism.MOId = mits.MOId;
                    mism.MOItemTaskStepId = mits.MOItemTaskId;
                    mism.PlannedDateFrom = mits.PlannedDateFrom != null ? ((DateTime)mits.PlannedDateFrom).ToString("yyyy-MM-dd HH:mm:ss") : "";
                    mism.PlannedDateTo = mits.PlannedDateTo != null ? ((DateTime)mits.PlannedDateTo).ToString("yyyy-MM-dd HH:mm:ss") : "";
                    MO mo = orbitEntity.MO.Where(p => p.MOId == mits.MOId).FirstOrDefault();
                    if (mo.SAP_VDATU != null)
                        mism.SapDeliveryDate = mo.SAP_VDATU;
                    else
                        mism.SapDeliveryDate = "";
                    mism.MOName = mo.MOName;

                    Specification sp = orbitEntity.Specification.Where(p => p.SpecificationId == mits.SpecificationID).FirstOrDefault();
                    if (sp != null)
                    {
                        SpecificationRoot sprt = orbitEntity.SpecificationRoot.Where(p => p.SpecificationRootId == sp.SpecificationRootId).FirstOrDefault();
                        if (sprt != null)
                            mism.Specification = sprt.SpecificationName;
                        else
                            mism.Specification = "";
                    }

                    Venture vt = orbitEntity.Venture.Where(p => p.VentureId == ventureId).FirstOrDefault();
                    if (vt != null)
                    {
                        mism.Venture = vt.VentureName;
                    }
                    else
                        mism.Venture = "";

                    Product pro = orbitEntity.Product.Where(p => p.ProductId == mits.ProductId).FirstOrDefault();
                    if (pro != null)
                    {
                        mism.ProductName = pro.ProductDescription;

                        ProductRoot pr = orbitEntity.ProductRoot.Where(p => p.ProductRootId == pro.ProductRootId).FirstOrDefault();
                        if (pr != null)
                        {
                            mism.ProductId = pr.ProductName;
                        }

                    }

                    mism.TaskStatus = mits.TaskStatus;

                    manager.Add(mism);

                    mism.Name = mism.ProductName;
                    var timespan1 = mits.PlannedDateFrom.Value - startTime;
                    mism.Start = timespan1;
                    var timespan2 = mits.PlannedDateTo.Value - startTime;
                    mism.End = timespan2;
                    mism.Duration = mism.End - mism.Start;

                    moItemStepModelList.Add(mism);
                }
            }

            return moItemStepModelList;
        }

        /// <summary>
        /// 根据工厂获取生产经营体
        /// </summary>
        /// <returns></returns>
        public List<Venture> GetVentureByFactory()
        {
            List<Venture> venturelist = new List<Venture>();
            if (FactoryID == "FAC1000000SF")
                venturelist = orbitEntity.Venture.Where(p => p.FactoryId == FactoryID && p.WorkcenterId != null && p.WorkcenterId != "").ToList();
            else if(FactoryID == "FAC1000000SG")
                venturelist = orbitEntity.Venture.Where(p => p.FactoryId == FactoryID).ToList();

            return venturelist;
        } 

        /// <summary>
        /// 获取对应经营体能力
        /// </summary>
        /// <param name="specificationId"></param>
        /// <param name="ventureId"></param>
        /// <returns></returns>
        public double GetVentureCapById(string specificationId,string ventureId)
        {
            var sv = orbitEntity.Specification_Venture.Where(p => p.SpecificationId == specificationId && p.VentureId == ventureId).FirstOrDefault();

            double cap = 0;

            if(sv != null)
            {
                try
                {
                    cap = (int)sv.ResourcesAmount * (int)sv.DayCapacity;
                }
                catch(Exception e)
                {
                    throw new Exception(specificationId + "工序下的" + ventureId + "经营体能力未维护，请维护后再试！");
                }
            }

            return cap;
        }

        public string ExcelImport(List<MOItemTask> mitList)
        {

            foreach(var mit in mitList)
            {
                var mit_db = orbitEntity.MOItemTask.Where(p => p.MOItemTaskName == mit.MOItemTaskName).FirstOrDefault();

                if(mit_db == null)
                {
                    return "工序ID为：" + mit.MOItemTaskId + "的工序计划在系统中未能找到，请检查后重新导入！";
                }
                else
                {
                    mit_db.PlannedDateFrom = mit.PlannedDateFrom;
                    mit_db.PlannedDateTo = mit.PlannedDateTo;
                    mit_db.VentureId = mit.VentureId;
                    mit_db.TaskStatusId = "01";
                    mit_db.TaskStatus = "已排产";

                    orbitEntity.SaveChanges();
                }
            }

            return "导入成功";
        }

        public string GetProcessListTimeText(string moid)
        {
            List<MOItemTask> mitList = GetStepByMoId(moid);

            string text = "";

            foreach (var mit in mitList)
            {
                Specification sp = orbitEntity.Specification.Where(p => p.SpecificationId == mit.SpecificationID).FirstOrDefault();
                SpecificationRoot sprt2 = orbitEntity.SpecificationRoot.Where(p => p.SpecificationRootId == sp.SpecificationRootId).FirstOrDefault();

                if (sprt2.SpecificationName.Length >= 4)
                    text += sprt2.SpecificationName.Substring(0, 4);
                else
                    text += sprt2.SpecificationName;

                text += ":" + mit.PlannedDateFrom.Value.ToString("yyyy-MM-dd") + "\t\t";
            }
            return text;
        }

        public List<ReportModel> GetSpecificationCapReport(string moid)
        {
            List<MOItemTask> mitList = GetStepByMoId(moid);

            List<ReportModel> reportModelList = new List<ReportModel>();


            foreach (var mit in mitList)
            {
                ReportModel reportModel = new ReportModel("周");

                double moPlnFinish = GetSpecificationCapValue((DateTime)mit.PlannedDateFrom, null, mit.SpecificationID) - GetSpecificationRealValue((DateTime)mit.PlannedDateFrom, mit.PlannedDateTo, mit.SpecificationID);
                double moRealFinish = GetSpecificationRealValue((DateTime)mit.PlannedDateFrom, mit.PlannedDateTo, mit.SpecificationID);
                double moRate = GetStandardWorkingTimeByProductId_SpecificationId(mit.ProductId, mit.SpecificationID);


                Specification sp = orbitEntity.Specification.Where(p => p.SpecificationId == mit.SpecificationID).FirstOrDefault();
                SpecificationRoot sprt2 = orbitEntity.SpecificationRoot.Where(p => p.SpecificationRootId == sp.SpecificationRootId).FirstOrDefault();

                if (sprt2.SpecificationName.Length >= 4)
                    reportModel.Name = sprt2.SpecificationName.Substring(0, 4);
                else
                    reportModel.Name = sprt2.SpecificationName;
                reportModel.Plan = moPlnFinish;
                reportModel.Real = moRealFinish;
                reportModel.Rate = moRate;

                reportModelList.Add(reportModel);
            }

            return reportModelList;
        }

        public double GetSpecificationCapValue(DateTime date, DateTime? endDate, string specificationId)
        {
            double scv = 0;
            if (endDate != null)
            {
                TimeSpan ts = (DateTime)endDate - date.AddMinutes(-10);
                int day = ts.Days;

                scv = GetSpecifacationCap(specificationId) * day;
            }
            else
            {
                scv = GetSpecifacationCap(specificationId) * 1;
            }

            return scv;
        }

        public double GetSpecificationRealValue(DateTime date, DateTime? endDate, string specificationId)
        {
            double realvalue = 0;
            List<MOItemTask> moitemList = new List<MOItemTask>();
            if (endDate != null)
            {
                List<MOItemTask> moitemtempList = orbitEntity.MOItemTask.Where(p => p.SpecificationID == specificationId && p.PlannedDateTo >= date && p.PlannedDateTo <= endDate).ToList();

                foreach (var moitem in moitemtempList)
                {
                    MO mo = orbitEntity.MO.Where(p => p.MOId == moitem.MOId).FirstOrDefault();
                    if (mo != null && mo.FactoryId == this.FactoryID)
                    {
                        moitemList.Add(moitem);
                    }
                }

            }
            else
            {
                DateTime end = date.AddDays(1);
                List<MOItemTask> moitemtempList = orbitEntity.MOItemTask.Where(p => p.SpecificationID == specificationId && p.PlannedDateTo >= date && p.PlannedDateTo <= end).ToList();

                foreach (var moitem in moitemtempList)
                {
                    MO mo = orbitEntity.MO.Where(p => p.MOId == moitem.MOId).FirstOrDefault();
                    if (mo != null && mo.FactoryId == this.FactoryID)
                    {
                        moitemList.Add(moitem);
                    }
                }

            }

            foreach (var moitem in moitemList)
            {
                realvalue += GetStandardWorkingTimeByProductId_SpecificationId(moitem.ProductId, moitem.SpecificationID);
            }

            return Math.Round(realvalue / 60, 2);
        }

        public double GetSpecifacationCap(string specificationId)
        {
            double specifacationCap = 0;
            List<Specification_Venture> svlist = orbitEntity.Specification_Venture.Where(p => p.SpecificationId == specificationId).ToList();

            foreach (var sv in svlist)
            {
                if (sv != null && sv.ResourcesAmount != null && sv.DayCapacity != null)
                {
                    specifacationCap += (int)sv.ResourcesAmount * (int)sv.DayCapacity;
                }
            }

            return specifacationCap;
        }
    }

}
