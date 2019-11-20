using PlanSource.Model;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ScheduleCore.Model
{
    public class StepModel
    {
        static string connectString = "metadata=res://*/Model.PlanModel.csdl|res://*/Model.PlanModel.ssdl|res://*/Model.PlanModel.msl;provider=System.Data.SqlClient;provider connection string=\"data source=192.168.0.54;initial catalog=OrBitMOM_Kocel;persist security info=True;user id=wangjunjie;password=123456;MultipleActiveResultSets=True;App=EntityFramework\"";

        public string SpecificationId { get; set; }
        public List<VentureModel> VentureModelList { get; set; }
        OrBitMOM_KocelEntities1 orbitdb = new OrBitMOM_KocelEntities1(connectString);

        /// <summary>
        /// 添加Step到开始时间最靠前的经营体List里
        /// </summary>
        /// <param name="step"></param>
        /// <param name="standardWorkingTime">标准工时</param>
        /// <param name="lastTime">最后时间，如果不为null，放置在List中的位置必须在此时间之后</param>
        public DateTime? AddStep(MOItemTask step,double standardWorkingTime,DateTime? lastTime,bool isFirstStep,bool isFuji)
        {
            DateTime? EndTime = null;

            if (lastTime != null && !isFirstStep && isFuji)
            {
                lastTime = new DateTime(lastTime.Value.Year, lastTime.Value.Month, lastTime.Value.Day, 9, 0, 0);
                lastTime = lastTime.Value.AddDays(1);
            }


            List<string> ventureIdList = GetVentureIDBYProductFamily(step.ProductId,step.SpecificationID);

            ////寻找时间片满足的经营体
            //VentureModel ventureModel = null;
            //if(ventureIdList.Any())
            //{
            //    ventureModel = this.VentureModelList.Where(p => ventureIdList.Contains(p.VentureId) && p.VenTureDayModelList.Where(a => a.Timeslice > standardWorkingTime && a.LastCap >= standardWorkingTime).Any()).FirstOrDefault();
            //}
            //else
            //{
            //    ventureModel = this.VentureModelList.Where(p => p.VenTureDayModelList.Where(a => a.Timeslice > standardWorkingTime && a.LastCap >= standardWorkingTime).Any()).FirstOrDefault();                
            //}

            ////存在于时间片的经营体 就把工序插入
            //if(ventureModel != null)
            //{
            //    var vdm = ventureModel.VenTureDayModelList.Where(p => p.TimesliceStartTime > lastTime && p.LastCap >= standardWorkingTime).FirstOrDefault();

            //    var PlannedDateFrom = GetWorkTime((DateTime)vdm.TimesliceStartTime, ventureModel.VentureId);
            //    EndTime = ventureModel.AddStep(step, PlannedDateFrom, standardWorkingTime);
            //}

            ////寻找最后时间靠前的经营体
            //VentureModel ventureModel2 = null;
            //if (ventureIdList.Any())
            //{
            //    ventureModel2 = this.VentureModelList.Where(p => ventureIdList.Contains(p.VentureId) && p.VenTureDayModelList.Where(a => !a.StepList.Any()).Any()).FirstOrDefault();
            //    if (ventureModel2 == null)
            //        ventureModel2 = this.VentureModelList.Where(p => ventureIdList.Contains(p.VentureId)).OrderBy(p => p.VenTureDayModelList.OrderBy(a=> a.LastEndTime)).FirstOrDefault();                    
            //}
            //else
            //{
            //    ventureModel2 = this.VentureModelList.Where(p =>  p.VenTureDayModelList.Where(a => !a.StepList.Any()).Any()).FirstOrDefault();
            //    if (ventureModel2 == null)
            //        ventureModel2 = this.VentureModelList.OrderBy(p => p.VenTureDayModelList.OrderBy(a => a.LastEndTime)).FirstOrDefault();
            //}

            //if(ventureModel2 != null)
            //{
            //    var vdm = ventureModel2.VenTureDayModelList.Where(p => p.LastEndTime > lastTime).FirstOrDefault();
            //    if (vdm != null)
            //    {
            //        var PlannedDateFrom = GetWorkTime((DateTime)vdm.LastEndTime, ventureModel2.VentureId);

            //        EndTime = ventureModel2.AddStep(step, PlannedDateFrom, standardWorkingTime);
            //    }
            //    else
            //    {
            //        var PlannedDateFrom = GetWorkTime((DateTime)lastTime, ventureModel2.VentureId);

            //        EndTime = ventureModel2.AddStep(step, PlannedDateFrom, standardWorkingTime);
            //    }
            //}

            VentureModel ventureModel2 = null;
            if (ventureIdList.Any())
            {
                var minPlanCapCount = this.VentureModelList.Where(p => ventureIdList.Contains(p.VentureId)).Min( p => p.PlanCapCount);
                ventureModel2 = this.VentureModelList.Where(p => ventureIdList.Contains(p.VentureId) && p.PlanCapCount == minPlanCapCount).FirstOrDefault();
            }
            else
            {
                var minPlanCapCount = this.VentureModelList.Min(p => p.PlanCapCount);
                ventureModel2 = this.VentureModelList.Where(p => p.PlanCapCount == minPlanCapCount).FirstOrDefault();
            }

            if (ventureModel2 != null)
            {

                var PlannedDateFrom = GetWorkTime((DateTime)lastTime, ventureModel2.VentureId);

                EndTime = ventureModel2.AddStep(step, PlannedDateFrom, standardWorkingTime);

            }


            if (ventureModel2 == null)
            {
                throw new Exception(SpecificationId + "工序未绑定经营体，请检查！");
            }

            //首先比较时间片，如果不满足就比较最后时间，加入满足的经营体
            //if (ventureModel != null)
            //{
            //    if (lastTime != null)
            //    {
            //        if(ventureModel.VenTureDayModelList.Where(p => p.TimesliceStartTime > lastTime && p.LastCap >= standardWorkingTime).Any())
            //        {
                        
            //            step.PlannedDateFrom = GetWorkTime((DateTime)ventureModel.TimesliceEndTime, ventureModel.VentureId);
            //            DateTime temp = ((DateTime)step.PlannedDateFrom).AddMinutes(standardWorkingTime);
            //            step.PlannedDateTo = temp;

            //            EndTime = ventureModel.AddStep(step, ventureModel.VenTureDayModelList.Where(p => p.TimesliceStartTime > lastTime && p.LastCap >= standardWorkingTime).FirstOrDefault().TimesliceStartTime,standardWorkingTime);
            //        }
            //        else
            //        {
            //            if (ventureModel2.endTime > lastTime)
            //            {
            //                step.PlannedDateFrom = GetWorkTime((DateTime)ventureModel2.endTime, ventureModel2.VentureId);
            //                DateTime temp = ((DateTime)step.PlannedDateFrom).AddMinutes(standardWorkingTime);
            //                step.PlannedDateTo = temp;
            //            }
            //            else
            //            {
            //                step.PlannedDateFrom = GetWorkTime((DateTime)lastTime, ventureModel2.VentureId);
            //                DateTime temp = ((DateTime)step.PlannedDateFrom).AddMinutes(standardWorkingTime);
            //                step.PlannedDateTo = temp;
            //            }

            //            EndTime = ventureModel2.AddStep(step);
            //        }
            //    }
            //    else
            //    {
            //        step.PlannedDateFrom = GetWorkTime((DateTime)ventureModel.TimesliceEndTime, ventureModel.VentureId);
            //        DateTime temp = ((DateTime)step.PlannedDateFrom).AddMinutes(standardWorkingTime);
            //        step.PlannedDateTo = temp;

            //        EndTime = ventureModel.AddStep(step);
            //    }
            //}
            //else
            //{
               
            //    if (lastTime != null)
            //    {
            //        if (ventureModel2.endTime > lastTime)
            //        {
            //            step.PlannedDateFrom = GetWorkTime((DateTime)ventureModel2.endTime, ventureModel2.VentureId);
            //            DateTime temp = ((DateTime)step.PlannedDateFrom).AddMinutes(standardWorkingTime);
            //            step.PlannedDateTo = temp;
            //        }
            //        else
            //        {
            //            step.PlannedDateFrom = GetWorkTime((DateTime)lastTime, ventureModel2.VentureId);
            //            DateTime temp = ((DateTime)step.PlannedDateFrom).AddMinutes(standardWorkingTime);
            //            step.PlannedDateTo = temp;
            //        }
            //    }
            //    else
            //    {
            //        if(ventureModel2.endTime != null)
            //        {
            //            step.PlannedDateFrom = GetWorkTime((DateTime)ventureModel2.endTime, ventureModel2.VentureId);
            //            DateTime temp = ((DateTime)step.PlannedDateFrom).AddMinutes(standardWorkingTime);
            //            step.PlannedDateTo = temp;
            //        }
            //        else
            //        {
            //            step.PlannedDateFrom = GetWorkTime(DateTime.Now, ventureModel2.VentureId);
            //            DateTime temp = ((DateTime)step.PlannedDateFrom).AddMinutes(standardWorkingTime);
            //            step.PlannedDateTo = temp;
            //        }
                    
            //    }

            //    EndTime = ventureModel2.AddStep(step);
            //}


            return EndTime;
        }

        /// <summary>
        /// 查找该工序下是否存在MoId一至的任务单
        /// </summary>
        /// <param name="MoId"></param>
        /// <returns>如果存在返回计划结束时间，否则返回null</returns>
        public DateTime? FindStepEndTime(string MoId)
        {
            foreach(var venture in this.VentureModelList)
            {
                var step = venture.StepList.Where(p =>  p.MOId == MoId).FirstOrDefault();
                if (step != null)
                {
                    return step.PlannedDateTo;
                }
            }

            return null;
        }


        /// <summary>
        /// 获取该工序下所有经营体计划的和
        /// </summary>
        /// <returns></returns>
        public List<MOItemTask> GetAllMOItemTaskStep()
        {
            List<MOItemTask> mits = new List<MOItemTask>();

            foreach (var ventureModel in VentureModelList)
            {
                foreach(var v in ventureModel.VenTureDayModelList)
                {
                    mits.AddRange(v.StepList);
                }
                
            }

            return mits;
        }


        ///// <summary>
        ///// 找到Moitemid对应的计划及后续计划移除出经营体列表，并返回该计划列表
        ///// </summary>
        ///// <param name="mitsId"></param>
        ///// <returns></returns>
        //public List<MOItemTask> RemoveMOItemTaskById(string mitsId)
        //{
        //    List<MOItemTask> mitsList = new List<MOItemTask>();

        //    foreach(var ventureModel in VentureModelList)
        //    {
        //        mitsList = ventureModel.RemoveMOItemTaskById(mitsId);
        //        if(mitsList.Any())
        //        {
        //            return mitsList;
        //        }
        //    }

        //    return null;
        //}

        /// <summary>
        /// 判断当前时间是否是工作时间
        /// </summary>
        /// <param name="startTime"></param>
        /// <returns></returns>
        DateTime GetWorkTime(DateTime startTime,string ventureId)
        {
            WorkCenterCalendar noWorkCalender = orbitdb.WorkCenterCalendar.Where(p => DbFunctions.TruncateTime(p.NoWorkBeginDate) <= startTime && DbFunctions.TruncateTime(p.NoWorkEndDate) >= startTime).FirstOrDefault();

            DateTime workTime;

            if (noWorkCalender != null)
                workTime = (DateTime)noWorkCalender.NoWorkEndDate;
            else
                workTime = startTime;

            List<Venture_Shift> vfList = orbitdb.Venture_Shift.Where(p => p.VentureId == ventureId).ToList();



            if (!vfList.Any())
            {
                Venture v = orbitdb.Venture.Where(p => p.VentureId == ventureId).FirstOrDefault();
                throw new Exception(v.VentureName + "(" + v.VentureId + ")" + "经营体与工厂日历关系未维护！");
            }

            var shiftIdList = vfList.Select(p => p.ShiftId).ToList();


            var shiftList = orbitdb.Shift.Where(p => shiftIdList.Contains(p.ShiftId)).ToList();

            var minTimeFrom = shiftList.Max(p => p.TimeFrom);

            Shift shift = shiftList.Where(p => p.TimeFrom == minTimeFrom).FirstOrDefault();

            if (shift == null)
                throw new Exception("工厂日历未维护！");

            string[] minsecond = shift.TimeFrom.Split(new char[] { ':' });

            DateTime workStartTime = new DateTime(startTime.Year, startTime.Month, startTime.Day, int.Parse(minsecond[0]), int.Parse(minsecond[1]), 0);
            DateTime workEndTime = workStartTime.AddMinutes(Convert.ToDouble(shift.TimeRange));

            if(workTime < workStartTime || workTime > workEndTime)
            {
                workTime = workStartTime.AddDays(1);

                noWorkCalender = orbitdb.WorkCenterCalendar.Where(p => DbFunctions.TruncateTime(p.NoWorkBeginDate) >= workTime || DbFunctions.TruncateTime(p.NoWorkEndDate) <= workTime).FirstOrDefault();

                if (noWorkCalender != null)
                {
                    TimeSpan? ts = noWorkCalender.NoWorkEndDate - noWorkCalender.NoWorkBeginDate;
                    int day = ((TimeSpan)ts).Days;

                    workTime = workStartTime.AddDays(day+1);
                }                                               
            }            

            //return 
            return workTime;
        }

        /// <summary>
        /// 根据产品代码获取产品族对应的经营体ID
        /// </summary>
        /// <param name="ProductId"></param>
        /// <returns></returns>
        List<string> GetVentureIDBYProductFamily(string ProductId,string specificationId)
        {
            List<ProductRoot> productRootList = orbitdb.ProductRoot.Where(p => p.DefaultProductId == ProductId).ToList();

            string productfamiliId = productRootList.FirstOrDefault().ProductFamilyId;

            List<VentureProductFamily> vpfList = orbitdb.VentureProductFamily.Where(p => p.ProductFamilyId == productfamiliId).ToList();

            List<string> vpfVentureIdList = new List<string>();

            if (vpfList.Any())
            {
                vpfVentureIdList = vpfList.Select(p => p.VentureId).ToList();

                //for(int i =0;i< vpfVentureIdList.Count;i++)
                //{
                //    Specification_Venture sv = orbitdb.Specification_Venture.Where(p => p.VentureId == vpfVentureIdList[i] && p.ResourcesAmount == 0 && p.DayCapacity == 0).FirstOrDefault();
                //    if (sv != null)
                //        vpfVentureIdList.Remove(vpfVentureIdList[i]);
                //}
            }
            else
            {
                vpfVentureIdList = new List<string>();
            }

            List<Specification_Venture> svList = orbitdb.Specification_Venture.Where(p => vpfVentureIdList.Contains(p.VentureId) && p.SpecificationId == specificationId && p.DayCapacity != 0 && p.ResourcesAmount != 0).ToList();

            if (svList.Any())
                return svList.Select(p => p.VentureId).ToList();
            else
                return new List<string>();
        }
    }
}
