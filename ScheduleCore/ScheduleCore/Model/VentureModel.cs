using PlanSource.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.Entity;

namespace ScheduleCore.Model
{
    public class VentureModel
    {

        public VentureModel(string SpecificationId,string VentureId,double ventureCap, List<MOItemTask> StepList)
        {
            this.SpecificationId = SpecificationId;
            this.VentureId = VentureId;
            this.VentureCap = ventureCap * 60;
            this.StepList = StepList;

            this.VenTureDayModelList = new List<VenTureDayModel>();

            SetPlanCapCount();
            DispatchStep();
            //SetTimeslice();
        }

        public string SpecificationId { get; set; }
        public string VentureId { get; set; }
        //public DateTime? endTime { get; set; }
        //public double Timeslice { get; set; }
        //public DateTime? TimesliceEndTime { get; set; }
        public List<MOItemTask> StepList { get; set; }

        public List<VenTureDayModel> VenTureDayModelList { get; set; }

        ///以经营体能力排产

        /// <summary>
        ///经营体能力，定值
        /// </summary>
        public double VentureCap { get; set; }

        public double PlanCapCount { get; set; }

        ///// <summary>
        ///// 已经排到的最后一天
        ///// </summary>
        //public DateTime? CapEndTime { get; set; }
        ///// <summary>
        ///// 最后一天剩余的能力值
        ///// </summary>
        //public double LastCap { get; set; }

        ///// <summary>
        ///// 添加Step到数组中，并记录最后的完成时间
        ///// </summary>
        ///// <param name="step"></param>
        //public DateTime? AddStep(MOItemTask step)
        //{
        //    DateTime? EndTime = null ;
        //    if (!StepList.Where(p => p.MOItemTaskId == step.MOItemTaskId).Any())
        //    {
        //        StepList.Add(step);

        //        this.endTime = step.PlannedDateTo;

        //        EndTime = this.endTime;
        //    }

        //    SetTimeslice();

        //    return EndTime;
        //}

        //public void AddStep(MOItemTask step,double cap)
        //{
        //    double sum = LastCap + cap;
        //    double maxcap = VentureCap * 1.2;
        //    if (sum <= maxcap)
        //    {
        //        LastCap = VentureCap - sum;                
        //    }
        //    else
        //    {
        //        LastCap = VentureCap - cap;
        //        if(CapEndTime == null)
        //        {
        //            CapEndTime = DateTime.Now;
        //        }

        //        do
        //        {
        //            CapEndTime = CapEndTime.Value.AddDays(1);
        //        }
        //        while (IsWorkTime(CapEndTime.Value));
        //    }

        //    step.PlannedDateFrom = new DateTime(CapEndTime.Value.Year, CapEndTime.Value.Month, CapEndTime.Value.Day, 8, 0, 0);
        //    step.PlannedDateTo = new DateTime(CapEndTime.Value.Year, CapEndTime.Value.Month, CapEndTime.Value.Day, 17, 0, 0);


        //    StepList.Add(step);
        //}

        public DateTime? AddStep(MOItemTask step, DateTime? lastTime, double cap)
        {
            if(lastTime != null)
            {
                var workTime = VenTureDayModel.GetWorkTime((DateTime)lastTime, VentureId);

                var venTureDayModel = this.VenTureDayModelList.Where(p => p.DayDate.Year == workTime.Year && p.DayDate.Month == workTime.Month && p.DayDate.Day == workTime.Day && p.LastCap > 0).FirstOrDefault();               

                if (venTureDayModel == null)
                {
                    this.VenTureDayModelList.Add(new VenTureDayModel(workTime, VentureCap, new List<MOItemTask>()));

                    venTureDayModel = this.VenTureDayModelList.Where(p => p.DayDate >= lastTime).FirstOrDefault();
                }

                if (venTureDayModel.LastCap < cap && VentureCap * 0.25 < cap)
                {

                }

                var retModel = venTureDayModel.AddStep(step, cap, lastTime,this.VentureId);

                StepList.Add(retModel.MOItemTask);

                return retModel.EndTime;
            }
            else
            {
                return null;
            }
        }

        public void DispatchStep()
        {
            foreach(var mit in StepList)
            {
                var vdm = VenTureDayModelList.Where(p => p.Year == mit.PlannedDateFrom.Value.Year && p.Month == mit.PlannedDateFrom.Value.Month && p.Day == mit.PlannedDateFrom.Value.Day).FirstOrDefault();
                if (vdm == null)
                {
                    vdm = new VenTureDayModel((DateTime)mit.PlannedDateFrom, this.VentureCap, new List<MOItemTask>());

                    this.VenTureDayModelList.Add(vdm);

                    vdm.StepList.Add(mit);
                }
                else
                {
                    vdm.StepList.Add(mit);
                }
            }

            foreach(var vdm in this.VenTureDayModelList)
            {
                vdm.SetTimeslice();
            }
        }

        public void SetPlanCapCount()
        {
            this.PlanCapCount = 0;

            foreach(var mit in this.StepList)
            {
                int cap = VenTureDayModel.GetStandardWorkingTimeByProductId_SpecificationId(mit.ProductId, mit.SpecificationID);

                this.PlanCapCount += cap;
            }
        }

        //public List<MOItemTask> GetVentureStepList()
        //{
        //    var ventureStepList = new List<MOItemTask>();

        //    foreach(var v in VenTureDayModelList)
        //    {
        //        ventureStepList.AddRange(v.StepList);
        //    }

        //    return ventureStepList;
        //}

        ///// <summary>
        ///// 返回经营体是否当天能排入这个计划
        ///// </summary>
        ///// <param name="step"></param>
        ///// <param name="cap"></param>
        ///// <returns></returns>
        //public bool isCanPayload(double cap)
        //{
        //    double sum = LastCap + cap;
        //    double maxcap = VentureCap * 1.2;
        //    if (sum <= maxcap)
        //    {
        //        return true;
        //    }
        //    else
        //    {
        //        return false;
        //    }
        //}

        ///// <summary>
        ///// 计算List中最大的时间片
        ///// </summary>
        //private void SetTimeslice()
        //{
        //    //List中至少有连个对象才能计算时间片
        //    if(this.StepList.Any() && this.StepList.Count>1)
        //    {
        //        for(int i=0;i<this.StepList.Count-1;i++)
        //        {
        //            DateTime? endTime = this.StepList[i].PlannedDateTo;
        //            DateTime? startTime = this.StepList[i+1].PlannedDateFrom;

        //            if(endTime != null && startTime != null )
        //            {
        //                TimeSpan time = (DateTime)startTime - (DateTime)endTime;
        //                double temptimeslice = time.TotalMinutes;
        //                if (Timeslice< temptimeslice)
        //                {
        //                    Timeslice = temptimeslice;
        //                    TimesliceEndTime = endTime;
        //                }                      
        //            }
        //        }
        //    }
        //    else
        //    {
        //        this.Timeslice = 0;
        //    }
        //}

        ///// <summary>
        ///// 找到Moitemid对应的计划及后续计划移除出经营体列表，并返回该计划列表
        ///// </summary>
        ///// <param name="mitsId"></param>
        ///// <returns></returns>
        //public List<MOItemTask> RemoveMOItemTaskById(string mitsId)
        //{
        //    List<MOItemTask> returnMitsList = new List<MOItemTask>();

        //    if (StepList.Where(p => p.MOItemTaskId == mitsId).Any())
        //    {
        //        MOItemTask step = StepList.Where(p => p.MOItemTaskId == mitsId).FirstOrDefault();

        //        int index = StepList.IndexOf(step);

        //        returnMitsList = StepList.Skip(index).Take(StepList.Count - index).ToList();

        //        //排除不是已登记状态的计划
        //        for (int i= 0;i < returnMitsList.Count;i++)
        //        {
        //            if (returnMitsList[i].TaskStatus != "已登记")
        //                returnMitsList.Remove(returnMitsList[i]);
        //        }

        //        //StepList.RemoveRange(index, StepList.Count - index);
        //        foreach(var returnMits in returnMitsList)
        //        {
        //            StepList.Remove(returnMits);
        //        }
        //    }

        //    //重新计划时间片
        //    SetTimeslice();

        //    return returnMitsList;
        //}

        ///// <summary>
        ///// 判断当前时间是否是工作时间
        ///// </summary>
        ///// <param name="startTime"></param>
        ///// <returns></returns>
        //bool IsWorkTime(DateTime startTime)
        //{
        //    return orbitdb.WorkCenterCalendar.Where(p => DbFunctions.TruncateTime(p.NoWorkBeginDate) >= startTime || DbFunctions.TruncateTime(p.NoWorkEndDate) <= startTime).Any();
        //}


        public class VenTureDayModel
        {

            public VenTureDayModel(DateTime day, double ventureCap, List<MOItemTask> milist)
            {
                this.DayDate = day;
                this.Year = this.DayDate.Year;
                this.Month = this.DayDate.Month;
                this.Day = this.DayDate.Day;

                this.VentureCap = ventureCap;
                this.LastCap = this.VentureCap;

                this.StepList = milist;
            }

            /// <summary>
            /// 年
            /// </summary>
            public int Year { get; set; }

            /// <summary>
            /// 月
            /// </summary>
            public int Month { get; set; }

            /// <summary>
            /// 日
            /// </summary>
            public int Day { get; set; }

            /// <summary>
            /// 所在天时间的Date
            /// </summary>
            public DateTime DayDate { get; set; }

            /// <summary>
            ///经营体能力，定值
            /// </summary>
            public double VentureCap { get; set; }

            /// <summary>
            /// 最后的结束时间
            /// </summary>
            public DateTime? LastEndTime { get; set; }

            /// <summary>
            /// 间隔的时间片
            /// </summary>
            public double Timeslice { get; set; }

            /// <summary>
            /// 时间片开始时间
            /// </summary>
            public DateTime? TimesliceStartTime { get; set; }

            /// <summary>
            /// 最后一天剩余的能力值
            /// </summary>
            public double LastCap { get; set; }

            /// <summary>
            /// 排产的工序列表
            /// </summary>
            public List<MOItemTask> StepList { get; set; }


            /// <summary>
            /// 计算List中最大的时间片
            /// </summary>
            public void SetTimeslice()
            {
                this.LastEndTime = StepList.Max(p => p.PlannedDateTo);

                //List中至少有连个对象才能计算时间片
                if (this.StepList.Any() && this.StepList.Count > 1)
                {
                    for (int i = 0; i < this.StepList.Count - 1; i++)
                    {
                        DateTime? endTime = this.StepList[i].PlannedDateTo;
                        DateTime? startTime = this.StepList[i + 1].PlannedDateFrom;

                        if (endTime != null && startTime != null)
                        {
                            TimeSpan time = (DateTime)startTime - (DateTime)endTime;
                            double temptimeslice = time.TotalMinutes;
                            if (Timeslice < temptimeslice)
                            {
                                Timeslice = temptimeslice;
                                TimesliceStartTime = endTime;
                            }
                        }
                    }
                }
                else
                {
                    this.Timeslice = 0;
                    TimesliceStartTime = null;
                }
                SetLastCap();
            }

            public void SetLastCap()
            {
                foreach (var mit in StepList)
                {
                    int cap = GetStandardWorkingTimeByProductId_SpecificationId(mit.ProductId,mit.SpecificationID);

                    LastCap -= cap;

                    if (LastCap < 0)
                        LastCap = 0;
                }
            }

            /// <summary>
            /// 找到Moitemid对应的计划及后续计划移除出经营体列表，并返回该计划列表
            /// </summary>
            /// <param name="mitsId"></param>
            /// <returns></returns>
            public List<MOItemTask> RemoveMOItemTaskById(string mitsId, double cap)
            {
                List<MOItemTask> returnMitsList = new List<MOItemTask>();

                if (StepList.Where(p => p.MOItemTaskId == mitsId).Any())
                {
                    MOItemTask step = StepList.Where(p => p.MOItemTaskId == mitsId).FirstOrDefault();

                    int index = StepList.IndexOf(step);

                    returnMitsList = StepList.Skip(index).Take(StepList.Count - index).ToList();

                    //排除不是已登记状态的计划
                    for (int i = 0; i < returnMitsList.Count; i++)
                    {
                        if (returnMitsList[i].TaskStatus != "已登记")
                            returnMitsList.Remove(returnMitsList[i]);
                    }

                    //StepList.RemoveRange(index, StepList.Count - index);
                    foreach (var returnMits in returnMitsList)
                    {
                        StepList.Remove(returnMits);
                    }
                }

                //重新计划时间片
                SetTimeslice();

                return returnMitsList;
            }

            /// <summary>
            /// 返回经营体是否当天能排入这个计划
            /// </summary>
            /// <param name="step"></param>
            /// <param name="cap"></param>
            /// <returns></returns>
            public bool isCanPayload(double cap)
            {
                double sum = LastCap + cap;
                double maxcap = VentureCap * 1.2;
                if (sum <= maxcap)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }

            /// <summary>
            /// 添加Step到数组中，并记录最后的完成时间
            /// </summary>
            /// <param name="step"></param>
            public RetureVentureDayModel AddStep(MOItemTask step, double cap,DateTime? priviousLastTime,string VentureID)
            {
                RetureVentureDayModel retModel = new RetureVentureDayModel();

                DateTime? EndTime = null;
                if (!StepList.Where(p => p.MOItemTaskId == step.MOItemTaskId).Any())
                {
                    //if(this.LastEndTime == null)
                    //{
                    //    step.PlannedDateFrom = new DateTime(priviousLastTime.Value.Year, priviousLastTime.Value.Month, priviousLastTime.Value.Day, 8,30,0) ;
                    //}
                    //else
                    //{
                    //    step.PlannedDateFrom = this.LastEndTime;
                    //}

                    step.PlannedDateFrom = new DateTime(priviousLastTime.Value.Year, priviousLastTime.Value.Month, priviousLastTime.Value.Day, 8, 30, 0);

                    DateTime temp = ((DateTime)step.PlannedDateFrom).AddMinutes(cap);
                    step.PlannedDateTo = temp;

                    step.TaskStatus = "被排产";
                    step.VentureId = VentureID;

                    StepList.Add(step);

                    this.LastEndTime = step.PlannedDateTo;

                    EndTime = this.LastEndTime;
                }

                //SetTimeslice();

                SetLastCap();

                retModel.EndTime = EndTime;
                retModel.MOItemTask = step;

                return retModel;
            }

            public static int GetStandardWorkingTimeByProductId_SpecificationId(string ProductId, string SpecificationId)
            {
                ProductSpecification ps = SchedulePlan.GetEntity().ProductSpecification.Where(p => p.ProductId == ProductId && p.SpecificationId == SpecificationId).FirstOrDefault();

                if (ps != null && ps.StandardWorkingTime != null)
                    return (int)ps.StandardWorkingTime;
                else
                    return 0;
            }

            public static DateTime GetWorkTime(DateTime startTime, string ventureId)
            {
                WorkCenterCalendar noWorkCalender = SchedulePlan.GetEntity().WorkCenterCalendar.Where(p => DbFunctions.TruncateTime(p.NoWorkBeginDate) <= startTime && DbFunctions.TruncateTime(p.NoWorkEndDate) >= startTime).FirstOrDefault();

                DateTime workTime;

                if (noWorkCalender != null)
                    workTime = (DateTime)noWorkCalender.NoWorkEndDate;
                else
                    workTime = startTime;

                List<Venture_Shift> vfList = SchedulePlan.GetEntity().Venture_Shift.Where(p => p.VentureId == ventureId).ToList();



                if (!vfList.Any())
                {
                    Venture v = SchedulePlan.GetEntity().Venture.Where(p => p.VentureId == ventureId).FirstOrDefault();
                    throw new Exception(v.VentureName + "(" + v.VentureId + ")" + "经营体与工厂日历关系未维护！");
                }

                var shiftIdList = vfList.Select(p => p.ShiftId).ToList();


                var shiftList = SchedulePlan.GetEntity().Shift.Where(p => shiftIdList.Contains(p.ShiftId)).ToList();

                var minTimeFrom = shiftList.Max(p => p.TimeFrom);

                Shift shift = shiftList.Where(p => p.TimeFrom == minTimeFrom).FirstOrDefault();

                if (shift == null)
                    throw new Exception("工厂日历未维护！");

                string[] minsecond = shift.TimeFrom.Split(new char[] { ':' });

                DateTime workStartTime = new DateTime(startTime.Year, startTime.Month, startTime.Day, int.Parse(minsecond[0]), int.Parse(minsecond[1]), 0);
                DateTime workEndTime = workStartTime.AddMinutes(Convert.ToDouble(shift.TimeRange));

                if (workTime < workStartTime || workTime > workEndTime)
                {
                    workTime = workStartTime.AddDays(1);

                    noWorkCalender = SchedulePlan.GetEntity().WorkCenterCalendar.Where(p => DbFunctions.TruncateTime(p.NoWorkBeginDate) >= workTime || DbFunctions.TruncateTime(p.NoWorkEndDate) <= workTime).FirstOrDefault();

                    if (noWorkCalender != null)
                    {
                        TimeSpan? ts = noWorkCalender.NoWorkEndDate - noWorkCalender.NoWorkBeginDate;
                        int day = ((TimeSpan)ts).Days;

                        workTime = workStartTime.AddDays(day + 1);
                    }
                }

                //return 
                return workTime;
            }
        }

        public class RetureVentureDayModel
        {
            public DateTime? EndTime { get; set; }

            public MOItemTask MOItemTask { get; set; }
        }

    }

    
}
