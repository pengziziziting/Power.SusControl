﻿using Newtonsoft.Json;
using Power.Business;
using Power.Controls.PMS;
using Power.Global;
using Power.Message;
using Power.Service.MailService;
using Power.WorkFlows;
using Power.WorkFlows.WorkManage;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Power.SusControl
{
    public class SUSAPIControl : BaseControl
    {
        public const string EEstimateDate = "预估完成日期", EActualDate = "实际完成日期",EPlanDate= "计划完成日期";
        public void setContentType()
        {
            if ( this.Context != null && this.Context.Response != null )
            {
                this.Context.Response.ContentType = "application/json; charset=utf-8";
            }
        }
        [Action(Authorize = false)]
        public string Test(string value)
        {
            NewLife.Log.XTrace.WriteLine("value");
            setContentType();
            ViewResultModel result = ViewResultModel.Create(true, "测试接口连通");
            return result.ToJson();
        }
        #region 采购计划监控
        /// <summary>
        /// 创建 采购计划执行监控
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        [Action]
        public string CreatePurPlanControl()
        {
            setContentType();
            ViewResultModel result = ViewResultModel.Create(false, "创建采购计划执行监控");
            try
            {
                string _keywordplan = "Sus_MakePlans", _keywordcontrol = "Sus_Pur_PlanControl", _keywordcontroldtl = "Sus_Pur_PlanControlDtl";
                var planOpt = BusinessFactory.CreateBusinessOperate(_keywordplan);
                var planList = planOpt.FindAll("EpsProjId='" + this.session.EpsProjId + "'", "Sequ", "");
                var oldControlList = BusinessFactory.CreateBusinessOperate(_keywordcontrol).FindAll("EpsProjId", this.session.EpsProjId);
                if ( planList.Count == 0 )
                    throw new Exception("采购计划不存在");
                var controlId = Guid.NewGuid();
                var saveList = BusinessFactory.CreateBusinessOperate(_keywordcontroldtl).FindAll("1=0", "Sequ", "");
                var periods = new string[] { "持续时间", "计划完成日期", "预估完成日期", "实际完成日期", "状态" };
                var offset = 0;
                foreach ( var row in planList )
                {
                    int sequ = Convert.ToInt32(row["Sequ"]);
                    int dtlSequ = sequ + offset;
                    offset += 5;
                    for ( int i = 0; i < periods.Length; i++ )
                    {
                        var saveBo = BusinessFactory.CreateBusiness(_keywordcontroldtl);
                        saveBo.SetItem("Id", Guid.NewGuid());
                        saveBo.SetItem("MasterId", controlId);
                        saveBo.SetItem("TempId", row["TempId"]);
                        saveBo.SetItem("Code", row["Code"]);
                        saveBo.SetItem("Name", row["Name"]);
                        saveBo.SetItem("Level", row["Level"]);
                        saveBo.SetItem("DesignCycle", row["DesignCycle"]);
                        saveBo.SetItem("FabricationCycle", row["FabricationCycle"]);
                        saveBo.SetItem("PurchasingEngineer", row["PurchasingEngineer"]);
                        saveBo.SetItem("TechnicalEngineer", row["TechnicalEngineer"]);
                        saveBo.SetItem("Period", periods[i]);
                        saveBo.SetItem("Sequ", dtlSequ + i);
                        saveBo.SetItem("NewDeliveryDate", row["NewDeliveryDate"]);
                        if ( periods[i] == "计划完成日期" || periods[i] == "预估完成日期" )
                        {
                            saveBo.SetItem("DesignDate", row["DesignDate"]);
                            saveBo.SetItem("DeliveryDate", row["DeliveryDate"]);
                            saveBo.SetItem("Step0", DateToString(row["Step0"]));
                            saveBo.SetItem("Step1", DateToString(row["Step1"]));
                            saveBo.SetItem("Step2", DateToString(row["Step2"]));
                            saveBo.SetItem("Step3", DateToString(row["Step3"]));
                            saveBo.SetItem("Step4", DateToString(row["Step4"]));
                            saveBo.SetItem("Step5", DateToString(row["Step5"]));
                            saveBo.SetItem("Step6", DateToString(row["Step6"]));
                            saveBo.SetItem("Step7", DateToString(row["Step7"]));
                            saveBo.SetItem("Step8", DateToString(row["Step8"]));
                        }
                        saveList.Add(saveBo);
                    }
                }
                var Project = BusinessFactory.CreateBusinessOperate("Project").FindByKey(this.session.EpsProjId);
                var controlBo = BusinessFactory.CreateBusiness(_keywordcontrol);
                controlBo.SetItem("Id", controlId);
                controlBo.SetItem("ProjectCode", Project["project_shortname"]);
                controlBo.SetItem("ProjectName", Project["project_name"]);
                controlBo.SetItem("ProjectManager", Project["Pro_manager_name"]);
                controlBo.SetItem("Address", Project["project_address"]);
                using ( var tran = new XCode.EntityTransaction(planOpt.GetEntityOperate()) )
                {
                    oldControlList.Delete();//删除以前的版本记录
                    controlBo.Save(System.ComponentModel.DataObjectMethodType.Insert);
                    int count = saveList.Save(true);
                    result.data.Add("count", count);
                    tran.Commit();
                }
                result.success = true;
            }
            catch ( Exception ex )
            {
                result.message = ex.Message;
            }
            return result.ToJson();
        }
        /// <summary>
        /// 获取采购计划执行监控信息
        /// </summary>
        /// <param name="id">采购计划执行监控id</param>
        /// <returns></returns>
        [Action]
        public string GetPurPlanControlInfo(string id)
        {
            setContentType();
            ViewResultModel result = ViewResultModel.Create(false, "创建采购计划执行监控");
            try
            {
                string _keywordcontrol = "Sus_Pur_PlanControl";
                var planOpt = BusinessFactory.CreateBusinessOperate(_keywordcontrol);
                if ( string.IsNullOrEmpty(id) )
                {
                    var plan = planOpt.FindAll("EpsProjId='" + this.session.EpsProjId + "'", "", "").FirstOrDefault();
                    if ( plan == null )
                        throw new Exception("找不到项目下的采购计划执行监控");
                    else
                        result.data.Add("value", plan);
                }
                else
                {
                    var plan = planOpt.FindByKey(id);
                    if ( plan == null )
                        throw new Exception("找不到项目下的采购计划执行监控");
                    else
                        result.data.Add("value", plan);
                }
                result.success = true;
            }
            catch ( Exception ex )
            {
                result.message = ex.Message;
            }
            return result.ToJson();
        }
        #endregion
        #region 获取实际完成时间
        [Action(Authorize = true)]
        public string GetActualFinishDate(string planId)
        {
            setContentType();
            ViewResultModel result = ViewResultModel.Create(true, "获取实际完成时间");
            result.data.Add("planId", planId);
            var sqlText = "select t1.ApprDate as Step1,t2.ApprDate as Step2,t3.ApprDate as Step3,t4.ApprDate as Step4,t5.ApprDate as Step5,t6.ApprDate as Step6,t7.ApprDate as Step7,t8.ApprDate as Step8,t1.Id,t1.OwnProjName as ProjectName,t1.DeviceCode from Sus_TechnicalBook t1 left join Sus_Bid_Inquiry t2 on (t1.DeviceCode=t2.DeviceCode and t1.EpsProjId=t2.EpsProjId and t2.Status=50)  left join PS_BID_BidOpen t3 on (t1.DeviceCode=t3.DeviceCode and t1.EpsProjId=t3.EpsProjId and t3.Status>30 and t3.Type='技术') left join Sus_Pur_ExpertReview t4 on (t1.DeviceCode=t4.DeviceCode and t1.EpsProjId=t4.EpsProjId and t4.Status=50) left join PS_BID_BidOpen t5 on (t1.DeviceCode=t3.DeviceCode and t1.EpsProjId=t5.EpsProjId and t5.Status>30 and t5.Type='商务') left join PS_BID_BidReview t6 on (t1.DeviceCode=t6.DeviceCode and t1.EpsProjId=t6.EpsProjId and t6.Status=50) left join PS_CM_SubContract t7 on (t1.DeviceCode=t7.DeviceCode and t1.EpsProjId=t7.EpsProjId and t7.Status=50) left join Contract_registration t8 on (t1.DeviceCode=t8.device_number and t1.EpsProjId=t8.EpsProjId and t8.Status=50) where t1.DeviceCode is not null and t1.Status=50";
            try
            {
                if ( !string.IsNullOrEmpty(planId) )
                {
                    sqlText += "and t1.EpsProjId='" + this.session.EpsProjId + "'";
                }
                var dt = XCode.DataAccessLayer.DAL.QuerySQL(sqlText);
                string _keywordplan = "Sus_MakePlans", _keywordcontrol = "Sus_Pur_PlanControl", _keywordcontroldtl = "Sus_Pur_PlanControlDtl";
                var planBo = Power.Business.BusinessFactory.CreateBusinessOperate(_keywordcontrol).FindByKey(planId);
                if ( planBo == null )
                    throw new Exception("计划不存在");
                var planDtlList = Power.Business.BusinessFactory.CreateBusinessOperate(_keywordcontroldtl).FindAll("MasterId", planId);
                if ( planDtlList == null || planDtlList.Count == 0 )
                    throw new Exception("采购计划详情不存在");
                NewLife.Log.XTrace.WriteException("获取采购计划详情成功");
                var actualFinishList = new List<Models.PurPlanRow>();
                result.list = actualFinishList;
                int rowCount = 0;
                //取出实际完成时间的行.
                foreach ( System.Data.DataRow dr in dt.Rows )
                {
                    rowCount++;
                    var row = new Models.PurPlanRow(dr["Step1"], dr["Step2"], dr["Step3"], dr["Step4"], dr["Step5"], dr["Step6"], dr["Step7"], dr["Step8"]);
                    row.Id = Convert.ToString(dr["Id"]);
                    row.DeviceCode = Convert.ToString(dr["DeviceCode"]);
                    row.ProjectName = Convert.ToString(dr["ProjectName"]);
                    actualFinishList.Add(row);
                    // NewLife.Log.XTrace.WriteException("添加实际完成时间成功"+row.DeviceCode);
                    //找出设备编码相同的更新实际时间
                    var actualDateBo = planDtlList.Where(s => s["Code"].Equals(row.DeviceCode) && s["Period"].Equals(EActualDate)).FirstOrDefault();
                    var estimateDateBo = planDtlList.Where(s => s["Code"].Equals(row.DeviceCode) && s["Period"].Equals(EEstimateDate)).FirstOrDefault();
                    var planDateBo = planDtlList.Where(s => s["Code"].Equals(row.DeviceCode) && s["Period"].Equals(EPlanDate)).FirstOrDefault();
                    if ( actualDateBo == null || estimateDateBo == null || planDateBo == null )
                    {
                        result.data.Add("null_" + rowCount, true);
                        continue;
                    }
                    /* var planDateRow = new Models.PurPlanRow(planDateBo["Step1"], planDateBo["Step2"], planDateBo["Step3"], planDateBo["Step4"], planDateBo["Step5"], planDateBo["Step6"], planDateBo["Step7"], planDateBo["Step8"]);*/
                    int startIndex = 0;
                    for ( int i = 1; i < 9; i++ )
                    {
                        var field = "Step" + i;
                        if ( row[i] != null && row[i].HasValue )
                        {
                            if ( startIndex == 0 )
                            {
                                startIndex = i;
                            }
                            actualDateBo.SetItem(field, row[i].ToString());
                        }
                        //NewLife.Log.XTrace.WriteException("更新实际完成时间成功" + field);
                    }
                    var stepDates = new Dictionary<string, DateTime>();
                    result.data.Add("startIndex_"+ rowCount, startIndex);
                    if ( startIndex > 0 )
                        calcNextStage(startIndex, row, planDateBo, stepDates);
                    if ( stepDates.ContainsKey("Step6") )
                    {
                        var designDate = stepDates["Step6"].AddDays(Convert.ToInt32(estimateDateBo["DesignCycle"]));
                        stepDates.Add("DesignDate", designDate);
                    }
                    //预估交货时间
                    if ( stepDates.ContainsKey("Step8") )
                    {
                        var deliveryDate = stepDates["Step8"].AddDays(Convert.ToInt32(estimateDateBo["FabricationCycle"]));
                        stepDates.Add("DeliveryDate", deliveryDate);
                    }
                    foreach (string key in stepDates.Keys )
                    {
                        estimateDateBo.SetItem(key,stepDates[key].ToString("yyyy-MM-dd"));//更新预估时间
                        //预估提资时间                       
                    }

                    result.data.Add("dates_"+ rowCount, stepDates);
                    actualDateBo.UpdateSelf();
                    estimateDateBo.UpdateSelf();
                }
                result.success = true;
            }
            catch ( Exception ex )
            {
                result.message = ex.Message;
                result.success = false;
            }
            return result.ToJson();
        }
        #endregion
        string DateToString(object input)
        {
            DateTime dt = Convert.ToDateTime(input);
            if ( dt == new DateTime() )
                return null;
            else
                return dt.ToString("yyyy-MM-dd");
        }
        Dictionary<string, string> boName = new Dictionary<string, string>()
        {
            { "Sus_TechnicalBook","技术规格书"},
             { "Sus_Bid_Inquiry","发标"},//项目管理-采购管理-采买管理-招标询价
              { "PS_BidOpen","发技术清标"},//开标记录 技术/商务
               { "Sus_Pur_ExpertReview","技术评标完成"},//项目管理-采购管理-采买管理-技术评标
                { "PS_BidReview","定标"},//项目管理-采购管理-采买管理-定标评审
            { "PS_SubContract","技术协议"},
            { "Contract_registration","合同签订"}

        };
        /// <summary>
        /// 更新下一阶段的预估完成时间
        /// </summary>
        /// <param name="actual">实际时间</param>
        /// <param name="planDateBo">计划时间</param>
        /// <param name="datas"></param>
        void calcNextStage(int curStage, Models.PurPlanRow actual, IBaseBusiness planDateBo, Dictionary<string, DateTime> datas)
        {
            var nextStageIndex = curStage + 1;
            var nextStage = "Step" + nextStageIndex;
            NewLife.Log.XTrace.WriteLine("curstage:{0},nextStage:{1}", curStage, nextStage);
            //超出索引返回,下一阶段有实际完成时间就跳出,下一阶段没有实际完成时间
            if ( nextStageIndex > 8 )
                return;

            else
            {   //如果当前阶段有实际完成时间,计算预估完成时间合理性
                NewLife.Log.XTrace.WriteLine("判断计划时间合理性:" + Convert.ToString(planDateBo[nextStage]));
                if ( actual[curStage] != null && actual[curStage].HasValue )
                {
                    if ( !datas.ContainsKey("Step" + curStage) )
                        datas.Add("Step" + curStage, actual[curStage].Value);
                    else
                        datas["Step" + curStage] = actual[curStage].Value;
                    //下一阶段有实际完成时间,不必计算下一阶段
                    if ( actual[nextStageIndex] != null && actual[nextStageIndex].HasValue )//
                    {
                        NewLife.Log.XTrace.WriteLine("跳出进入阶段" + nextStageIndex);
                        datas.Add(nextStage, actual[nextStageIndex].Value);
                    }
                    else
                    {
                        var planDate = Convert.ToDateTime(planDateBo[nextStage].ToString());
                        var currentActualDate = actual[curStage].Value;
                        var maxintervalfield = "Stage" + curStage + "Max";
                        var minintervalfield = "Stage" + curStage + "Min";
                        var diffvalue = planDate.Subtract(currentActualDate).TotalDays;
                        var minInterval = Convert.ToInt32(planDateBo[minintervalfield]);
                        var maxInterval = Convert.ToInt32(planDateBo[maxintervalfield]);
                        if ( diffvalue < minInterval )
                        {
                            datas.Add(nextStage, currentActualDate.AddDays(minInterval));
                        }
                        else if ( diffvalue > maxInterval )
                        {
                            datas.Add(nextStage, currentActualDate.AddDays(maxInterval));
                        }
                        else
                            datas.Add(nextStage, planDate);
                    }
                    calcNextStage(nextStageIndex, actual, planDateBo, datas);
                }
                else
                {
                    NewLife.Log.XTrace.WriteLine("连续计算计划时间");
                    if ( datas.ContainsKey("Step" + curStage) )
                    {
                        var nextplanDate = Convert.ToDateTime(planDateBo[nextStage].ToString());
                        var currentPlanDate = datas["Step" + curStage];
                        var maxintervalfield = "Stage" + curStage + "Max";
                        var minintervalfield = "Stage" + curStage + "Min";
                        var diffvalue = nextplanDate.Subtract(currentPlanDate).TotalDays;
                        var minInterval = Convert.ToInt32(planDateBo[minintervalfield]);
                        var maxInterval = Convert.ToInt32(planDateBo[maxintervalfield]);
                        if ( diffvalue < minInterval )
                        {
                            datas.Add(nextStage, currentPlanDate.AddDays(minInterval));
                        }
                        else if ( diffvalue > maxInterval )
                        {
                            datas.Add(nextStage, currentPlanDate.AddDays(maxInterval));
                        }
                        else
                            datas.Add(nextStage, nextplanDate);
                    }
                    calcNextStage(nextStageIndex, actual, planDateBo, datas);
                }
            }
            //下一阶段如果实际时间没值就更新预估完成时间

        }
    }
}
