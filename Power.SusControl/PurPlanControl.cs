using Newtonsoft.Json;
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
                        saveBo.SetItem("DesignDate", row["DesignDate"]);
                        saveBo.SetItem("DeliveryDate", row["DeliveryDate"]);
                        saveBo.SetItem("NewDeliveryDate", row["NewDeliveryDate"]);
                        if ( periods[i] == "计划完成日期" || periods[i] == "预估完成日期" )
                        {
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
        string DateToString(object input)
        {
            DateTime dt = Convert.ToDateTime(input);
            if ( dt == new DateTime() )
                return null;
            else
                return dt.ToString("yyyy-MM-dd");
        }
    }
}
