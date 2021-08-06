using Aspose.Cells;
using Newtonsoft.Json;
using Power.Business;
using Power.Controls.PMS;
using Power.Global;
using Power.ISystems.IMessageService.Excel;
using Power.Message;
using Power.Service.MailService;
using Power.WorkFlows;
using Power.WorkFlows.WorkManage;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using JWT;
using JWT.Algorithms;
using JWT.Serializers;
using Power.Controls.SystemCESE.entity;
using System.Web.Script.Serialization;
using XCode.DataAccessLayer;

namespace Power.SusControl
{
    public class SusControl : BaseControl
    {
        #region 个人中心 我的任务
        // Power.Controls.PMS.MessageControl
        [ActionAttribute]
        public string MyTaskInfos(string types, string index, string size, string swhere, string humanid = "")
        {
            ViewResultModel viewResultModel = ViewResultModel.Create(true, "");
            bool flag = string.IsNullOrEmpty(humanid);
            if (flag)
            {
                humanid = base.session.HumanId;
            }
            bool flag2 = swhere == "";
            if (flag2)
            {
                swhere = " 1=1";
            }
            int num = int.Parse(index);
            int num2 = int.Parse(size);
            string text = PowerGlobal.CheckWhere(swhere);
            bool flag3 = !string.IsNullOrEmpty(text);
            if (flag3)
            {
                throw new Exception(text);
            }
            num *= num2;
            StringBuilder stringBuilder = new StringBuilder();
            bool flag4 = types == null || types == "";
            if (flag4)
            {
                types = "infos";
            }
            List<string> list = types.ToLower().Split(new char[]
            {
        ','
            }).ToList<string>();
            string a = "SQL";
            bool flag6 = list.Contains("infos");
            if (flag6)
            {
                stringBuilder.Length = 0;
                stringBuilder.AppendFormat("RegHumId='{0}' and Status = '0' and " + swhere, humanid);
                bool flag7 = a == "Orcl";
                int num3;
                if (flag7)
                {
                    DataTable dt = BusinessFactory.CreateBusinessOperate("Sus_Pur_FeedbackRegister").FindAllByTable(stringBuilder.ToString(), "\"BidClosingDate\" Desc", "Id as \"Id\",Title as \"Title\",Code as \"Code\",BidClosingDate as \"BidClosingDate\",RegHumName as \"RegHumName\",SubmitDate as \"SubmitDate\",Versions as \"Versions\",Title as \"HtmlPath\",OwnProjName as \"OwnProjName\"", num, num2, SearchFlag.IgnoreRight);
                    viewResultModel.data.Add("infos", BusiHelper.DataTableToHashtable(dt));
                    num3 = BusinessFactory.CreateBusinessOperate("Sus_Pur_FeedbackRegister").FindCount(stringBuilder.ToString(), "BidClosingDate Desc", "Id,Title,BidClosingDate as BidClosingDate,Code,RegHumName,SubmitDate,Title as HtmlPath,Versions,OwnProjName", 0, 0, SearchFlag.IgnoreRight);
                }
                else
                {
                    DataTable dt2 = BusinessFactory.CreateBusinessOperate("Sus_Pur_FeedbackRegister").FindAllByTable(stringBuilder.ToString(), "BidClosingDate Desc", "Id,Title,BidClosingDate as BidClosingDate,Code,RegHumName,SubmitDate,Title as HtmlPath,Versions,OwnProjName", num, num2, SearchFlag.IgnoreRight);
                    viewResultModel.data.Add("infos", BusiHelper.DataTableToHashtable(dt2));
                    num3 = BusinessFactory.CreateBusinessOperate("Sus_Pur_FeedbackRegister").FindCount(stringBuilder.ToString(), "BidClosingDate Desc", "Id,Title,BidClosingDate as BidClosingDate,Code,RegHumName,SubmitDate,Title as HtmlPath,Versions,OwnProjName", 0, 0, SearchFlag.IgnoreRight);
                }
                viewResultModel.data.Add("infostotalcount", num3);
            }
            bool flag8 = list.Contains("actived");
            if (flag8)
            {
                stringBuilder.Length = 0;
                stringBuilder.AppendFormat("RegHumId='{0}' and  Status in ('35','50') and " + swhere, humanid);
                bool flag9 = a == "Orcl";
                int num3;
                if (flag9)
                {
                    DataTable dt3 = BusinessFactory.CreateBusinessOperate("Sus_Pur_FeedbackRegister").FindAllByTable(stringBuilder.ToString(), "\"BidClosingDate\" Desc", "Id as \"Id\",Title as \"Title\",Code as \"Code\",BidClosingDate as \"BidClosingDate\",RegHumName as \"RegHumName\",SubmitDate as \"SubmitDate\",Versions as \"Versions\",Title as \"HtmlPath\",OwnProjName as \"OwnProjName\"", num, num2, SearchFlag.IgnoreRight);
                    viewResultModel.data.Add("actived", BusiHelper.DataTableToHashtable(dt3));
                    num3 = BusinessFactory.CreateBusinessOperate("Sus_Pur_FeedbackRegister").FindCount(stringBuilder.ToString(), "BidClosingDate Desc", "Id,Title,BidClosingDate as BidClosingDate,Code,RegHumName,SubmitDate,Title as HtmlPath,Versions,OwnProjName", 0, 0, SearchFlag.IgnoreRight);
                }
                else
                {
                    DataTable dt4 = BusinessFactory.CreateBusinessOperate("Sus_Pur_FeedbackRegister").FindAllByTable(stringBuilder.ToString(), "BidClosingDate Desc", "Id,Title,BidClosingDate as BidClosingDate,Code,RegHumName,SubmitDate,Title as HtmlPath,Versions,OwnProjName", num, num2, SearchFlag.IgnoreRight);
                    viewResultModel.data.Add("actived", BusiHelper.DataTableToHashtable(dt4));
                    num3 = BusinessFactory.CreateBusinessOperate("Sus_Pur_FeedbackRegister").FindCount(stringBuilder.ToString(), "BidClosingDate Desc", "Id,Title,BidClosingDate as BidClosingDate,Code,RegHumName,SubmitDate,Title as HtmlPath,Versions,OwnProjName", 0, 0, SearchFlag.IgnoreRight);
                }
                viewResultModel.data.Add("activedtotalcount", num3);
            }
            return viewResultModel.ToJson();
        }
        #endregion
        #region 生成编码
        [ActionAttribute(Authorize = false)]
        public string BuildNumber(string Code)
        {
            ViewResultModel re = ViewResultModel.Create(true, "");
            string code = "1";
            String strsql = "select max(Code) as maxcode from SB_SupplierRegistration where  Code like '" + Code + "-%'";
            DataTable dtTemp = XCode.DataAccessLayer.DAL.QuerySQL(strsql);

            if (dtTemp.Rows.Count > 0 && dtTemp.Rows[0]["maxcode"] != DBNull.Value)
            {
                string leng = dtTemp.Rows[0]["maxcode"].ToString();
                code = (int.Parse(leng.Substring(2, leng.Length - 2)) + 1).ToString();
            }

            while (code.Length < 4)
                code = "0" + code;
            code = Code + '-' + code;
            re.data.Add("values", code);
            return re.ToJson();
        }

        #endregion

        #region 生成评标记录
        [ActionAttribute(Authorize = false)]
        public string SelectModule(string Id, string keyvalue)
        {
            Dictionary<String, Guid> tempids = new Dictionary<String, Guid>();
            Power.Business.IBusinessList list = Power.Business.BusinessFactory.CreateBusinessOperate("Sus_Base_budget_1").FindAll("MasterId", Id, Business.SearchFlag.Default);
            Power.Business.IBusinessList budget = Power.Business.BusinessFactory.CreateBusinessOperate("Sus_budget_1").FindAll("MasterId", keyvalue, Business.SearchFlag.Default);
            if (budget.Count() > 0)
            {
                budget.Delete();
            }
            foreach (Power.Business.IBaseBusiness item in list)
            {
                if (tempids.ContainsKey(item["Id"].ToString()) == false)
                {
                    tempids.Add(item["Id"].ToString(), Guid.NewGuid());

                }
                if (XCode.Common.Helper.IsNullKey(item["ParentId"]) == false && tempids.ContainsKey(item["ParentId"].ToString()) == false)
                {
                    tempids.Add(item["ParentId"].ToString(), Guid.NewGuid());
                }
                Power.Business.IBaseBusiness budget_1 = Power.Business.BusinessFactory.CreateBusiness("Sus_budget_1");
                budget_1.SetItem("Id", tempids[item["Id"].ToString()]);
                if (XCode.Common.Helper.IsNullKey(item["ParentId"]) == false)
                    budget_1.SetItem("ParentId", tempids[item["ParentId"].ToString()]);
                else
                    budget_1.SetItem("ParentId", null);
                budget_1.SetItem("MasterId", keyvalue);
                budget_1.SetItem("Code", item["Code"]);
                budget_1.SetItem("Content", item["Name"]);
                budget_1.Save(System.ComponentModel.DataObjectMethodType.Insert);
            }
            return "true";

        }

        #endregion

        #region 三级审核根据Id来升版内容
        [ActionAttribute]
        public string Sus_ApprTableCreate(Guid Id)
        {
            Power.Global.ViewResultModel result = Global.ViewResultModel.Create(false, "");
            Power.Business.IBusinessOperate busiOpt = BusinessFactory.CreateBusinessOperate("Sus_ApprTable");
            Power.Business.IBaseBusiness busi = busiOpt.FindByKey(Id);

            Power.Business.IBaseBusiness tempbusi = Sus_ApprTableVersion(this.session, Id);

            //获取到更新后主表数据的Id
            result.data.Add("value", tempbusi["Id"]);

            result.success = true;
            result.message = "升版成功！";
            return result.ToJson();
        }
        #endregion

        #region 三级审核,版本号升级
        public IBaseBusiness Sus_ApprTableVersion(Power.IBaseCore.ISession session, Guid Id)
        {
            #region 复制主表并修改版本号 + 1
            Power.Business.IBaseBusiness busi = this.Sus_ApprTableClone(session, Id);
            string version = busi["Version"].ToString();
            var str = RemoveLastChar(version, ".0");
            //将字母转换为数字后+1
            int nversion = char.Parse(str) + 1;
            //+1后在转换为字母
            string temp = (char)nversion + ".0";

            busi.SetItem("Version", temp);
            #endregion

            #region 复制附件表的数据
            List<Hashtable> doclistnew = new List<Hashtable>();
            IBusinessList docfilelist = BusinessFactory.CreateBusinessOperate("DocFile").FindAll("FolderId", Id, SearchFlag.IgnoreRight);
            foreach (Power.Business.IBaseBusiness item in docfilelist)
            {
                Hashtable hash1 = item._ConvertToHashtable();
                hash1["Id"] = Guid.NewGuid();
                hash1["FolderId"] = busi["Id"];
                doclistnew.Add(hash1);
            }
            #endregion

            #region 保存数据库
            using (var trans = new XCode.EntityTransaction(XCode.DataAccessLayer.DAL.Create()))
            {
                //主表添加
                busi.Save(System.ComponentModel.DataObjectMethodType.Insert);

                //附件表判断后添加
                docfilelist.Save(true);
                if (doclistnew.Count > 0)
                {
                    Power.Business.IBaseBusiness docfile = Power.Business.BusinessFactory.CreateBusiness("DocFile");
                    foreach (Hashtable docitem in doclistnew)
                        docfile.Save(docitem, System.ComponentModel.DataObjectMethodType.Insert);
                }

                trans.Commit();
            }
            #endregion

            return busi;
        }

        #endregion

        #region 三级审核主表复制修改字段
        public Power.Business.IBaseBusiness Sus_ApprTableClone(Power.IBaseCore.ISession session, Guid Id)
        {
            IBusinessOperate busiOpt = BusinessFactory.CreateBusinessOperate("Sus_ApprTable");
            Power.Business.IBaseBusiness busi = busiOpt.FindByKey(Id);

            #region 生成新记录
            busi.SetItem("Id", Guid.NewGuid());
            busi.SetItem("Status", 0);
            if (session != null)
            {

                busi.SetItem("RegHumName", session.HumanName);
            }

            busi.SetItem("RegDate", DateTime.Now);
            #endregion

            return busi;
        }
        #endregion

        #region 通用升版功能
        [ActionAttribute]
        /// <summary>
        /// 通用版本升版功能
        /// </summary>
        /// <param name="josn">打开的表单Id</param>
        /// <param name="id">主键</param>
        /// <param name="updCoulom">更新字段信息</param>
        /// <returns></returns>
        public string setUpgrade(string openfromId, string id, string field)
        {
            ViewResultModel rlt = ViewResultModel.Create(true, "");
            string widgetSql = string.Format(@"select ExtJson from pb_widget 
                                            where Id = '{0}'", openfromId);
            DataTable widgetList = XCode.DataAccessLayer.DAL.QuerySQL(widgetSql);
            foreach (DataRow row in widgetList.Rows)
            {
                Hashtable has = Newtonsoft.Json.JsonConvert.DeserializeObject<Hashtable>(row["ExtJson"].ToString());
                if (has != null && has["config"] != null && has["config"].ToString() != "")
                {
                    Hashtable configList = Newtonsoft.Json.JsonConvert.DeserializeObject<Hashtable>(has["config"].ToString());
                    ConfigChildren keyWordList = Newtonsoft.Json.JsonConvert.DeserializeObject<ConfigChildren>(configList["joindata"].ToString());

                    XCode.IEntity mainPlan = Power.Business.BusinessFactory.CreateBusinessOperate(keyWordList.KeyWord).GetEntityOperate().FindByKey(id);
                    if (mainPlan == null || (mainPlan["Status"].ToString() != "50" && mainPlan["Status"].ToString() != "35"))
                    {
                        throw new Exception("当前数据状态未批准，请批准。");
                    }
                    XCode.IEntity oldMainPlan = Power.Business.BusinessFactory.CreateBusinessOperate(keyWordList.KeyWord).GetEntityOperate().FindByKey(id);
                    Business.IBusinessList CountList = Business.BusinessFactory.CreateBusinessOperate(keyWordList.KeyWord)
                             .FindAll("EpsProjId = '" + mainPlan["EpsProjId"] + "'", "", "", 0, 0, Business.SearchFlag.IgnoreRight);

                    List<decimal> versionList = new List<decimal>();
                    foreach (Business.IBaseBusiness countList in CountList)
                    {
                        versionList.Add(countList[field] == null ? 0 : Convert.ToDecimal(mainPlan[field]));
                    }
                    string mainId = Guid.NewGuid().ToString();
                    decimal version = versionList.Max() + 1;
                    mainPlan.SetItem("Id", mainId);
                    mainPlan.SetItem("RegDate", DateTime.Now);
                    mainPlan.SetItem("ApprDate", DateTime.Now);
                    mainPlan.SetItem("ApprDate", null);
                    mainPlan.SetItem("ApprHumId", null);
                    mainPlan.SetItem("ApprHumName", null);
                    mainPlan.SetItem("Status", 0);
                    mainPlan.SetItem(field, version);
                    string strJson = JsonConvert.SerializeObject(mainPlan);
                    FormControl formControl = new FormControl();
                    string str = formControl.GetCode(keyWordList.KeyWord, strJson);
                    var serializer = new JavaScriptSerializer();
                    //将json字符转换为实体对象
                    ViewResultModel resulet = serializer.Deserialize<ViewResultModel>(str);
                    //Hashtable hashtable = Newtonsoft.Json.JsonConvert.DeserializeObject<Hashtable>(Newtonsoft.Json.JsonConvert.SerializeObject(resulet.data["value"]));
                    string CodeStr = Newtonsoft.Json.JsonConvert.SerializeObject(resulet.data["value"]);
                    string cs = CodeStr.Substring(1, CodeStr.Length - 2);
                    Hashtable hashtable = Newtonsoft.Json.JsonConvert.DeserializeObject<Hashtable>(cs);
                    if (hashtable != null)
                    {
                        mainPlan.SetItem(hashtable["code"].ToString(), hashtable["value"]);
                    }
                    mainPlan.Insert();
                    if (keyWordList.children.Count > 0)
                    {
                        setChildrenUpgrade(keyWordList.children, mainPlan, oldMainPlan, mainId, oldMainPlan["Id"].ToString());
                    }
                    //附件复制
                    XCode.IEntityList fileList = Power.Business.BusinessFactory.CreateBusinessOperate("DocFile")
                        .GetEntityOperate().FindAll("FolderId", id);
                    foreach (XCode.IEntity ent in fileList)
                    {
                        ent.SetItem("Id", Guid.NewGuid());
                        ent.SetItem("FolderId", mainId);
                        ent.Insert();
                    }
                    rlt.data.Add("value", mainId);
                }
                else
                {
                    rlt.data.Add("value", "");
                }
            }
            rlt.success = true;
            rlt.message = "成功";
            return rlt.ToJson();
        }

        /// <summary>
        /// 更新子表数据信息
        /// </summary>
        /// <param name="configChildrens">配置文件信息</param>
        /// <param name="mainPlan">新生成的数据</param>
        /// <param name="oldMainPlan">历史数据</param>
        public void setChildrenUpgrade(List<ConfigChildren> configChildrens, XCode.IEntity mainPlan,
                    XCode.IEntity oldMainPlan, string newmainId, string oldMainId, Hashtable hashtable = null)
        {
            if (configChildrens != null)
            {
                foreach (ConfigChildren children in configChildrens)
                {
                    string swhere = " 1=1 ";
                    if (children.filter != null)
                    {
                        foreach (DictionaryEntry f2 in children.filter) //过滤信息
                        {
                            string filter_Key = f2.Key.ToString();
                            string filter_Value = f2.Value.ToString();
                            if (filter_Key != "" && filter_Key != null)
                            {
                                swhere += " and " + filter_Key + "= '" + filter_Value + "'";
                            }
                        }
                    }
                    if (children.swhere != "" && children.swhere != null)
                    {
                        swhere += " and " + children.swhere.ToString();
                    }
                    if (hashtable != null)
                    {
                        swhere += " and MasterId = '" + hashtable["OldMasterId"] + "'";
                    }
                    foreach (DictionaryEntry fl in children.fields) //关联信息
                    {
                        string Key = fl.Key.ToString(); //子表外键
                        string Value = fl.Value.ToString(); //主表主键
                        Object mainId = null;
                        if (hashtable != null)
                        {
                            mainId = hashtable[Value];
                        }
                        else
                        {
                            mainId = mainPlan[Value];
                        }

                        if (children.KeyWordType == "ViewEntity")
                        {

                            /*XCode.IEntityList childrenList =
                             (XCode.IEntityList)Power.Business.ViewEntity.ViewEntityFactory.CreateViewEntity(children.KeyWord).
                             LoadDataList(Key + "= '" + oldMainPlan[Value] + "' and " + swhere, "", null, 0, 0);
                            foreach (XCode.IEntity ent in childrenList)
                            {
                                XCode.IEntityList oldent =
                                      (XCode.IEntityList)Power.Business.ViewEntity.ViewEntityFactory.CreateViewEntity(children.KeyWord)
                                      .LoadDataList(Key + " = '" + ent[Value] + "'", "", null, 0, 0);
                                setChildrenUpgrade(children.children, ent, oldent[0]);
                            }*/
                            Power.Business.ViewEntity.IViewEntityList childrenList = Power.Business.ViewEntity.ViewEntityFactory.CreateViewEntity(children.KeyWord).LoadDataList(Key + "= '" + oldMainPlan[Value] + "' and " + swhere, "", null, 0, 0);
                            foreach (Business.ViewEntity.IViewEntity row in childrenList)
                            {
                                Hashtable viewHash = new Hashtable();
                                //将查询结果转换为Hashtable
                                DataTable data = row.LoadDataTableList();
                                DataColumnCollection columns = data.Columns;
                                foreach (object o in columns)
                                {
                                    if (o.ToString() == "MasterId")
                                    {
                                        viewHash.Add("MasterId", newmainId);
                                    }
                                    else
                                    {
                                        viewHash.Add(o.ToString(), row[o.ToString()]);
                                    }
                                }
                                //视图中无外键，默认添加外键
                                if (!columns.Contains("MasterId"))
                                {
                                    viewHash.Add("MasterId", newmainId);
                                }
                                viewHash.Add("OldMasterId", oldMainId);
                                XCode.IEntity entities = null;
                                setChildrenUpgrade(children.children, entities, entities, newmainId, oldMainId, viewHash);
                            }
                        }
                        else
                        {
                            Object mainIdOld = null;
                            if (hashtable != null)
                            {
                                mainIdOld = hashtable[Value];
                            }
                            else
                            {
                                mainIdOld = oldMainPlan[Value];
                            }
                            XCode.IEntityList childrenList = Power.Business.BusinessFactory.CreateBusinessOperate(children.KeyWord)
                            .GetEntityOperate().FindAll(Key + " = '" + mainIdOld + "' and " + swhere, "", "", 0, 0);
                            foreach (XCode.IEntity ent in childrenList)
                            {
                                XCode.IEntity oldent = Power.Business.BusinessFactory.CreateBusinessOperate(children.KeyWord).GetEntityOperate().FindByKey(ent["Id"]);
                                var entId = Guid.NewGuid();
                                ent.SetItem("Id", entId);
                                ent.SetItem(Key, mainId);
                                if (hashtable != null)
                                {
                                    ent.SetItem("MasterId", hashtable["MasterId"]);
                                }
                                ent.Insert();
                                setChildrenUpgrade(children.children, ent, oldent, newmainId, oldMainId);
                            }

                        }
                    }
                }
            }

        }
        #endregion

        #region  通过登录人编号寻找登录人工号
        [ActionAttribute]
        public string GetHumanWorkCode(string HumanCode)
        {
            string sql = "select * from PB_Human_Log where code = '" + HumanCode + "'";
            DataTable tb = XCode.DataAccessLayer.DAL.QuerySQL(string.Format(sql));
            string str = "";
            if (tb.Rows.Count > 0)
            {
                foreach (DataRow item in tb.Rows)
                {
                    str = item["workcode"].ToString();
                }
            }
            return str;
        }
        #endregion

        #region 技术开标 商务专用  通过KeyValue取消生效 
        [ActionAttribute]
        public string UpdateDataStatus(string KeyValue)
        {
            string sql = "update PS_BID_BidOpen set Status = '0' where Id = '" + KeyValue + "'; select 1 as num;";
            DataTable tb = XCode.DataAccessLayer.DAL.QuerySQL(string.Format(sql));
            return "success";
        }
        #endregion

        #region  通过表单KeyValue获取附件信息
        [ActionAttribute]
        public string GetDocFiles(string KeyValue)
        {
            string sql = "select * from PB_DocFiles where FolderId = '" + KeyValue + "'";
            DataTable tb = XCode.DataAccessLayer.DAL.QuerySQL(string.Format(sql));
            Power.Global.ViewResultModel result = Power.Global.ViewResultModel.Create(true, "");
            result.data.Add("data", Power.Global.BusiHelper.DataTableToHashtable(tb));
            return result.ToJson();
        }
        #endregion

        #region  招标询价 将版本号 带入附件文件版本中
        [ActionAttribute]
        public string SetDocFilesFileVersion(string KeyValue, string FileVersion)
        {
            string sql = "update PB_DocFiles set FileVersion = '" + FileVersion + "' where FolderId = '" + KeyValue + "';select 1 as num;";
            DataTable tb = XCode.DataAccessLayer.DAL.QuerySQL(string.Format(sql));
            Power.Global.ViewResultModel result = Power.Global.ViewResultModel.Create(true, "");
            result.data.Add("data", Power.Global.BusiHelper.DataTableToHashtable(tb));
            return result.ToJson();
        }
        #endregion


        #region 根据Id来升版内容
        [ActionAttribute]
        public string BudgetCreate(Guid Id)
        {
            Power.Global.ViewResultModel result = Global.ViewResultModel.Create(false, "");
            Power.Business.IBusinessOperate busiOpt = BusinessFactory.CreateBusinessOperate("Sus_budget");
            Power.Business.IBaseBusiness busi = busiOpt.FindByKey(Id);

            Power.Business.IBaseBusiness tempbusi = BudgetVersion(this.session, Id);

            //获取到更新后主表数据的Id
            result.data.Add("value", tempbusi["Id"]);

            result.success = true;
            result.message = "升版成功！";
            return result.ToJson();
        }
        #endregion

        #region 预算操作,版本号升级
        public IBaseBusiness BudgetVersion(Power.IBaseCore.ISession session, Guid Id)
        {
            #region 复制主表并修改版本号 + 1
            Power.Business.IBaseBusiness busi = this.Clone(session, Id);
            string version = busi["Version"].ToString();
            var str = RemoveLastChar(version, ".0");
            //将字母转换为数字后+1
            int nversion = char.Parse(str) + 1;
            //+1后在转换为字母
            string temp = (char)nversion + ".0";
            busi.SetItem("Version", temp);
            #endregion

            #region 复制子表的数据
            ////树形菜单节点Id
            Dictionary<string, Guid> a = new Dictionary<string, Guid>();
            //a.Add("ParentId", Guid.NewGuid());

            List<Hashtable> listnew = new List<Hashtable>();
            IBusinessList Suslist = BusinessFactory.CreateBusinessOperate("Sus_budget_1").FindAll("MasterId", Id, SearchFlag.IgnoreRight);
            foreach (Power.Business.IBaseBusiness item in Suslist)
            {
                Hashtable hash1 = item._ConvertToHashtable();
                if (a.ContainsKey(item["Id"].ToString()) == false)
                {
                    a.Add(item["Id"].ToString(), Guid.NewGuid());
                }
                if (XCode.Common.Helper.IsNullKey(item["ParentId"]) == false && a.ContainsKey(item["ParentId"].ToString()) == false)
                {
                    a.Add(item["ParentId"].ToString(), Guid.NewGuid());
                }

                hash1["Id"] = a[hash1["Id"].ToString()];
                if (XCode.Common.Helper.IsNullKey(item["ParentId"]) == false)
                {
                    hash1["ParentId"] = a[item["ParentId"].ToString()];
                }
                else
                {
                    hash1["ParentId"] = null;
                }
                hash1["MasterId"] = busi["Id"];
                //// hash1["ParentId"] = hash1["Id"];
                listnew.Add(hash1);
            }
            #endregion

            #region 复制附件表的数据
            List<Hashtable> doclistnew = new List<Hashtable>();
            IBusinessList docfilelist = BusinessFactory.CreateBusinessOperate("DocFile").FindAll("FolderId", Id, SearchFlag.IgnoreRight);
            foreach (Power.Business.IBaseBusiness item in docfilelist)
            {
                Hashtable hash1 = item._ConvertToHashtable();
                hash1["Id"] = Guid.NewGuid();
                hash1["FolderId"] = busi["Id"];
                doclistnew.Add(hash1);
            }
            #endregion

            #region 保存数据库
            using (var trans = new XCode.EntityTransaction(XCode.DataAccessLayer.DAL.Create()))
            {
                //主表添加
                busi.Save(System.ComponentModel.DataObjectMethodType.Insert);

                //字表判断后添加
                Suslist.Save(true);
                if (listnew.Count > 0)
                {
                    Power.Business.IBaseBusiness bud = Power.Business.BusinessFactory.CreateBusiness("Sus_budget_1");
                    foreach (Hashtable item in listnew)
                        bud.Save(item, System.ComponentModel.DataObjectMethodType.Insert);
                }

                //附件表判断后添加
                docfilelist.Save(true);
                if (doclistnew.Count > 0)
                {
                    Power.Business.IBaseBusiness docfile = Power.Business.BusinessFactory.CreateBusiness("DocFile");
                    foreach (Hashtable docitem in doclistnew)
                        docfile.Save(docitem, System.ComponentModel.DataObjectMethodType.Insert);
                }

                trans.Commit();
            }
            #endregion

            return busi;
        }

        #endregion

        #region 主表复制修改字段
        public Power.Business.IBaseBusiness Clone(Power.IBaseCore.ISession session, Guid Id)
        {
            IBusinessOperate busiOpt = BusinessFactory.CreateBusinessOperate("Sus_budget");
            Power.Business.IBaseBusiness busi = busiOpt.FindByKey(Id);

            #region 生成新记录
            busi.SetItem("Id", Guid.NewGuid());
            busi.SetItem("Status", 0);
            if (session != null)
            {

                busi.SetItem("RegHumName", session.HumanName);
            }

            busi.SetItem("RegDate", DateTime.Now);
            #endregion

            return busi;
        }
        #endregion

        #region 供应商年审模板导入
        [ActionAttribute]
        public string YearCarefulModule(string Id, string keyvalue)
        {
            Power.Global.ViewResultModel result = Power.Global.ViewResultModel.Create(true, "");
            Dictionary<String, Guid> tempids = new Dictionary<String, Guid>();
            Power.Business.IBusinessList list = Power.Business.BusinessFactory.CreateBusinessOperate("Sus_Suppliers_TemplateList").FindAll("MasterId", Id, Business.SearchFlag.Default);
            Power.Business.IBusinessList budget = Power.Business.BusinessFactory.CreateBusinessOperate("Sus_Suppliers_AnnualAuditList").FindAll("MasterId", keyvalue, Business.SearchFlag.Default);
            if (budget.Count() > 0)
            {
                budget.Delete();
            }
            foreach (Power.Business.IBaseBusiness item in list)
            {
                if (tempids.ContainsKey(item["Id"].ToString()) == false)
                {
                    tempids.Add(item["Id"].ToString(), Guid.NewGuid());

                }
                if (XCode.Common.Helper.IsNullKey(item["ParentId"]) == false && tempids.ContainsKey(item["ParentId"].ToString()) == false)
                {
                    tempids.Add(item["ParentId"].ToString(), Guid.NewGuid());
                }
                Power.Business.IBaseBusiness budget_1 = Power.Business.BusinessFactory.CreateBusiness("Sus_Suppliers_AnnualAuditList");
                budget_1.SetItem("Id", tempids[item["Id"].ToString()]);
                if (XCode.Common.Helper.IsNullKey(item["ParentId"]) == false)
                    budget_1.SetItem("ParentId", tempids[item["ParentId"].ToString()]);
                else
                    budget_1.SetItem("ParentId", null);
                budget_1.SetItem("MasterId", keyvalue);
                budget_1.SetItem("Code", item["Code"]);
                budget_1.SetItem("Items", item["Name"]);
                budget_1.SetItem("Details", item["Items"]);
                budget_1.SetItem("TempMemo", item["Memo"]);
                budget_1.Save(System.ComponentModel.DataObjectMethodType.Insert);
            }
            return result.ToJson();


        }
        #endregion

        #region 邮件内容格式转换
        private String processMessageContent(String text, Power.IBaseCore.ISession session, Power.Business.IBaseBusiness busi, Power.Business.IBusinessOperate busiOpt)
        {
            text = Power.Business.Common.Helper.ReplaceSessionParams(session, text, null);

            foreach (KeyValuePair<string, Power.Business.BusinessProperty> item in busiOpt.AllPropertyList)
            {

                if (busi[item.Key] != null)
                    text = text.Replace("[" + busiOpt.KeyWord + "." + item.Key + "]", busi[item.Key].ToString());


            }
            return text;
        }
        #endregion

        #region 邮件群发送功能
        [ActionAttribute]
        public string SendEmail(string Id)
        {
            ViewResultModel result = ViewResultModel.Create(true, "");

            Power.Business.IBusinessOperate listOpt = Power.Business.BusinessFactory.CreateBusinessOperate("Sus_Bid_InquiryList_Short");
            Power.Business.IBusinessList list = listOpt.FindAll("MasterId", Id, Business.SearchFlag.IgnoreRight);
            Power.Business.IBusinessOperate bidOpt = Power.Business.BusinessFactory.CreateBusinessOperate("Sus_Bid_Inquiry");
            Power.Business.IBaseBusiness bidMaster = bidOpt.FindByKey(Id);
            Power.Global.MailDataPack model = new Power.Global.MailDataPack();
            model.useparam = Power.Global.EUseMailDefaultParma.UseSmtpUser; //使用Power.ConfigEdit.exe 中配置的发送邮件参数

            #region 获取邮件模板
            String swhere = "BaseDataId in (select x1.Id from PB_BaseData x1 where x1.DataType= 'Sus_Bid_Email_Template')";
            Power.Business.IBusinessList basedataList = Power.Business.BusinessFactory.CreateBusinessOperate("BaseDataList").FindAll(swhere, "", "", 0, 0, SearchFlag.IgnoreRight);
            String mailtitle = "";
            String mailcontent = "";
            foreach (Power.Business.IBaseBusiness item in basedataList)
            {
                if (item["Code"] != null && item["Code"].ToString() == "Title")
                {
                    mailtitle = item["Name"].ToString();
                }
                if (item["Code"] != null && item["Code"].ToString() == "Content")
                {
                    mailcontent = item["Name"].ToString();
                }
            }
            #endregion

            //设置邮件的收件人
            string address = "";
            string subject = "";
            string content = "";
            string dateTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            //传入多个邮箱，
            foreach (Power.Business.IBaseBusiness it in list)
            {
                if ((int)it["isPass"] == 1)
                {

                    continue;
                }
                else
                {
                    if (it["Email"] != null && it["IsSupplier"].ToString().Equals("1"))
                    {
                        string[] email = it["Email"].ToString().Split(',');

                        foreach (string itme in email)
                        {
                            address += "'" + itme + "'";
                        }


                        #region 替换模板参数
                        subject = processMessageContent(mailtitle, this.session, bidMaster, bidOpt);
                        subject = processMessageContent(subject, this.session, it, listOpt);
                        content = processMessageContent(mailcontent, this.session, bidMaster, bidOpt);
                        content = processMessageContent(content, this.session, it, listOpt);
                        #endregion

                        //收件人地址
                        model.msg_to = address;
                        //邮件标题
                        model.msg_subject = subject;
                        //邮件内容
                        model.msg_content = content;

                        String errorinfo = "";
                        if (Power.Service.MailService.MailBLL.SendMail(model, out errorinfo))
                        {
                            //发送邮件成功,修改数据
                            Power.Business.IBusinessOperate busiOpt = BusinessFactory.CreateBusinessOperate("Sus_Bid_InquiryList_Short");

                            Power.Business.IBaseBusiness busi = busiOpt.FindByKey(it["Id"]);
                            busi.SetItem("isPass", 1);
                            busi.SetItem("PassDate", dateTime);
                            busi.Save(System.ComponentModel.DataObjectMethodType.Update);

                            address = "";
                            //subject = "";

                        }
                        else
                        {
                            //发送邮件失败，错误原因在 errorinfo 里面
                            result.success = false;
                            result.message = "邮件发送失败!";

                        }
                    }
                }
                continue;
            }
            return result.ToJson();
        }
        #endregion

        #region 再次报价

        [ActionAttribute]
        public string QuoteAgain(string Id)
        {
            ViewResultModel result = ViewResultModel.Create(true, "");

            Power.Business.IBaseBusiness Inquiry = Power.Business.BusinessFactory.CreateBusinessOperate("Sus_Bid_Inquiry").FindByKey(Id);
            Power.Business.IBusinessOperate listOpt = Power.Business.BusinessFactory.CreateBusinessOperate("Sus_Bid_InquiryList_Short");
            Power.Business.IBusinessList list = listOpt.FindAll("MasterId", Id, Business.SearchFlag.IgnoreRight);
            Power.Business.IBusinessOperate bidOpt = Power.Business.BusinessFactory.CreateBusinessOperate("Sus_Bid_Inquiry");
            Power.Business.IBaseBusiness bidMaster = bidOpt.FindByKey(Id);
            Power.Global.MailDataPack model = new Power.Global.MailDataPack();
            model.useparam = Power.Global.EUseMailDefaultParma.UseSmtpUser; //使用Power.ConfigEdit.exe 中配置的发送邮件参数

            #region 判断是否批准,如果不为批准,不允许再次报价 
            if (Inquiry["Status"].ToString() != null && int.Parse(Inquiry["Status"].ToString()) == 50)
            {
                #region 获取邮件模板
                String swhere = "BaseDataId in (select x1.Id from PB_BaseData x1 where x1.DataType= 'Sus_Bid_Email_Template')";
                Power.Business.IBusinessList basedataList = Power.Business.BusinessFactory.CreateBusinessOperate("BaseDataList").FindAll(swhere, "", "", 0, 0, SearchFlag.IgnoreRight);
                String mailtitle = "";
                String mailcontent = "";
                foreach (Power.Business.IBaseBusiness item in basedataList)
                {
                    if (item["Code"] != null && item["Code"].ToString() == "Title")
                    {
                        mailtitle = item["Name"].ToString();
                    }
                    if (item["Code"] != null && item["Code"].ToString() == "Content")
                    {
                        mailcontent = item["Name"].ToString();
                    }
                }
                #endregion

                #region 执行存储过程
                string StrSql = "exec p_Sus_Bid_FeedbackRegister_crt '" + Id + "'";
                XCode.DataAccessLayer.DAL.Create().Execute(StrSql);
                #endregion

                #region 循环发送邮件
                //设置邮件的收件人
                string address = "";
                string subject = "";
                string content = "";
                string dateTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                //传入多个邮箱，
                foreach (Power.Business.IBaseBusiness it in list)
                {
                    if (it["Email"] != null && it["IsSupplier"].ToString().Equals("1"))
                    {
                        string[] email = it["Email"].ToString().Split(',');

                        foreach (string itme in email)
                        {
                            address += "" + itme + "";
                        }


                        #region 替换模板参数
                        subject = processMessageContent(mailtitle, this.session, bidMaster, bidOpt);
                        subject = processMessageContent(subject, this.session, it, listOpt);
                        content = processMessageContent(mailcontent, this.session, bidMaster, bidOpt);
                        content = processMessageContent(content, this.session, it, listOpt);
                        #endregion

                        //收件人地址
                        model.msg_to = address;
                        //邮件标题
                        model.msg_subject = subject;
                        //邮件内容
                        model.msg_content = content;

                        String errorinfo = "";
                        if (Power.Service.MailService.MailBLL.SendMail(model, out errorinfo))
                        {
                            //发送邮件成功,修改数据
                            Power.Business.IBusinessOperate busiOpt = BusinessFactory.CreateBusinessOperate("Sus_Bid_InquiryList_Short");

                            Power.Business.IBaseBusiness busi = busiOpt.FindByKey(it["Id"]);
                            busi.SetItem("isPass", 1);
                            busi.SetItem("PassDate", dateTime);
                            busi.Save(System.ComponentModel.DataObjectMethodType.Update);

                            address = "";

                        }
                        else
                        {
                            //发送邮件失败，错误原因在 errorinfo 里面
                            result.success = false;
                            result.message = "邮件发送失败!";
                        }
                    }
                }
                #endregion
            }
            else
            {
                result.success = false;
                result.message = "只有批准后的数据才能再次发起报价！";
            }
            #endregion
            return result.ToJson();
        }
        #endregion

        #region 供应商新增
        [ActionAttribute]
        public string SupplierChange(string Id)
        {
            Power.Global.ViewResultModel result = Power.Global.ViewResultModel.Create(true, "");
            XCode.DataAccessLayer.DAL dal = XCode.DataAccessLayer.DAL.Create();
            string strSql = "select * from SB_SupplierRegistration where Status = 50 and Sup_HumanId='" + Id + "'";
            DataTable tb = XCode.DataAccessLayer.DAL.QuerySQL(string.Format(strSql));
            result.data.Add("values", Power.Global.BusiHelper.DataTableToHashtable(tb));
            return result.ToJson();
        }
        #endregion

        #region 评审专家是否生成
        /// <summary>
        /// 
        /// </summary>
        /// <param name="">判断评审专家是否生成</param>        
        /// <returns></returns>
        [ActionAttribute]
        public string AssessmentSpecialistYES(string Id, string FormId)
        {
            Power.Global.ViewResultModel result = Power.Global.ViewResultModel.Create(true, "");
            Power.Business.IBaseBusiness busi = Power.Business.BusinessFactory.CreateBusinessOperate("PS_BidOpen").FindByKey(Id);
            string strwhere = "Select * From Sus_Pur_ExpertReview Where InquiryCode  = '" + busi["BidInquiryCode"] + "'";
            DataTable tb1 = XCode.DataAccessLayer.DAL.QuerySQL(string.Format(strwhere));
            if (tb1.Rows.Count <= 0)
            {
                result.success = false;
                result.message = "您还未生成评审专家!";
                return result.ToJson();
            }
            else
            {

            }
            result.success = true;
            return result.ToJson();
        }
        #endregion

        #region 评审专家
        /// <summary>
        /// 
        /// </summary>
        /// <param name="Id">评标记录id </param>        
        /// <param name="FormId">专家评审表单id</param>
        /// <returns></returns>
        [ActionAttribute]
        public string AssessmentSpecialist(string Id, string FormId)
        {
            Power.Global.ViewResultModel result = Power.Global.ViewResultModel.Create(true, "");

            Power.Business.IBusinessOperate listOpt = Power.Business.BusinessFactory.CreateBusinessOperate("Sus_Pur_BidOpen_Assessors");
            Power.Business.IBusinessList list = listOpt.FindAll("MasterId", Id, Business.SearchFlag.IgnoreRight);

            Power.Business.IBaseBusiness busi = Power.Business.BusinessFactory.CreateBusinessOperate("PS_BidOpen").FindByKey(Id);

            Power.Business.IBaseBusiness busin = Power.Business.BusinessFactory.CreateBusiness("Sus_Pur_ExpertReview");


            string strwhere = "Select * From Sus_Pur_ExpertReview Where InquiryCode  = '" + busi["BidInquiryCode"] + "'";
            DataTable tb1 = XCode.DataAccessLayer.DAL.QuerySQL(string.Format(strwhere));
            #region  判断询价单编号是否生成
            if (tb1.Rows.Count <= 0)
            {
                #region 查找模板记录

                //查找模板数据
                string swhere = "MasterId in (select x1.TemplateId from PS_BID_BidOpen x1 where x1.Id = '" + Id + "')";
                XCode.IEntityList TemplateList = Power.Business.BusinessFactory.CreateBusinessOperate("Sus_Bid_TemplateList").GetEntityOperate().FindAll(swhere, "Sequ", "", 0, 0);
                #endregion

                WorkFlowManager workflowManager = new WorkFlowManager(this.session, this.Context);
                Dictionary<string, string> senduser = new Dictionary<string, string>();

                string strwhere1 = "select * from Sus_Pur_BidOpen_Assessors where MasterId='" + Id + "'";
                DataTable tb2 = XCode.DataAccessLayer.DAL.QuerySQL(string.Format(strwhere1));
                //判断是否有选专家
                if (tb2.Rows.Count > 0)
                {
                    foreach (Power.Business.IBaseBusiness it in list)
                    {

                        #region 专家库添加专家评审数据

                        #region 1、生成评审主表
                        Guid guid = Guid.NewGuid();
                        busin.SetItem("Id", guid);
                        busin.SetItem("InquiryName", busi["BidInquiryTitle"]);
                        busin.SetItem("InquiryCode", busi["BidInquiryCode"]);
                        busin.SetItem("DeviceId", busi["DeviceId"]);
                        busin.SetItem("DeviceCode", busi["DeviceCode"]);
                        busin.SetItem("Memo", it["Assessors"].ToString());
                        busin.SetItem("Name", busi["BidInquiryTitle"]);
                        busin.SetItem("Status", "20");
                        busin.SetItem("InquiryId", busi["BidInquiry_Guid"]);
                        busin.SetItem("ProjCode", session.EpsProjCode);
                        busin.SetItem("ProjName", session.EpsProjName);
                        busin.Save(System.ComponentModel.DataObjectMethodType.Insert);
                        #endregion
                        #region 2、生成评审内容

                        #region 2.1 内容供应商


                        //string strsql = "select x2.* from PS_BID_BidOpen x1 "
                        //    + " join Sus_Pur_FeedbackRegister x2 on x1.BidInquiry_Guid = x2.Inquiry_Guid"
                        //    + " where x1.Id = '92b43c86-1144-4245-af40-1073a56f12bc'";

                        string strsql = "select x2.RegHumId as RegHumId,x2.RegHumName as RegHumName,Sum(x2.SumPirce) as Score from Sus_Pur_ExpertReview x1 "
                             + " join Sus_Pur_FeedbackRegister x2 on x2.Inquiry_Guid = x1.InquiryId"
                              + " where x1.Id = '" + guid + "'  and x2.SumPirce <>0 "
                              + " and Versions=(select Max(Versions) from Sus_Pur_FeedbackRegister where Inquiry_Guid='" + busi["BidInquiry_Guid"] + "' ) "
                             + " group by x1.Id,x2.RegHumId,x2.RegHumName";
                        DataTable dtSupply = XCode.DataAccessLayer.DAL.QuerySQL(strsql);
                        foreach (DataRow temprow in dtSupply.Rows)
                        {
                            //专家评审，供应商子表
                            Power.Business.IBaseBusiness expertListItem = Power.Business.BusinessFactory.CreateBusiness("Sus_Pur_ExpertReviewList");
                            Guid listItemId = Guid.NewGuid();
                            expertListItem.SetItem("Id", listItemId);
                            expertListItem.SetItem("MasterId", guid);
                            expertListItem.SetItem("SupplierId", temprow["RegHumId"]);
                            expertListItem.SetItem("SupplierName", temprow["RegHumName"]);
                            expertListItem.SetItem("TotlePrice", temprow["Score"]);


                            expertListItem.Save(System.ComponentModel.DataObjectMethodType.Insert);
                            #region 2.1.1 供应商对应的评审内容
                            Dictionary<String, Guid> tempids = new Dictionary<String, Guid>();
                            foreach (XCode.IEntity item in TemplateList)
                            {
                                if (tempids.ContainsKey(item["Id"].ToString()) == false)
                                {
                                    tempids.Add(item["Id"].ToString(), Guid.NewGuid());

                                }
                                if (XCode.Common.Helper.IsNullKey(item["ParentId"]) == false && tempids.ContainsKey(item["ParentId"].ToString()) == false)
                                {
                                    tempids.Add(item["ParentId"].ToString(), Guid.NewGuid());
                                }

                                Power.Business.IBaseBusiness expertListItem2 = Power.Business.BusinessFactory.CreateBusiness("Sus_Pur_ExpertReviewList2");
                                expertListItem2.SetItem("Id", tempids[item["Id"].ToString()]);
                                if (XCode.Common.Helper.IsNullKey(item["ParentId"]) == false)
                                    expertListItem2.SetItem("ParentId", tempids[item["ParentId"].ToString()]);
                                else
                                    expertListItem2.SetItem("ParentId", null);
                                expertListItem2.SetItem("MasterId", listItemId);
                                expertListItem2.SetItem("Weight", item["Weight"]);
                                expertListItem2.SetItem("Items", item["Items"]);
                                expertListItem2.SetItem("Name", item["Name"]);
                                //.....
                                expertListItem2.Save(System.ComponentModel.DataObjectMethodType.Insert);
                            }
                            #endregion
                        }
                        #endregion

                        #endregion

                        #region 3、生成技术文件
                        TechnicalFileAdd(guid, busi["BidInquiry_Guid"].ToString());
                        qingbiaoTechnicalFileAdd(guid, Id);
                        #endregion

                        #endregion
                        #region 发送流程
                        senduser.Clear();
                        senduser.Add(it["AssessorsId"].ToString(), it["Assessors"].ToString());
                        workflowManager.autoStartWorkFlow(FormId, "Sus_Pur_ExpertReview", guid.ToString(), senduser);
                        #endregion

                    }
                }
                else
                {
                    result.success = false;
                    result.message = "生成专家评审信息的时候,至少需要选择一位专家!";
                    return result.ToJson();
                }

                #endregion
            }
            else
            {
                result.success = false;
                result.message = "询价单编号：" + busi["BidInquiryCode"] + "的专家数据已生成，无需再次生成！";
                return result.ToJson();
            }
            result.success = true;
            result.message = "专家评审信息生成成功！";
            return result.ToJson();
        }

        #endregion

        #region 技术文件添加
        private void TechnicalFileAdd(Guid Id, string BidInquiry_Guid)
        {
            string strSql = "select x2.* from Sus_Pur_FeedbackRegister x1 "
                          + " join PB_DocFiles x2 on x1.Id=x2.FolderId  "
                          + " where x1.Inquiry_Guid = '" + BidInquiry_Guid + "' and x2.SN='技术文件' "
                          + " and x1.Versions = (select max(Versions) from Sus_Pur_FeedbackRegister where Inquiry_Guid = '" + BidInquiry_Guid + "')";
            DataTable dtData = XCode.DataAccessLayer.DAL.QuerySQL(string.Format(strSql));
            Power.Business.IBaseBusiness busin = Power.Business.BusinessFactory.CreateBusiness("DocFile");
            foreach (DataRow row in dtData.Rows)
            {
                Guid docId = Guid.NewGuid();
                busin.SetItem("Id", docId);
                busin.SetItem("FolderId", Id);
                busin.SetItem("BOKeyWord", "Sus_Pur_ExpertReview");
                busin.SetItem("Code", row["Code"]);
                busin.SetItem("Name", row["Name"]);
                busin.SetItem("FileExt", row["FileExt"]);
                busin.SetItem("FileSize", row["FileSize"]);
                busin.SetItem("FileVersion", row["FileVersion"]);
                busin.SetItem("SecretLevel", row["SecretLevel"]);
                busin.SetItem("EncodeFlag", row["EncodeFlag"]);
                busin.SetItem("EncodeMethod", row["EncodeMethod"]);
                busin.SetItem("SourceFileId", row["SourceFileId"]);
                busin.SetItem("TemplateId", row["TemplateId"]);
                busin.SetItem("ServerUrl", row["ServerUrl"]);
                busin.SetItem("SN", row["SN"]);
                busin.SetItem("PublishFlag", row["PublishFlag"]);
                busin.SetItem("HandOverFlag", row["HandOverFlag"]);
                busin.SetItem("ArriveDate", row["ArriveDate"]);
                busin.SetItem("PlanDate", row["PlanDate"]);
                busin.SetItem("RequireDate", row["RequireDate"]);
                busin.SetItem("Designer", row["Designer"]);
                busin.SetItem("DesignOrganize", row["DesignOrganize"]);
                busin.SetItem("Charger", row["Charger"]);
                busin.SetItem("ChargeOrganize", row["ChargeOrganize"]);
                busin.SetItem("PageCount", row["PageCount"]);
                busin.SetItem("Labels", row["Labels"]);
                busin.SetItem("NeedTransfer", row["NeedTransfer"]);
                busin.SetItem("BIMJson", row["BIMJson"]);
                busin.SetItem("CheckFlag", row["CheckFlag"]);
                busin.SetItem("CheckHumId", row["CheckHumId"]);
                busin.SetItem("CheckName", row["CheckName"]);
                busin.SetItem("CheckDate", row["CheckDate"]);
                busin.SetItem("DeletFlag", row["DeletFlag"]);
                busin.SetItem("DeletHumId", row["DeletHumId"]);
                busin.SetItem("DeletName", row["DeletName"]);
                busin.SetItem("DeletDate", row["DeletDate"]);
                busin.SetItem("Deliverable_guid", row["Deliverable_guid"]);
                busin.SetItem("Deliverable_name", row["Deliverable_name"]);
                busin.SetItem("VersionKeyValue", row["VersionKeyValue"]);

                busin.SetItem("RegHumId", row["RegHumId"]);
                busin.SetItem("RegHumName", row["RegHumName"]);
                busin.SetItem("RegDate", row["RegDate"]);
                busin.SetItem("UpdHumId", row["UpdHumId"]);
                busin.SetItem("UpdHumName", row["UpdHumName"]);
                busin.SetItem("UpdDate", row["UpdDate"]);
                busin.SetItem("Memo", row["Memo"]);
                busin.Save(System.ComponentModel.DataObjectMethodType.Insert);



            }
        }
        #endregion
        #region 清标文件添加
        private void qingbiaoTechnicalFileAdd(Guid Id, string BidId)
        {
            string strSql = "select x2.* from PS_BID_BidOpen x1 "
                          + " join PB_DocFiles x2 on x1.Id=x2.FolderId  "
                          + " where x1.Id = '" + BidId + "'  and x2.SN='清标附件' ";
            DataTable dtData = XCode.DataAccessLayer.DAL.QuerySQL(string.Format(strSql));
            Power.Business.IBaseBusiness busin = Power.Business.BusinessFactory.CreateBusiness("DocFile");
            foreach (DataRow row in dtData.Rows)
            {
                Guid docId = Guid.NewGuid();
                busin.SetItem("Id", docId);
                busin.SetItem("FolderId", Id);
                busin.SetItem("BOKeyWord", "Sus_Pur_ExpertReview");
                busin.SetItem("Code", row["Code"]);
                busin.SetItem("Name", row["Name"]);
                busin.SetItem("FileExt", row["FileExt"]);
                busin.SetItem("FileSize", row["FileSize"]);
                busin.SetItem("FileVersion", row["FileVersion"]);
                busin.SetItem("SecretLevel", row["SecretLevel"]);
                busin.SetItem("EncodeFlag", row["EncodeFlag"]);
                busin.SetItem("EncodeMethod", row["EncodeMethod"]);
                busin.SetItem("SourceFileId", row["SourceFileId"]);
                busin.SetItem("TemplateId", row["TemplateId"]);
                busin.SetItem("ServerUrl", row["ServerUrl"]);
                busin.SetItem("SN", row["SN"]);
                busin.SetItem("PublishFlag", row["PublishFlag"]);
                busin.SetItem("HandOverFlag", row["HandOverFlag"]);
                busin.SetItem("ArriveDate", row["ArriveDate"]);
                busin.SetItem("PlanDate", row["PlanDate"]);
                busin.SetItem("RequireDate", row["RequireDate"]);
                busin.SetItem("Designer", row["Designer"]);
                busin.SetItem("DesignOrganize", row["DesignOrganize"]);
                busin.SetItem("Charger", row["Charger"]);
                busin.SetItem("ChargeOrganize", row["ChargeOrganize"]);
                busin.SetItem("PageCount", row["PageCount"]);
                busin.SetItem("Labels", row["Labels"]);
                busin.SetItem("NeedTransfer", row["NeedTransfer"]);
                busin.SetItem("BIMJson", row["BIMJson"]);
                busin.SetItem("CheckFlag", row["CheckFlag"]);
                busin.SetItem("CheckHumId", row["CheckHumId"]);
                busin.SetItem("CheckName", row["CheckName"]);
                busin.SetItem("CheckDate", row["CheckDate"]);
                busin.SetItem("DeletFlag", row["DeletFlag"]);
                busin.SetItem("DeletHumId", row["DeletHumId"]);
                busin.SetItem("DeletName", row["DeletName"]);
                busin.SetItem("DeletDate", row["DeletDate"]);
                busin.SetItem("Deliverable_guid", row["Deliverable_guid"]);
                busin.SetItem("Deliverable_name", row["Deliverable_name"]);
                busin.SetItem("VersionKeyValue", row["VersionKeyValue"]);

                busin.SetItem("RegHumId", row["RegHumId"]);
                busin.SetItem("RegHumName", row["RegHumName"]);
                busin.SetItem("RegDate", row["RegDate"]);
                busin.SetItem("UpdHumId", row["UpdHumId"]);
                busin.SetItem("UpdHumName", row["UpdHumName"]);
                busin.SetItem("UpdDate", row["UpdDate"]);
                busin.SetItem("Memo", row["Memo"]);
                busin.Save(System.ComponentModel.DataObjectMethodType.Insert);



            }
        }
        #endregion

        #region 评标报告数据生成
        [ActionAttribute]
        public string InquiryWizard(string BidInquiryId)
        {
            Power.Global.ViewResultModel result = Power.Global.ViewResultModel.Create(true, "");
            string strsql = "select y.SupplierId,y.SupplierName,y.AdminName,y.TotlePrice,y.SumPirce,y.Versions,sum(y.Score)/(select COUNT(*) from Sus_Pur_ExpertReview where InquiryId='" + BidInquiryId + "') as Score from Sus_Bid_Inquiry x "
                          + "inner join (select MasterId, SupplierName, SupplierId, TotlePrice, Score, b.AdminName, c.InquiryId,e.SumPirce,e.Versions from Sus_Pur_ExpertReviewList a "
                          + "inner join SB_SupplierRegistration b on a.SupplierId= b.Sup_HumanId "
                          + "inner join Sus_Pur_ExpertReview c on a.MasterId = c.Id   left join Sus_Pur_FeedbackRegister e on e.Inquiry_Guid = c.InquiryId  and b.Sup_HumanId = e.RegHumId  "
                          + "group by a.MasterId, a.SupplierName, a.SupplierId, a.TotlePrice, Score, "
                          + "b.AdminName, c.InquiryId,e.SumPirce,e.Versions) y on x.Id = y.InquiryId "
                          + "where x.Id = '" + BidInquiryId + "' "
                          + "group by x.Id,y.SupplierName,y.AdminName,y.TotlePrice,y.SupplierId,y.SumPirce,y.Versions   order by y.Versions asc";
            DataTable tb = XCode.DataAccessLayer.DAL.QuerySQL(string.Format(strsql));
            result.data.Add("values", Power.Global.BusiHelper.DataTableToHashtable(tb));
            return result.ToJson();
        }

        #endregion

        #region 按条件过滤查询显示
        [ActionAttribute]
        public string ConditionalInquiry1(string Whether)
        {
            Power.Global.ViewResultModel result = Power.Global.ViewResultModel.Create(true, "");
            if (Whether == "是否招标")
            {
                string strSql = "select * from V_BiddingStatusTracking where IsTenderFile=1";
                DataReturned(result, strSql);
            }
            else if (Whether == "是否询价反馈")
            {
                string strSql = "select * from V_BiddingStatusTracking where WhetherFeedback=1";
                DataReturned(result, strSql);
            }
            else if (Whether == "是否开标")
            {
                string strSql = "select * from V_BiddingStatusTracking where IsOpen=1";
                DataReturned(result, strSql);
            }
            else if (Whether == "是否评审")
            {
                string strSql = "select * from V_BiddingStatusTracking where IsReview=1";
                DataReturned(result, strSql);
            }
            else if (Whether == "是否评标")
            {
                string strSql = "select * from V_BiddingStatusTracking where IsReviewReport=1";
                DataReturned(result, strSql);
            }
            else if (Whether == "是否中标")
            {
                string strSql = "select * from V_BiddingStatusTracking where IsnoticeOfAward=1";
                DataReturned(result, strSql);
            }
            else
            {
                result.success = false;
                result.message = "需要查询的值有误";
            }

            return result.ToJson();
        }
        #endregion

        #region 按条件点击按钮过滤查询显示
        [ActionAttribute]
        public string ConditionalInquiry2(string num1, string num2)
        {
            Power.Global.ViewResultModel result = Power.Global.ViewResultModel.Create(true, "");
            if (num1 == "编号")
            {
                string strSql = "select * from V_BiddingStatusTracking where Code like '%" + num2 + "%'";
                DataReturned(result, strSql);
            }
            else if (num1 == "标题")
            {
                string strSql = "select * from V_BiddingStatusTracking where Title like '%" + num2 + "%'";
                DataReturned(result, strSql);
            }
            else if (num1 == "录入日期")
            {
                string strSql = "select * from V_BiddingStatusTracking where RegDate like '%" + num2 + "%'";
                DataReturned(result, strSql);
            }
            else if (num1 == "录入人名称")
            {
                string strSql = "select * from V_BiddingStatusTracking where RegHumName like '%" + num2 + "%'";
                DataReturned(result, strSql);
            }
            else
            {
                result.success = false;
                result.message = "需要查询的值有误";
            }

            return result.ToJson();
        }
        #endregion

        #region 招标询价商务清标选择数据带过来
        /// <param name="Id">招标询价Id</param>
        /// <param name="GuideId">向导Id</param>
        /// <returns></returns>
        [ActionAttribute]
        public string WizardDataSelection(string Id, string GuideId, string flag)
        {
            Power.Global.ViewResultModel result = Power.Global.ViewResultModel.Create(true, "");
            #region  商务清标
            //if (flag == "C")
            //{
            //    #region 商务清标
            //    //创建键值对存储Id和ParentId
            //    Dictionary<String, Guid> tempids = new Dictionary<String, Guid>();
            //    //取招标询价主表数据
            //    Power.Business.IBaseBusiness bus = Power.Business.BusinessFactory.CreateBusinessOperate("Sus_Bid_Inquiry").FindByKey(Id);
            //    //根据招标询价Id获取到商务清标的值
            //    Power.Business.IBusinessList budget = Power.Business.BusinessFactory.CreateBusinessOperate("Sus_Bid_InquiryList_C").FindAll("MasterId", Id, Business.SearchFlag.Default);

            //    #region JSON格式字符串转换成对象
            //    var srt = JsonConvert.DeserializeObject<string[]>(GuideId);
            //    string sqlSrt = "";
            //    foreach (object itme in srt)
            //    {
            //        sqlSrt += "'" + itme + "'" + ",";
            //    }
            //    sqlSrt = sqlSrt.TrimEnd(',');
            //    #endregion

            //    #region 1、查询出向导中所有选中的数据
            //    string strsql = "select * from V_Sus_parameterTemplate where Id in({0})";
            //    DataTable dtData = XCode.DataAccessLayer.DAL.QuerySQL(string.Format(strsql, sqlSrt));
            //    #endregion
            //    #region 找到所有根节点
            //    List<String> listRoot = new List<string>();
            //    #region 判断是否有数据，有则赋值
            //    if (dtData.Rows.Count == 0)
            //    {
            //        result.success = false;
            //        result.message = "请选择一条数据！";
            //        return result.ToJson();
            //    }
            //    else
            //    {
            //        DataTable dtC = Power.Business.BusinessFactory.CreateBusinessOperate("Sus_Bid_InquiryList_C").FindAllByTable(string.Format("MasterId='{0}'", Id), "", "NodeId", 0, 0, SearchFlag.IgnoreRight);
            //        foreach (DataRow row in dtData.Rows)
            //        {
            //            DataRow[] rowsSelect = dtC.Select(String.Format("NodeId='{0}'", row["Id"]));
            //            if (rowsSelect != null && rowsSelect.Length != 0)
            //            {
            //                continue;
            //            }
            //            //1、insert 左边
            //            if (tempids.ContainsKey(row["Id"].ToString()) == false)
            //            {
            //                tempids.Add(row["Id"].ToString(), Guid.NewGuid());
            //            }
            //            if (XCode.Common.Helper.IsNullKey(row["ParentId"]) == false && tempids.ContainsKey(row["ParentId"].ToString()) == false)
            //            {
            //                tempids.Add(row["ParentId"].ToString(), Guid.NewGuid());
            //            }

            //            //2、insert 右边
            //            Power.Business.IBusinessList list = Power.Business.BusinessFactory.CreateBusinessOperate("Sus_parameterTemplateList").FindAll("MasterId", row["Id"].ToString(), Business.SearchFlag.Default);
            //            string temp = "";
            //            foreach (Power.Business.IBaseBusiness item in list)
            //            {
            //                string temp2 = "";
            //                if (!(XCode.Common.Helper.IsNullKey(item["property"])))
            //                {
            //                    temp2 += string.Format("{0}{1}", item["property"], "    ");
            //                    if (!(XCode.Common.Helper.IsNullKey(item["SimpleValue"])))
            //                    {
            //                        temp2 = "";
            //                        temp2 += string.Format("{0};{1}{2}", item["property"], item["SimpleValue"], "    ");
            //                        if (!(XCode.Common.Helper.IsNullKey(item["Unit"])))
            //                        {
            //                            temp2 = "";
            //                            temp2 += string.Format("{0};{1};{2}{3}", item["property"], item["SimpleValue"], item["Unit"], "    ");
            //                        }
            //                    }else
            //                    {
            //                        if (!(XCode.Common.Helper.IsNullKey(item["Unit"])))
            //                        {
            //                            temp2 = "";
            //                            temp2 += string.Format("{0};{1}{2}", item["property"], item["Unit"], "    ");
            //                        }
            //                    }
            //                }
            //                else if (!(XCode.Common.Helper.IsNullKey(item["SimpleValue"])))
            //                {
            //                    temp2 += string.Format("{0}{1}", item["SimpleValue"], "    ");
            //                    if (!(XCode.Common.Helper.IsNullKey(item["Unit"])))
            //                    {
            //                        temp2 = "";
            //                        temp2 += string.Format("{0};{1}{2}", item["SimpleValue"], item["Unit"], "    ");
            //                    }
            //                }
            //                else if (!(XCode.Common.Helper.IsNullKey(item["Unit"])))
            //                {
            //                    temp2 += string.Format("{0}{1}", item["Unit"], "    ");
            //                }
            //                //根节点不需要insert
            //                if (listRoot.Contains(row["Id"].ToString().ToLower()))
            //                    continue;
            //                //赋值给上面的变量
            //                temp += temp2;
            //            }
            //            //取招标询价的商务子表信息
            //            Power.Business.IBaseBusiness businList = Power.Business.BusinessFactory.CreateBusiness("Sus_Bid_InquiryList_C");
            //            businList.SetItem("Id", tempids[row["Id"].ToString()]);
            //            businList.SetItem("MasterId", bus["Id"].ToString());
            //            if (XCode.Common.Helper.IsNullKey(row["ParentId"]) == false)
            //                businList.SetItem("ParentId", tempids[row["ParentId"].ToString()]);
            //            else
            //                businList.SetItem("ParentId", null);
            //            temp = temp.TrimEnd("    ");
            //            businList.SetItem("Code", row["code"]);
            //            businList.SetItem("Name", row["name"]);
            //            businList.SetItem("Dept", row["Unit"]);
            //            businList.SetItem("NodeId", row["Id"]);
            //            businList.SetItem("Parameter", temp);
            //            businList.SetItem("Sequ", row["Sequ"]);
            //            businList.Save(System.ComponentModel.DataObjectMethodType.Insert);
            //            temp = "";
            //            //tempids.Clear();

            //        }
            //    }
            //    #endregion
            //    #endregion
            //    #endregion
            //   

            //}
            #endregion
            #region  商务清标

            if (flag == "C")
            {
                Dictionary<String, Guid> tempids = new Dictionary<String, Guid>();
                Dictionary<String, Guid> tempidsList = new Dictionary<String, Guid>();

                #region JSON格式字符串转换成对象
                var srt = JsonConvert.DeserializeObject<string[]>(GuideId);
                string sqlSrt = "";
                foreach (object itme in srt)
                {
                    sqlSrt += "'" + itme + "'" + ",";
                }
                sqlSrt = sqlSrt.TrimEnd(',');
                #endregion

                #region 查询出向导中所有选中主表的数据
                string strsql = "select * from V_Sus_parameterTemplate where Id in({0})";
                DataTable dtData = XCode.DataAccessLayer.DAL.QuerySQL(string.Format(strsql, sqlSrt));
                #endregion
                #region 找到所有根节点
                List<String> listRoot = new List<string>();
                #region 判断是否有数据，有则赋值
                if (dtData.Rows.Count == 0)
                {
                    result.success = false;
                    result.message = "请选择一条数据！";
                    return result.ToJson();
                }
                else
                {
                    #region 2.1、取数
                    DataTable dtC = Power.Business.BusinessFactory.CreateBusinessOperate("Sus_Bid_InquiryList_C").FindAllByTable(string.Format("MasterId='{0}'", Id), "", "NodeId", 0, 0, SearchFlag.IgnoreRight);
                    foreach (DataRow row in dtData.Rows)
                    {
                        if (tempids.ContainsKey(row["Id"].ToString()) == false)
                        {
                            tempids.Add(row["Id"].ToString(), Guid.NewGuid());
                        }
                        if (XCode.Common.Helper.IsNullKey(row["ParentId"]) == false && tempids.ContainsKey(row["ParentId"].ToString()) == false)
                        {
                            tempids.Add(row["ParentId"].ToString(), Guid.NewGuid());
                        }
                        DataRow[] rowsSelect = dtC.Select(String.Format("NodeId='{0}'", row["Id"]));
                        if (rowsSelect != null && rowsSelect.Length != 0)
                        {
                            continue;
                        }
                        Power.Business.IBaseBusiness TBT = Power.Business.BusinessFactory.CreateBusiness("Sus_Bid_InquiryList_C");
                        Power.Business.IBusinessList list = Power.Business.BusinessFactory.CreateBusinessOperate("Sus_parameterTemplateList").FindAll("MasterId", row["Id"].ToString(), Business.SearchFlag.Default);
                        string temp = "";
                        foreach (Power.Business.IBaseBusiness item in list)
                        {
                            string temp2 = "";
                            if (!(XCode.Common.Helper.IsNullKey(item["property"])))
                            {
                                temp2 += string.Format("{0}{1}", item["property"], "    ");
                                if (!(XCode.Common.Helper.IsNullKey(item["SimpleValue"])))
                                {
                                    temp2 = "";
                                    temp2 += string.Format("{0};{1}{2}", item["property"], item["SimpleValue"], "    ");
                                    if (!(XCode.Common.Helper.IsNullKey(item["Unit"])))
                                    {
                                        temp2 = "";
                                        temp2 += string.Format("{0};{1};{2}{3}", item["property"], item["SimpleValue"], item["Unit"], "    ");
                                    }
                                    if (!(XCode.Common.Helper.IsNullKey(item["Amount"])))
                                    {
                                        temp2 = "";
                                        temp2 += string.Format("{0};{1};{2};{3}{4}", item["property"], item["SimpleValue"], item["Amount"], item["Unit"], "    ");
                                    }
                                }
                                else
                                {
                                    if (!(XCode.Common.Helper.IsNullKey(item["Unit"])))
                                    {
                                        temp2 = "";
                                        temp2 += string.Format("{0};{1}{2}", item["property"], item["Unit"], "    ");
                                    }
                                }
                            }
                            else if (!(XCode.Common.Helper.IsNullKey(item["SimpleValue"])))
                            {
                                temp2 += string.Format("{0}{1}", item["SimpleValue"], "    ");
                                if (!(XCode.Common.Helper.IsNullKey(item["Unit"])))
                                {
                                    temp2 = "";
                                    temp2 += string.Format("{0};{1}{2}", item["SimpleValue"], item["Unit"], "    ");
                                }
                            }
                            else if (!(XCode.Common.Helper.IsNullKey(item["Unit"])))
                            {
                                temp2 += string.Format("{0}{1}", item["Unit"], "    ");
                            }

                            //赋值给上面的变量
                            temp += temp2;
                        }


                        TBT.SetItem("Id", tempids[row["Id"].ToString()]);
                        if (XCode.Common.Helper.IsNullKey(row["ParentId"]) == false)
                            TBT.SetItem("ParentId", tempids[row["ParentId"].ToString()]);
                        else
                            TBT.SetItem("ParentId", null);
                        TBT.SetItem("MasterId", Id);
                        temp = temp.TrimEnd("    ");
                        TBT.SetItem("Code", row["code"]);
                        TBT.SetItem("Name", row["name"]);
                        TBT.SetItem("Dept", row["Unit"]);
                        TBT.SetItem("Parameter", temp);
                        TBT.Save(System.ComponentModel.DataObjectMethodType.Insert);
                        #endregion
                        #endregion
                        #region 2.2、取数
                        Power.Business.IBusinessList TemplateList = Power.Business.BusinessFactory.CreateBusinessOperate("Sus_parameterTemplateList").FindAll("MasterId", row["Id"], Business.SearchFlag.Default);
                        foreach (Power.Business.IBaseBusiness item in TemplateList)
                        {
                            if (tempidsList.ContainsKey(item["Id"].ToString()) == false)
                            {
                                tempidsList.Add(item["Id"].ToString(), Guid.NewGuid());
                            }
                            if (XCode.Common.Helper.IsNullKey(item["ParentId"]) == false && tempidsList.ContainsKey(item["ParentId"].ToString()) == false)
                            {
                                tempidsList.Add(item["ParentId"].ToString(), Guid.NewGuid());
                            }
                            Power.Business.IBaseBusiness TBTList = Power.Business.BusinessFactory.CreateBusiness("Sus_Bid_InquiryList_C_Del");
                            TBTList.SetItem("Id", tempidsList[item["Id"].ToString()]);
                            if (XCode.Common.Helper.IsNullKey(item["ParentId"]) == false)
                                TBTList.SetItem("ParentId", tempidsList[item["ParentId"].ToString()]);
                            else
                                TBTList.SetItem("ParentId", null);
                            TBTList.SetItem("MasterId", tempids[row["Id"].ToString()]);
                            TBTList.SetItem("property", item["property"]);
                            TBTList.SetItem("SimpleValue", item["SimpleValue"]);
                            TBTList.SetItem("Amount", item["Amount"]);
                            TBTList.SetItem("Unit", item["Unit"]);
                            TBTList.SetItem("Memo", item["Memo"]);
                            TBTList.Save(System.ComponentModel.DataObjectMethodType.Insert);
                        }
                    }
                    #endregion

                    return result.ToJson();
                }
            }
            #endregion
            #endregion
            else if (flag == "D")
            {
                #region 设计评审选择商务清标
                //创建键值对存储Id和ParentId
                Dictionary<String, Guid> tempids = new Dictionary<String, Guid>();
                //取招标询价主表数据
                Power.Business.IBaseBusiness bus = Power.Business.BusinessFactory.CreateBusinessOperate("Sus_TechnicalBook").FindByKey(Id);
                //根据招标询价Id获取到商务清标的值
                Power.Business.IBusinessList budget = Power.Business.BusinessFactory.CreateBusinessOperate("Sus_TechnicalBook_T").FindAll("MasterId", Id, Business.SearchFlag.Default);

                #region JSON格式字符串转换成对象
                var srt = JsonConvert.DeserializeObject<string[]>(GuideId);
                string sqlSrt = "";
                foreach (object itme in srt)
                {
                    sqlSrt += "'" + itme + "'" + ",";
                }
                sqlSrt = sqlSrt.TrimEnd(',');
                #endregion

                #region 1、查询出向导中所有选中的数据
                string strsql = "select * from V_Sus_parameterTemplate where Id in({0})";
                DataTable dtData = XCode.DataAccessLayer.DAL.QuerySQL(string.Format(strsql, sqlSrt));
                #endregion

                #region 判断是否有数据，有则赋值
                if (dtData.Rows.Count == 0)
                {
                    result.success = false;
                    result.message = "请选择一条数据！";
                    return result.ToJson();
                }
                else
                {
                    DataTable dtC = Power.Business.BusinessFactory.CreateBusinessOperate("Sus_TechnicalBook_T").FindAllByTable(string.Format("MasterId='{0}'", Id), "", "NodeId", 0, 0, SearchFlag.IgnoreRight);
                    foreach (DataRow row in dtData.Rows)
                    {
                        DataRow[] rowsSelect = dtC.Select(String.Format("NodeId='{0}'", row["Id"]));
                        if (rowsSelect != null && rowsSelect.Length != 0)
                        {
                            continue;
                        }
                        //1、insert 左边
                        if (tempids.ContainsKey(row["Id"].ToString()) == false)
                        {
                            tempids.Add(row["Id"].ToString(), Guid.NewGuid());
                        }
                        if (XCode.Common.Helper.IsNullKey(row["ParentId"]) == false && tempids.ContainsKey(row["ParentId"].ToString()) == false)
                        {
                            tempids.Add(row["ParentId"].ToString(), Guid.NewGuid());
                        }

                        //2、insert 右边
                        Power.Business.IBusinessList list = Power.Business.BusinessFactory.CreateBusinessOperate("Sus_parameterTemplateList").FindAll("MasterId", row["Id"].ToString(), Business.SearchFlag.Default);
                        string temp = "";

                        foreach (Power.Business.IBaseBusiness item in list)
                        {
                            string temp2 = "";
                            if (!(XCode.Common.Helper.IsNullKey(item["property"])))
                            {
                                temp2 += string.Format("{0}{1}", item["property"], "    ");
                                if (!(XCode.Common.Helper.IsNullKey(item["SimpleValue"])))
                                {
                                    temp2 = "";
                                    temp2 += string.Format("{0};{1}{2}", item["property"], item["SimpleValue"], "    ");
                                    if (!(XCode.Common.Helper.IsNullKey(item["Unit"])))
                                    {
                                        temp2 = "";
                                        temp2 += string.Format("{0};{1};{2}{3}", item["property"], item["SimpleValue"], item["Unit"], "    ");
                                    }
                                }
                                else
                                {
                                    if (!(XCode.Common.Helper.IsNullKey(item["Unit"])))
                                    {
                                        temp2 = "";
                                        temp2 += string.Format("{0};{1}{2}", item["property"], item["Unit"], "    ");
                                    }
                                }
                            }
                            else if (!(XCode.Common.Helper.IsNullKey(item["SimpleValue"])))
                            {
                                temp2 += string.Format("{0}{1}", item["SimpleValue"], "    ");
                                if (!(XCode.Common.Helper.IsNullKey(item["Unit"])))
                                {
                                    temp2 = "";
                                    temp2 += string.Format("{0};{1}{2}", item["SimpleValue"], item["Unit"], "    ");
                                }
                            }
                            else if (!(XCode.Common.Helper.IsNullKey(item["Unit"])))
                            {
                                temp2 += string.Format("{0}{1}", item["Unit"], "    ");
                            }

                            //赋值给上面的变量
                            temp += temp2;

                        }

                        //取招标询价的商务子表信息
                        Power.Business.IBaseBusiness businList = Power.Business.BusinessFactory.CreateBusiness("Sus_TechnicalBook_T");
                        businList.SetItem("Id", tempids[row["Id"].ToString()]);
                        businList.SetItem("MasterId", bus["Id"].ToString());
                        if (XCode.Common.Helper.IsNullKey(row["ParentId"]) == false)
                            businList.SetItem("ParentId", tempids[row["ParentId"].ToString()]);
                        else
                            businList.SetItem("ParentId", null);
                        temp = temp.TrimEnd(":    ");
                        businList.SetItem("Code", row["code"]);
                        businList.SetItem("Name", row["name"]);
                        businList.SetItem("Dept", row["Unit"]);
                        businList.SetItem("NodeId", row["Id"]);
                        businList.SetItem("Specification", temp);
                        businList.SetItem("Sequ", row["Sequ"]);

                        businList.Save(System.ComponentModel.DataObjectMethodType.Insert);
                        temp = "";
                        //tempids.Clear();

                    }
                }
                #endregion
                #endregion
            }

            else if (flag == "T")
            {
                #region 技术清标
                //创建键值对存储Id和ParentId
                Dictionary<String, Guid> tempids = new Dictionary<String, Guid>();
                //取招标询价主表数据
                Power.Business.IBaseBusiness bus = Power.Business.BusinessFactory.CreateBusinessOperate("Sus_Bid_Inquiry").FindByKey(Id);
                //根据招标询价Id获取到商务清标的值
                Power.Business.IBusinessList budget = Power.Business.BusinessFactory.CreateBusinessOperate("Sus_Bid_InquiryList_T").FindAll("MasterId", Id, Business.SearchFlag.Default);
                #region JSON格式字符串转换成对象
                var srt = JsonConvert.DeserializeObject<string[]>(GuideId);
                string sqlSrt = "";
                foreach (object itme in srt)
                {
                    sqlSrt += "'" + itme + "'" + ",";
                }
                sqlSrt = sqlSrt.TrimEnd(',');
                #endregion

                #region 1、查询出向导中所有选中的数据
                string strsql = "select * from V_Sus_parameterTemplate where Id in({0})";
                DataTable dtData = XCode.DataAccessLayer.DAL.QuerySQL(string.Format(strsql, sqlSrt));
                #endregion

                #region 判断是否有数据，有则赋值
                if (dtData.Rows.Count == 0)
                {
                    result.success = false;
                    result.message = "请选择一条数据！";
                    return result.ToJson();
                }
                else
                {
                    DataTable dtC = Power.Business.BusinessFactory.CreateBusinessOperate("Sus_Bid_InquiryList_T").FindAllByTable(string.Format("MasterId='{0}'", Id), "", "NodeId", 0, 0, SearchFlag.IgnoreRight);
                    foreach (DataRow row in dtData.Rows)
                    {
                        DataRow[] rowsSelect = dtC.Select(String.Format("NodeId='{0}'", row["Id"]));
                        if (rowsSelect != null && rowsSelect.Length != 0)
                        {
                            continue;
                        }
                        //1、insert 左边
                        if (tempids.ContainsKey(row["Id"].ToString()) == false)
                        {
                            tempids.Add(row["Id"].ToString(), Guid.NewGuid());
                        }
                        if (XCode.Common.Helper.IsNullKey(row["ParentId"]) == false && tempids.ContainsKey(row["ParentId"].ToString()) == false)
                        {
                            tempids.Add(row["ParentId"].ToString(), Guid.NewGuid());
                        }

                        //2、insert 右边
                        Power.Business.IBusinessList list = Power.Business.BusinessFactory.CreateBusinessOperate("Sus_parameterTemplateList").FindAll("MasterId", row["Id"].ToString(), Business.SearchFlag.Default);
                        string temp = "";
                        foreach (Power.Business.IBaseBusiness item in list)
                        {
                            string temp2 = "";
                            if (!(XCode.Common.Helper.IsNullKey(item["property"])))
                            {
                                temp2 += string.Format("{0}{1}", item["property"], "    ");
                                if (!(XCode.Common.Helper.IsNullKey(item["SimpleValue"])))
                                {
                                    temp2 = "";
                                    temp2 += string.Format("{0};{1}{2}", item["property"], item["SimpleValue"], "    ");
                                    if (!(XCode.Common.Helper.IsNullKey(item["Unit"])))
                                    {
                                        temp2 = "";
                                        temp2 += string.Format("{0};{1};{2}{3}", item["property"], item["SimpleValue"], item["Unit"], "    ");
                                    }
                                }
                                else
                                {
                                    if (!(XCode.Common.Helper.IsNullKey(item["Unit"])))
                                    {
                                        temp2 = "";
                                        temp2 += string.Format("{0};{1}{2}", item["property"], item["Unit"], "    ");
                                    }
                                }
                            }
                            else if (!(XCode.Common.Helper.IsNullKey(item["SimpleValue"])))
                            {
                                temp2 += string.Format("{0}{1}", item["SimpleValue"], "    ");
                                if (!(XCode.Common.Helper.IsNullKey(item["Unit"])))
                                {
                                    temp2 = "";
                                    temp2 += string.Format("{0};{1}{2}", item["SimpleValue"], item["Unit"], "    ");
                                }
                            }
                            else if (!(XCode.Common.Helper.IsNullKey(item["Unit"])))
                            {
                                temp2 += string.Format("{0}{1}", item["Unit"], "    ");
                            }

                            //赋值给上面的变量
                            temp += temp2;
                        }
                        //取招标询价的商务子表信息
                        Power.Business.IBaseBusiness businList = Power.Business.BusinessFactory.CreateBusiness("Sus_Bid_InquiryList_T");
                        businList.SetItem("Id", tempids[row["Id"].ToString()]);
                        businList.SetItem("MasterId", bus["Id"].ToString());
                        if (XCode.Common.Helper.IsNullKey(row["ParentId"]) == false)
                            businList.SetItem("ParentId", tempids[row["ParentId"].ToString()]);
                        else
                            businList.SetItem("ParentId", null);
                        temp = temp.TrimEnd("    ");
                        businList.SetItem("Code", row["code"]);
                        businList.SetItem("Name", row["name"]);
                        businList.SetItem("Dept", row["Unit"]);
                        businList.SetItem("NodeId", row["Id"]);
                        businList.SetItem("Specification", temp);
                        businList.SetItem("Sequ", row["Sequ"]);
                        businList.Save(System.ComponentModel.DataObjectMethodType.Insert);
                        temp = "";
                        //tempids.Clear();
                    }
                }
                #endregion
                #endregion
            }
            else
            {
                result.success = false;
                result.message = flag + "无法识别！";
            }
            return result.ToJson();
        }
        #endregion

        #region 中标通知书发送邮件
        [ActionAttribute]
        public string NoticeOfAwardEmail(string Id)
        {
            Power.Global.ViewResultModel result = Power.Global.ViewResultModel.Create(true, "");
            Power.Business.IBusinessOperate bidOpt = Power.Business.BusinessFactory.CreateBusinessOperate("PS_BidWinNotice");
            Power.Business.IBaseBusiness bidMaster = bidOpt.FindByKey(Id);
            Power.Business.IBaseBusiness busin = Power.Business.BusinessFactory.CreateBusiness("PowerMessage");
            Power.Global.MailDataPack model = new Power.Global.MailDataPack();
            model.useparam = Power.Global.EUseMailDefaultParma.UseSmtpUser; //使用Power.ConfigEdit.exe 中配置的发送邮件参数

            #region 获取邮件模板
            String swhere = "BaseDataId in (select x1.Id from PB_BaseData x1 where x1.DataType= 'Bid_Otific_Mail_Template')";
            Power.Business.IBusinessList basedataList = Power.Business.BusinessFactory.CreateBusinessOperate("BaseDataList").FindAll(swhere, "", "", 0, 0, SearchFlag.IgnoreRight);
            String mailtitle = "";
            String mailcontent = "";
            foreach (Power.Business.IBaseBusiness item in basedataList)
            {
                if (item["Code"] != null && item["Code"].ToString() == "Title")
                {
                    mailtitle = item["Name"].ToString();
                }
                if (item["Code"] != null && item["Code"].ToString() == "Content")
                {
                    mailcontent = item["Name"].ToString();
                }
            }
            #endregion

            //设置邮件的收件人
            string address = "";
            string subject = "";
            string content = "";
            if (!(XCode.Common.Helper.IsNullKey(bidMaster["LinkEmail"]))/* && int.Parse(bidMaster["Status"].ToString()) == 50*/)
            {
                address += bidMaster["LinkEmail"];
                #region 替换模板参数

                #region 替换模板参数
                subject = processMessageContent(mailtitle, this.session, bidMaster, bidOpt);
                subject = processMessageContent(subject, this.session, bidMaster, bidOpt);
                content = processMessageContent(mailcontent, this.session, bidMaster, bidOpt);
                content = processMessageContent(content, this.session, bidMaster, bidOpt);
                #endregion
                #endregion
                //收件人地址
                model.msg_to = address;
                //邮件标题
                model.msg_subject = subject;
                //邮件内容
                model.msg_content = content;

                String errorinfo = "";
                if (Power.Service.MailService.MailBLL.SendMail(model, out errorinfo))
                {
                    address = "";
                    string dateTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    busin.SetItem("Id", Guid.NewGuid());
                    busin.SetItem("Title", subject);
                    busin.SetItem("FromDate", dateTime);
                    busin.SetItem("ToHumanId", bidMaster["WinBidSupplierId"]);
                    busin.SetItem("ToHumanName", bidMaster["WinBidSupplier"]);
                    busin.SetItem("FromHumanId", bidMaster["RegHumId"]);
                    busin.SetItem("FromHumanName", bidMaster["RegHumName"]);
                    busin.SetItem("MessageType", "notify");
                    busin.SetItem("KeyValue", bidMaster["Id"]);
                    busin.SetItem("IsMail", 0);
                    busin.SetItem("IsTextMessage", 0);
                    busin.SetItem("IsPowerMessage", 1);
                    busin.SetItem("IsDeviceMessage", 0);
                    busin.SetItem("ContentText", content);
                    busin.Save(System.ComponentModel.DataObjectMethodType.Insert);
                    result.success = true;
                    result.message = "邮件和系统消息都已发送成功!";
                }
                else
                {
                    //发送邮件失败，错误原因在 errorinfo 里面
                    result.success = false;
                    result.message = "邮件发送失败!";
                    result.message = errorinfo;
                }
            }
            return result.ToJson();
        }

        #endregion
        #region 合同付款发送邮件
        [ActionAttribute]
        public string PS_SubContractApplyMoneyNewEmail(string Id)
        {
            Power.Global.ViewResultModel result = Power.Global.ViewResultModel.Create(true, "");
            Power.Business.IBusinessOperate bidOpt = Power.Business.BusinessFactory.CreateBusinessOperate("PS_SubContractApplyMoneyNew");
            Power.Business.IBaseBusiness bidMaster = bidOpt.FindByKey(Id);
            Power.Business.IBaseBusiness busin = Power.Business.BusinessFactory.CreateBusiness("PowerMessage");
            Power.Global.MailDataPack model = new Power.Global.MailDataPack();
            model.useparam = Power.Global.EUseMailDefaultParma.UseSmtpUser; //使用Power.ConfigEdit.exe 中配置的发送邮件参数

            #region 获取邮件模板
            String swhere = "BaseDataId in (select x1.Id from PB_BaseData x1 where x1.DataType= 'Bid_Otific_Mail_Template')";
            Power.Business.IBusinessList basedataList = Power.Business.BusinessFactory.CreateBusinessOperate("BaseDataList").FindAll(swhere, "", "", 0, 0, SearchFlag.IgnoreRight);
            String mailtitle = "";
            String mailcontent = "";
            foreach (Power.Business.IBaseBusiness item in basedataList)
            {
                if (item["Code"] != null && item["Code"].ToString() == "Title")
                {
                    mailtitle = item["Name"].ToString();
                }
                if (item["Code"] != null && item["Code"].ToString() == "Content")
                {
                    mailcontent = item["Name"].ToString();
                }
            }
            #endregion

            //设置邮件的收件人
            string address = "";
            string subject = "";
            string content = "";
            if (!(XCode.Common.Helper.IsNullKey(bidMaster["Email"])) && int.Parse(bidMaster["Status"].ToString()) == 50)
            {
                address += bidMaster["Email"];
                #region 替换模板参数

                #region 替换模板参数
                subject = processMessageContent(mailtitle, this.session, bidMaster, bidOpt);
                subject = processMessageContent(subject, this.session, bidMaster, bidOpt);
                content = processMessageContent(mailcontent, this.session, bidMaster, bidOpt);
                content = processMessageContent(content, this.session, bidMaster, bidOpt);
                #endregion
                #endregion
                //收件人地址
                model.msg_to = address;
                //邮件标题
                model.msg_subject = subject;
                //邮件内容
                model.msg_content = content;

                String errorinfo = "";
                if (Power.Service.MailService.MailBLL.SendMail(model, out errorinfo))
                {
                    address = "";
                    string dateTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    busin.SetItem("Id", Guid.NewGuid());
                    busin.SetItem("Title", subject);
                    //busin.SetItem("FromDate", dateTime);
                    //busin.SetItem("ToHumanId", bidMaster["WinBidSupplierId"]);
                    //busin.SetItem("ToHumanName", bidMaster["WinBidSupplier"]);
                    //busin.SetItem("FromHumanId", bidMaster["RegHumId"]);
                    //busin.SetItem("FromHumanName", bidMaster["RegHumName"]);
                    //busin.SetItem("MessageType", "notify");
                    //busin.SetItem("KeyValue", bidMaster["Id"]);
                    //busin.SetItem("IsMail", 0);
                    //busin.SetItem("IsTextMessage", 0);
                    //busin.SetItem("IsPowerMessage", 1);
                    //busin.SetItem("IsDeviceMessage", 0);
                    //busin.SetItem("ContentText", content);
                    busin.Save(System.ComponentModel.DataObjectMethodType.Insert);
                    result.success = true;
                    result.message = "邮件和系统消息都已发送成功!";
                }
                else
                {
                    //发送邮件失败，错误原因在 errorinfo 里面
                    result.success = false;
                    result.message = "邮件发送失败!";
                    result.message = errorinfo;
                }
            }
            return result.ToJson();
        }

        #endregion
        #region 公用发送邮件
        [ActionAttribute]
        public string PS_Email(string Id, string Keyword, string Email)
        {
            Power.Global.ViewResultModel result = Power.Global.ViewResultModel.Create(true, "");
            Power.Business.IBusinessOperate bidOpt = Power.Business.BusinessFactory.CreateBusinessOperate(Keyword);
            Power.Business.IBaseBusiness bidMaster = bidOpt.FindByKey(Id);
            Power.Business.IBaseBusiness busin = Power.Business.BusinessFactory.CreateBusiness("PowerMessage");
            Power.Global.MailDataPack model = new Power.Global.MailDataPack();
            model.useparam = Power.Global.EUseMailDefaultParma.UseSmtpUser; //使用Power.ConfigEdit.exe 中配置的发送邮件参数

            #region 获取邮件模板
            String swhere = "BaseDataId in (select x1.Id from PB_BaseData x1 where x1.DataType= 'EmailBH')";
            Power.Business.IBusinessList basedataList = Power.Business.BusinessFactory.CreateBusinessOperate("BaseDataList").FindAll(swhere, "", "", 0, 0, SearchFlag.IgnoreRight);
            String mailtitle = "";
            String mailcontent = "";
            foreach (Power.Business.IBaseBusiness item in basedataList)
            {
                if (item["Code"] != null && item["Code"].ToString() == "Title")
                {
                    mailtitle = item["Name"].ToString();
                }
                if (item["Code"] != null && item["Code"].ToString() == "Content")
                {
                    mailcontent = item["Name"].ToString();
                }
            }
            #endregion

            //设置邮件的收件人
            string address = "";
            string subject = "";
            string content = "";
            if (!(XCode.Common.Helper.IsNullKey(Email)))
            {
                address += Email;
                #region 替换模板参数

                #region 替换模板参数

                subject = processMessageContent(mailtitle, this.session, bidMaster, bidOpt);
                subject = processMessageContent(subject, this.session, bidMaster, bidOpt);
                content = processMessageContent(mailcontent, this.session, bidMaster, bidOpt);
                content = processMessageContent(content, this.session, bidMaster, bidOpt);
                #endregion
                #endregion
                //收件人地址
                model.msg_to = address;
                //邮件标题
                model.msg_subject = subject;
                //邮件内容
                model.msg_content = content;

                String errorinfo = "";
                if (Power.Service.MailService.MailBLL.SendMail(model, out errorinfo))
                {
                    address = "";
                    string dateTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    busin.SetItem("Id", Guid.NewGuid());
                    busin.SetItem("Title", subject);
                    //busin.SetItem("FromDate", dateTime);
                    //busin.SetItem("ToHumanId", bidMaster["WinBidSupplierId"]);
                    //busin.SetItem("ToHumanName", bidMaster["WinBidSupplier"]);
                    //busin.SetItem("FromHumanId", bidMaster["RegHumId"]);
                    //busin.SetItem("FromHumanName", bidMaster["RegHumName"]);
                    //busin.SetItem("MessageType", "notify");
                    //busin.SetItem("KeyValue", bidMaster["Id"]);
                    //busin.SetItem("IsMail", 0);
                    //busin.SetItem("IsTextMessage", 0);
                    //busin.SetItem("IsPowerMessage", 1);
                    //busin.SetItem("IsDeviceMessage", 0);
                    //busin.SetItem("ContentText", content);
                    busin.Save(System.ComponentModel.DataObjectMethodType.Insert);
                    result.success = true;
                    result.message = "邮件和系统消息都已发送成功!";
                }
                else
                {
                    //发送邮件失败，错误原因在 errorinfo 里面
                    result.success = false;
                    result.message = "邮件发送失败!";
                    result.message = errorinfo;
                }
            }
            return result.ToJson();
        }
        #endregion
        public DataTable ConvertExcelFileToDataTable(string serverUrl)
        {
            DataTable dt = null;
            Power.Service.FileService.FtpUploadFile powerService = new Power.Service.FileService.FtpUploadFile();
            powerService.Ip = Power.Global.PowerGlobal.FTPIp;
            powerService.Port = int.Parse(Power.Global.PowerGlobal.FTPPort);
            powerService.UserId = Power.Global.PowerGlobal.FTPUserId;
            powerService.UserPwd = Power.Global.PowerGlobal.FTPUserPwd;
            powerService.UsePassive = Power.Global.PowerGlobal.FTPUsePassive;
            try
            {
                int bufferSize = 2048;
                byte[] buffer = new byte[bufferSize];
                using (MemoryStream tmpFileStream = (MemoryStream)powerService.DownLoadFile(serverUrl))
                {
                    Workbook excelDesginer = new Aspose.Cells.Workbook(tmpFileStream);
                    dt = excelDesginer.Worksheets[0].Cells.ExportDataTable(0, 0,
                                                       excelDesginer.Worksheets[0].Cells.MaxDataRow + 1,
                                                      excelDesginer.Worksheets[0].Cells.MaxColumn + 1);
                    foreach (DataColumn item in dt.Columns)
                    {
                        item.ColumnName = dt.Rows[0][dt.Columns.IndexOf(item)].ToString();
                    }
                    if (dt.Rows.Count > 0)
                        dt.Rows[0].Delete();
                }
            }
            catch (Exception e)
            {
                NewLife.Log.XTrace.WritePMSLog(string.Format("文件导入-错误信息:{0}", e.Message), NewLife.Log.LogOperator.Export);
            }
            finally
            {
                powerService.DeleteFile(serverUrl);
            }
            return dt;
        }

        #region 数据导入
        [ActionAttribute]
        public string GoodsImport(Guid fileid)
        {

            Power.Global.ViewResultModel result = Power.Global.ViewResultModel.Create(true, "");
            String filename = getLocalFileName(fileid);
            if (String.IsNullOrEmpty(filename))
            {
                result.success = false;
                result.message = "获取上传的导入文件失败";
                return result.ToJson();
            }
            //string xlsfilename = @"D:\软件\qq接收文件\概算工程量清单EXECL导入模板(B1190单元).xlsx";
            if (!System.IO.File.Exists(filename))
                throw new Exception("Excel文件不存在.");
            DataSet data = new DataSet();
            Workbook workbook = new Workbook(filename);
            foreach (Worksheet ws in workbook.Worksheets)
            {
                Cells ce = ws.Cells;
                if (ce.MaxDataRow == -1 || ce.MaxDataColumn == -1)
                    continue;
                try
                {
                    DataTable dtTemp = ce.ExportDataTable(0, 0, ce.MaxDataRow + 1, ce.MaxDataColumn + 1, true);
                    dtTemp.TableName = ws.Name;
                    data.Tables.Add(dtTemp);
                    if (dtTemp.TableName == "反馈商务信息")
                    {
                        foreach (DataRow row in dtTemp.Rows)
                        {
                            Power.Business.IBaseBusiness budget = Power.Business.BusinessFactory.CreateBusinessOperate("Sus_Pur_RegisterList_C").FindByKey(row["Id"]);
                            if (budget != null)
                            {

                                if (!(XCode.Common.Helper.IsNullKey(row[6])))
                                {
                                    budget.SetItem("Specification", row[6]);
                                }
                                if (!(XCode.Common.Helper.IsNullKey(row[7])))
                                {
                                    budget.SetItem("Price", row[7]);
                                }

                                if (!(XCode.Common.Helper.IsNullKey(row[9])))
                                {
                                    budget.SetItem("Offer", row[9]);
                                }
                                if (!(XCode.Common.Helper.IsNullKey(row[10])))
                                {
                                    budget.SetItem("Amount", row[10]);
                                }

                                if (!(XCode.Common.Helper.IsNullKey(row[11])))
                                {
                                    budget.SetItem("Memo", row[11]);
                                }
                                budget.Save(System.ComponentModel.DataObjectMethodType.Update);
                            }
                            else
                            {
                                result.success = false;
                                result.message = "请勿修改Excel表格Id列的值,否则导入的时候无法进行数据的匹配！";
                                return result.ToJson();
                            }
                        }
                    }
                    else if (dtTemp.TableName == "反馈技术信息")
                    {

                        foreach (DataRow row in dtTemp.Rows)
                        {
                            Power.Business.IBaseBusiness budget = Power.Business.BusinessFactory.CreateBusinessOperate("Sus_Pur_RegisterList_T").FindByKey(row["Id"]);
                            if (budget != null)
                            {
                                if (!(XCode.Common.Helper.IsNullKey(row[6])))
                                {
                                    budget.SetItem("SupplierModel", row[6]);
                                }
                                if (!(XCode.Common.Helper.IsNullKey(row[7])))
                                {
                                    budget.SetItem("SupplierNum", row[7]);
                                }
                                if (!(XCode.Common.Helper.IsNullKey(row[8])))
                                {
                                    budget.SetItem("Place", row[8]);
                                }
                                if (!(XCode.Common.Helper.IsNullKey(row[9])))
                                {
                                    budget.SetItem("Memo", row[9]);
                                }
                                budget.Save(System.ComponentModel.DataObjectMethodType.Update);
                            }
                            else
                            {
                                result.success = false;
                                result.message = "请勿修改Excel表格Id列的值,否则导入的时候无法进行数据的匹配！";
                                return result.ToJson();
                            }
                        }
                    }
                    else if (dtTemp.TableName == "招标商务信息")
                    {
                        foreach (DataRow row in dtTemp.Rows)
                        {
                            Power.Business.IBaseBusiness budget = Power.Business.BusinessFactory.CreateBusinessOperate("Sus_Bid_InquiryList_C").FindByKey(row["Id"]);
                            if (budget != null)
                            {
                                if (!(XCode.Common.Helper.IsNullKey(row[1])))
                                {
                                    budget.SetItem("Code", row[1]);
                                }
                                if (!(XCode.Common.Helper.IsNullKey(row[2])))
                                {
                                    budget.SetItem("Name", row[2]);
                                }
                                if (!(XCode.Common.Helper.IsNullKey(row[3])))
                                {
                                    budget.SetItem("Dept", row[3]);
                                }
                                if (!(XCode.Common.Helper.IsNullKey(row[4])))
                                {
                                    budget.SetItem("Parameter", row[4]);
                                }
                                if (!(XCode.Common.Helper.IsNullKey(row[5])))
                                {
                                    budget.SetItem("Memo", row[5]);
                                }
                                budget.Save(System.ComponentModel.DataObjectMethodType.Update);
                            }


                            else
                            {
                                result.success = false;
                                result.message = "请勿修改Excel表格Id列的值,否则导入的时候无法进行数据的匹配！";
                                return result.ToJson();
                            }
                        }
                    }
                    else if (dtTemp.TableName == "招标技术信息")
                    {
                        foreach (DataRow row in dtTemp.Rows)
                        {
                            Power.Business.IBaseBusiness budget = Power.Business.BusinessFactory.CreateBusinessOperate("Sus_Bid_InquiryList_T").FindByKey(row["Id"]);
                            if (budget != null)
                            {
                                if (!(XCode.Common.Helper.IsNullKey(row[1])))
                                {
                                    budget.SetItem("Code", row[1]);
                                }
                                if (!(XCode.Common.Helper.IsNullKey(row[2])))
                                {
                                    budget.SetItem("Name", row[2]);
                                }
                                if (!(XCode.Common.Helper.IsNullKey(row[3])))
                                {
                                    budget.SetItem("Specification", row[3]);
                                }
                                if (!(XCode.Common.Helper.IsNullKey(row[4])))
                                {
                                    budget.SetItem("Numbers", row[4]);
                                }
                                if (!(XCode.Common.Helper.IsNullKey(row[5])))
                                {
                                    budget.SetItem("Dept", row[5]);
                                }
                                if (!(XCode.Common.Helper.IsNullKey(row[6])))
                                {
                                    budget.SetItem("Memo", row[6]);
                                }
                                budget.Save(System.ComponentModel.DataObjectMethodType.Update);
                            }
                            else
                            {
                                result.success = false;
                                result.message = "请勿修改Excel表格Id列的值,否则导入的时候无法进行数据的匹配！";
                                return result.ToJson();
                            }
                        }
                    }
                    else if (dtTemp.TableName == "技术规格信息")
                    {
                        foreach (DataRow row in dtTemp.Rows)
                        {
                            Power.Business.IBaseBusiness budget = Power.Business.BusinessFactory.CreateBusinessOperate("Sus_TechnicalBook_T").FindByKey(row["Id"]);
                            if (budget != null)
                            {
                                if (!(XCode.Common.Helper.IsNullKey(row[1])))
                                {
                                    budget.SetItem("Code", row[1]);
                                }
                                if (!(XCode.Common.Helper.IsNullKey(row[2])))
                                {
                                    budget.SetItem("Name", row[2]);
                                }
                                if (!(XCode.Common.Helper.IsNullKey(row[3])))
                                {
                                    budget.SetItem("Specification", row[3]);
                                }
                                if (!(XCode.Common.Helper.IsNullKey(row[4])))
                                {
                                    budget.SetItem("Numbers", row[4]);
                                }
                                if (!(XCode.Common.Helper.IsNullKey(row[5])))
                                {
                                    budget.SetItem("Dept", row[5]);
                                }
                                if (!(XCode.Common.Helper.IsNullKey(row[6])))
                                {
                                    budget.SetItem("Memo", row[6]);
                                }
                                budget.Save(System.ComponentModel.DataObjectMethodType.Update);
                            }
                            else
                            {
                                result.success = false;
                                result.message = "请确保你的导入页签名字是：技术规格信息！";
                                return result.ToJson();
                            }
                        }
                    }
                    else
                    {
                        result.success = false;
                        result.message = "导入Excel文件页签名称不匹配！";
                        return result.ToJson();
                    }
                }

                catch (CellsException cex)
                {
                    string message = cex.Message;
                    string[] ms = message.Split(' ');
                    string errorcells = "";
                    for (int i = 0; i < ms.Length; i++)
                    {
                        if (ms[i].ToLower() == "cell" && i + 2 < ms.Length && ms[i + 2] == "should".ToLower())
                            errorcells = ms[i + 1].ToUpper();
                    }
                    if (errorcells != "")
                        throw new Exception("单元格【" + errorcells + "】格式不正确");
                    else
                        throw new Exception("单元格格式不正确：" + message);
                }
                finally
                {

                    if (File.Exists(filename))
                    {
                        File.Delete(filename);
                    }

                }
            }

            return result.ToJson();
        }

        #endregion

        #region 导入方法
        private String getLocalFileName(Guid Id)
        {
            ViewResultModel result = ViewResultModel.Create(false, "");

            #region  获取文件路径和大小
            IBaseBusiness docfile = BusinessFactory.CreateBusinessOperate("DocFile").FindByKey(Id);
            if (docfile == null)
            {
                result.message = "对应的文档不存在。";
                return result.ToJson();
            }
            string strmd5 = docfile["ServerUrl"].ToString() + docfile["FileSize"].ToString();
            string suffix = docfile["FileExt"].ToString();
            string fileName = docfile["Name"].ToString();
            string md5 = NewLife.Security.DataHelper.MD5String(strmd5);
            String sTemplatePath = AppDomain.CurrentDomain.BaseDirectory + "/App_Data/manualCache/";
            string md5filename = sTemplatePath + md5 + suffix;
            #endregion

            #region  查询文件夹中是否存在,不存在下载文件
            //string sTemplatePath = AppDomain.CurrentDomain.BaseDirectory + "/App_Data/";
            if (!File.Exists(md5filename))
            {
                //获取全局参数配置中的文件上传方式
                if (Directory.Exists(sTemplatePath) == false)
                    Directory.CreateDirectory(sTemplatePath);
                //string DocUploadType = "";
                //外部站点访问时，可以直接传递session，否则session为空，这2个取值有问题
                //DocUploadType = Power.Global.PowerGlobal.GetConfigRunTimeValue("FtpConfig", "DocUploadType", session).ToString();
                //找到eps根节点下的 ftp 配置参数
                //StringBuilder sb = new StringBuilder();
                //sb.Append("select code,Value from PB_ConfigRunTime x1 ")
                //  .Append(" join(select * from pln_project where parent_guid not in (select project_guid from pln_project)) x2 on x1.EpsProjId = x2.project_guid ")
                //  .Append("where x1.ConfigTypeCode = 'FtpConfig'");
                //DataTable dt = XCode.DataAccessLayer.DAL.QuerySQL(sb.ToString());

                //DataRow row = dt.Select("code='Ip'").FirstOrDefault();
                //if (row == null || XCode.Common.Helper.IsNullKey(row["Value"]))
                //{
                //    result.message = "eps根节点ftp地址未配置";
                //    return result.ToJson();
                //}
                //string Ip = row["Value"].ToString();

                //row = dt.Select("code='Port'").FirstOrDefault();
                //if (row == null)
                //{
                //    result.message = "eps根节点ftp端口未配置";
                //    return result.ToJson();
                //}
                //string Port = XCode.Common.Helper.IsNullKey(row["Value"]) ? "21" : row["Value"].ToString();

                //row = dt.Select("code='UserId'").FirstOrDefault();
                //if (row == null || XCode.Common.Helper.IsNullKey(row["Value"]))
                //{
                //    result.message = "eps根节点ftp用户未配置";
                //    return result.ToJson();
                //}
                //string UserId = row["Value"].ToString();

                //row = dt.Select("code='UserPwd'").FirstOrDefault();
                //if (row == null)
                //{
                //    result.message = "eps根节点ftp地址未配置";
                //    return result.ToJson();
                //}
                //string UserPwd = XCode.Common.Helper.IsNullKey(row["Value"]) ? "" : NewLife.Security.DataHelper.DESDecrypt(row["Value"].ToString()); ;

                try
                {
                    //switch (DocUploadType.ToLower().Trim())
                    //{
                    ////    case "ftp":
                    string Ip = Power.Global.PowerGlobal.FTPIp;
                    string Port = Power.Global.PowerGlobal.FTPPort;
                    string UserId = Power.Global.PowerGlobal.FTPUserId;
                    string UserPwd = Power.Global.PowerGlobal.FTPUserPwd;
                    Power.Global.FtpHelper.FtpfileDownLoad(Ip + ":" + Port, docfile["ServerUrl"].ToString(), md5filename, UserId, UserPwd);
                    //break;
                    //    default:
                    //        result.message = "仅支持ftp方式存储文件";
                    //        return result.ToJson();
                    //}
                }
                catch (Exception ex)
                {
                    result.message = "文件下载失败：" + ex.Message;
                    return result.ToJson();
                }
            }
            #endregion

            //this.Context.Response.AddHeader("Content-Disposition", "attachment;filename=" + fileName + suffix);
            //this.Context.Response.AddHeader("title", fileName);
            //this.Context.Response.AddHeader("Content-Transfer-Encoding", "binary");
            ////this.Context.Response.ContentType = contentType;
            //this.Context.Response.WriteFile(md5filename);
            //this.Context.ApplicationInstance.CompleteRequest();
            return md5filename;

        }
        #endregion

        #region 导出反馈技术信息
        [ActionAttribute]
        public string DerivedFeedbackTechnique(string Id)
        {
            ViewResultModel result = ViewResultModel.Create(true, "");
            XCode.DataAccessLayer.DAL dal = XCode.DataAccessLayer.DAL.Create();
            string sql = string.Format("select Id,Code,Name,NormModel,NormNum,Dept,SupplierModel,SupplierNum,Place,Memo from Sus_Pur_RegisterList_T x1 where x1.MasterId "
                                      + " in(select Id from Sus_Pur_FeedbackRegister where Id='{0}')order by Sequ", Id);
            DataTable dt;
            dt = dal.Session.Query(sql).Tables[0];
            ExcelExportMessageServiceModel messageInputDto = new ExcelExportMessageServiceModel();
            //dt.Columns.Remove("Id");

            dt.Columns["Id"].ColumnName = "Id";
            dt.Columns["Code"].ColumnName = "编码";
            dt.Columns["Name"].ColumnName = "名称";
            dt.Columns["NormModel"].ColumnName = "技术参数";
            dt.Columns["NormNum"].ColumnName = "标准数量";
            dt.Columns["Dept"].ColumnName = "标准单位";
            dt.Columns["SupplierModel"].ColumnName = "规格型号";
            dt.Columns["SupplierNum"].ColumnName = "数量";

            dt.Columns["Place"].ColumnName = "产地/品牌";
            dt.Columns["Memo"].ColumnName = "备注";

            messageInputDto.datatable = dt;
            messageInputDto.fileName = "反馈技术信息导出";
            messageInputDto.menuid = "";
            messageInputDto.headcolor = "";
            messageInputDto.headfontcolor = "";
            messageInputDto.ispdf = false;

            var message = new Power.Message.MessageArg<ExcelExportMessageServiceModel, string>(null, Power.Message.MessageTypes.Other, "Power.Control.StdExcel.DataTableToExcel", messageInputDto);

            Power.Control.StdExcel.StdExcel.DataTableToExcel(message);

            return result.ToJson();
        }
        #endregion

        #region 导出反馈商务信息
        [ActionAttribute]
        public string DerivedFeedbackRegisterList_C(string Id)
        {
            ViewResultModel result = ViewResultModel.Create(true, "");
            XCode.DataAccessLayer.DAL dal = XCode.DataAccessLayer.DAL.Create();
            string sql = string.Format("select Id,Code,Name,NormModel,NormNum,Dept,Specification,Price,Amount,Offer,Memo from Sus_Pur_RegisterList_C x1 where x1.MasterId "
                                      + " in(select Id from Sus_Pur_FeedbackRegister where Id='{0}')order by Sequ", Id);
            DataTable dt;
            dt = dal.Session.Query(sql).Tables[0];
            ExcelExportMessageServiceModel messageInputDto = new ExcelExportMessageServiceModel();
            //dt.Columns.Remove("Id");

            dt.Columns["Id"].ColumnName = "Id";
            dt.Columns["Code"].ColumnName = "编码";
            dt.Columns["Name"].ColumnName = "名称";
            dt.Columns["NormModel"].ColumnName = "技术参数";
            dt.Columns["NormNum"].ColumnName = "标准数量";
            dt.Columns["Dept"].ColumnName = "标准单位";
            dt.Columns["Specification"].ColumnName = "规格型号";
            dt.Columns["Price"].ColumnName = "单价（元）";

            dt.Columns["Amount"].ColumnName = "数量";
            dt.Columns["Offer"].ColumnName = "报价金额（元）";
            dt.Columns["Memo"].ColumnName = "备注";


            messageInputDto.datatable = dt;
            messageInputDto.fileName = "反馈商务信息导出";
            messageInputDto.menuid = "";
            messageInputDto.headcolor = "";
            messageInputDto.headfontcolor = "";
            messageInputDto.ispdf = false;

            var message = new Power.Message.MessageArg<ExcelExportMessageServiceModel, string>(null, Power.Message.MessageTypes.Other, "Power.Control.StdExcel.DataTableToExcel", messageInputDto);

            Power.Control.StdExcel.StdExcel.DataTableToExcel(message);

            return result.ToJson();
        }
        #endregion

        #region 供应商审核页面模糊查询
        [ActionAttribute]
        public string FuzzySupplyContent(string fields, string Values, string Result)
        {
            Power.Global.ViewResultModel result = Power.Global.ViewResultModel.Create(true, "");
            string strSql = "";
            switch (fields)
            {
                case "Status":
                    if (Values == "0" || Values == "20" || Values == "35"
                        || Values == "40" || Values == "50")
                    {
                        if (!(XCode.Common.Helper.IsNullKey(Result)))
                        {
                            strSql = "select * from SB_SupplierRegistration where Status='" + Values + "' and Result='" + Result + "' and Status=50";
                        }
                        else
                        {
                            strSql = "select * from SB_SupplierRegistration where Status='" + Values + "'";
                        }

                        DataReturned(result, strSql);
                    }
                    else
                    {
                        strSql = "select * from SB_SupplierRegistration where Status='66'";
                        DataReturned(result, strSql);
                    }
                    break;
                case "Code":
                    if (!(XCode.Common.Helper.IsNullKey(Result)))
                    {
                        strSql = "select * from SB_SupplierRegistration where Code like '%" + Values + "%' and Result='" + Result + "' and Status=50";
                    }
                    else
                    {
                        strSql = "select * from SB_SupplierRegistration where Code like '%" + Values + "%'";
                    }

                    DataReturned(result, strSql);
                    break;
                case "Title":
                    if (!(XCode.Common.Helper.IsNullKey(Result)))
                    {
                        strSql = "select * from SB_SupplierRegistration where Title like '%" + Values + "%' and Result='" + Result + "' and Status=50";
                    }
                    else
                    {
                        strSql = "select * from SB_SupplierRegistration where Title like '%" + Values + "%'";
                    }

                    DataReturned(result, strSql);
                    break;
                case "LegalPerson":
                    if (!(XCode.Common.Helper.IsNullKey(Result)))
                    {
                        strSql = "select * from SB_SupplierRegistration where LegalPerson like '%" + Values + "%' and Result='" + Result + "' and Status=50";
                    }
                    else
                    {
                        strSql = "select * from SB_SupplierRegistration where LegalPerson like '%" + Values + "%'";
                    }

                    DataReturned(result, strSql);
                    break;
                case "Owned":
                    if (!(XCode.Common.Helper.IsNullKey(Result)))
                    {
                        strSql = "select * from SB_SupplierRegistration where Owned like '%" + Values + "%' and Result='" + Result + "' and Status=50";
                    }
                    else
                    {
                        strSql = "select * from SB_SupplierRegistration where Owned like '%" + Values + "%'";
                    }

                    DataReturned(result, strSql);
                    break;
                case "Manufacturer":
                    if (!(XCode.Common.Helper.IsNullKey(Result)))
                    {
                        strSql = "select * from SB_SupplierRegistration where Manufacturer like '%" + Values + "%' and Result='" + Result + "' and Status=50";
                    }
                    else
                    {
                        strSql = "select * from SB_SupplierRegistration where Manufacturer like '%" + Values + "%'";
                    }

                    DataReturned(result, strSql);
                    break;
                case "MainSupplies":
                    if (!(XCode.Common.Helper.IsNullKey(Result)))
                    {
                        strSql = "select * from SB_SupplierRegistration where (MainSupplies like '%" + Values + "%' or SubSupplies like'%" + Values + "%') and Result='" + Result + "' and Status=50";
                    }
                    else
                    {
                        strSql = "select * from SB_SupplierRegistration where MainSupplies like '%" + Values + "%' or SubSupplies like'%" + Values + "%'";
                    }

                    DataReturned(result, strSql);
                    break;
                case "SubSupplies":
                    if (!(XCode.Common.Helper.IsNullKey(Result)))
                    {
                        strSql = "select * from SB_SupplierRegistration where (SubSupplies like '%" + Values + "%' or MainSupplies like '%" + Values + "%') and Result='" + Result + "' and Status=50";
                    }
                    else
                    {
                        strSql = "select * from SB_SupplierRegistration where SubSupplies like '%" + Values + "%' or MainSupplies like '%" + Values + "%'";
                    }

                    DataReturned(result, strSql);
                    break;
                case "OwnPerson":
                    if (!(XCode.Common.Helper.IsNullKey(Result)))
                    {
                        strSql = "select * from SB_SupplierRegistration where OwnPerson like '%" + Values + "%' and Result='" + Result + "' and Status=50";
                    }
                    else
                    {
                        strSql = "select * from SB_SupplierRegistration where OwnPerson like '%" + Values + "%'";
                    }

                    DataReturned(result, strSql);
                    break;
                case "PoserType":
                    if (!(XCode.Common.Helper.IsNullKey(Result)))
                    {
                        strSql = "select * from SB_SupplierRegistration where PoserType like '%" + Values + "%' and Result='" + Result + "' and Status=50";
                    }
                    else
                    {
                        strSql = "select * from SB_SupplierRegistration where PoserType like '%" + Values + "%'";
                    }

                    DataReturned(result, strSql);
                    break;
                case "AdminEmail":
                    if (!(XCode.Common.Helper.IsNullKey(Result)))
                    {
                        strSql = "select * from SB_SupplierRegistration where AdminEmail like '%" + Values + "%' and Result='" + Result + "' and Status=50";
                    }
                    else
                    {
                        strSql = "select * from SB_SupplierRegistration where AdminEmail like '%" + Values + "%'";
                    }

                    DataReturned(result, strSql);
                    break;
                case "CompAddress":
                    if (!(XCode.Common.Helper.IsNullKey(Result)))
                    {
                        strSql = "select * from SB_SupplierRegistration where CompAddress like '%" + Values + "%' and Result='" + Result + "' and Status=50";
                    }
                    else
                    {
                        strSql = "select * from SB_SupplierRegistration where CompAddress like '%" + Values + "%'";
                    }

                    DataReturned(result, strSql);
                    break;
                case "Types":
                    if (!(XCode.Common.Helper.IsNullKey(Result)))
                    {
                        strSql = "select * from SB_SupplierRegistration where Types like '%" + Values + "%' and Result='" + Result + "' and Status=50";
                    }
                    else
                    {
                        strSql = "select * from SB_SupplierRegistration where Types like '%" + Values + "%'";
                    }

                    DataReturned(result, strSql);
                    break;
                default:
                    if (!(XCode.Common.Helper.IsNullKey(Result)))
                    {
                        strSql = "select * from SB_SupplierRegistration where Result='" + Result + "' and Status=50";
                    }
                    else
                    {
                        strSql = "select * from SB_SupplierRegistration";
                    }
                    DataReturned(result, strSql);
                    break;
            }
            return result.ToJson();
        }
        #endregion

        #region 查询结果返回前端
        private void DataReturned(Power.Global.ViewResultModel result, string strSql)
        {
            DataTable tb = XCode.DataAccessLayer.DAL.QuerySQL(string.Format(strSql));
            result.data.Add("values", Power.Global.BusiHelper.DataTableToHashtable(tb));
        }
        #endregion
        #region 设备模板导入
        [ActionAttribute]
        public string DevicePropertyImport(Guid fileid, string Id)
        {
            Power.Global.ViewResultModel result = Power.Global.ViewResultModel.Create(true, "");
            String filename = getLocalFileName(fileid);
            if (String.IsNullOrEmpty(filename))
            {
                result.success = false;
                result.message = "获取上传的导入文件失败";
                return result.ToJson();
            }
            if (!System.IO.File.Exists(filename))
                throw new Exception("Excel文件不存在.");
            DataSet data = new DataSet();
            Workbook workbook = new Workbook(filename);
            foreach (Worksheet ws in workbook.Worksheets)
            {
                Cells ce = ws.Cells;
                if (ce.MaxDataRow == -1 || ce.MaxDataColumn == -1)
                    continue;
                try
                {
                    Dictionary<string, Guid> tempids = new Dictionary<string, Guid>();//存储键值对Code为键Id为值。
                    DataTable dtTemp = ce.ExportDataTable(0, 0, ce.MaxDataRow + 1, ce.MaxDataColumn + 1, true);
                    dtTemp.TableName = ws.Name;
                    data.Tables.Add(dtTemp);
                    if (dtTemp.TableName == "设备属性导入")
                    {
                        foreach (DataRow row in dtTemp.Rows)
                        {
                            Power.Business.IBaseBusiness busin = Power.Business.BusinessFactory.CreateBusiness("Sus_parameterTemplateList");
                            string leng = row[5].ToString();
                            string PreviousValue = leng.LastIndexOf(".").ToString();

                            if (PreviousValue.Equals("-1"))
                            {
                                Guid NewId = Guid.NewGuid();
                                busin.SetItem("Id", NewId);
                                busin.SetItem("ParentId", "00000000-0000-0000-0000-000000000000");
                                busin.SetItem("MasterId", Id);
                                tempids.Add(row[5].ToString(), NewId);
                                if (!(XCode.Common.Helper.IsNullKey(row[0])))
                                {
                                    busin.SetItem("property", row[0]);
                                }
                                if (!(XCode.Common.Helper.IsNullKey(row[1])))
                                {
                                    busin.SetItem("SimpleValue", row[1]);
                                }
                                if (!(XCode.Common.Helper.IsNullKey(row[2])))
                                {
                                    busin.SetItem("Amount", row[2]);
                                }
                                if (!(XCode.Common.Helper.IsNullKey(row[3])))
                                {
                                    busin.SetItem("Unit", row[3]);
                                }
                                if (!(XCode.Common.Helper.IsNullKey(row[4])))
                                {
                                    busin.SetItem("Memo", row[4]);
                                }
                                busin.Save(System.ComponentModel.DataObjectMethodType.Insert);
                            }
                            else if (int.Parse(PreviousValue) > 0)
                            {
                                Guid NewId = Guid.NewGuid();
                                string a = leng.Substring(0, int.Parse(PreviousValue));
                                tempids.Add(row[5].ToString(), NewId);
                                busin.SetItem("Id", NewId);
                                busin.SetItem("ParentId", tempids[a]);//用Code找到Id赋值给ParentId
                                busin.SetItem("MasterId", Id);

                                if (!(XCode.Common.Helper.IsNullKey(row[0])))
                                {
                                    busin.SetItem("property", row[0]);
                                }
                                if (!(XCode.Common.Helper.IsNullKey(row[1])))
                                {
                                    busin.SetItem("SimpleValue", row[1]);
                                }
                                if (!(XCode.Common.Helper.IsNullKey(row[2])))
                                {
                                    busin.SetItem("Amount", row[2]);
                                }
                                if (!(XCode.Common.Helper.IsNullKey(row[3])))
                                {
                                    busin.SetItem("Unit", row[3]);
                                }
                                if (!(XCode.Common.Helper.IsNullKey(row[4])))
                                {
                                    busin.SetItem("Memo", row[4]);
                                }
                                busin.Save(System.ComponentModel.DataObjectMethodType.Insert);
                            }
                            else
                            {
                                result.success = false;
                                result.message = "编号列必须输入按规格输入,否则无法生成树结构显示！";
                                result.ToJson();
                            }
                        }
                    }
                    else
                    {
                        result.success = false;
                        result.message = "导入Excel文件页签名称不匹配！";
                        return result.ToJson();
                    }
                }
                catch (CellsException cex)
                {
                    string message = cex.Message;
                    string[] ms = message.Split(' ');
                    string errorcells = "";
                    for (int i = 0; i < ms.Length; i++)
                    {
                        if (ms[i].ToLower() == "cell" && i + 2 < ms.Length && ms[i + 2] == "should".ToLower())
                            errorcells = ms[i + 1].ToUpper();
                    }
                    if (errorcells != "")
                        throw new Exception("单元格【" + errorcells + "】格式不正确");
                    else
                        throw new Exception("单元格格式不正确：" + message);
                }
                finally
                {

                    if (File.Exists(filename))
                    {
                        File.Delete(filename);
                    }

                }
            }

            return result.ToJson();
        }
        #endregion
        [ActionAttribute]
        public string ImportClassTmpDatas(string jsonData)
        {
            string strLog = "{'result':'true','msg':'模板数据导入成功'}";
            try
            {
                Hashtable hts = (Hashtable)Newtonsoft.Json.JsonConvert.DeserializeObject<Hashtable>(jsonData);
                string serverUrl = string.Format("/{0}/{1}/{2}", hts["keyword"], hts["Id"], hts["tmpFileName"]);
                DataTable dt = ConvertExcelFileToDataTable(serverUrl);
                if (dt != null && dt.Rows.Count <= 0)
                    return "{'result':'false','msg':'Excel模板数据为空'}";
                foreach (DataColumn item in dt.Columns)
                {
                    switch (item.ColumnName)
                    {

                        case "编码":
                            item.ColumnName = "Code";
                            break;
                        case "名称":
                            item.ColumnName = "Name";
                            break;
                        case "单位":
                            item.ColumnName = "Dept";
                            break;

                        case "数量":
                            item.ColumnName = "Numbers";
                            break;
                        case "技术参数":
                            item.ColumnName = "Parameter";
                            break;
                        case "备注":
                            item.ColumnName = "Memo";
                            break;
                        default:
                            break;
                    }
                }
                //导入数据库 
                IBusinessOperate busC = BusinessFactory.CreateBusinessOperate("Sus_Bid_InquiryList_C");
                IBusinessList dtCListOLd = busC.FindAll("MasterId", Convert.ToString(hts["Id"]), SearchFlag.IgnoreRight);
                if (dtCListOLd.Count > 0)
                    dtCListOLd.Delete();  //清除原始分类

                //转换Exlcel模板数据为业务对象集合
                // dt.DefaultView.Sort = "LongCode asc";
                // dt = dt.DefaultView.ToTable();
                IBusinessList dtCList = busC.ConvertDataTableToEntity(dt);
                if (dtCList.Count == 0)
                    return "{\"result\":\"false\",\"msg\":格式错误的物资分类模板\"\"}";
                for (int i = 0; i < dtCList.Count; i++)
                {
                    if (StrTorF(dtCList[i]["Name"]))
                    {
                        dtCList.Remove(dtCList[i]);
                        i--;
                    }
                    else
                    {
                        dtCList[i].SetItem("MasterId", hts["Id"]);

                    }


                }
                dtCList.Save(true);
            }
            catch (Exception e)
            {
                strLog = "{'result':'false','msg':'" + e.Message + "'}";
            }
            return strLog;
        }
        /// <summary>
        /// 招标技术信息导入
        /// </summary>
        /// <param name="jsonData"></param>
        /// <returns></returns>
        [ActionAttribute]
        public string ImportSus_Bid_InquiryList_T(string jsonData)
        {
            string strLog = "{'result':'true','msg':'模板数据导入成功'}";
            try
            {
                Hashtable hts = (Hashtable)Newtonsoft.Json.JsonConvert.DeserializeObject<Hashtable>(jsonData);
                string serverUrl = string.Format("/{0}/{1}/{2}", hts["keyword"], hts["Id"], hts["tmpFileName"]);
                DataTable dt = ConvertExcelFileToDataTable(serverUrl);
                if (dt != null && dt.Rows.Count <= 0)
                    return "{'result':'false','msg':'Excel模板数据为空'}";
                foreach (DataColumn item in dt.Columns)
                {
                    switch (item.ColumnName)
                    {

                        case "编码":
                            item.ColumnName = "Code";
                            break;
                        case "名称":
                            item.ColumnName = "Name";
                            break;
                        case "单位":
                            item.ColumnName = "Dept";
                            break;

                        case "数量":
                            item.ColumnName = "Numbers";
                            break;
                        case "技术参数":
                            item.ColumnName = "Specification";
                            break;
                        case "备注":
                            item.ColumnName = "Memo";
                            break;
                        default:
                            break;
                    }
                }
                //导入数据库 
                IBusinessOperate busC = BusinessFactory.CreateBusinessOperate("Sus_Bid_InquiryList_T");
                IBusinessList dtCListOLd = busC.FindAll("MasterId", Convert.ToString(hts["Id"]), SearchFlag.IgnoreRight);
                if (dtCListOLd.Count > 0)
                    dtCListOLd.Delete();  //清除原始分类

                //转换Exlcel模板数据为业务对象集合
                // dt.DefaultView.Sort = "LongCode asc";
                // dt = dt.DefaultView.ToTable();
                IBusinessList dtCList = busC.ConvertDataTableToEntity(dt);
                if (dtCList.Count == 0)
                    return "{\"result\":\"false\",\"msg\":格式错误的物资分类模板\"\"}";
                for (int i = 0; i < dtCList.Count; i++)
                {
                    if (StrTorF(dtCList[i]["Name"]))
                    {
                        dtCList.Remove(dtCList[i]);
                        i--;
                    }
                    else
                    {
                        dtCList[i].SetItem("MasterId", hts["Id"]);

                    }


                }
                dtCList.Save(true);
            }
            catch (Exception e)
            {
                strLog = "{'result':'false','msg':'" + e.Message + "'}";
            }
            return strLog;
        }
        /// <summary>
        /// 判断字符串是否是GUID
        /// </summary>
        /// <param name="strSrc"></param>
        /// <returns></returns>
        private bool StrTorF(object obj)
        {
            string s = Convert.ToString(obj).Trim();
            bool flg = string.IsNullOrEmpty(s);
            if (!flg)
            {
                if (IsGuidByParse(s))
                {
                    flg = s == Guid.Empty.ToString();
                }
            }
            return flg;
        }
        private bool IsGuidByParse(string strSrc)
        {
            Guid g = Guid.Empty;
            return Guid.TryParse(strSrc, out g);
        }
        /// <summary>
        /// 移除字符串末尾指定字符
        /// </summary>
        /// <param name="str">需要移除的字符串</param>
        /// <param name="value">指定字符</param>
        /// <returns>移除后的字符串</returns>
        private string RemoveLastChar(string str, string value)
        {
            try
            {
                int _finded = str.LastIndexOf(value);
                if (_finded != -1)
                {
                    return str.Substring(0, _finded);
                }
                return str;
            }
            catch
            {
                return str;
            }
        }
        public DataTable ReslutData(string KeyWord, string KeyValue)
        {

            // 上传的数据
            DataTable tempFile = new DataTable();

            //找到目标文件对象
            //System.Web.HttpPostedFile uploadFile = this.Context.Request.Files["PCPath"];

            // 如果有文件, 则读取文件信息
            // if (uploadFile.ContentLength > 0)
            //{
            //System.IO.Stream fileDataStream = uploadFile.InputStream;
            //int fileLength = uploadFile.ContentLength;
            // byte[] fileData = new byte[fileLength];
            //通过KeyWord、Keyword找到PB_DocFiles对应的数据
            string[] keys = { "BOKeyWord", "FolderId" };
            string[] values = { KeyWord, KeyValue };
            Power.Systems.Systems.DocFileBO docfile = Power.Business.BusinessFactory.CreateBusinessOperate("DocFile").FindByKey(keys, values) as Power.Systems.Systems.DocFileBO;

            string Ip = Power.Global.PowerGlobal.GetConfigRunTimeValue("FtpConfig", "Ip").ToString();
            string Port = Power.Global.PowerGlobal.GetConfigRunTimeValue("FtpConfig", "Port").ToString();
            string UserId = Power.Global.PowerGlobal.GetConfigRunTimeValue("FtpConfig", "UserId").ToString();
            string UserPwd = Power.Global.PowerGlobal.GetConfigRunTimeValue("FtpConfig", "UserPwd").ToString();//UserPwd
            string filePath = "ftp://" + Ip + ":" + Port + docfile.ServerUrl;
            System.IO.MemoryStream memory = new System.IO.MemoryStream();
            byte[] bt = Power.Global.FtpHelper.FtpfileDownLoad(filePath, UserId, UserPwd).GetBuffer();
            foreach (byte item in bt)
            {
                memory.WriteByte(item);
            }

            //System.IO.StreamWriter writer = new System.IO.StreamWriter(memory);
            string fileurl = AppDomain.CurrentDomain.BaseDirectory + "\\" + docfile.Name + docfile.FileExt;
            System.IO.FileStream file = new System.IO.FileStream(AppDomain.CurrentDomain.BaseDirectory + "\\" + docfile.Name + docfile.FileExt, System.IO.FileMode.CreateNew);
            memory.WriteTo(file);

            file.Dispose();
            //writer.Dispose();
            memory.Dispose();
            //把文件流填充到数组   
            //fileDataStream.Read(fileData, 0, fileLength);
            //解码二进制数组
            DataSet ds = Power.Global.PowerGlobal.Office.ExcelToDataSet(AppDomain.CurrentDomain.BaseDirectory + "\\" + docfile.Name + docfile.FileExt);
            //tempFile = GetExcelDatatable(AppDomain.CurrentDomain.BaseDirectory + "\\" + docfile.Name + docfile.FileExt, "dt1");
            tempFile = ds.Tables[0];

            if (System.IO.File.Exists(fileurl))
                System.IO.File.Delete(fileurl);
            docfile.Delete();
            // uploadFile.SaveAs(string.Format("{0}{1}{2}", tempFile, "Images/", uploadFile.FileName));
            // }
            //tempFile = System.IO.File.ReadAllText(@"C:\Data.txt");

            //从ftp下载文件到本地服务器

            //读取服务器上的文件
            return tempFile;
        }
        #region 招标文件商务信息导入
        [ActionAttribute]
        public string Sus_Bid_InquiryList_CImport(Guid fileid, string Id)
        {
            //先删除原来的
            Power.Business.BusinessFactory.CreateBusinessOperate("Sus_Bid_InquiryList_C").FindAll("MasterId", Id, Power.Business.SearchFlag.IgnoreRight).Delete();//此处修改关键词、外键
            Power.Global.ViewResultModel result = Power.Global.ViewResultModel.Create(true, "");
            String filename = getLocalFileName(fileid);
            if (String.IsNullOrEmpty(filename))
            {
                result.success = false;
                result.message = "获取上传的导入文件失败";
                return result.ToJson();
            }
            if (!System.IO.File.Exists(filename))
                throw new Exception("Excel文件不存在.");
            DataSet data = new DataSet();
            Workbook workbook = new Workbook(filename);
            foreach (Worksheet ws in workbook.Worksheets)
            {
                Cells ce = ws.Cells;
                if (ce.MaxDataRow == -1 || ce.MaxDataColumn == -1)
                    continue;
                try
                {
                    Dictionary<string, Guid> tempids = new Dictionary<string, Guid>();//存储键值对Code为键Id为值。
                    DataTable dtTemp = ce.ExportDataTable(0, 0, ce.MaxDataRow + 1, ce.MaxDataColumn + 1, true);
                    dtTemp.TableName = ws.Name;
                    data.Tables.Add(dtTemp);
                    if (dtTemp.TableName == "商务信息导入")
                    {
                        foreach (DataRow row in dtTemp.Rows)
                        {
                            Power.Business.IBaseBusiness busin = Power.Business.BusinessFactory.CreateBusiness("Sus_Bid_InquiryList_C");
                            string leng = row[6].ToString();
                            string PreviousValue = leng.LastIndexOf(".").ToString();

                            if (PreviousValue.Equals("-1"))
                            {
                                Guid NewId = Guid.NewGuid();
                                busin.SetItem("Id", NewId);
                                busin.SetItem("ParentId", "00000000-0000-0000-0000-000000000000");
                                busin.SetItem("MasterId", Id);
                                tempids.Add(row[6].ToString(), NewId);
                                if (!(XCode.Common.Helper.IsNullKey(row[0])))
                                {
                                    busin.SetItem("property", row[0]);
                                }
                                if (!(XCode.Common.Helper.IsNullKey(row[0])))
                                {
                                    busin.SetItem("Code", row[0]);
                                }
                                if (!(XCode.Common.Helper.IsNullKey(row[1])))
                                {
                                    busin.SetItem("Name", row[1]);
                                }
                                if (!(XCode.Common.Helper.IsNullKey(row[2])))
                                {
                                    busin.SetItem("Dept", row[2]);
                                }
                                if (!(XCode.Common.Helper.IsNullKey(row[3])))
                                {
                                    busin.SetItem("Numbers", row[3]);
                                }
                                if (!(XCode.Common.Helper.IsNullKey(row[4])))
                                {
                                    busin.SetItem("Parameter", row[4]);
                                }
                                if (!(XCode.Common.Helper.IsNullKey(row[5])))
                                {
                                    busin.SetItem("Memo", row[5]);
                                }
                                busin.Save(System.ComponentModel.DataObjectMethodType.Insert);
                            }
                            else if (int.Parse(PreviousValue) > 0)
                            {
                                Guid NewId = Guid.NewGuid();
                                string a = leng.Substring(0, int.Parse(PreviousValue));
                                tempids.Add(row[6].ToString(), NewId);
                                busin.SetItem("Id", NewId);
                                busin.SetItem("ParentId", tempids[a]);//用Code找到Id赋值给ParentId
                                busin.SetItem("MasterId", Id);

                                if (!(XCode.Common.Helper.IsNullKey(row[0])))
                                {
                                    busin.SetItem("property", row[0]);
                                }
                                if (!(XCode.Common.Helper.IsNullKey(row[0])))
                                {
                                    busin.SetItem("Code", row[0]);
                                }
                                if (!(XCode.Common.Helper.IsNullKey(row[1])))
                                {
                                    busin.SetItem("Name", row[1]);
                                }
                                if (!(XCode.Common.Helper.IsNullKey(row[2])))
                                {
                                    busin.SetItem("Dept", row[2]);
                                }
                                if (!(XCode.Common.Helper.IsNullKey(row[3])))
                                {
                                    busin.SetItem("Numbers", row[3]);
                                }
                                if (!(XCode.Common.Helper.IsNullKey(row[4])))
                                {
                                    busin.SetItem("Parameter", row[4]);
                                }
                                if (!(XCode.Common.Helper.IsNullKey(row[5])))
                                {
                                    busin.SetItem("Memo", row[5]);
                                }
                                busin.Save(System.ComponentModel.DataObjectMethodType.Insert);
                            }
                            else
                            {
                                result.success = false;
                                result.message = "编号列必须输入按规格输入,否则无法生成树结构显示！";
                                result.ToJson();
                            }
                        }
                    }
                    else
                    {
                        result.success = false;
                        result.message = "导入Excel文件页签名称不匹配！";
                        return result.ToJson();
                    }
                }
                catch (CellsException cex)
                {
                    string message = cex.Message;
                    string[] ms = message.Split(' ');
                    string errorcells = "";
                    for (int i = 0; i < ms.Length; i++)
                    {
                        if (ms[i].ToLower() == "cell" && i + 2 < ms.Length && ms[i + 2] == "should".ToLower())
                            errorcells = ms[i + 1].ToUpper();
                    }
                    if (errorcells != "")
                        throw new Exception("单元格【" + errorcells + "】格式不正确");
                    else
                        throw new Exception("单元格格式不正确：" + message);
                }
                finally
                {

                    if (File.Exists(filename))
                    {
                        File.Delete(filename);
                    }

                }
            }

            return result.ToJson();

        }
        #endregion
        //#region  技术规格书评审
        //[ActionAttribute]
        //public string TechnicalSpecificationReview(string Id, string GuideId)
        //{
        //    Power.Global.ViewResultModel result = Power.Global.ViewResultModel.Create(true, "");

        //    Dictionary<String, Guid> tempids = new Dictionary<String, Guid>();
        //    Dictionary<String, Guid> tempidsList = new Dictionary<String, Guid>();

        //    #region JSON格式字符串转换成对象
        //    var srt = JsonConvert.DeserializeObject<string[]>(GuideId);
        //    string sqlSrt = "";
        //    foreach (object itme in srt)
        //    {
        //        sqlSrt += "'" + itme + "'" + ",";
        //    }
        //    sqlSrt = sqlSrt.TrimEnd(',');
        //    #endregion

        //    #region 查询出向导中所有选中主表的数据
        //    string strsql = "select * from V_Sus_parameterTemplate where Id in({0})";
        //    DataTable dtData = XCode.DataAccessLayer.DAL.QuerySQL(string.Format(strsql, sqlSrt));
        //    #endregion

        //    #region 判断是否有数据，有则赋值
        //    if (dtData.Rows.Count == 0)
        //    {
        //        result.success = false;
        //        result.message = "请选择一条数据！";
        //        return result.ToJson();
        //    }
        //    else
        //    {
        //        #region 1、数据删除
        //        Power.Business.IBusinessList businBook_T = Power.Business.BusinessFactory.CreateBusinessOperate("Sus_TechnicalBook_T").FindAll("MasterId", Id, Business.SearchFlag.Default);
        //        if (businBook_T.Count > 0)
        //        {
        //            foreach (Power.Business.IBaseBusiness item in businBook_T)
        //            {
        //                Power.Business.IBusinessList businBook_T_Del = Power.Business.BusinessFactory.CreateBusinessOperate("Sus_TechnicalBook_T_Del").FindAll("MasterId", item["Id"].ToString(), Business.SearchFlag.Default);
        //                if (businBook_T_Del.Count > 0)
        //                {
        //                    businBook_T_Del.Delete();
        //                }
        //            }
        //            businBook_T.Delete();
        //        }
        //        #endregion

        //        #region 2.1、取数
        //        foreach (DataRow row in dtData.Rows)
        //        {
        //            if (tempids.ContainsKey(row["Id"].ToString()) == false)
        //            {
        //                tempids.Add(row["Id"].ToString(), Guid.NewGuid());
        //            }
        //            if (XCode.Common.Helper.IsNullKey(row["ParentId"]) == false && tempids.ContainsKey(row["ParentId"].ToString()) == false)
        //            {
        //                tempids.Add(row["ParentId"].ToString(), Guid.NewGuid());
        //            }
        //            Power.Business.IBaseBusiness TBT = Power.Business.BusinessFactory.CreateBusiness("Sus_TechnicalBook_T");
        //            TBT.SetItem("Id", tempids[row["Id"].ToString()]);
        //            if (XCode.Common.Helper.IsNullKey(row["ParentId"]) == false)
        //                TBT.SetItem("ParentId", tempids[row["ParentId"].ToString()]);
        //            else
        //                TBT.SetItem("ParentId", null);
        //            TBT.SetItem("MasterId", Id);
        //            TBT.SetItem("Code", row["code"]);
        //            TBT.SetItem("Name", row["name"]);
        //            TBT.SetItem("Dept", row["Unit"]);
        //            TBT.Save(System.ComponentModel.DataObjectMethodType.Insert);
        //            #endregion

        //            #region 2.2、取数
        //            Power.Business.IBusinessList TemplateList = Power.Business.BusinessFactory.CreateBusinessOperate("Sus_parameterTemplateList").FindAll("MasterId", row["Id"], Business.SearchFlag.Default);
        //            foreach (Power.Business.IBaseBusiness item in TemplateList)
        //            {
        //                if (tempidsList.ContainsKey(item["Id"].ToString()) == false)
        //                {
        //                    tempidsList.Add(item["Id"].ToString(), Guid.NewGuid());
        //                }
        //                if (XCode.Common.Helper.IsNullKey(item["ParentId"]) == false && tempidsList.ContainsKey(item["ParentId"].ToString()) == false)
        //                {
        //                    tempidsList.Add(item["ParentId"].ToString(), Guid.NewGuid());
        //                }
        //                Power.Business.IBaseBusiness TBTList = Power.Business.BusinessFactory.CreateBusiness("Sus_TechnicalBook_T_Del");
        //                TBTList.SetItem("Id", tempidsList[item["Id"].ToString()]);
        //                if (XCode.Common.Helper.IsNullKey(item["ParentId"]) == false)
        //                    TBTList.SetItem("ParentId", tempidsList[item["ParentId"].ToString()]);
        //                else
        //                    TBTList.SetItem("ParentId", null);
        //                TBTList.SetItem("MasterId", tempids[row["Id"].ToString()]);
        //                TBTList.SetItem("property", item["property"]);
        //                TBTList.SetItem("SimpleValue", item["SimpleValue"]);
        //                TBTList.SetItem("Amount", item["Amount"]);
        //                TBTList.SetItem("Unit", item["Unit"]);
        //                TBTList.SetItem("Memo", item["Memo"]);
        //                TBTList.Save(System.ComponentModel.DataObjectMethodType.Insert);
        //            }
        //        }
        //        #endregion 

        //        return result.ToJson();
        //    }
        //}
        //#endregion
        //#endregion
        #region  技术规格书评审
        [ActionAttribute]
        public string TechnicalSpecificationReview(string Id, string GuideId)
        {
            Power.Global.ViewResultModel result = Power.Global.ViewResultModel.Create(true, "");

            Dictionary<String, Guid> tempids = new Dictionary<String, Guid>();
            Dictionary<String, Guid> tempidsList = new Dictionary<String, Guid>();

            #region JSON格式字符串转换成对象
            var srt = JsonConvert.DeserializeObject<string[]>(GuideId);
            string sqlSrt = "";
            foreach (object itme in srt)
            {
                sqlSrt += "'" + itme + "'" + ",";
            }
            sqlSrt = sqlSrt.TrimEnd(',');
            #endregion

            #region 查询出向导中所有选中主表的数据
            string strsql = "select * from V_Sus_parameterTemplate where Id in({0})";
            DataTable dtData = XCode.DataAccessLayer.DAL.QuerySQL(string.Format(strsql, sqlSrt));
            #endregion
            #region 找到所有根节点
            List<String> listRoot = new List<string>();
            #region 判断是否有数据，有则赋值
            if (dtData.Rows.Count == 0)
            {
                result.success = false;
                result.message = "请选择一条数据！";
                return result.ToJson();
            }
            else
            {
                //#region 1、数据删除
                //Power.Business.IBusinessList businBook_T = Power.Business.BusinessFactory.CreateBusinessOperate("Sus_TechnicalBook_T").FindAll("MasterId", Id, Business.SearchFlag.Default);
                //if (businBook_T.Count > 0)
                //{
                //    foreach (Power.Business.IBaseBusiness item in businBook_T)
                //    {
                //        Power.Business.IBusinessList businBook_T_Del = Power.Business.BusinessFactory.CreateBusinessOperate("Sus_TechnicalBook_T_Del").FindAll("MasterId", item["Id"].ToString(), Business.SearchFlag.Default);
                //        if (businBook_T_Del.Count > 0)
                //        {
                //            businBook_T_Del.Delete();
                //        }
                //    }
                //    businBook_T.Delete();
                //}
                //#endregion

                #region 2.1、取数
                DataTable dtC = Power.Business.BusinessFactory.CreateBusinessOperate("Sus_TechnicalBook_T").FindAllByTable(string.Format("MasterId='{0}'", Id), "", "NodeId", 0, 0, SearchFlag.IgnoreRight);
                foreach (DataRow row in dtData.Rows)
                {
                    if (tempids.ContainsKey(row["Id"].ToString()) == false)
                    {
                        tempids.Add(row["Id"].ToString(), Guid.NewGuid());
                    }
                    if (XCode.Common.Helper.IsNullKey(row["ParentId"]) == false && tempids.ContainsKey(row["ParentId"].ToString()) == false)
                    {
                        tempids.Add(row["ParentId"].ToString(), Guid.NewGuid());
                    }
                    DataRow[] rowsSelect = dtC.Select(String.Format("NodeId='{0}'", row["Id"]));
                    if (rowsSelect != null && rowsSelect.Length != 0)
                    {
                        continue;
                    }
                    Power.Business.IBaseBusiness TBT = Power.Business.BusinessFactory.CreateBusiness("Sus_TechnicalBook_T");

                    TBT.SetItem("Id", tempids[row["Id"].ToString()]);
                    if (XCode.Common.Helper.IsNullKey(row["ParentId"]) == false)
                        TBT.SetItem("ParentId", tempids[row["ParentId"].ToString()]);
                    else
                        TBT.SetItem("ParentId", null);
                    TBT.SetItem("MasterId", Id);

                    TBT.SetItem("Code", row["code"]);
                    TBT.SetItem("Name", row["name"]);
                    TBT.SetItem("Dept", row["Unit"]);

                    TBT.Save(System.ComponentModel.DataObjectMethodType.Insert);
                    #endregion

                    #region 2.2、取数
                    Power.Business.IBusinessList TemplateList = Power.Business.BusinessFactory.CreateBusinessOperate("Sus_parameterTemplateList").FindAll("MasterId", row["Id"], Business.SearchFlag.Default);
                    foreach (Power.Business.IBaseBusiness item in TemplateList)
                    {
                        if (tempidsList.ContainsKey(item["Id"].ToString()) == false)
                        {
                            tempidsList.Add(item["Id"].ToString(), Guid.NewGuid());
                        }
                        if (XCode.Common.Helper.IsNullKey(item["ParentId"]) == false && tempidsList.ContainsKey(item["ParentId"].ToString()) == false)
                        {
                            tempidsList.Add(item["ParentId"].ToString(), Guid.NewGuid());
                        }
                        Power.Business.IBaseBusiness TBTList = Power.Business.BusinessFactory.CreateBusiness("Sus_TechnicalBook_T");
                        TBTList.SetItem("Id", tempidsList[item["Id"].ToString()]);
                        if (XCode.Common.Helper.IsNullKey(item["ParentId"]) == false)
                            TBTList.SetItem("ParentId", tempidsList[item["ParentId"].ToString()]);
                        else
                            TBTList.SetItem("ParentId", tempids[row["Id"].ToString()]);
                        TBTList.SetItem("MasterId", Id);
                        TBTList.SetItem("Name", item["property"]);
                        TBTList.SetItem("Specification", item["SimpleValue"]);
                        TBTList.SetItem("Numbers", item["Amount"]);
                        TBTList.SetItem("Dept", item["Unit"]);
                        TBTList.SetItem("Memo", item["Memo"]);
                        TBTList.SetItem("Sequ", item["Sequ"]);
                        TBTList.Save(System.ComponentModel.DataObjectMethodType.Insert);
                    }
                }
                #endregion 

                return result.ToJson();
            }
        }
        #endregion
        #endregion
        #endregion


        #region  招标询价商务信息
        [ActionAttribute]
        public string Sus_Bid_InquiryList_C(string Id, string GuideId)
        {
            Power.Global.ViewResultModel result = Power.Global.ViewResultModel.Create(true, "");


            Dictionary<String, Guid> tempids = new Dictionary<String, Guid>();
            Dictionary<String, Guid> tempidsList = new Dictionary<String, Guid>();

            #region JSON格式字符串转换成对象
            var srt = JsonConvert.DeserializeObject<string[]>(GuideId);
            string sqlSrt = "";
            foreach (object itme in srt)
            {
                sqlSrt += "'" + itme + "'" + ",";
            }
            sqlSrt = sqlSrt.TrimEnd(',');
            #endregion

            #region 查询出向导中所有选中主表的数据
            string strsql = "select * from V_Sus_parameterTemplate where Id in({0})";
            DataTable dtData = XCode.DataAccessLayer.DAL.QuerySQL(string.Format(strsql, sqlSrt));
            #endregion
            #region 找到所有根节点
            List<String> listRoot = new List<string>();
            #region 判断是否有数据，有则赋值
            if (dtData.Rows.Count == 0)
            {
                result.success = false;
                result.message = "请选择一条数据！";
                return result.ToJson();
            }
            else
            {
                //#region 1、数据删除
                //Power.Business.IBusinessList businBook_T = Power.Business.BusinessFactory.CreateBusinessOperate("Sus_TechnicalBook_T").FindAll("MasterId", Id, Business.SearchFlag.Default);
                //if (businBook_T.Count > 0)
                //{
                //    foreach (Power.Business.IBaseBusiness item in businBook_T)
                //    {
                //        Power.Business.IBusinessList businBook_T_Del = Power.Business.BusinessFactory.CreateBusinessOperate("Sus_TechnicalBook_T_Del").FindAll("MasterId", item["Id"].ToString(), Business.SearchFlag.Default);
                //        if (businBook_T_Del.Count > 0)
                //        {
                //            businBook_T_Del.Delete();
                //        }
                //    }
                //    businBook_T.Delete();
                //}
                //#endregion

                #region 2.1、取数
                DataTable dtC = Power.Business.BusinessFactory.CreateBusinessOperate("Sus_Bid_InquiryList_C").FindAllByTable(string.Format("MasterId='{0}'", Id), "", "NodeId", 0, 0, SearchFlag.IgnoreRight);
                foreach (DataRow row in dtData.Rows)
                {
                    if (tempids.ContainsKey(row["Id"].ToString()) == false)
                    {
                        tempids.Add(row["Id"].ToString(), Guid.NewGuid());
                    }
                    if (XCode.Common.Helper.IsNullKey(row["ParentId"]) == false && tempids.ContainsKey(row["ParentId"].ToString()) == false)
                    {
                        tempids.Add(row["ParentId"].ToString(), Guid.NewGuid());
                    }
                    DataRow[] rowsSelect = dtC.Select(String.Format("NodeId='{0}'", row["Id"]));
                    if (rowsSelect != null && rowsSelect.Length != 0)
                    {
                        continue;
                    }
                    Power.Business.IBaseBusiness TBT = Power.Business.BusinessFactory.CreateBusiness("Sus_Bid_InquiryList_C");

                    TBT.SetItem("Id", tempids[row["Id"].ToString()]);
                    if (XCode.Common.Helper.IsNullKey(row["ParentId"]) == false)
                        TBT.SetItem("ParentId", tempids[row["ParentId"].ToString()]);
                    else
                        TBT.SetItem("ParentId", null);
                    TBT.SetItem("MasterId", Id);

                    TBT.SetItem("Code", row["code"]);
                    TBT.SetItem("Name", row["name"]);
                    TBT.SetItem("Dept", row["Unit"]);

                    TBT.Save(System.ComponentModel.DataObjectMethodType.Insert);
                    #endregion

                    #region 2.2、取数
                    Power.Business.IBusinessList TemplateList = Power.Business.BusinessFactory.CreateBusinessOperate("Sus_parameterTemplateList").FindAll("MasterId", row["Id"], Business.SearchFlag.Default);
                    foreach (Power.Business.IBaseBusiness item in TemplateList)
                    {
                        if (tempidsList.ContainsKey(item["Id"].ToString()) == false)
                        {
                            tempidsList.Add(item["Id"].ToString(), Guid.NewGuid());
                        }
                        if (XCode.Common.Helper.IsNullKey(item["ParentId"]) == false && tempidsList.ContainsKey(item["ParentId"].ToString()) == false)
                        {
                            tempidsList.Add(item["ParentId"].ToString(), Guid.NewGuid());
                        }
                        Power.Business.IBaseBusiness TBTList = Power.Business.BusinessFactory.CreateBusiness("Sus_Bid_InquiryList_C");
                        TBTList.SetItem("Id", tempidsList[item["Id"].ToString()]);
                        if (XCode.Common.Helper.IsNullKey(item["ParentId"]) == false)
                            TBTList.SetItem("ParentId", tempidsList[item["ParentId"].ToString()]);
                        else
                            TBTList.SetItem("ParentId", tempids[row["Id"].ToString()]);
                        TBTList.SetItem("MasterId", Id);
                        TBTList.SetItem("Name", item["property"]);
                        TBTList.SetItem("Parameter", item["SimpleValue"]);
                        TBTList.SetItem("Numbers", item["Amount"]);
                        TBTList.SetItem("Dept", item["Unit"]);
                        TBTList.SetItem("Memo", item["Memo"]);
                        TBTList.SetItem("Sequ", item["Sequ"]);
                        TBTList.Save(System.ComponentModel.DataObjectMethodType.Insert);
                    }
                }
                #endregion 

                return result.ToJson();
            }
        }
        #endregion
        #endregion
        #endregion


        #region  招标询价技术信息
        [ActionAttribute]
        public string Sus_Bid_InquiryList_T(string Id, string GuideId)
        {
            Power.Global.ViewResultModel result = Power.Global.ViewResultModel.Create(true, "");


            Dictionary<String, Guid> tempids = new Dictionary<String, Guid>();
            Dictionary<String, Guid> tempidsList = new Dictionary<String, Guid>();

            #region JSON格式字符串转换成对象
            var srt = JsonConvert.DeserializeObject<string[]>(GuideId);
            string sqlSrt = "";
            foreach (object itme in srt)
            {
                sqlSrt += "'" + itme + "'" + ",";
            }
            sqlSrt = sqlSrt.TrimEnd(',');
            #endregion

            #region 查询出向导中所有选中主表的数据
            string strsql = "select * from V_Sus_parameterTemplate where Id in({0})";
            DataTable dtData = XCode.DataAccessLayer.DAL.QuerySQL(string.Format(strsql, sqlSrt));
            #endregion
            #region 找到所有根节点
            List<String> listRoot = new List<string>();
            #region 判断是否有数据，有则赋值
            if (dtData.Rows.Count == 0)
            {
                result.success = false;
                result.message = "请选择一条数据！";
                return result.ToJson();
            }
            else
            {
                //#region 1、数据删除
                //Power.Business.IBusinessList businBook_T = Power.Business.BusinessFactory.CreateBusinessOperate("Sus_TechnicalBook_T").FindAll("MasterId", Id, Business.SearchFlag.Default);
                //if (businBook_T.Count > 0)
                //{
                //    foreach (Power.Business.IBaseBusiness item in businBook_T)
                //    {
                //        Power.Business.IBusinessList businBook_T_Del = Power.Business.BusinessFactory.CreateBusinessOperate("Sus_TechnicalBook_T_Del").FindAll("MasterId", item["Id"].ToString(), Business.SearchFlag.Default);
                //        if (businBook_T_Del.Count > 0)
                //        {
                //            businBook_T_Del.Delete();
                //        }
                //    }
                //    businBook_T.Delete();
                //}
                //#endregion

                #region 2.1、取数
                DataTable dtC = Power.Business.BusinessFactory.CreateBusinessOperate("Sus_Bid_InquiryList_T").FindAllByTable(string.Format("MasterId='{0}'", Id), "", "NodeId", 0, 0, SearchFlag.IgnoreRight);
                foreach (DataRow row in dtData.Rows)
                {
                    if (tempids.ContainsKey(row["Id"].ToString()) == false)
                    {
                        tempids.Add(row["Id"].ToString(), Guid.NewGuid());
                    }
                    if (XCode.Common.Helper.IsNullKey(row["ParentId"]) == false && tempids.ContainsKey(row["ParentId"].ToString()) == false)
                    {
                        tempids.Add(row["ParentId"].ToString(), Guid.NewGuid());
                    }
                    DataRow[] rowsSelect = dtC.Select(String.Format("NodeId='{0}'", row["Id"]));
                    if (rowsSelect != null && rowsSelect.Length != 0)
                    {
                        continue;
                    }
                    Power.Business.IBaseBusiness TBT = Power.Business.BusinessFactory.CreateBusiness("Sus_Bid_InquiryList_T");
                    Power.Business.IBusinessList list = Power.Business.BusinessFactory.CreateBusinessOperate("Sus_parameterTemplateList").FindAll("MasterId", row["Id"].ToString(), Business.SearchFlag.Default);
                    string temp = "";
                    foreach (Power.Business.IBaseBusiness item in list)
                    {
                        string temp2 = "";
                        if (!(XCode.Common.Helper.IsNullKey(item["property"])))
                        {
                            temp2 += string.Format("{0}{1}", item["property"], "    ");
                            if (!(XCode.Common.Helper.IsNullKey(item["SimpleValue"])))
                            {
                                temp2 = "";
                                temp2 += string.Format("{0};{1}{2}", item["property"], item["SimpleValue"], "    ");
                                if (!(XCode.Common.Helper.IsNullKey(item["Unit"])))
                                {
                                    temp2 = "";
                                    temp2 += string.Format("{0};{1};{2}{3}", item["property"], item["SimpleValue"], item["Unit"], "    ");
                                }
                                if (!(XCode.Common.Helper.IsNullKey(item["Amount"])))
                                {
                                    temp2 = "";
                                    temp2 += string.Format("{0};{1};{2};{3}{4}", item["property"], item["SimpleValue"], item["Amount"], item["Unit"], "    ");
                                }
                            }
                            else
                            {
                                if (!(XCode.Common.Helper.IsNullKey(item["Unit"])))
                                {
                                    temp2 = "";
                                    temp2 += string.Format("{0};{1}{2}", item["property"], item["Unit"], "    ");
                                }
                            }
                        }
                        else if (!(XCode.Common.Helper.IsNullKey(item["SimpleValue"])))
                        {
                            temp2 += string.Format("{0}{1}", item["SimpleValue"], "    ");
                            if (!(XCode.Common.Helper.IsNullKey(item["Unit"])))
                            {
                                temp2 = "";
                                temp2 += string.Format("{0};{1}{2}", item["SimpleValue"], item["Unit"], "    ");
                            }
                        }
                        else if (!(XCode.Common.Helper.IsNullKey(item["Unit"])))
                        {
                            temp2 += string.Format("{0}{1}", item["Unit"], "    ");
                        }

                        //赋值给上面的变量
                        temp += temp2;
                    }
                    //根节点不需要insert
                    if (listRoot.Contains(row["Id"].ToString().ToLower()))
                        continue;

                    TBT.SetItem("Id", tempids[row["Id"].ToString()]);
                    if (XCode.Common.Helper.IsNullKey(row["ParentId"]) == false)
                        TBT.SetItem("ParentId", tempids[row["ParentId"].ToString()]);
                    else
                        TBT.SetItem("ParentId", null);
                    TBT.SetItem("MasterId", Id);
                    temp = temp.TrimEnd("    ");
                    TBT.SetItem("Code", row["code"]);
                    TBT.SetItem("Name", row["name"]);
                    TBT.SetItem("Dept", row["Unit"]);
                    TBT.SetItem("Specification", temp);
                    TBT.Save(System.ComponentModel.DataObjectMethodType.Insert);
                    #endregion

                    #region 2.2、取数
                    Power.Business.IBusinessList TemplateList = Power.Business.BusinessFactory.CreateBusinessOperate("Sus_parameterTemplateList").FindAll("MasterId", row["Id"], Business.SearchFlag.Default);
                    foreach (Power.Business.IBaseBusiness item in TemplateList)
                    {
                        if (tempidsList.ContainsKey(item["Id"].ToString()) == false)
                        {
                            tempidsList.Add(item["Id"].ToString(), Guid.NewGuid());
                        }
                        if (XCode.Common.Helper.IsNullKey(item["ParentId"]) == false && tempidsList.ContainsKey(item["ParentId"].ToString()) == false)
                        {
                            tempidsList.Add(item["ParentId"].ToString(), Guid.NewGuid());
                        }
                        Power.Business.IBaseBusiness TBTList = Power.Business.BusinessFactory.CreateBusiness("Sus_Bid_InquiryList_T_Del");
                        TBTList.SetItem("Id", tempidsList[item["Id"].ToString()]);
                        if (XCode.Common.Helper.IsNullKey(item["ParentId"]) == false)
                            TBTList.SetItem("ParentId", tempidsList[item["ParentId"].ToString()]);
                        else
                            TBTList.SetItem("ParentId", null);
                        TBTList.SetItem("MasterId", tempids[row["Id"].ToString()]);
                        TBTList.SetItem("property", item["property"]);
                        TBTList.SetItem("SimpleValue", item["SimpleValue"]);
                        TBTList.SetItem("Amount", item["Amount"]);
                        TBTList.SetItem("Unit", item["Unit"]);
                        TBTList.SetItem("Memo", item["Memo"]);
                        TBTList.Save(System.ComponentModel.DataObjectMethodType.Insert);
                    }
                }
                #endregion 

                return result.ToJson();
            }
        }
        #endregion
        #endregion
        #endregion
        #region  招标询价技术信息
        [ActionAttribute]
        public string Sus_Bid_InquiryList_TNew(string Id, string GuideId)
        {
            Power.Global.ViewResultModel result = Power.Global.ViewResultModel.Create(true, "");


            Dictionary<String, Guid> tempids = new Dictionary<String, Guid>();
            Dictionary<String, Guid> tempidsList = new Dictionary<String, Guid>();

            #region JSON格式字符串转换成对象
            var srt = JsonConvert.DeserializeObject<string[]>(GuideId);
            string sqlSrt = "";
            foreach (object itme in srt)
            {
                sqlSrt += "'" + itme + "'" + ",";
            }
            sqlSrt = sqlSrt.TrimEnd(',');
            #endregion

            #region 查询出向导中所有选中主表的数据
            string strsql = "select * from V_Sus_parameterTemplate where Id in({0})";
            DataTable dtData = XCode.DataAccessLayer.DAL.QuerySQL(string.Format(strsql, sqlSrt));
            #endregion
            #region 找到所有根节点
            List<String> listRoot = new List<string>();
            #region 判断是否有数据，有则赋值
            if (dtData.Rows.Count == 0)
            {
                result.success = false;
                result.message = "请选择一条数据！";
                return result.ToJson();
            }
            else
            {
                //#region 1、数据删除
                //Power.Business.IBusinessList businBook_T = Power.Business.BusinessFactory.CreateBusinessOperate("Sus_TechnicalBook_T").FindAll("MasterId", Id, Business.SearchFlag.Default);
                //if (businBook_T.Count > 0)
                //{
                //    foreach (Power.Business.IBaseBusiness item in businBook_T)
                //    {
                //        Power.Business.IBusinessList businBook_T_Del = Power.Business.BusinessFactory.CreateBusinessOperate("Sus_TechnicalBook_T_Del").FindAll("MasterId", item["Id"].ToString(), Business.SearchFlag.Default);
                //        if (businBook_T_Del.Count > 0)
                //        {
                //            businBook_T_Del.Delete();
                //        }
                //    }
                //    businBook_T.Delete();
                //}
                //#endregion

                #region 2.1、取数
                DataTable dtC = Power.Business.BusinessFactory.CreateBusinessOperate("Sus_Bid_InquiryList_T").FindAllByTable(string.Format("MasterId='{0}'", Id), "", "NodeId", 0, 0, SearchFlag.IgnoreRight);
                foreach (DataRow row in dtData.Rows)
                {
                    if (tempids.ContainsKey(row["Id"].ToString()) == false)
                    {
                        tempids.Add(row["Id"].ToString(), Guid.NewGuid());
                    }
                    if (XCode.Common.Helper.IsNullKey(row["ParentId"]) == false && tempids.ContainsKey(row["ParentId"].ToString()) == false)
                    {
                        tempids.Add(row["ParentId"].ToString(), Guid.NewGuid());
                    }
                    DataRow[] rowsSelect = dtC.Select(String.Format("NodeId='{0}'", row["Id"]));
                    if (rowsSelect != null && rowsSelect.Length != 0)
                    {
                        continue;
                    }
                    Power.Business.IBaseBusiness TBT = Power.Business.BusinessFactory.CreateBusiness("Sus_Bid_InquiryList_T");

                    TBT.SetItem("Id", tempids[row["Id"].ToString()]);
                    if (XCode.Common.Helper.IsNullKey(row["ParentId"]) == false)
                        TBT.SetItem("ParentId", tempids[row["ParentId"].ToString()]);
                    else
                        TBT.SetItem("ParentId", null);
                    TBT.SetItem("MasterId", Id);
                    TBT.SetItem("Code", row["code"]);
                    TBT.SetItem("Name", row["name"]);
                    TBT.SetItem("Dept", row["Unit"]);

                    TBT.Save(System.ComponentModel.DataObjectMethodType.Insert);
                    #endregion

                    #region 2.2、取数
                    Power.Business.IBusinessList TemplateList = Power.Business.BusinessFactory.CreateBusinessOperate("Sus_parameterTemplateList").FindAll("MasterId", row["Id"], Business.SearchFlag.Default);
                    foreach (Power.Business.IBaseBusiness item in TemplateList)
                    {
                        if (tempidsList.ContainsKey(item["Id"].ToString()) == false)
                        {
                            tempidsList.Add(item["Id"].ToString(), Guid.NewGuid());
                        }
                        if (XCode.Common.Helper.IsNullKey(item["ParentId"]) == false && tempidsList.ContainsKey(item["ParentId"].ToString()) == false)
                        {
                            tempidsList.Add(item["ParentId"].ToString(), Guid.NewGuid());
                        }
                        Power.Business.IBaseBusiness TBTList = Power.Business.BusinessFactory.CreateBusiness("Sus_Bid_InquiryList_T");
                        TBTList.SetItem("Id", tempidsList[item["Id"].ToString()]);
                        if (XCode.Common.Helper.IsNullKey(item["ParentId"]) == false)
                            TBTList.SetItem("ParentId", tempidsList[item["ParentId"].ToString()]);
                        else
                            TBTList.SetItem("ParentId", tempids[row["Id"].ToString()]);
                        TBTList.SetItem("MasterId", Id);
                        TBTList.SetItem("Name", item["property"]);
                        TBTList.SetItem("Specification", item["SimpleValue"]);
                        TBTList.SetItem("Numbers", item["Amount"]);
                        TBTList.SetItem("Dept", item["Unit"]);
                        TBTList.SetItem("Memo", item["Memo"]);
                        TBTList.SetItem("Sequ", item["Sequ"]);
                        TBTList.Save(System.ComponentModel.DataObjectMethodType.Insert);
                    }
                }
                #endregion 

                return result.ToJson();
            }
        }
        #endregion
        #endregion
        #endregion
        #region  招标询价技术信息导入到商务信息
        [ActionAttribute]
        public string SP_Sus_Bid_InquiryList_C_Insert(string Id, string GuideId)
        {
            Power.Global.ViewResultModel result = Power.Global.ViewResultModel.Create(true, "");

            Dictionary<String, Guid> tempids = new Dictionary<String, Guid>();

            #region JSON格式字符串转换成对象
            var srt = JsonConvert.DeserializeObject<string[]>(GuideId);
            string sqlSrt = "";
            foreach (object itme in srt)
            {
                sqlSrt += "'" + itme + "'" + ",";
            }
            sqlSrt = sqlSrt.TrimEnd(',');
            #endregion

            #region 查询出向导中所有选中主表的数据
            string strsql = "select * from Sus_Bid_InquiryList_T where Id in({0})";
            DataTable dtData = XCode.DataAccessLayer.DAL.QuerySQL(string.Format(strsql, sqlSrt));
            #endregion
            #region 找到所有根节点
            List<String> listRoot = new List<string>();
            #region 判断是否有数据，有则赋值
            if (dtData.Rows.Count == 0)
            {
                result.success = false;
                result.message = "请选择一条数据！";
                return result.ToJson();
            }

            #region 2.1、取数
            DataTable dtC = Power.Business.BusinessFactory.CreateBusinessOperate("Sus_Bid_InquiryList_C").FindAllByTable(string.Format("MasterId='{0}'", Id), "", "NodeId", 0, 0, SearchFlag.IgnoreRight);
            foreach (DataRow row in dtData.Rows)
            {
                if (tempids.ContainsKey(row["Id"].ToString()) == false)
                {
                    tempids.Add(row["Id"].ToString(), Guid.NewGuid());
                }
                if (XCode.Common.Helper.IsNullKey(row["ParentId"]) == false && tempids.ContainsKey(row["ParentId"].ToString()) == false)
                {
                    tempids.Add(row["ParentId"].ToString(), Guid.NewGuid());
                }
                DataRow[] rowsSelect = dtC.Select(String.Format("NodeId='{0}'", row["Id"]));
                if (rowsSelect != null && rowsSelect.Length != 0)
                {
                    continue;
                }
                Power.Business.IBaseBusiness TBT = Power.Business.BusinessFactory.CreateBusiness("Sus_Bid_InquiryList_C");

                TBT.SetItem("Id", tempids[row["Id"].ToString()]);
                if (XCode.Common.Helper.IsNullKey(row["ParentId"]) == false)
                    TBT.SetItem("ParentId", tempids[row["ParentId"].ToString()]);
                else
                    TBT.SetItem("ParentId", null);
                TBT.SetItem("MasterId", Id);

                TBT.SetItem("Code", row["Code"]);
                TBT.SetItem("Name", row["Name"]);
                TBT.SetItem("Parameter", row["Specification"]);
                TBT.SetItem("Numbers", row["Numbers"]);
                TBT.SetItem("Dept", row["Dept"]);
                TBT.SetItem("Sequ", row["Sequ"]);

                TBT.Save(System.ComponentModel.DataObjectMethodType.Insert);
                #endregion

            }
            return result.ToJson();
        }
        #region 技术规格书导入
        [ActionAttribute]
        public string Sus_TechnicalBook_T_ImPort(Guid fileid, string Id)
        {
            Power.Global.ViewResultModel result = Power.Global.ViewResultModel.Create(true, "");
            String filename = getLocalFileName(fileid);
            if (String.IsNullOrEmpty(filename))
            {
                result.success = false;
                result.message = "获取上传的导入文件失败";
                return result.ToJson();
            }
            if (!System.IO.File.Exists(filename))
                throw new Exception("Excel文件不存在.");
            DataSet data = new DataSet();
            Workbook workbook = new Workbook(filename);
            foreach (Worksheet ws in workbook.Worksheets)
            {
                Cells ce = ws.Cells;
                if (ce.MaxDataRow == -1 || ce.MaxDataColumn == -1)
                    continue;
                try
                {
                    Dictionary<string, Guid> tempids = new Dictionary<string, Guid>();//存储键值对Code为键Id为值。
                    DataTable dtTemp = ce.ExportDataTable(0, 0, ce.MaxDataRow + 1, ce.MaxDataColumn + 1, true);
                    dtTemp.TableName = ws.Name;
                    data.Tables.Add(dtTemp);
                    if (dtTemp.TableName == "技术规格信息")
                    {
                        foreach (DataRow row in dtTemp.Rows)
                        {
                            Power.Business.IBaseBusiness busin = Power.Business.BusinessFactory.CreateBusiness("Sus_TechnicalBook_T");
                            string leng = row[6].ToString();
                            string PreviousValue = leng.LastIndexOf(".").ToString();

                            if (PreviousValue.Equals("-1"))
                            {
                                Guid NewId = Guid.NewGuid();
                                busin.SetItem("Id", NewId);
                                busin.SetItem("ParentId", "00000000-0000-0000-0000-000000000000");
                                busin.SetItem("MasterId", Id);
                                tempids.Add(row[6].ToString(), NewId);

                                if (!(XCode.Common.Helper.IsNullKey(row[0])))
                                {
                                    busin.SetItem("Code", row[0]);
                                }
                                if (!(XCode.Common.Helper.IsNullKey(row[1])))
                                {
                                    busin.SetItem("Name", row[1]);
                                }
                                if (!(XCode.Common.Helper.IsNullKey(row[2])))
                                {
                                    busin.SetItem("Specification", row[2]);
                                }
                                if (!(XCode.Common.Helper.IsNullKey(row[3])))
                                {
                                    busin.SetItem("Numbers", row[3]);
                                }
                                if (!(XCode.Common.Helper.IsNullKey(row[4])))
                                {
                                    busin.SetItem("Dept", row[4]);
                                }
                                if (!(XCode.Common.Helper.IsNullKey(row[5])))
                                {
                                    busin.SetItem("Memo", row[5]);
                                }
                                busin.Save(System.ComponentModel.DataObjectMethodType.Insert);
                            }
                            else if (int.Parse(PreviousValue) > 2)
                            {
                                Guid NewId = Guid.NewGuid();
                                string a = leng.Substring(0, int.Parse(PreviousValue));
                                tempids.Add(row[6].ToString(), NewId);
                                busin.SetItem("Id", NewId);
                                busin.SetItem("ParentId", tempids[a]);//用Code找到Id赋值给ParentId
                                busin.SetItem("MasterId", Id);


                                if (!(XCode.Common.Helper.IsNullKey(row[0])))
                                {
                                    busin.SetItem("Code", row[0]);
                                }
                                if (!(XCode.Common.Helper.IsNullKey(row[1])))
                                {
                                    busin.SetItem("Name", row[1]);
                                }
                                if (!(XCode.Common.Helper.IsNullKey(row[2])))
                                {
                                    busin.SetItem("Specification", row[2]);
                                }
                                if (!(XCode.Common.Helper.IsNullKey(row[3])))
                                {
                                    busin.SetItem("Numbers", row[3]);
                                }
                                if (!(XCode.Common.Helper.IsNullKey(row[4])))
                                {
                                    busin.SetItem("Dept", row[4]);
                                }
                                if (!(XCode.Common.Helper.IsNullKey(row[5])))
                                {
                                    busin.SetItem("Memo", row[5]);
                                }
                                busin.Save(System.ComponentModel.DataObjectMethodType.Insert);
                            }
                            else
                            {
                                result.success = false;
                                result.message = "长编号列必须输入按规格输入,否则无法生成树结构显示！";
                                result.ToJson();
                            }
                        }
                    }
                    else
                    {
                        result.success = false;
                        result.message = "导入Excel文件页签名称不匹配！必须是技术规格信息！";
                        return result.ToJson();
                    }
                }
                catch (CellsException cex)
                {
                    string message = cex.Message;
                    string[] ms = message.Split(' ');
                    string errorcells = "";
                    for (int i = 0; i < ms.Length; i++)
                    {
                        if (ms[i].ToLower() == "cell" && i + 2 < ms.Length && ms[i + 2] == "should".ToLower())
                            errorcells = ms[i + 1].ToUpper();
                    }
                    if (errorcells != "")
                        throw new Exception("单元格【" + errorcells + "】格式不正确");
                    else
                        throw new Exception("单元格格式不正确：" + message);
                }
                finally
                {

                    if (File.Exists(filename))
                    {
                        File.Delete(filename);
                    }

                }
            }

            return result.ToJson();
        }



        /// <summary>
        /// app单点登录
        /// </summary>
        /// <param name="code"></param>
        /// <returns></returns>
        [ActionAttribute(Authorize = false)]
        public string getAppMoblie(string code)
        {
            string sql = "select * from PB_User where code = '" + code + "' ";
            DataTable userList = XCode.DataAccessLayer.DAL.QuerySQL(sql);
            string pwd = "";
            if (userList.Rows[0]["PassWord"] != null)
            {
                pwd = userList.Rows[0]["PassWord"].ToString();
            }
            Power.Controls.Action.ILoginAction loginAct = new Power.Controls.Action.LoginAction();
            Power.Global.ViewResultModel result = loginAct.Login(code, pwd, "zh-CN");



            if (result.success && result.data["sessionid"] != null)
            {

                var token = this.IssueToken(result.data["sessionid"], result.data["humanid"], code);
                Power.Global.PowerGlobal.getSession(result.data["sessionid"].ToString(), true, true);
                Power.IBaseCore.ISession ss = Power.Global.PowerGlobal.getSession(result.data["sessionid"].ToString());
                var rs = GetLogUserdata(ss);
                result.data.Add("token", token);
                result.data.Add("UserData", rs.data);
            }

            return result.ToJson();
        }

        //加密密钥
        private static string secret = "41894BD9548041D4AC2E0E4EF13DFEC8";
        //颁发签证
        public string IssueToken(object sessionid, object humanid, object usercode)
        {
            var now = DateTime.UtcNow;
            var t = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
            var iat = (long)(now - t).TotalSeconds;//签发时间戳
            var exp = (long)(now.AddDays(20) - t).TotalSeconds;//过期时间戳(20天过期)
            var jti = Guid.NewGuid();//JWT唯一身份标识(可存入数据库回避重放攻击)
            var payload = new Dictionary<string, object>
            {
                { "iss", "普华科技" },
                { "sub", "PowerMobile" },
                { "iat", iat },
                { "exp", exp },
                { "jti", jti },
                { "usc", usercode },
                { "humanid", humanid },
                { "sessionid", sessionid },
            };
            IJwtAlgorithm algorithm = new HMACSHA256Algorithm();
            IJsonSerializer serializer = new JsonNetSerializer();
            IBase64UrlEncoder urlEncoder = new JwtBase64UrlEncoder();
            IJwtEncoder encoder = new JwtEncoder(algorithm, serializer, urlEncoder);
            return encoder.Encode(payload, secret);
        }

        /// <summary>
        /// SS读取userData
        /// </summary>
        /// <param name="ss"></param>
        /// <returns></returns>
        private Power.Global.ViewResultModel GetLogUserdata(Power.IBaseCore.ISession ss)
        {
            Power.Global.ViewResultModel result = Power.Global.ViewResultModel.Create(true, "");
            result.data.Add("lang", ss.Language);
            result.data.Add("epsProjId", ss.EpsProjId);
            result.data.Add("epsProjName", ss.EpsProjName);
            result.data.Add("sessionId", ss.SessionId);
            result.data.Add("userName", ss.UserName);
            result.data.Add("userCode", ss.UserCode);
            result.data.Add("posiId", ss.PosiId);
            result.data.Add("posiName", ss.PosiName);
            result.data.Add("humLogo", "");
            result.data.Add("humName", ss.HumanName);
            result.data.Add("humId", ss.HumanId);

            return result;
        }
        #endregion
    }
    #endregion
    #endregion
    #endregion





}


