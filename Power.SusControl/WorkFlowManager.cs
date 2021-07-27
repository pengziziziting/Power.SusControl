using Newtonsoft.Json.Linq;
using Power.Controls.PMS;
using Power.Global;
using Power.IBaseCore;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;

namespace Power.SusControl
{
    public class WorkFlowManager
    {
        private APIControl apiControl;

        public WorkFlowManager(ISession session, HttpContext context)
        {
            apiControl = new APIControl();
            apiControl.session = session;
            apiControl.Context = context;
        }

        /// <summary>
        /// 获取流程列表
        /// </summary>
        /// <param name="FormId"></param>
        /// <param name="KeyWord"></param>
        /// <param name="KeyValue"></param>
        /// <returns></returns>
        public List<Hashtable> getWorkFlowList(String FormId, String KeyWord, String KeyValue)
        {
            SelectWorkFlowJson selectInfo = new SelectWorkFlowJson(FormId, KeyWord, KeyValue);
            ViewResultModel result = execAPIMessage(selectInfo);
            if (result.success)
            {

                return jArray2List(result.data["WorkFlowList"]);

            }
            return null;
        }

        /// <summary>
        /// 激活流程
        /// </summary>
        /// <param name="workfowid"></param>
        /// <param name="version"></param>
        /// <param name="FormId"></param>
        /// <param name="KeyWord"></param>
        /// <param name="KeyValue"></param>
        /// <returns></returns>
        public ViewResultModel activeWorkFlow(String workfowid, String version, String FormId, String KeyWord, String KeyValue)
        {
            ActiveWorkFlowJson activeInfo = new ActiveWorkFlowJson(workfowid, version, FormId, KeyWord, KeyValue);
            ViewResultModel result = execAPIMessage(activeInfo);
            return result;
        }
        /// <summary>
        /// 表单发起流程
        /// </summary>
        /// <param name="FormId"></param>
        /// <param name="KeyWord"></param>
        /// <param name="KeyValue"></param>
        /// <param name="userList">humanid,humanname</param>
        /// <returns></returns>
        public bool autoStartWorkFlow(String FormId, String KeyWord, String KeyValue, Dictionary<String, String> userList)
        {
            List<Hashtable> workflowlist = getWorkFlowList(FormId, KeyWord, KeyValue);
            if (workflowlist == null || workflowlist.Count == 0)
                return false;
            Hashtable workflowItem = workflowlist[0];
            String workflowid = workflowItem["WorkFlowID"].ToString();
            String workflowversion = workflowItem["Version"].ToString();
            ViewResultModel actiResult = activeWorkFlow(workflowid, workflowversion, FormId, KeyWord, KeyValue);
            if (actiResult.success == false)
                return false;
            //提取激活流程后的信息
            //1、下一个节点名称
            List<Hashtable> nextList = jArray2List(actiResult.data["NextNodeList"]);
            String nextNodeCode = "";
            if (nextList != null && nextList.Count != 0)
            {
                nextNodeCode = nextList[0]["NodeCode"].ToString();
            }
            Hashtable Current = jObject2Hashtable(actiResult.data["Current"]);
            //2、拼接 SelectedNode
            List<Hashtable> SelectedNode = new List<Hashtable>();
            Hashtable selectNode = new Hashtable();
            SelectedNode.Add(selectNode);
            selectNode.Add("CopyUserList", new ArrayList());
            selectNode.Add("NodeCode", nextNodeCode);
            List<Hashtable> SendUserList = new List<Hashtable>();
            selectNode.Add("SendUserList", SendUserList);
            foreach (KeyValuePair<string, string> p in userList)
            {
                Hashtable senduser = new Hashtable();
                SendUserList.Add(senduser);
                senduser.Add("UserID", p.Key);
                senduser.Add("SourceUserID", p.Key);
                senduser.Add("UserName", p.Value);
                senduser.Add("SourceUserName", p.Value);
                senduser.Add("DeptPositionID", "");
                senduser.Add("DeptPositionName", "");
                senduser.Add("SourceMode", "30");
                senduser.Add("PlanEndDate", Current["PlanEndDate"]);
            }

            //{
            //    "SendUserList": [
            //            {
            //                "UserID": "7b09b7e0-a5f0-4b07-ab28-3fb90fd6dad5",
            //                "SourceUserID": "7b09b7e0-a5f0-4b07-ab28-3fb90fd6dad5",
            //                "UserName": "ceshi1",
            //                "SourceUserName": "ceshi1",
            //                "DeptPositionID": "1aa0183a-fbf4-4e44-81af-ef091f656ea9",
            //                "DeptPositionName": "项目经理",
            //                "SourceMode": 40,
            //                "PlanEndDate": "2018-09-04T15:26:41"
            //            }
            //        ],
            //        "CopyUserList": [],
            //        "NodeCode": "demo_node_2"
            //    }

            //送审
            SendWorkFlowJons sendInfo = new SendWorkFlowJons(Current, SelectedNode);
            ViewResultModel result = execAPIMessage(sendInfo);

            return result.success;
        }

        private Hashtable jObject2Hashtable(Object jobject)
        {
            if (jobject == null || jobject.GetType().Name != "JObject")
                return null;
            Hashtable temp = new Hashtable();
            JObject item = (JObject)jobject;
            foreach (var obj in item)
                temp.Add(obj.Key, obj.Value);
            return temp;
        }

        private List<Hashtable> jArray2List(Object array)
        {
            if (array == null || array.GetType().Name != "JArray")
                return null;
            JArray list = (JArray)array;
            List<Hashtable> rest = new List<Hashtable>();
            foreach (JObject item in list)
            {
                Hashtable temp = jObject2Hashtable(item);
                if (temp != null)
                    rest.Add(temp);
            }
            return rest;
        }


        private ViewResultModel execAPIMessage(WorkFlowJson workflow)
        {
            String strjson = apiControl.APIMessage(workflow.toString());
            ViewResultModel result = Newtonsoft.Json.JsonConvert.DeserializeObject<ViewResultModel>(strjson);
            return result;
        }

        private class WorkFlowJson
        {

            public WorkFlowJson()
            {
                this.MessageCode = "Power.WorkFlows.Actions.RecvFlowOperate";
                this.data = new Hashtable();
            }
            public String MessageCode { get; set; }

            public Hashtable data { get; set; }

            //        {
            //    "MessageCode": "Power.WorkFlows.Actions.RecvFlowOperate",
            //    "data": {
            //        "FormId": "dc82d741-0edb-4ab9-9d1e-7e3d4c975bde",
            //        "KeyValue": "511cd072-7cf0-4f7a-863d-7a54f68937ef",
            //        "KeyWord": "Sus_Pur_ExpertReview",
            //        "SequeID": "-1",
            //        "FlowOperate": "SelectFlow"
            //    }
            //}

            public String toString()
            {
                return Newtonsoft.Json.JsonConvert.SerializeObject(this);
            }
        }

        private class SelectWorkFlowJson : WorkFlowJson
        {
            public SelectWorkFlowJson(string FormId, string KeyWord, string KeyValue)
            {
                this.data.Add("FormId", FormId);
                this.data.Add("KeyValue", KeyValue);
                this.data.Add("KeyWord", KeyWord);
                this.data.Add("SequeID", "-1");
                this.data.Add("FlowOperate", "SelectFlow");
            }
        }

        private class ActiveWorkFlowJson : WorkFlowJson
        {

            //    "MessageCode": "Power.WorkFlows.Actions.RecvFlowOperate",
            //    "data": {
            //        "Current": {
            //            "WorkFlowID": "664d3fa4-2b48-4923-bb3a-aff260c42ee8",
            //            "Version": "1.0.0.1",
            //            "FormId": "dc82d741-0edb-4ab9-9d1e-7e3d4c975bde",
            //            "KeyWord": "Sus_Pur_ExpertReview",
            //            "KeyValue": "511cd072-7cf0-4f7a-863d-7a54f68937ef",
            //            "WorkInfoID": "b5cb6abd-15f9-b926-da02-8adf59226e3e"
            //        },
            //        "FlowOperate": "Active"
            //    }
            //}
            public ActiveWorkFlowJson(String workfowid, String version, String FormId, String KeyWord, String KeyValue)
            {
                Hashtable Current = new Hashtable();
                this.data.Add("Current", Current);
                this.data.Add("FlowOperate", "Active");
                Current.Add("WorkFlowID", workfowid);
                Current.Add("Version", version);
                Current.Add("WorkInfoID", Guid.NewGuid());
                Current.Add("FormId", FormId);
                Current.Add("KeyWord", KeyWord);
                Current.Add("KeyValue", KeyValue);
            }
        }


        private class SendWorkFlowJons : WorkFlowJson
        {

            //{
            //    "OpenTrans": "true",
            //    "Active": {
            //        "MessageCode": "Power.WorkFlows.Actions.RecvFlowOperate",
            //        "data": {
            //            "Current": {
            //                "WorkInfoID": "0b6deb1b-d51a-4c43-b539-2f9ba090a82f",
            //                "WorkFlowID": "664d3fa4-2b48-4923-bb3a-aff260c42ee8",
            //                "Version": "1.0.0.1",
            //                "NodeCode": "demo_node_1",
            //                "NodeName": "开始",
            //                "SequeID": 1,
            //                "PlanEndDate": "2018-09-05T14:41:05",
            //                "FormId": null,
            //                "KeyWord": null,
            //                "KeyValue": null,
            //                "RecordStatus": 15
            //            },
            //            "SelectedNode": [
            //                {
            //                    "SendUserList": [
            //                        {
            //                            "UserID": "7b09b7e0-a5f0-4b07-ab28-3fb90fd6dad5",
            //                            "SourceUserID": "7b09b7e0-a5f0-4b07-ab28-3fb90fd6dad5",
            //                            "UserName": "ceshi1",
            //                            "SourceUserName": "ceshi1",
            //                            "DeptPositionID": "1aa0183a-fbf4-4e44-81af-ef091f656ea9",
            //                            "DeptPositionName": "项目经理",
            //                            "SourceMode": 40,
            //                            "PlanEndDate": "2018-09-04T15:26:41"
            //                        }
            //                    ],
            //                    "CopyUserList": [],
            //                    "NodeCode": "demo_node_2"
            //                }
            //            ],
            //            "MindInfo": "",
            //            "VoteValue": "",
            //            "VoteText": "",
            //            "FlowOperate": "Send"
            //        }
            //    }
            //}
            public SendWorkFlowJons(Hashtable Current, List<Hashtable> SelectedNode)
            {
                this.data.Add("Current", Current);
                this.data.Add("SelectedNode", SelectedNode);
                this.data.Add("MindInfo", "");
                this.data.Add("VoteValue", "");
                this.data.Add("VoteText", "");
                this.data.Add("FlowOperate", "Send");

            }

            public string toString()
            {
                Hashtable result = new Hashtable();
                result.Add("OpenTrans", "true");
                result.Add("Active", this);
                return Newtonsoft.Json.JsonConvert.SerializeObject(result);
            }
               




   
            }

    }
}
