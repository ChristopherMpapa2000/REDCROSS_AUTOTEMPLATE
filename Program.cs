using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Net;
using System.Net.Mail;
using Newtonsoft.Json.Linq;
using WOLF_START_MigrateDAR;
using Microsoft.Exchange.WebServices.Data;
using System.Text;
using System.IO;
using Attachment = System.Net.Mail.Attachment;
using WolfApprove.Model;
using System.Globalization;
using Newtonsoft.Json;

namespace AutoTemplate
{
    class Program
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(typeof(Program));

        private static string dbConnectionString
        {
            get
            {
                var ServarName = ConfigurationManager.AppSettings["ServarName"];
                var Database = ConfigurationManager.AppSettings["Database"];
                var Username_database = ConfigurationManager.AppSettings["Username_database"];
                var Password_database = ConfigurationManager.AppSettings["Password_database"];
                var dbConnectionString = $"data source={ServarName};initial catalog={Database};persist security info=True;user id={Username_database};password={Password_database};Connection Timeout=200";

                if (!string.IsNullOrEmpty(dbConnectionString))
                {
                    return dbConnectionString;
                }
                return "";
            }
        }
        private static int iIntervalTime
        {
            //ตั้งค่าเวลา
            get
            {
                var IntervalTime = ConfigurationManager.AppSettings["IntervalTimeMinute"];
                if (!string.IsNullOrEmpty(IntervalTime))
                {
                    return Convert.ToInt32(IntervalTime);
                }
                return -10;
            }
        }
        static void Main(string[] args)
        {
            try
            {
                log4net.Config.XmlConfigurator.Configure();
                log.Info("====== Start Process AutoTemplate ====== : " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                log.Info(string.Format("Run batch as :{0}", System.Security.Principal.WindowsIdentity.GetCurrent().Name));

                DbwolfDataContext db = new DbwolfDataContext(dbConnectionString);
                if (db.Connection.State == ConnectionState.Open)
                {
                    db.Connection.Close();
                    db.Connection.Open();
                }
                db.Connection.Open();
                db.CommandTimeout = 0;

                GetTemplate(db);
            }
            catch (Exception ex)
            {
                Console.WriteLine(":ERROR");
                Console.WriteLine("exit 1");

                log.Error(":ERROR");
                log.Error("message: " + ex.Message);
                log.Error("Exit ERROR");
            }
            finally
            {
                log.Info("====== End Process Process AutoTemplate ====== : " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));

            }
        }
        public static void GetTemplate(DbwolfDataContext db)
        {
            List<TRNMemo> lstmemo = new List<TRNMemo>();
            List<ViewEmployee> listemp = db.ViewEmployees.Where(v => v.IsActive == true).ToList();
            List<MSTRole> lisRoles = db.MSTRoles.ToList();

            var MemoId_TEST = ConfigurationManager.AppSettings["MemoId_TEST"];
            if (!string.IsNullOrEmpty(MemoId_TEST))
            {
                lstmemo = db.TRNMemos.Where(x => x.MemoId == Convert.ToInt32(MemoId_TEST)).ToList();
            }
            else
            {
                var Templateform = ConfigurationManager.AppSettings["Templateform"];
                int Templateform_id = db.MSTTemplates.Where(f => f.DocumentCode == Templateform && f.IsActive == true).Select(a => a.TemplateId).FirstOrDefault();
                lstmemo = db.TRNMemos.Where(x => x.TemplateId == Templateform_id && x.StatusName == "Completed" && x.ModifiedDate >= DateTime.Now.AddMinutes(iIntervalTime)).ToList();
            }
            if (lstmemo.Count > 0)
            {
                foreach (var itemmemo in lstmemo)
                {
                    Console.WriteLine("Start Getdata : " + itemmemo.MemoId);
                    log.Info("Start Getdata : " + itemmemo.MemoId);

                    string Subject = "";
                    string DateStart = "";
                    string DateEnd = "";
                    List<object> List_of_procurement = new List<object>();
                    JObject jsonAdvanceForm = JsonUtils.createJsonObject(itemmemo.MAdvancveForm);
                    JArray itemsArray = (JArray)jsonAdvanceForm["items"];
                    foreach (JObject jItems in itemsArray)
                    {

                        JArray jLayoutArray = (JArray)jItems["layout"];
                        if (jLayoutArray.Count >= 1)
                        {
                            JObject jTemplateL = (JObject)jLayoutArray[0]["template"];
                            JObject jData = (JObject)jLayoutArray[0]["data"];
                            if ((String)jTemplateL["label"] == "ชื่อฟอร์มข้อมูลประเมินผู้ค้า")
                            {
                                Subject = jData["value"].ToString();
                            }
                            if ((String)jTemplateL["label"] == "ระยะเวลาการประเมิน")
                            {
                                DateStart = jData["value"].ToString();
                            }
                            if ((String)jTemplateL["label"] == "รายการผู้ค้าที่จัดซื้อจัดจ้าง")
                            {

                                foreach (JArray row in jData["row"])
                                {
                                    List<object> rowObject = new List<object>();
                                    foreach (JObject item in row)
                                    {
                                        rowObject.Add(item["value"].ToString());
                                    }
                                    List_of_procurement.Add(rowObject);
                                }
                            }
                            if (jLayoutArray.Count > 1)
                            {
                                JObject jTemplateR = (JObject)jLayoutArray[1]["template"];
                                JObject jData2 = (JObject)jLayoutArray[1]["data"];
                                if ((String)jTemplateR["label"] == "วันที่สิ้นสุด")
                                {
                                    DateEnd = jData2["value"].ToString();
                                }
                            }
                        }
                    }
                    if (List_of_procurement.Count() > 0)
                    {
                        foreach (var Eitem in List_of_procurement)
                        {
                            dynamic item = Eitem;
                            string Trader_type = item[0];
                            string Company = item[1];
                            string[] List_DepartmentCode = item[2].Split(',');
                            if (Trader_type == "ผู้ขายสินค้า")
                            {
                                log.Info("StartMemo TemplateFM02");
                                var TemplateToFM02 = ConfigurationManager.AppSettings["TemplateToFM02"];
                                MSTTemplate TemplatFM02 = db.MSTTemplates.Where(f => f.DocumentCode == TemplateToFM02 && f.IsActive == true).FirstOrDefault();
                                StartAutoTemplate(TemplatFM02, db, Company, List_DepartmentCode, listemp, lisRoles);
                                log.Info("EndMemo TemplateFM02");
                            }
                            if (Trader_type == "ผู้เช่า / ให้บริการ")
                            {
                                log.Info("StartMemo TemplateFM03");
                                var TemplateToFM03 = ConfigurationManager.AppSettings["TemplateToFM03"];
                                MSTTemplate TemplateFM03 = db.MSTTemplates.Where(f => f.DocumentCode == TemplateToFM03 && f.IsActive == true).FirstOrDefault();
                                StartAutoTemplate(TemplateFM03, db, Company, List_DepartmentCode, listemp, lisRoles);
                                log.Info("EndMemo TemplateFM02");
                            }
                        }
                    }
                    log.Info("----------------------------------------------------------------------------");
                }
            }
            else
            {
                Console.WriteLine("FM-01 : " + lstmemo.Count);
                log.Info("FM-01 : " + lstmemo.Count);
            }
        }
        public static void StartAutoTemplate(MSTTemplate sTemplateTo, DbwolfDataContext db, string Company, string[] List_DepartmentCode, List<ViewEmployee> listemp, List<MSTRole> lisRoles)
        {

            List<string> ListMAdvancveform = new List<string>();
            List<int> listEmp_id = new List<int>();

            foreach (var item in List_DepartmentCode)
            {
                int Role_id = lisRoles.Where(r => r.NameEn.Contains(item + " ผู้ประเมิน") || r.NameTh.Contains(item + " ผู้ประเมิน")).Select(s => s.RoleId).FirstOrDefault();
                if (Role_id != null)
                {
                    var Emp_id = db.MSTUserPermissions.Where(u => u.RoleId == Role_id).Select(e => e.EmployeeId).ToList();
                    listEmp_id.AddRange(Emp_id);
                }
            }
            listEmp_id = listEmp_id.GroupBy(x => x).Select(g => g.First()).ToList();
            foreach (var item in listEmp_id)
            {
                #region ชื่อสัญญา
                var TemplateContract = ConfigurationManager.AppSettings["TemplateContract"];
                var TemplatFM02 = db.MSTTemplates.Where(f => f.DocumentCode == TemplateContract && f.IsActive == true).FirstOrDefault();
                var listmemoform = db.TRNMemoForms.Where(x => x.TemplateId == TemplatFM02.TemplateId && (x.obj_label.Contains("ชื่อผู้ประกอบการ คู่สัญญา") || x.obj_label.Contains("ชื่อผู้ประกอบการ"))).ToList();
                if (listmemoform.Count > 0)
                {
                    var ListContract = listmemoform.Where(x => x.obj_value == Company).GroupBy(x => x.MemoId).Select(g => g.Key).ToList();
                    if (ListContract.Count > 0)
                    {
                        foreach (var items in ListContract)
                        {
                            var NameContract = db.TRNMemoForms.Where(x => x.MemoId == items && x.obj_label.Contains("ชื่อสัญญา")).Select(o => o.obj_value).FirstOrDefault();
                            if (NameContract != null)
                            {
                                var emp = listemp.Where(e => e.EmployeeId == item && e.IsActive == true).FirstOrDefault();
                                if (emp != null)
                                {
                                    string MAdvancveform = StrMAdvancveform(sTemplateTo, Company, emp, NameContract: NameContract);
                                    if (!string.IsNullOrEmpty(MAdvancveform))
                                    {
                                        InsertTrnMemo(MAdvancveform, db, sTemplateTo, Company, emp);
                                    }
                                }
                                else
                                {
                                    log.Info("Not have EmployeeID: " + item);
                                }
                            }
                            else
                            {
                                var emp = listemp.Where(e => e.EmployeeId == item && e.IsActive == true).FirstOrDefault();
                                if (emp != null)
                                {
                                    string MAdvancveform = StrMAdvancveform(sTemplateTo, Company, emp);
                                    if (!string.IsNullOrEmpty(MAdvancveform))
                                    {
                                        InsertTrnMemo(MAdvancveform, db, sTemplateTo, Company, emp);
                                    }
                                }
                                else
                                {
                                    log.Info("Not have EmployeeID: " + item);
                                }
                            }
                        }
                    }
                    else
                    {
                        var emp = listemp.Where(e => e.EmployeeId == item && e.IsActive == true).FirstOrDefault();
                        if (emp != null)
                        {
                            string MAdvancveform = StrMAdvancveform(sTemplateTo, Company, emp);
                            if (!string.IsNullOrEmpty(MAdvancveform))
                            {
                                InsertTrnMemo(MAdvancveform, db, sTemplateTo, Company, emp);
                            }
                        }
                        else
                        {
                            log.Info("Not have EmployeeID: " + item);
                        }
                    }
                }
                else
                {
                    var emp = listemp.Where(e => e.EmployeeId == item && e.IsActive == true).FirstOrDefault();
                    if (emp != null)
                    {
                        string MAdvancveform = StrMAdvancveform(sTemplateTo, Company, emp);
                        if (!string.IsNullOrEmpty(MAdvancveform))
                        {
                            InsertTrnMemo(MAdvancveform, db, sTemplateTo, Company, emp);
                        }
                    }
                    else
                    {
                        log.Info("Not have EmployeeID: " + item);
                    }
                }
                #endregion
            }
        }

        public static string StrMAdvancveform(MSTTemplate sTemplateToFM, string Company, ViewEmployee emp,string NameContract ="")
        {
            string currentDate = DateTime.Now.ToString("dd MMM yyyy", new CultureInfo("en-US"));
            JObject jsonAdvanceForm = JsonUtils.createJsonObject(sTemplateToFM.AdvanceForm);
            JArray itemsArray = (JArray)jsonAdvanceForm["items"];
            foreach (JObject jItems in itemsArray)
            {
                JArray jLayoutArray = (JArray)jItems["layout"];
                if (jLayoutArray.Count >= 1)
                {
                    JObject jTemplateL = (JObject)jLayoutArray[0]["template"];
                    JObject jData = (JObject)jLayoutArray[0]["data"];
                    if ((String)jTemplateL["label"] == "ชื่อเรื่อง")
                    {
                        string Subject = "";
                        if (jData != null)
                        {
                            if (sTemplateToFM.DocumentCode == "FM-02")
                            {
                                Subject = "แบบประเมินผู้ค้าประเภทผู้ขายสินค้า";
                            }
                            else if (sTemplateToFM.DocumentCode == "FM-03")
                            {
                                Subject = "แบบประเมินผู้ค้าประเภทผู้เช่า / ให้บริการ";
                            }
                            jData["value"] = $"{Subject} : {Company}";
                        }
                    }
                    if ((String)jTemplateL["label"] == "ชื่อผู้ประเมิน")
                    {
                        if (jData != null)
                        {
                            jData["value"] = emp.NameTh;
                        }
                    }
                    if ((String)jTemplateL["label"] == "ตำแหน่ง")
                    {
                        if (jData != null)
                        {
                            jData["value"] = emp.PositionNameTh;
                        }
                    }
                    if ((String)jTemplateL["label"] == "ประเมิน ณ วันที่")
                    {
                        if (jData != null)
                        {
                            jData["value"] = currentDate;
                        }
                    }
                    if ((String)jTemplateL["label"] == "ชื่อผู้เช่า / ให้บริการ")
                    {
                        if (jData != null)
                        {
                            jData["value"] = Company;
                        }
                    }
                    if ((String)jTemplateL["label"] == "ชื่อผู้ขายสินค้า")
                    {
                        if (jData != null)
                        {
                            jData["value"] = Company;
                        }
                    }
                    if ((String)jTemplateL["label"] == "แบบรายการประเมินผู้ขาย")
                    {
                        if (jData != null)
                        {
                            if (sTemplateToFM.DocumentCode == "FM-02")
                            {
                                jData["value"] = "แบบรายการประเมินผู้ขาย";
                            }
                            else if (sTemplateToFM.DocumentCode == "FM-03")
                            {
                                jData["value"] = "แบบรายการประเมินผู้เช่า / ให้บริการ";
                            }
                        }
                    }
                    if (jLayoutArray.Count > 1)
                    {
                        JObject jTemplateR = (JObject)jLayoutArray[1]["template"];
                        JObject jData2 = (JObject)jLayoutArray[1]["data"];
                        if ((String)jTemplateR["label"] == "หน่วยงาน")
                        {
                            if (jData2 != null)
                            {
                                jData2["value"] = emp.DepartmentNameTh;
                            }
                        }
                        if ((String)jTemplateR["label"] == "ชื่อสัญญา")
                        {
                            if (jData2 != null)
                            {
                                if (!string.IsNullOrEmpty(NameContract))
                                {
                                    jData2["value"] = NameContract;
                                }
                            }
                        }
                        if ((String)jTemplateR["label"] == "ประเภทผู้ค้า")
                        {
                            if (jData2 != null)
                            {
                                if (sTemplateToFM.DocumentCode == "FM-02")
                                {
                                    jData["value"] = "ผู้ขายสินค้า";
                                }
                                else if (sTemplateToFM.DocumentCode == "FM-03")
                                {
                                    jData["value"] = "ผู้เช่า / ให้บริการ";
                                }
                            }
                        }
                    }
                }
            }
            return JsonConvert.SerializeObject(jsonAdvanceForm);
        }
        public static void InsertTrnMemo(string MAdvancveform, DbwolfDataContext db, MSTTemplate Template, string Company, ViewEmployee emp)
        {
            List<MSTCompany> lstCompany = db.MSTCompanies.ToList();
            string guids = Guid.NewGuid().ToString().Replace("-", "");
            string company = ConfigurationManager.AppSettings["company"];
            string Status = ConfigurationManager.AppSettings["Status"];

            TRNMemo objMemo = new TRNMemo();
            objMemo.StatusName = Status;
            objMemo.PersonWaitingId = emp.EmployeeId;
            objMemo.PersonWaiting = emp.DivisionNameTh;
            objMemo.CreatedDate = DateTime.Now;
            objMemo.CreatedBy = emp.NameEn;
            objMemo.CreatorId = emp.EmployeeId;
            objMemo.CNameTh = emp.NameTh;
            objMemo.CNameEn = emp.NameEn;
            objMemo.CPositionId = emp.PositionId;
            objMemo.CPositionTh = emp.PositionNameTh;
            objMemo.CPositionEn = emp.PositionNameEn;
            objMemo.CDepartmentId = emp.DepartmentId;
            objMemo.CDepartmentTh = emp.DepartmentNameTh;
            objMemo.CDepartmentEn = emp.DepartmentNameEn;

            objMemo.RequesterId = emp.EmployeeId;
            objMemo.RNameTh = emp.NameTh;
            objMemo.RNameEn = emp.NameEn;
            objMemo.RPositionId = emp.PositionId;
            objMemo.RPositionTh = emp.PositionNameTh;
            objMemo.RPositionEn = emp.PositionNameEn;
            objMemo.RDepartmentId = emp.DepartmentId;
            objMemo.RDepartmentTh = emp.DepartmentNameTh;
            objMemo.RDepartmentEn = emp.DepartmentNameEn;

            objMemo.ModifiedDate = DateTime.Now;
            objMemo.ModifiedBy = objMemo.ModifiedBy;
            objMemo.TemplateId = Template.TemplateId;
            objMemo.TemplateName = Template.TemplateName;
            objMemo.GroupTemplateName = Template.GroupTemplateName;
            objMemo.RequestDate = DateTime.Now;
            var CurrentCom = lstCompany.Find(a => a.CompanyId == Convert.ToInt32(company));
            objMemo.CompanyId = CurrentCom.CompanyId;
            objMemo.CompanyName = CurrentCom.NameTh;

            objMemo.MAdvancveForm = MAdvancveform;
            objMemo.TAdvanceForm = MAdvancveform;
            objMemo.MemoSubject = Template.TemplateName + " : " + Company;
            objMemo.TemplateSubject = Template.TemplateName + " : " + Company;
            objMemo.TemplateDetail = guids;
            objMemo.ProjectID = 0;
            objMemo.DocumentCode = GenControlRunning(emp, Template.DocumentCode, objMemo, db);
            objMemo.DocumentNo = objMemo.DocumentCode;
            db.TRNMemos.InsertOnSubmit(objMemo);
            db.SubmitChanges();
            Console.WriteLine("GenerateTrnMemo success : Memoid >> " + objMemo.MemoId);
            log.Info("GenerateTrnMemo success : Memoid >> " + objMemo.MemoId);

            if (emp != null)
            {
                int sequence = 1;
                TRNLineApprove approve = new TRNLineApprove();
                approve.LineApproveId = 0;
                approve.MemoId = objMemo.MemoId;
                approve.Seq = sequence;
                approve.EmployeeId = emp.EmployeeId;
                approve.EmployeeCode = emp.EmployeeCode;
                approve.NameTh = emp.NameTh;
                approve.NameEn = emp.NameEn;
                approve.PositionTH = emp.PositionNameTh;
                approve.PositionEN = emp.PositionNameEn;
                approve.SignatureId = 2019;
                approve.SignatureTh = "อนุมัติ";
                approve.SignatureEn = "Approved";
                approve.IsActive = true;
                db.TRNLineApproves.InsertOnSubmit(approve);
                db.SubmitChanges();
                sequence++;
                log.Info("InsertTRNLineapprove success");
                log.Info("--------------------------");
            }
            else
            {
                log.Info("InsertTRNLineapprove Not have Employee " + emp.NameTh + " Memoid >> " + objMemo.MemoId);
                log.Info("--------------------------");
            }
        }
        public static string GenControlRunning(ViewEmployee Emp, string DocumentCode, TRNMemo objTRNMemo, DbwolfDataContext db)
        {
            string TempCode = DocumentCode;
            String sPrefixDocNo = $"{TempCode}-{DateTime.Now.Year.ToString()}-";
            int iRunning = 1;
            List<TRNMemo> temp = db.TRNMemos.Where(a => a.DocumentNo.ToUpper().Contains(sPrefixDocNo.ToUpper())).ToList();
            if (temp.Count > 0)
            {
                String sLastDocumentNo = temp.OrderBy(a => a.DocumentNo).Last().DocumentNo;
                if (!String.IsNullOrEmpty(sLastDocumentNo))
                {
                    List<String> list_LastDocumentNo = sLastDocumentNo.Split('-').ToList();

                    if (list_LastDocumentNo.Count >= 3)
                    {
                        iRunning = checkDataIntIsNull(list_LastDocumentNo[list_LastDocumentNo.Count - 1]) + 1;
                    }
                }
            }

            String sDocumentNo = $"{sPrefixDocNo}{iRunning.ToString().PadLeft(6, '0')}";
            return sDocumentNo;
        }
        public static string DocNoGenerate(string FixDoc, string DocCode, string CCode, string DCode, string DSCode, DbwolfDataContext db)
        {
            string sDocumentNo = "";
            int iRunning;
            if (!string.IsNullOrWhiteSpace(FixDoc))
            {
                string y4 = DateTime.Now.ToString("yyyy");
                string y2 = DateTime.Now.ToString("yy");
                string CompanyCode = CCode;
                string DepartmentCode = DCode;
                string DivisionCode = DSCode;
                string FixCode = FixDoc;
                FixCode = FixCode.Replace("[CompanyCode]", CompanyCode);
                FixCode = FixCode.Replace("[DepartmentCode]", DepartmentCode);
                FixCode = FixCode.Replace("[DocumentCode]", DocCode);
                FixCode = FixCode.Replace("[DivisionCode]", DivisionCode);

                FixCode = FixCode.Replace("[YYYY]", y4);
                FixCode = FixCode.Replace("[YY]", y2);
                sDocumentNo = FixCode;
                List<TRNMemo> tempfixDoc = db.TRNMemos.Where(a => a.DocumentNo.ToUpper().Contains(sDocumentNo.ToUpper())).ToList();


                List<TRNMemo> tempfixDocByYear = db.TRNMemos.ToList();

                tempfixDocByYear = tempfixDocByYear.FindAll(a => a.DocumentNo != ("Auto Generate") & Convert.ToDateTime(a.RequestDate).Year.ToString().Equals(y4)).ToList();
                if (tempfixDocByYear.Count > 0)
                {
                    tempfixDocByYear = tempfixDocByYear.OrderByDescending(a => a.MemoId).ToList();

                    String sLastDocumentNofix = tempfixDocByYear.First().DocumentNo;
                    if (!String.IsNullOrEmpty(sLastDocumentNofix))
                    {
                        List<String> list_LastDocumentNofix = sLastDocumentNofix.Split('-').ToList();
                        //Arm Edit 2020-05-18 Bug If Prefix have '-' will no running because list_LastDocumentNo.Count > 3

                        if (list_LastDocumentNofix.Count >= 3)
                        {
                            iRunning = checkDataIntIsNull(list_LastDocumentNofix[list_LastDocumentNofix.Count - 1]) + 1;
                            sDocumentNo = $"{sDocumentNo}-{iRunning.ToString().PadLeft(6, '0')}";
                        }
                    }
                }
                else
                {
                    sDocumentNo = $"{sDocumentNo}-{1.ToString().PadLeft(6, '0')}";

                }
            }
            return sDocumentNo;

        }
        public static int checkDataIntIsNull(object Input)
        {
            int Results = 0;
            if (Input != null)
                int.TryParse(Input.ToString().Replace(",", ""), out Results);

            return Results;
        }
    }
}
