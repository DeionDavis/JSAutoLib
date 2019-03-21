using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Configuration;
using Microsoft.VisualStudio.TestTools.UITesting;
using System.Threading;
using AutoItX3Lib;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using System.Drawing.Imaging;
using System.Drawing;
using CommonLibrary.Exceptions;
using CommonLibrary.DataDrivenTesting;
using CommonLibrary.KeywordDrivenTesting;
using CommonLibrary.Operations;
using CommonLibrary.Reports;
using CommonLibrary.Writedata;
using System.Diagnostics;
using OperationLibrary;
using Microsoft.VisualStudio.TestTools.UITest.Extension;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using System.Globalization;
using CommonLibrary.CommonLanguageReader;
using System.Text.RegularExpressions;
using System.Linq.Expressions;
using System.Windows.Forms;

namespace CommonLibrary.Log
{
    public class LoginOperatrion
    {
        private string ConfigPath = string.Empty;
        private string Url = string.Empty;
        string path = string.Empty;
        string url = string.Empty;
        public static int len = 0, min = 0, mid = 0, max = 0;
        DataTable dt = new DataTable();
        public static string ProjectName;
        public static string LogPath = string.Empty;
        public static string SecurityCode = string.Empty;
        public static string WareHouseNo = string.Empty;
        public static string DetaildReportStatus = string.Empty;
        public static string batchforReport = string.Empty;
        public static string requiredPlant = string.Empty;
        public string UsertypeLogin = string.Empty;
        public int numberOfIteration = 0;
        public string Init = string.Empty;
        private static string _numbers = "0123456789";
        Random random = new Random();
        AutoItX3 auto = new AutoItX3();
        LogLanguageTemplete lang = new LogLanguageTemplete();
        PerformOperation Pop = new PerformOperation();
        ReportGeneration genreport = new ReportGeneration();
        WriteAndReadData datawrite = new WriteAndReadData();
        CommonLanguageTemplateReader languageResource = new CommonLanguageTemplateReader();


        #region InititalOperations
        /// <summary>
        /// This will  launch the browser with specified URL.
        /// </summary>
        /// <param name="project"></param>
        public void Initialize(string project)
        {
            ProjectName = project;
            BrowserWindow window = new BrowserWindow();
            window.SearchProperties[BrowserWindow.PropertyNames.ClassName] = "IEFrame";
            UITestControlCollection wndcollection = window.FindMatchingControls();
            foreach (UITestControl control in wndcollection)
            {
                if (control is BrowserWindow)
                {
                    ((BrowserWindow)control).Close();
                }
            }
            //Reading the data from the excel file from the path given in the project.
            string[] Allkeys = ConfigurationManager.AppSettings.AllKeys;
            if (Allkeys.Contains<string>("LoginDetails"))
            {
                path = ConfigurationManager.AppSettings["LoginDetails"];
            }
            else
            {
                path = ConfigurationManager.AppSettings["PlantLoginDetails"];
            }
            Init = ConfigurationManager.AppSettings["Initialization"];
            ExcelDataTable.PopulateBatchWiseData(path);
            ExcelDataTable.PopulateInCollection(Init + "\\GlobalSettings.xlsx");
            batchforReport = ConfigurationManager.AppSettings["Batch"];
            len = Convert.ToInt16(ExcelDataTable.ReadBatchData(batchforReport, "Execution"));
            DetaildReportStatus = ExcelDataTable.ReadData(1, "DetailedReportGeneration");
            for (int i = 0; i <= len; i++)
            {
                url = ExcelDataTable.ReadData(1, "Url");
            }
            BrowserWindow.Launch(url);
        }
        #endregion

        #region Login & Logout
        /// <summary>
        /// This will help to login to super admin.
        /// </summary>
        public void SuperLogin()
        {
            min = Convert.ToInt16(ExcelDataTable.ReadData(1, "Minimum"));
            mid = Convert.ToInt16(ExcelDataTable.ReadData(1, "Medium"));
            max = Convert.ToInt16(ExcelDataTable.ReadData(1, "Maximum"));
            LogLanguageTemplete.messageResource(ExcelDataTable.ReadBatchData(batchforReport, "languageCode"));
            requiredPlant = ExcelDataTable.ReadBatchData(batchforReport, "Plant");
            LogPath = ConfigurationManager.AppSettings["LogOperation"];
            Thread.Sleep(max);
            dt.Clear();

            dt = ExcelKeywordTable.ExcelData(LogPath + "\\KeywordDrivenData.xlsx", "SuperAdminLogin");

            Thread.Sleep(max);
            BrowserWindow window = new BrowserWindow();
            window.Maximized = true;
            Thread.Sleep(max);
            Pop.OperationStart(dt.Rows[0]["Keyword"].ToString(), dt.Rows[0]["TypeOfControl"].ToString(), dt.Rows[0]["ControlKeyword"].ToString(), dt.Rows[0]["ControlKeywordValue"].ToString(), ExcelDataTable.ReadBatchData(batchforReport, dt.Rows[0]["DataRefferencekeyword"].ToString()), dt.Rows[0]["TypeOfWindow"].ToString(), dt.Rows[0]["Step No"].ToString(), dt.Rows[0]["Description"].ToString());
            Pop.OperationStart(dt.Rows[1]["Keyword"].ToString(), dt.Rows[1]["TypeOfControl"].ToString(), dt.Rows[1]["ControlKeyword"].ToString(), dt.Rows[1]["ControlKeywordValue"].ToString(), ExcelDataTable.ReadBatchData(batchforReport, dt.Rows[1]["DataRefferencekeyword"].ToString()), dt.Rows[1]["TypeOfWindow"].ToString(), dt.Rows[1]["Step No"].ToString(), dt.Rows[1]["Description"].ToString());
            Pop.OperationStart(dt.Rows[2]["Keyword"].ToString(), dt.Rows[2]["TypeOfControl"].ToString(), dt.Rows[2]["ControlKeyword"].ToString(), dt.Rows[2]["ControlKeywordValue"].ToString(), dt.Rows[2]["DataRefferencekeyword"].ToString(), dt.Rows[2]["TypeOfWindow"].ToString(), dt.Rows[2]["Step No"].ToString(), dt.Rows[2]["Description"].ToString());
            Thread.Sleep(max * 2);
            string Login = window.ExecuteScript("var data = document.getElementById('ContentPlaceHolder_lblErrorMessage');return data;").ToString();
            if (Login != string.Empty)
            {
                string LoginValidation = window.ExecuteScript("var data = document.getElementById('ContentPlaceHolder_lblErrorMessage').innerHTML;return data;").ToString();
                if (LoginValidation != "")
                {
                    Thread.Sleep(max);
                    Pop.OperationStart(dt.Rows[1]["Keyword"].ToString(), dt.Rows[1]["TypeOfControl"].ToString(), dt.Rows[1]["ControlKeyword"].ToString(), dt.Rows[1]["ControlKeywordValue"].ToString(), ExcelDataTable.ReadBatchData(batchforReport, dt.Rows[1]["DataRefferencekeyword"].ToString()), dt.Rows[1]["TypeOfWindow"].ToString(), dt.Rows[1]["Step No"].ToString(), dt.Rows[1]["Description"].ToString());
                    Thread.Sleep(mid);
                    Pop.OperationStart(dt.Rows[2]["Keyword"].ToString(), dt.Rows[2]["TypeOfControl"].ToString(), dt.Rows[2]["ControlKeyword"].ToString(), dt.Rows[2]["ControlKeywordValue"].ToString(), dt.Rows[2]["DataRefferencekeyword"].ToString(), dt.Rows[2]["TypeOfWindow"].ToString(), dt.Rows[2]["Step No"].ToString(), dt.Rows[2]["Description"].ToString());
                    Thread.Sleep(max * 2);
                    Login = window.ExecuteScript("var data = document.getElementById('ContentPlaceHolder_lblErrorMessage');return data;").ToString();
                    if (Login != string.Empty)
                    {
                        LoginValidation = window.ExecuteScript("var data = document.getElementById('ContentPlaceHolder_lblErrorMessage').innerHTML;return data;").ToString();
                        if (LoginValidation != "")
                        {
                            string screenShotName = "LoginFailed";
                            Operation.ErrorScreenPath = screenShot(screenShotName);
                            Operation.FailerReason = "Login Failed";
                            genreport.FileCorreptionCheck();
                            Assert.Fail("Login Failed");
                        }
                        else
                        {
                            Thread.Sleep(max);
                            Pop.OperationStart(dt.Rows[3]["Keyword"].ToString(), dt.Rows[3]["TypeOfControl"].ToString(), dt.Rows[3]["ControlKeyword"].ToString(), dt.Rows[3]["ControlKeywordValue"].ToString(), dt.Rows[3]["DataRefferencekeyword"].ToString(), dt.Rows[3]["TypeOfWindow"].ToString(), dt.Rows[3]["Step No"].ToString(), dt.Rows[3]["Description"].ToString());
                            generateSecurityCode(ExcelDataTable.ReadBatchData(batchforReport, dt.Rows[0]["DataRefferencekeyword"].ToString()));
                            Pop.OperationStart("EnterText", dt.Rows[4]["TypeOfControl"].ToString(), dt.Rows[4]["ControlKeyword"].ToString(), dt.Rows[4]["ControlKeywordValue"].ToString(), SecurityCode, dt.Rows[4]["TypeOfWindow"].ToString(), dt.Rows[4]["Step No"].ToString(), dt.Rows[4]["Description"].ToString());
                            Pop.OperationStart(dt.Rows[5]["Keyword"].ToString(), dt.Rows[5]["TypeOfControl"].ToString(), dt.Rows[5]["ControlKeyword"].ToString(), dt.Rows[5]["ControlKeywordValue"].ToString(), dt.Rows[5]["DataRefferencekeyword"].ToString(), dt.Rows[5]["TypeOfWindow"].ToString(), dt.Rows[5]["Step No"].ToString(), dt.Rows[5]["Description"].ToString());
                            Thread.Sleep(max * 2);
                            string SecurityMsg = window.ExecuteScript("var data = document.getElementById('spnSecurityCodeErrorMsg');return data;").ToString();
                            if (SecurityMsg != string.Empty)
                            {
                                string SecurityMsgValidation = window.ExecuteScript("var data = document.getElementById('spnSecurityCodeErrorMsg').innerHTML;return data;").ToString();
                                if (SecurityMsgValidation == lang.Msg_WrongSecurityCode)
                                {
                                    string screenShotName = "WrongSecurityCodeGenerated";
                                    Operation.ErrorScreenPath = screenShot(screenShotName);
                                    Operation.FailerReason = "Wrong Security Code Generated, Login Failed";
                                    genreport.FileCorreptionCheck();
                                }
                                Assert.AreEqual("", SecurityMsgValidation, "Wrong Security Code Generated, Login Failed");
                            }
                            Pop.OperationStart(dt.Rows[6]["Keyword"].ToString(), dt.Rows[6]["TypeOfControl"].ToString(), dt.Rows[6]["ControlKeyword"].ToString(), dt.Rows[6]["ControlKeywordValue"].ToString(), dt.Rows[6]["DataRefferencekeyword"].ToString(), dt.Rows[6]["TypeOfWindow"].ToString(), dt.Rows[6]["Step No"].ToString(), dt.Rows[6]["Description"].ToString());
                        }
                    }
                }
                else
                {
                    Thread.Sleep(max);
                    Pop.OperationStart(dt.Rows[3]["Keyword"].ToString(), dt.Rows[3]["TypeOfControl"].ToString(), dt.Rows[3]["ControlKeyword"].ToString(), dt.Rows[3]["ControlKeywordValue"].ToString(), dt.Rows[3]["DataRefferencekeyword"].ToString(), dt.Rows[3]["TypeOfWindow"].ToString(), dt.Rows[3]["Step No"].ToString(), dt.Rows[3]["Description"].ToString());
                    generateSecurityCode(ExcelDataTable.ReadBatchData(batchforReport, dt.Rows[0]["DataRefferencekeyword"].ToString()));
                    Pop.OperationStart("EnterText", dt.Rows[4]["TypeOfControl"].ToString(), dt.Rows[4]["ControlKeyword"].ToString(), dt.Rows[4]["ControlKeywordValue"].ToString(), SecurityCode, dt.Rows[4]["TypeOfWindow"].ToString(), dt.Rows[4]["Step No"].ToString(), dt.Rows[4]["Description"].ToString());
                    Pop.OperationStart(dt.Rows[5]["Keyword"].ToString(), dt.Rows[5]["TypeOfControl"].ToString(), dt.Rows[5]["ControlKeyword"].ToString(), dt.Rows[5]["ControlKeywordValue"].ToString(), dt.Rows[5]["DataRefferencekeyword"].ToString(), dt.Rows[5]["TypeOfWindow"].ToString(), dt.Rows[5]["Step No"].ToString(), dt.Rows[5]["Description"].ToString());
                    Thread.Sleep(max * 2);
                    string SecurityMsg = window.ExecuteScript("var data = document.getElementById('spnSecurityCodeErrorMsg');return data;").ToString();
                    if (SecurityMsg != string.Empty)
                    {
                        string SecurityMsgValidation = window.ExecuteScript("var data = document.getElementById('spnSecurityCodeErrorMsg').innerHTML;return data;").ToString();
                        if (SecurityMsgValidation == lang.Msg_WrongSecurityCode)
                        {
                            string screenShotName = "WrongSecurityCodeGenerated";
                            Operation.ErrorScreenPath = screenShot(screenShotName);
                            Operation.FailerReason = "Wrong Security Code Generated, Login Failed";
                            genreport.FileCorreptionCheck();
                        }
                        Assert.AreEqual("", SecurityMsgValidation, "Wrong Security Code Generated, Login Failed");
                    }
                    Pop.OperationStart(dt.Rows[6]["Keyword"].ToString(), dt.Rows[6]["TypeOfControl"].ToString(), dt.Rows[6]["ControlKeyword"].ToString(), dt.Rows[6]["ControlKeywordValue"].ToString(), dt.Rows[6]["DataRefferencekeyword"].ToString(), dt.Rows[6]["TypeOfWindow"].ToString(), dt.Rows[6]["Step No"].ToString(), dt.Rows[6]["Description"].ToString());
                }
            }
        }
        /// <summary>
        /// This will perform the Login operations that specified in the Common keyword drivers. for login refer Sheet Name:"Logout" 
        /// </summary>
        public void Logout()
        {
            BrowserWindow window = new BrowserWindow();
            Thread.Sleep(mid);
            dt.Clear();
            dt = ExcelKeywordTable.ExcelData(LogPath + "\\KeywordDrivenData.xlsx", "Logout");
            Thread.Sleep(max);
            Pop.OperationStart(dt.Rows[0]["Keyword"].ToString(), dt.Rows[0]["TypeOfControl"].ToString(), dt.Rows[0]["ControlKeyword"].ToString(), dt.Rows[0]["ControlKeywordValue"].ToString(), dt.Rows[0]["DataRefferencekeyword"].ToString(), dt.Rows[0]["TypeOfWindow"].ToString(), dt.Rows[0]["Step No"].ToString(), dt.Rows[0]["Description"].ToString());
            Thread.Sleep(max);
            try
            {
                Pop.OperationStart(dt.Rows[1]["Keyword"].ToString(), dt.Rows[1]["TypeOfControl"].ToString(), dt.Rows[1]["ControlKeyword"].ToString(), dt.Rows[1]["ControlKeywordValue"].ToString(), dt.Rows[1]["DataRefferencekeyword"].ToString(), dt.Rows[1]["TypeOfWindow"].ToString(), dt.Rows[1]["Step No"].ToString(), dt.Rows[1]["Description"].ToString());
            }
            catch (Exception e) { }
            Thread.Sleep(mid * 2);

            Pop.OperationStart(dt.Rows[2]["Keyword"].ToString(), dt.Rows[2]["TypeOfControl"].ToString(), dt.Rows[2]["ControlKeyword"].ToString(), dt.Rows[2]["ControlKeywordValue"].ToString(), dt.Rows[2]["DataRefferencekeyword"].ToString(), dt.Rows[2]["TypeOfWindow"].ToString(), dt.Rows[2]["Step No"].ToString(), dt.Rows[2]["Description"].ToString());

            string AddingSuccess = Pop.WebGetControlData(dt.Rows[3]["TypeOfControl"].ToString(), dt.Rows[3]["ControlKeyword"].ToString(), dt.Rows[3]["ControlKeywordValue"].ToString(), "LoginPageNotOpen", "Logout Failed", dt.Rows[3]["Step No"].ToString(), dt.Rows[3]["Description"].ToString(), dt.Rows[3]["Keyword"].ToString()).Trim();
            if (AddingSuccess != null)
            {
                string[] msg = AddingSuccess.Split('.');
                AddingSuccess = msg[0];
            }
            if (AddingSuccess != lang.Msg_LogoutSuccessMessage)
            {
                string screenShotName = "LoginPageNotOpen";
                Operation.ErrorScreenPath = screenShot(screenShotName);
                Operation.FailerReason = "Logout Failed";
                genreport.FileCorreptionCheck();
            }
            Assert.AreEqual(lang.Msg_LogoutSuccessMessage, AddingSuccess, "Logout Failed");

            Thread.Sleep(mid);
            Pop.OperationStart(dt.Rows[4]["Keyword"].ToString(), dt.Rows[4]["TypeOfControl"].ToString(), dt.Rows[4]["ControlKeyword"].ToString(), dt.Rows[4]["ControlKeywordValue"].ToString(), dt.Rows[4]["DataRefferencekeyword"].ToString(), dt.Rows[4]["TypeOfWindow"].ToString(), dt.Rows[4]["Step No"].ToString(), dt.Rows[4]["Description"].ToString());
            genreport.FileCorreptionCheck();
        }

        public void generateSecurityCode(string username)
        {
            SecurityCode = DateTime.Today.ToString("MMM") + DateTime.Today.Day.ToString("00") + username.Substring(0, 3);
        }
        #endregion

        #region Plant wise Functions
        /// <summary>
        /// This will perform the Login operations based on Plant this is used in Plant Specific Batch only that specified in the Common keyword drivers. for login refer Sheet Name:"Login" 
        /// </summary>
        /// <param name="Usertype">Regional User or Plant User</param>
        /// <param name="Iteration">Number of execution of the test case</param>
        public void PlantLogin(string Usertype, int Iteration)
        {
            min = Convert.ToInt16(ExcelDataTable.ReadData(1, "Minimum"));
            mid = Convert.ToInt16(ExcelDataTable.ReadData(1, "Medium"));
            max = Convert.ToInt16(ExcelDataTable.ReadData(1, "Maximum"));
            string PlantPath = ConfigurationManager.AppSettings["PlantLoginDetails"];
            ExcelDataTable.PopulatePlantWiseData(PlantPath);
            numberOfIteration = Iteration;
            LogLanguageTemplete.messageResource(ExcelDataTable.ReadBatchData(batchforReport, "languageCode"));
            if (Usertype == "PlantLogin")
            {
                requiredPlant = ExcelDataTable.ReadBatchData(batchforReport, "Plant_" + Iteration);
            }
            else if (Usertype == "PlantRegionalLogin")
            {
                requiredPlant = ExcelDataTable.ReadBatchData(batchforReport, "Plant Name");
            }          
            //Reading the excel file for common operations like Login, Logout, Manage language, Navigation.
            LogPath = ConfigurationManager.AppSettings["LogOperation"];
            Thread.Sleep(max);
            UsertypeLogin = Usertype;
            dt.Clear();
            if (Usertype == "PlantLogin")
            {
                dt = ExcelKeywordTable.ExcelData(LogPath + "\\KeywordDrivenData.xlsx", "Login");
            }
            else if (Usertype == "PlantRegionalLogin")
            {
                dt = ExcelKeywordTable.ExcelData(LogPath + "\\KeywordDrivenData.xlsx", "RegionalLogin");
            }
            BrowserWindow window = new BrowserWindow();
            window.Maximized = true;
            Thread.Sleep(max);
            Pop.OperationStart(dt.Rows[0]["Keyword"].ToString(), dt.Rows[0]["TypeOfControl"].ToString(), dt.Rows[0]["ControlKeyword"].ToString(), dt.Rows[0]["ControlKeywordValue"].ToString(), ExcelDataTable.ReadPlantLoginInfo(requiredPlant, dt.Rows[0]["DataRefferencekeyword"].ToString()), dt.Rows[0]["TypeOfWindow"].ToString(), dt.Rows[0]["Step No"].ToString(), dt.Rows[0]["Description"].ToString());
            Pop.OperationStart(dt.Rows[1]["Keyword"].ToString(), dt.Rows[1]["TypeOfControl"].ToString(), dt.Rows[1]["ControlKeyword"].ToString(), dt.Rows[1]["ControlKeywordValue"].ToString(), ExcelDataTable.ReadPlantLoginInfo(requiredPlant, dt.Rows[1]["DataRefferencekeyword"].ToString()), dt.Rows[1]["TypeOfWindow"].ToString(), dt.Rows[1]["Step No"].ToString(), dt.Rows[1]["Description"].ToString());
            Pop.OperationStart(dt.Rows[2]["Keyword"].ToString(), dt.Rows[2]["TypeOfControl"].ToString(), dt.Rows[2]["ControlKeyword"].ToString(), dt.Rows[2]["ControlKeywordValue"].ToString(), dt.Rows[2]["DataRefferencekeyword"].ToString(), dt.Rows[2]["TypeOfWindow"].ToString(), dt.Rows[2]["Step No"].ToString(), dt.Rows[2]["Description"].ToString());
            Thread.Sleep(max * 2);
            string Login = window.ExecuteScript("var data = document.getElementById('ContentPlaceHolder_lblErrorMessage');return data;").ToString();
            if (Login != string.Empty)
            {
                string LoginValidation = window.ExecuteScript("var data = document.getElementById('ContentPlaceHolder_lblErrorMessage').innerHTML;return data;").ToString();
                if (LoginValidation != "")
                {
                    Thread.Sleep(max);
                    Pop.OperationStart(dt.Rows[1]["Keyword"].ToString(), dt.Rows[1]["TypeOfControl"].ToString(), dt.Rows[1]["ControlKeyword"].ToString(), dt.Rows[1]["ControlKeywordValue"].ToString(), ExcelDataTable.ReadPlantLoginInfo(requiredPlant, dt.Rows[1]["DataRefferencekeyword"].ToString()), dt.Rows[1]["TypeOfWindow"].ToString(), dt.Rows[1]["Step No"].ToString(), dt.Rows[1]["Description"].ToString());
                    Thread.Sleep(mid);
                    Pop.OperationStart(dt.Rows[2]["Keyword"].ToString(), dt.Rows[2]["TypeOfControl"].ToString(), dt.Rows[2]["ControlKeyword"].ToString(), dt.Rows[2]["ControlKeywordValue"].ToString(), dt.Rows[2]["DataRefferencekeyword"].ToString(), dt.Rows[2]["TypeOfWindow"].ToString(), dt.Rows[2]["Step No"].ToString(), dt.Rows[2]["Description"].ToString());
                    Thread.Sleep(max * 2);
                    Login = window.ExecuteScript("var data = document.getElementById('ContentPlaceHolder_lblErrorMessage');return data;").ToString();
                    if (Login != string.Empty)
                    {
                        LoginValidation = window.ExecuteScript("var data = document.getElementById('ContentPlaceHolder_lblErrorMessage').innerHTML;return data;").ToString();
                        if (LoginValidation != lang.Msg_LogoutSuccessMessage)
                        {
                            string screenShotName = "Plant_LoginFailed";
                            Operation.ErrorScreenPath = screenShot(screenShotName);
                            Operation.FailerReason = requiredPlant + " Plant Login Failed";
                            genreport.FileCorreptionCheck();
                        }
                        Assert.AreEqual("", LoginValidation, requiredPlant + " Plant Login Failed");
                    }
                    else
                    {
                        Thread.Sleep(max);
                        auto.Send("{F5}");
                        Pop.OperationStart(dt.Rows[3]["Keyword"].ToString(), dt.Rows[3]["TypeOfControl"].ToString(), dt.Rows[3]["ControlKeyword"].ToString(), dt.Rows[3]["ControlKeywordValue"].ToString(), dt.Rows[3]["DataRefferencekeyword"].ToString(), dt.Rows[3]["TypeOfWindow"].ToString(), dt.Rows[3]["Step No"].ToString(), dt.Rows[3]["Description"].ToString());
                    }
                }
            }
            else
            {
                Thread.Sleep(max);
                auto.Send("{F5}");
                Pop.OperationStart(dt.Rows[3]["Keyword"].ToString(), dt.Rows[3]["TypeOfControl"].ToString(), dt.Rows[3]["ControlKeyword"].ToString(), dt.Rows[3]["ControlKeywordValue"].ToString(), dt.Rows[3]["DataRefferencekeyword"].ToString(), dt.Rows[3]["TypeOfWindow"].ToString(), dt.Rows[3]["Step No"].ToString(), dt.Rows[3]["Description"].ToString());
            }
        }
        #endregion

        #region LinkNavigations
        /// <summary>
        /// This will refer the links specified in the excel sheet "LinksData" in the Common keyword drivers excel.
        /// </summary>
        /// <param name="keyword">This will accept the reference keyword to the link specified in the LinksData excel sheet.</param>
        public void ModuleNavigation(string keyword)
        {
            ExcelKeywordTable.getKewordData(LogPath + "\\KeywordDrivenData.xlsx", "ModuleData");
            Thread.Sleep(mid);
            Pop.OperationStart(ExcelKeywordTable.ReadKeywordData(keyword, "Keyword"), ExcelKeywordTable.ReadKeywordData(keyword, "TypeOfControl"), ExcelKeywordTable.ReadKeywordData(keyword, "ControlKeyword"), ExcelKeywordTable.ReadKeywordData(keyword, "ControlKeywordValue"), ExcelKeywordTable.ReadKeywordData(keyword, "DataRefferencekeyword"), ExcelKeywordTable.ReadKeywordData(keyword, "TypeOfWindow"), ExcelKeywordTable.ReadKeywordData(keyword, "Step No"), ExcelKeywordTable.ReadKeywordData(keyword, "Description"));
            Thread.Sleep(mid);
        }

        /// <summary>
        /// This will refer the links specified in the excel sheet "ModuleData" in the Common keyword drivers excel.
        /// </summary>
        /// <param name="keyword">This will accept the reference keyword to the link specified in the ModuleData excel sheet.</param>
        public void SubModuleNavigation(string keyword)
        {
            ExcelKeywordTable.getKewordData(LogPath + "\\KeywordDrivenData.xlsx", "SubModuleData");
            auto.Send("{DOWN 2}");
            Thread.Sleep(mid);
            Pop.OperationStart(ExcelKeywordTable.ReadKeywordData(keyword, "Keyword"), ExcelKeywordTable.ReadKeywordData(keyword, "TypeOfControl"), ExcelKeywordTable.ReadKeywordData(keyword, "ControlKeyword"), ExcelKeywordTable.ReadKeywordData(keyword, "ControlKeywordValue"), ExcelKeywordTable.ReadKeywordData(keyword, "DataRefferencekeyword"), ExcelKeywordTable.ReadKeywordData(keyword, "TypeOfWindow"), ExcelKeywordTable.ReadKeywordData(keyword, "Step No"), ExcelKeywordTable.ReadKeywordData(keyword, "Description"));
            Thread.Sleep(mid);
        }

        /// <summary>
        /// This will refer the links specified in the excel sheet "SubModuleData" in the Common keyword drivers excel.
        /// </summary>
        /// <param name="keyword">This will accept the reference keyword to the link specified in the SubModuleData excel sheet.</param>
        public void LinkNavigation(string keyword)
        {
            ExcelKeywordTable.getKewordData(LogPath + "\\KeywordDrivenData.xlsx", "LinksData");
            auto.Send("{DOWN 2}");
            Thread.Sleep(mid);
            string[] IdData = keyword.Split('+');
            if (IdData.Count() > 1)
            {
                if (IdData[1].StartsWith("Rec_"))
                {
                    keyword = IdData[0].Replace("&", Operation.recordedData[IdData[1].Replace("Rec_", string.Empty)]);
                }
                if (IdData[1].Contains(":"))
                {
                    WriteAndReadData.DataFilePath = ConfigurationManager.AppSettings["ExcelDataFile"] + "\\DynamicData" + "\\" + Operation.Batch + ".xls";
                    ExcelDataTable.PopulateRecordData(WriteAndReadData.DataFilePath);
                    keyword = IdData[0].Replace("&", Pop.getDataFromDynamicExcel(IdData[1]));
                }
                else
                {
                    keyword = IdData[0].Replace("&", ExcelDataTable.ReadData(1, IdData[1]));
                }
            }

            Pop.OperationStart(ExcelKeywordTable.ReadKeywordData(keyword, "Keyword"), ExcelKeywordTable.ReadKeywordData(keyword, "TypeOfControl"), ExcelKeywordTable.ReadKeywordData(keyword, "ControlKeyword"), ExcelKeywordTable.ReadKeywordData(keyword, "ControlKeywordValue"), ExcelKeywordTable.ReadKeywordData(keyword, "DataRefferencekeyword"), ExcelKeywordTable.ReadKeywordData(keyword, "TypeOfWindow"), ExcelKeywordTable.ReadKeywordData(keyword, "Step No"), ExcelKeywordTable.ReadKeywordData(keyword, "Description"));
            Thread.Sleep(mid);
        }
        #endregion

        #region validateLog
        /// <summary>
        /// This will validate the log file downloaded after uploading any data
        /// </summary>
        /// <param name="validMsg">Here you want to specify the message to validate.</param>
        /// <returns>This will return the true or false value.</returns>
        public bool validateLOGcheckFail(string validMsg)
        {
            string fileName = string.Empty;
            try
            {
                Thread.Sleep(max * 3);
                Process[] notepads = Process.GetProcessesByName("notepad");
                Thread.Sleep(max * 3);
                fileName = "\\" + notepads[0].MainWindowTitle;
                //Using below code we can access special folders in windows like documents, favorites, or library folders 
                fileName = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + fileName;
                //replacing text from the path that fetched 
                fileName = fileName.Replace(" - Notepad", ".txt");
                fileName = fileName.Replace("Documents", "Downloads");
                FileStream file = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.Read);
                bool b1 = false;
                string[] lines = File.ReadAllLines(fileName);
                string[] Validlines = validMsg.Split(',');
                for (int i = 1; i < lines.Length; i++)
                {
                    if (Validlines.Length == 2)
                    {
                        if ((lines[i - 1].Equals(Validlines[0]) || lines[i].Equals(Validlines[1])))
                        {
                            b1 = true;
                        }
                    }
                    else if (Validlines.Length == 1)
                    {
                        if ((lines[i - 1].Equals(Validlines[0])))
                        {
                            b1 = true;
                        }
                        else if ((lines[i - 1].Equals(Validlines[0].Replace('0', '1'))))
                        {
                            b1 = false;
                            break;
                        }
                        else if ((lines[i - 1].Equals(Validlines[0].Replace('1', '0'))))
                        {
                            b1 = false;
                            break;
                        }
                        else if ((lines[i - 1].Contains(Validlines[0])))
                        {
                            b1 = true;
                        }
                    }
                }
                file.Close();
                return b1;
            }
            catch (FileNotFoundException e)
            {
                Operation.FailerReason = fileName + " not found in the specified location";
                Assert.Fail(fileName + " not found in the specified location");
                return false;
            }
        }
        #endregion

        #region PageNavigation
        /// <summary>
        /// This will navigate through the pages, and validate data is displaying or not.
        /// </summary>
        /// <param name="PageID">it contains table id and Page Navigation Controller ID separated by ':' , Ex: TableId:PageNavigataionControlerID</param>
        /// <param name="step">Step number</param>
        /// <param name="description">Description for the step number</param>
        /// <param name="AssertMsg">Assertion Message or error message you want display</param>
        /// <param name="AssertScreenshotName">Name of error screen shot image</param>
        public void NavigationPage(string PageID, string step, string description, string AssertScreenshotName, string AssertMsg)
        {
            BrowserWindow window = new BrowserWindow();
            string lastPage = string.Empty;
            int count = 0;
            int tablerow = 0;
            int tableCol = 0;
            string tbody = PageID.Split(':')[0];
            string divID = PageID.Split(':')[1];
            try
            {
                count = Convert.ToInt16(window.ExecuteScript("var data=$('#" + divID + " a').length;return data"));
            }
            catch (Exception e) { }
            if (count == 0)
            {
                tablerow = Convert.ToInt16(window.ExecuteScript("var tblRow=$('#" + tbody + " tr').length;return tblRow", "javascript"));
                tableCol = Convert.ToInt16(window.ExecuteScript("var tblCol=$('#" + tbody + " td').length;return tblCol", "javascript"));
                if (tablerow >= 1 && tableCol > 1)
                {
                    genreport.Reports(step, "Data Loaded in page successfully", "ValidatingPageData", true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                }
                else
                {
                    TakeScreenshot(step, AssertScreenshotName, AssertMsg, description);
                }
            }
            //  string idval = PageID.Split('_')[0];

            string idval = "ctl01";
            if (count >= 3)
            {
                if (count > 5)
                {
                    lastPage = idval + "_hrefLast";
                    count = 3;
                }
                else
                {
                    count = 3;
                }
            }
            for (int i = 0; i < count; i++)
            {

                window.ExecuteScript("document.getElementById('" + divID + "').getElementsByTagName('a')[" + i + "].scrollIntoView(true); ");
                Thread.Sleep(2500);
                window.ExecuteScript("$('#" + divID + " a')[" + i + "].click();", "javascript");
                Pop.OperationStart("Wait", "HtmlTable", "ID", tbody, "", "Web", step, description);
                tablerow = Convert.ToInt16(window.ExecuteScript("var tblRow=$('#" + tbody + " tr').length;return tblRow", "javascript"));
                tableCol = Convert.ToInt16(window.ExecuteScript("var tblCol=$('#" + tbody + " td').length;return tblCol", "javascript"));
                if (tablerow >= 1 && tableCol > 1)
                {
                    genreport.Reports(step, "Data Loaded in page " + i + " successfully", "ValidatingPageData", true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                }
                else
                {
                    TakeScreenshot(step, AssertScreenshotName, AssertMsg + i + " Failed", description);
                }
            }
            if (lastPage != string.Empty)
            {
                window.ExecuteScript("document.getElementById('" + lastPage + "').scrollIntoView(true);");
                // Pop.OperationStart("Click", "HtmlHyperlink", "ID", lastPage, "", "Web", step, description + " Last Page");
                window.ExecuteScript("$('#" + lastPage + "').click();", "javascript");
                Pop.OperationStart("Wait", "HtmlTable", "ID", tbody, "", "Web", step, description);
                Thread.Sleep(max);
                int tablerowL = Convert.ToInt16(window.ExecuteScript("var tblRow=$('#" + tbody + " tr').length;return tblRow", "javascript"));
                int tableColL = Convert.ToInt16(window.ExecuteScript("var tblCol=$('#" + tbody + " td').length;return tblCol", "javascript"));
                if (tablerowL >= 1 && tableColL > 1)
                {
                    genreport.Reports(step, "Data Loaded in Last page successfully", "ValidatingPageData", true, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                }
                else
                {
                    TakeScreenshot(step, AssertScreenshotName, AssertMsg, description);
                }
            }

        }
        #endregion

        #region BaseFunctions

        /// <summary>
        /// This will close the opened file.
        /// </summary>
        public void CloseFile()
        { 
            Thread.Sleep(2000);
            auto.Send("!{SPACE}");
            Thread.Sleep(min);
            auto.Send("{C}");
            Thread.Sleep(max);
        }

        /// <summary>
        /// This will take the screen shot and saved in the specified path.
        /// </summary>
        /// <param name="imgName">Name of the image will pass here.</param>
        /// <returns>This will return the path of the image file.</returns>
        public string screenShot(string imgName)
        {
            //Read the path from the App.config file to save the screen short.
            string path = ConfigurationManager.AppSettings["ScreenShot"];
            string Imagepath = string.Empty;
            path = path + @"\" + LoginOperatrion.ProjectName + "";
            Directory.CreateDirectory(path);
            try
            {

                Image image = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);
                using (Graphics g = Graphics.FromImage(image))
                {
                    g.CopyFromScreen(0, 0, 0, 0, Screen.PrimaryScreen.Bounds.Size);
                    if (requiredPlant != string.Empty)
                    {
                        image.Save(path + "\\" + imgName + "_" + requiredPlant + ".jpeg", ImageFormat.Jpeg);
                        Imagepath = path + "\\" + imgName + "_" + requiredPlant + ".jpeg";
                        image.Dispose();
                    }
                    else
                    {
                        image.Save(path + "\\" + imgName + ".jpeg", ImageFormat.Jpeg);
                        Imagepath = path + "\\" + imgName + ".jpeg";
                        image.Dispose();
                    }
                }
            }
            catch (Exception e) { }
            Console.WriteLine("Screen Shot is Available in " + path + "\t" + "Folder");
            return Imagepath;
        }

        /// <summary>
        /// This is used to click any button in the keyboard. and specify the which button you want to click.
        /// </summary>
        /// <param name="Key">representation of the button you want to click, for that refer "https://www.autoitscript.com/autoit3/docs/appendix/SendKeys.htm"</param>
        public void SendKey(string Key)
        {
            if (Key.Contains("Alt"))
            {
                auto.Send("!{" + Key.Split('+')[1] + "}");
            }
            else if (Key.Contains("Ctrl"))
            {
                auto.Send("^{" + Key.Split('+')[1] + "}");
            }
            else
            {
                auto.Send("{" + Key + "}");
            }
        }
        #endregion

        /// <summary>
        /// This will Scroll the page or table left or right according to the value given in the "Control Keyword Value"
        /// </summary>
        /// <param name="Direction">In "Control keyword" You have to specify "RIGHT" or "LEFT" as direction to scroll </param>
        /// <param name="count">This will specify the mouse wheel count you have to scroll in "ControlKeyword"</param>
        public void HorizantalScroll(string Direction, int count)
        {
            if(Direction == "RIGHT")
            {
                auto.Send("{RIGHT " + count + "}");
            }
            else if (Direction == "LEFT")
            {
                auto.Send("{LEFT " + count + "}");
            }
            else
            {
                throw new Exception("Specify the Direction to move the page");
            }
        }

        /// <summary>
        /// Send Text To a Control that is dynamic
        /// </summary>
        /// <param name="text">Test to be entered in the control</param>
        public void Sendtext(string text)
        {
            auto.Send(text,0);
        }

        /// <summary>
        /// This will get the current system date in the specific format specified in the "GlobalSettings" excel 
        /// </summary>
        /// <param name="dateformat">date for mat of the application.</param>
        /// <returns>return string value in a specified date format</returns>
        public string getSystemDate(string dateformat)
        {
            try
            {
                var date = System.DateTime.Today;
                string systemdate = date.ToString(dateformat);
                if (dateformat.Contains('/'))
                {
                    systemdate = systemdate.Replace('-', '/');
                }
                else if (dateformat.Contains('-'))
                {
                    systemdate = systemdate.Replace('-', '-');
                }
                return systemdate;
            }
            catch(Exception e)
            {
                return string.Empty;
            }
            
        }
        /// <summary>
        /// This will get date from test data from dynamic data and you can process the date. Like add and subtract days
        /// </summary>
        /// <param name="dateformat">This will take default date fort given in the Global settings</param>
        /// <param name="DayCount">You can specify the number of days to add to the given date.</param>
        /// <returns> this will return the processed date in the format given in the  global settings.</returns>
        public string getDate(string dateformat, string DayCount)
        {
            try
            {
                string[] dateDetails = DayCount.Split('|');
                var date = System.DateTime.Today;
                double days = 0;
                int extradayCount = 0;
                WriteAndReadData.DataFilePath = ConfigurationManager.AppSettings["ExcelDataFile"] + "\\DynamicData" + "\\" + batchforReport + ".xls";
                ExcelDataTable.PopulateRecordData(WriteAndReadData.DataFilePath);

                if (dateDetails[0] == "ADD")
                {
                    if (dateDetails[1].Contains('='))
                    {
                        days = Convert.ToDouble(System.Text.RegularExpressions.Regex.Replace(Pop.getDataFromDynamicExcel(dateDetails[1]), "[^0-9]+", string.Empty));
                    }
                    else
                    {
                        days = Convert.ToDouble(ExcelDataTable.ReadData(1, dateDetails[1]));
                    }
                }
                else if (dateDetails[0] == "SUB")
                {
                    if (dateDetails[1].Contains('='))
                    {
                        days = Convert.ToDouble(System.Text.RegularExpressions.Regex.Replace(Pop.getDataFromDynamicExcel(dateDetails[1]), "[^0-9]+", string.Empty));
                        days = days * -1;
                    }
                    else
                    {
                        days = Convert.ToDouble(ExcelDataTable.ReadData(1, dateDetails[1])) * -1;
                    }
                }
                else
                {
                    throw new Exception("Specified Date Operation is not proper");
                }
                if (dateDetails.Count() > 2)
                {
                    extradayCount = Convert.ToInt32(dateDetails[2]);
                    if (dateDetails[0] == "ADD")
                    {
                        date = date.AddDays(days+ extradayCount);
                    }
                    else if (dateDetails[0] == "SUB")
                    {
                        date = date.AddDays(days - extradayCount);
                    }
                }
                else
                {
                    date = date.AddDays(days);
                }

                string systemdate = date.ToString(dateformat);
                if (dateformat.Contains('/'))
                {
                    systemdate = systemdate.Replace('-', '/');
                }
                else if (dateformat.Contains('-'))
                {
                    systemdate = systemdate.Replace('-', '-');
                }
                return systemdate;
            }
            catch(Exception e)
            {
                return string.Empty;
            }
        }
        /// <summary>
        /// Adding the hour, minute, Day, Month, Year, seconds to the given date time 
        /// </summary>
        /// <param name="dateformat">Here we are passing the date format we are keeping for the system or application </param>
        /// <param name="Count">"+Sec" Specifies that the time should return including seconds if not it will not contain the seconds. </param>
        public string GetDateTime(string dateformat, string Count)
        {
            try
            {
                string[] dateDetails = Count.Split('|');
                DateTime dateData;
                string processedTime = string.Empty;
                string Operation = dateDetails[0];
                CultureInfo culture = CultureInfo.CreateSpecificCulture("en-US");

                if (dateDetails[1].Contains('='))
                {
                    DateTime.TryParse(Pop.getDataFromDynamicExcel(dateDetails[1]),culture,DateTimeStyles.None, out dateData);
                }
                else
                {
                    DateTime.TryParse(ExcelDataTable.ReadData(1, dateDetails[1]), culture, DateTimeStyles.None, out dateData);
                }


                if (Operation.Split(':')[0] == "ADD")
                {
                    dateData = AddCalculateDateTime(Operation.Split(':')[1], Convert.ToInt32(dateDetails[2].Split('+')[0]),dateData);
                    //processedTime = dateData.ToString(dateformat + " hh:mm:ss tt");
                }
                else if(Operation.Split(':')[0] == "SUB")
                {
                    dateData = SubtractCalculateDateTime(Operation.Split(':')[1], Convert.ToInt32(dateDetails[2].Split('+')[0]),dateData);
                    //processedTime = dateData.ToString(dateformat + " hh:mm:ss tt");
                }
                else
                {
                    processedTime = string.Empty;
                }
                if (Count.Contains("+Sec"))
                {
                    processedTime = dateData.ToString(dateformat + " hh:mm:ss tt");
                }
                else
                {
                    processedTime = dateData.ToString(dateformat + " hh:mm tt");
                }
                

                if (dateformat.Contains('/'))
                {
                    processedTime = processedTime.Replace('-', '/');
                }
                else if (dateformat.Contains('-'))
                {
                    processedTime = processedTime.Replace('-', '-');
                }
                return processedTime;
            }
            catch(Exception e)
            {
                throw new Exception("Date time conversion Failed");
            }
        }

        public DateTime AddCalculateDateTime(string Attrib, int Count, DateTime date)
        {
            switch (Attrib)
            {
                case "D":
                    return date.AddDays(Count);
                case "M":
                    return date.AddMonths(Count);
                case "Y":
                    return date.AddYears(Count);
                case "H":
                    return date.AddHours(Count);
                case "S":
                    return date.AddSeconds(Count);
                case "Min":
                    return date.AddMinutes(Count);
                default:
                    return DateTime.Now;
            }
        }
        public DateTime SubtractCalculateDateTime(string Attrib, int Count, DateTime date)
        {
            switch (Attrib)
            {
                case "D":
                    return date.AddDays(Count * -1);
                case "M":
                    return date.AddMonths(Count * -1);
                case "Y":
                    return date.AddYears(Count * -1);
                case "H":
                    return date.AddHours(Count * -1);
                case "S":
                    return date.AddSeconds(Count * -1);
                case "Min":
                    return date.AddMinutes(Count * -1);
                default:
                    return DateTime.Now;
            }
        }


        #region DownloadUpload
        /// <summary>
        /// this will help to toggle Download and Upload in page
        /// </summary>
        /// <param name="steps">Step number</param>
        /// <param name="description">Description for the step number</param>
        /// <param name="typeofoperation">Here specify the keyword to toggle Download and Upload Keyword: "DownloadUpload"</param>
        public void ClickDownloadUpload(string steps, string description, string typeofoperation)
        {
            try
            {
                BrowserWindow download = new BrowserWindow();
                download.ExecuteScript(@"var data =  myFunction();
                          function myFunction() {"
                          + "var data = 'Download & Upload';"
                          + "var count = document.getElementsByClassName('search-legend cursor-pointer').length;"
                          + "for(i = 0; i < count; i++){"
                          + "var dataname = document.getElementsByClassName('search-legend cursor-pointer')[0].innerText.trim();"
                          + "if(dataname === data){"
                          + "document.getElementsByTagName('legend')[i].click();break;}}}");
            }
            catch (Exception e)
            {
                Operation.FailerReason = "There is no Download and upload to navigate";
                genreport.Reports(steps, description, typeofoperation, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genreport.FileCorreptionCheck();
                throw new NoSuchControlTypeFound("There is no Download and upload to navigate");
            }
        }
        #endregion

        #region Select Default Value of External Receive Location
        /// <summary>
        /// Getting Default Receive location for the External ware house.based on  product type
        /// </summary>
        /// <param name="step">Step number</param>
        /// <param name="description">Step Description</param>
        /// <param name="keyword">GetDefaultLocation</param>
        /// <param name="Producttype">Reference to the product type(Data reference keyword)</param>
        /// <returns>return the Location for the Product type</returns>
        public string GetDfaultLocation(string step, string description, string keyword, string Producttype, string ErrorMessage)
        {
            try
            {
                BrowserWindow download = new BrowserWindow();
                string DefaultLoation = download.ExecuteScript(@"var data = getDefaultData();
                                function getDefaultData(){
                                var extLocation = document.getElementById('divExternalReceiveLocationInfoList').getElementsByClassName('pull-xs-left').length;
                                for(var i=0; i<extLocation; i++){
                                var productType = document.getElementById('select2-drpExtRecieveType' + i +  '-container').innerHTML;
                                if(productType == '" + Producttype + "'){"
                                + "var d = document.getElementById('select2-drpExtRecieveYardLocation' + i +'-container').innerHTML;"
                                + "break;}else {var d='';}}return d;}return data;").ToString();
                return DefaultLoation;
            }
            catch(Exception e)
            {
                Operation.FailerReason = ErrorMessage + " " + Producttype ;
                genreport.Reports(step, description, keyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genreport.FileCorreptionCheck();
                throw new NoSuchControlTypeFound(ErrorMessage + " " + Producttype);
            }
        }
        #endregion


        #region Select Default Value of External Receive Location
        /// <summary>
        /// Getting Default Receive location for the Internal warehouse based on  product type and Production line
        /// </summary>
        /// <param name="step">Step number</param>
        /// <param name="description">Step Description</param>
        /// <param name="keyword">GetInternDefaultLocation</param>
        /// <param name="Producttype">Reference to the product type and for the production line separated by '|' eg: "<Plant Input>|<Production Line>" (Data reference keyword)</param>
        /// <returns>return the internal receive Location for the Product type based on production line</returns>
        public string GetDfaultInterneralLocation(string step, string description, string keyword, string Producttype, string ErrorMessage)
        {
            try
            {
                BrowserWindow download = new BrowserWindow();
                string DefaultLoation = download.ExecuteScript(@"var data = getDefaultData();
                                function getDefaultData(){
                                var extLocation = document.getElementById('divReceiveLocationInfo').getElementsByClassName('pull-xs-left').length;
                                for(var i=0; i<extLocation; i++){
                                var productType = document.getElementById('select2-drpRecieveType' + i +  '-container').innerHTML;"
                                + "if(productType == '" + ExcelDataTable.ReadData(1, Producttype.Split('|')[0]) + "'){"
                                + "var PLine = document.getElementById('select2-drpRecievePLine' + i + '-container').innerHTML;"
                                + "if(PLine == '" + ExcelDataTable.ReadData(1, Producttype.Split('|')[1]) + "'){"
                                + "var d = document.getElementById('select2-drpRecieveYardLocation' + i +'-container').innerHTML;"
                                + "break;}else {var d='';}}}return d;}return data;").ToString();
                return DefaultLoation;
            }
            catch (Exception e)
            {
                Operation.FailerReason = ErrorMessage + " " + Producttype;
                genreport.Reports(step, description, keyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genreport.FileCorreptionCheck();
                throw new NoSuchControlTypeFound(ErrorMessage + " " + Producttype);
            }
        }
        #endregion

        /// <summary>
        /// This will clean up all the  browser data.
        /// </summary>
        public void CleanUp()
        {
            //Temporary Internet Files
            System.Diagnostics.Process.Start("rundll32.exe", "InetCpl.cpl,ClearMyTracksByProcess 8");
            //Cookies()
            System.Diagnostics.Process.Start("rundll32.exe", "InetCpl.cpl,ClearMyTracksByProcess 2");
            //History()
            System.Diagnostics.Process.Start("rundll32.exe", "InetCpl.cpl,ClearMyTracksByProcess 1");
            //Form(Data)
            System.Diagnostics.Process.Start("rundll32.exe", "InetCpl.cpl,ClearMyTracksByProcess 16");
            //Passwords
            System.Diagnostics.Process.Start("rundll32.exe", "InetCpl.cpl,ClearMyTracksByProcess 32");
            //Delete(All)
            System.Diagnostics.Process.Start("rundll32.exe", "InetCpl.cpl,ClearMyTracksByProcess 255");
            //Delete All – Also delete files and settings stored by add-ons
            System.Diagnostics.Process.Start("rundll32.exe", "InetCpl.cpl,ClearMyTracksByProcess 4351");
        }

        #region Delete Data By Navigate through the pages
        /// <summary>
        /// This will delete the data in the table by navigating through the pages.
        /// </summary>
        /// <param name="step">Step Number.</param>
        /// <param name="description">Description about the step.</param>
        /// <param name="PageID">Here this will accept the table footer id value of the DIV tag that contains the pages</param>
        /// <param name="data">Data You Want to delete or edit</param>
        /// <param name="typeofcontrol">Type of control of the data you want to delete</param>
        /// <param name="codinates">coordinate to move to a control.</param>
        public void NavigationPageDeleteData(string step, string description, string PageID, string data, string typeofcontrol, params int[] codinates)
        {
            BrowserWindow window = new BrowserWindow();
            string lastPage = string.Empty;
            bool availabledata = false;
            string tbody = PageID.Split(':')[0];
            string divID = PageID.Split(':')[1];
            int count = 0;
            try
            {
                try
                {
                    count = Convert.ToInt16(window.ExecuteScript("var data=$('#" + divID + " a').length;return data"));
                }
                catch (Exception e) { }
                //string idval = PageID.Split('_')[0];
                string idval = "ctl01";
                if (count > 3)
                {
                    if (count > 5)
                    {
                        lastPage = idval + "_hrefLast";
                        count = 6;
                    }
                    else
                    {
                        count = 4;
                    }
                }
                if (count > 0)
                {
                    for (int i = 1; i <= count; i++)
                    {
                        window.ExecuteScript("document.getElementById('" + divID + "').getElementsByTagName('a')[" + i + "].scrollIntoView(true); ");
                        Thread.Sleep(2500);
                        //window.ExecuteScript("$('#" + divID + " a')[" + i + "].click();", "javascript");
                        Pop.OperationStart("Wait", "HtmlTable", "ID", tbody, "", "Web", step, description);
                        Thread.Sleep(max);
                        Thread.Sleep(max);
                        try
                        {
                            Pop.OperationStart("ClickAction", typeofcontrol, "SearchText", data, "", "Web", step, description, codinates);
                            availabledata = true;
                            break;
                        }
                        catch (Exception e) { }
                    }
                }
                else
                {
                    try
                    {
                        Pop.OperationStart("ClickAction", typeofcontrol, "SearchText", data, "", "Web", step, description, codinates);
                        availabledata = true;
                    }
                    catch (Exception e) { }
                }
                if (!availabledata)
                {
                    Operation.FailerReason = "Given data is not available in the page";
                    genreport.Reports(step, description, "SearchDataOperation", false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                    Assert.Fail("Given data is not available in the page");
                }
            }
            catch (Exception e)
            {
                Operation.FailerReason = "Given data is not available in the page";
                genreport.Reports(step, description, "SearchDataOperation", false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genreport.FileCorreptionCheck();
                Assert.Fail("Given data is not available in the page");
            }
        }
        #endregion

        #region Launch SQLDeveloperApplication
        /// <summary>
        /// This will Launch The Test Application, Invoking Keyword..
        /// </summary>
        public void TestApplicationLaunch(string ControlKeyVal, string Application)
        {
            WinWindow wind = new WinWindow();
            wind.SearchProperties[WinWindow.PropertyNames.ClassName] = ControlKeyVal;
            UITestControlCollection wndctrl = wind.FindMatchingControls();
            string ExEpath = ExcelDataTable.ReadData(1, Application);
            ApplicationUnderTest application = ApplicationUnderTest.Launch(ExEpath);
            //application.Maximized = true;
        }
        #endregion

        /// <summary>
        /// This will generate a random number in string format and return.
        /// </summary>
        /// <param name="numberCount">here you can specify the number of characters or Digits for a number want</param>
        /// <returns></returns>
        #region RandomNumberGen
        public string RamdomBatch(int numberCount)
        {
            try
            {
                StringBuilder builder = new StringBuilder(6);
                string numberAsString = string.Empty;

                for (var i = 0; i < numberCount; i++)
                {
                    builder.Append(_numbers[random.Next(0, _numbers.Length)]);
                }

                numberAsString = builder.ToString();
                return numberAsString;
            }
            catch(Exception e)
            {
                Operation.FailerReason = "Random Number generation failed";
                genreport.FileCorreptionCheck();
                throw new NoSuchControlTypeFound("Random Number generation failed");
            }
            
        }
        #endregion

        ///<summary>
        ///This function will get the number of rows present in the particular table
        /// </summary>
        /// <param name="step">Step Number.</param>
        /// <param name="description">Description about the step.</param>
        /// <param name="TBodyID">this will accept tbody id value and count number of rows present in table</param>
        /// <param name="ControlKeyword"> Table Unique attribute like ID</param>
        /// <param name="TypeofControl">type of control to count </param>
        /// <param name="operationKeyword">this will be triggered using keyword Ex:TableRowCount </param>
        public void ValidateTableRowCount(string TypeControl, string ControlKeyword, string TBodyID, string StepId, string Description, string Operationkeyword, string AssertScreenshotName, string AssertionMsg)
        {
            BrowserWindow window = new BrowserWindow();
            
            int tRowCount = 0, tbody = 0;
            try
            {
                tRowCount = Convert.ToInt16(window.ExecuteScript("var count=document.getElementById('" + TBodyID + "').getElementsByTagName('tr').length; return count"));            
                tbody = Convert.ToInt16(window.ExecuteScript("var count=document.getElementById('" + TBodyID + "').getElementsByTagName('tbody').length; return count"));
                if (tbody != 0)
                {
                    if (tbody <= 2)
                    {
                        TakeScreenshot(StepId, AssertScreenshotName, AssertionMsg, Description);
                    }
                    else
                    {
                        genreport.Reports(StepId, Description, Operationkeyword, true, LoginOperatrion.batchforReport, DetaildReportStatus, "");
                    }
                }
                else if (tRowCount != 0)
                {
                    if (tRowCount < 2)
                    {
                        TakeScreenshot(StepId, AssertScreenshotName, AssertionMsg, Description);
                    }
                    else
                    {
                        genreport.Reports(StepId, Description, Operationkeyword, true, LoginOperatrion.batchforReport, DetaildReportStatus, "");
                    }
                }

            }
            
            catch (Exception e)
            {
                TakeScreenshot(StepId, AssertScreenshotName, AssertionMsg, Description);
            }
        }

        /// <summary>
        /// This will Get the Process Order Number for particular Product ID, from "View Production DetailsPage"
        /// </summary>
        /// <param name="StepNo">Step Number</param>
        /// <param name="StepDescription">Description for that Step</param>
        /// <param name="Keyword">"GetProcessOrderNo" This key Word will access to this function.</param>
        /// <returns>Returns 12 Digit Process Order Number</returns>
        public string GetProcessOrder(string StepNo, string StepDescription, string Keyword)
        {
            BrowserWindow window = new BrowserWindow();
            string ProcessOrder = string.Empty;
            try
            {
                ProcessOrder = window.ExecuteScript(@"var str = document.getElementById('form1').getAttribute('action');
                            var PO = str.substr(str.indexOf('&plineid')-12,12); return PO;").ToString();
                return ProcessOrder;
            }
            catch (Exception e)
            {
                Operation.FailerReason = "Could not find Process Order Number from the Production Details Page URL";
                genreport.Reports(StepNo, StepDescription, Keyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genreport.FileCorreptionCheck();
                throw new NoSuchControlTypeFound("Could not find Process Order Number from the Production Details Page URL");
            }
        }

        /// <summary>
        /// This will Get the Phase Number for particular SOP, from "SOP Details Page PHASE info tab"
        /// </summary>
        /// <param name="StepNo">Step Number</param>
        /// <param name="StepDescription">Description for that Step</param>
        /// <param name="Keyword">"GetPhaseNumber" This key Word will access to this function.</param>
        /// <returns>Returns 4 Digit Phase Number</returns>
        public string GetPhaseNumber(string StepNo, string StepDescription, string Keyword)
        {
            BrowserWindow window = new BrowserWindow();
            string PhaseNumber = string.Empty;
            try
            {
                PhaseNumber = window.ExecuteScript(@"var phDetails = document.getElementById('ucPhaseInfoControl_0_phaseNumberSpan').innerHTML;
                            var PhaseNo=phDetails.substr(phDetails.indexOf('Phase - ')-4,4); return PhaseNo;").ToString();
                return PhaseNumber;
            }
            catch (Exception e)
            {
                Operation.FailerReason = "Could not find Phase details in the SOP Details Page phase from Phase Info TAB";
                genreport.Reports(StepNo, StepDescription, Keyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genreport.FileCorreptionCheck();
                throw new NoSuchControlTypeFound("Could not find Phase details in the SOP Details Page phase from Phase Info TAB");
            }
        }


        /// <summary>
        /// TakeScreenshot will capture the screen when condition failed or assertion failed , it accepts two parameters ImageName and Error message
        /// </summary>
        /// <param name="step">Step number</param>
        /// <param name="ImageName">caption for screen shot  </param>
        /// <param name="message">Error message want to write </param>
        public void TakeScreenshot(string step, string ImageName, string message, string description)
        {
            BrowserWindow window = new BrowserWindow();
            string path = ConfigurationManager.AppSettings["ScreenShot"];
            path = path + @"\" + LoginOperatrion.ProjectName + "";
            Directory.CreateDirectory(path);
            try
            {
                Image image = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);
                using (Graphics g = Graphics.FromImage(image))
                {
                    image = window.CaptureImage();
                    image.Save(path + "\\" + ImageName + ".jpeg", ImageFormat.Jpeg);
                    Operation.ErrorScreenPath = path + ImageName + ".jpeg";
                    image.Dispose();
                }
            }
            catch (Exception v) { }
            Operation.FailerReason = message;
            genreport.Reports(step, description, "ValidatingPageData", false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
            Assert.Fail(message);

        }
        /// <summary>
        /// Bind Two data together and create the another data.
        /// </summary>
        /// <param name="Data">In "ControlKeywordValue" give what data you want to bind or the references</param>
        /// <param name="stepno">Step Number for the operation</param>
        /// <param name="description">Description for the step</param>
        /// <returns>This will return binded data.</returns>
        public string GetDataBinded(string Data, string stepno, string description)
        {
            string[] processdata = Data.Split('+');
            string processedData = string.Empty;
            if (processdata.Count() > 1)
            {
                for(int i =0; i<processdata.Count(); i++)
                {
                    if (processdata[i].Contains('='))
                    {
                        processedData = processedData + Pop.getDataFromDynamicExcel(processdata[i].ToString());
                    }
                    else
                    {
                        processedData = processedData + ExcelDataTable.ReadData(1, processdata[i].ToString());
                    }
                }
            }
            else
            {
                throw new Exception("Data to bind not given properly");
            }
            return processedData;
        }

        /// <summary>
        /// For calculating the formula and get the result.
        /// </summary>
        /// <param name="keyword">"CalcFormula" This keyword will used to process this function</param>
        /// <param name="Formula">Here the formula will give including the all reference to the data.</param>
        /// <param name="step">Step Number</param>
        /// <param name="description">Description for the step</param>
        /// <returns>This will return the result of the calculation</returns>
        public void CalculatePlateWeight(string keyword, string Formula, string step, string description, string DataRefferencekeyword)
        {
            var GeneratedFormula = string.Empty;
            try
            {
                if(Formula != string.Empty)
                {
                    string[] data = Formula.Split('|');
                    string[] values = data[1].Split(',');
                    for (int i = 0; i < values.Count(); i++)
                    {
                        if (i == 0)
                        {
                            if (values[i].StartsWith("Rec_"))
                            {
                                GeneratedFormula = data[0].Replace("[X" + i + "]", Operation.recordedData[values[i].Replace("Rec_", string.Empty) + ".0"]) .ToString();
                            }
                            else if (values[i].StartsWith("Reference"))
                            {
                                GeneratedFormula = data[0].Replace("[X" + i + "]", Pop.getDataFromDynamicExcel(values[i]) + ".0");
                            }
                            else
                            {
                                GeneratedFormula = data[0].Replace("[X" + i + "]", ExcelDataTable.ReadData(1, values[i]) + ".0") .ToString();
                            }
                        }
                        else
                        {
                            if (values[i].StartsWith("Rec_"))
                            {
                                GeneratedFormula = GeneratedFormula.Replace("[X" + i + "]", Operation.recordedData[values[i].Replace("Rec_", string.Empty) + ".0"]).ToString();
                            }
                            else if (values[i].StartsWith("Reference"))
                            {
                                GeneratedFormula = GeneratedFormula.Replace("[X" + i + "]", Pop.getDataFromDynamicExcel(values[i]) + ".0") ;
                            }
                            else
                            {
                                GeneratedFormula = GeneratedFormula.Replace("[X" + i + "]", ExcelDataTable.ReadData(1, values[i]) + ".0") .ToString();
                            }
                        }
                    }
                    double Result = Convert.ToDouble(new DataTable().Compute(GeneratedFormula, null));
                    datawrite.WriteExcel(DataRefferencekeyword.ToString().Split(':')[1], Result.ToString("0.000"), "ReadData", DataRefferencekeyword.Split(':')[0]);
                }
                else
                {
                    throw new Exception("Please specify the formula to do calculation");
                }
            }
            catch (Exception e)
            {
                Operation.FailerReason = "Not able to process the formula.";
                genreport.Reports(step, description, keyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genreport.FileCorreptionCheck();
                throw new NoSuchControlTypeFound("Not able to process the formula.");
            }
        }

        /// <summary>
        /// To get the Data from the Table order by recent date
        /// </summary>
        /// <param name="step">Step Number</param>
        /// <param name="description">Description for the step</param>
        /// <param name="keyword">"GetDataOrderByDate" This keyword will invoke the function</param>
        /// <param name="ControlKeywordValue">this will contain  the table id value and column for checking the order and column number for the value we want to take</param>
        /// <returns> this will return value for the recent date in the table</returns>
        public string Density(string step, string description, string keyword, string ControlKeywordValue)
        {
            BrowserWindow br = new BrowserWindow();
            int index = -1;
            DateTime d1, d2;
            string result = string.Empty;
            var compaireval = 0;
            try
            {
                int dCount = Convert.ToInt32(br.ExecuteScript(@"var dCount = document.getElementById('"+ControlKeywordValue.Split('|')[0]+"').getElementsByTagName('tr').length; return dCount;"));
                if (dCount > 1)
                {
                    for (int i = 0; i < dCount; i++)
                    {
                        if (index == -1)
                        {
                            int j = i + 1;
                            d1 = Convert.ToDateTime(br.ExecuteScript(@"var Fval = document.getElementById('"+ ControlKeywordValue.Split('|')[0] + "').getElementsByTagName('tr')['" + i + "'].getElementsByTagName('td')['" + ControlKeywordValue.Split('|')[1].Split('/')[0].ToString().Replace("ORD:", string.Empty) + "'].innerHTML; return Fval;"));
                            d2 = Convert.ToDateTime(br.ExecuteScript(@"var Fval = document.getElementById('"+ ControlKeywordValue.Split('|')[0] + "').getElementsByTagName('tr')['" + j + "'].getElementsByTagName('td')['" + ControlKeywordValue.Split('|')[1].Split('/')[0].ToString().Replace("ORD:", string.Empty) + "'].innerHTML; return Fval;"));
                            compaireval = DateTime.Compare(d1, d2);
                            if (compaireval == -1)
                            {
                                index = j;
                            }
                            else
                            {
                                index = i;
                            }
                        }
                        else
                        {
                            d1 = Convert.ToDateTime(br.ExecuteScript(@"var Fval = document.getElementById('"+ ControlKeywordValue.Split('|')[0] + "').getElementsByTagName('tr')['" + index + "'].getElementsByTagName('td')['" + ControlKeywordValue.Split('|')[1].Split('/')[0].ToString().Replace("ORD:", string.Empty) + "'].innerHTML; return Fval;"));
                            d2 = Convert.ToDateTime(br.ExecuteScript(@"var Fval = document.getElementById('"+ ControlKeywordValue.Split('|')[0] + "').getElementsByTagName('tr')['" + i + "'].getElementsByTagName('td')['" + ControlKeywordValue.Split('|')[1].Split('/')[0].ToString().Replace("ORD:", string.Empty) + "'].innerHTML; return Fval;"));
                            compaireval = DateTime.Compare(d1, d2);
                            if (compaireval == -1)
                            {
                                index = i;
                            }
                            else
                            {
                                continue;
                            }
                        }
                    }
                    result = br.ExecuteScript(@"var Fval = document.getElementById('"+ ControlKeywordValue.Split('|')[0] + "').getElementsByTagName('tr')['" + index + "'].getElementsByTagName('td')['"+ ControlKeywordValue.Split('|')[1].Split('/')[1].ToString().Replace("VAL:",string.Empty) + "'].innerHTML; return Fval;").ToString();
                }
                else
                {
                    result = br.ExecuteScript(@"var Fval = document.getElementById('"+ ControlKeywordValue.Split('|')[0] + "').getElementsByTagName('tr')['0'].getElementsByTagName('td')['" + ControlKeywordValue.Split('|')[1].Split('/')[1].ToString().Replace("VAL:", string.Empty) + "'].innerHTML; return Fval;").ToString();
                }
                return result;
            }
            catch(Exception e)
            {
                Operation.FailerReason = "Not able to find density details";
                genreport.Reports(step, description, keyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                TakeScreenshot(step, Operation.FailerReason, Operation.FailerReason, description);
                genreport.FileCorreptionCheck();
                throw new NoSuchControlTypeFound("Not able to find density details");
            }
        }
    }
}
