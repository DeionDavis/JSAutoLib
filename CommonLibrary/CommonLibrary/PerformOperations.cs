﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UITesting.HtmlControls;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using Microsoft.VisualStudio.TestTools.UITesting;
using CommonLibrary.DataDrivenTesting;
using CommonLibrary.CommonLanguageReader;
using CommonLibrary.Exceptions;
using CommonLibrary.Log;
using Microsoft.VisualStudio.TestTools.UITest.Extension;
using System.Drawing;
using System.Threading;
using AutoItX3Lib;
using System.Data;
using System.Diagnostics;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Drawing.Imaging;
using System.Configuration;
using OperationLibrary;
using CommonLibrary.Reports;
using System.IO;
using System.Windows.Forms;
using CommonLibrary.Writedata;
using System.Globalization;

namespace CommonLibrary.Operations
{
    public class PerformOperation
    {
        BrowserWindow window = new BrowserWindow();
        AutoItX3 auto = new AutoItX3();
        ReportGeneration genDetailedReport = new ReportGeneration();
        CommonLanguageTemplateReader languageResource = new CommonLanguageTemplateReader();
        public bool status = false;
        public void OperationStart(string _TypeOfOperation, string _TypeControl, string _ControlKeyword, string _ControlKeywordValue, string _DataRefferencekeyword, string _TypeOfWindow, string _StepId, string _Description, params int[] codinates)
        {
            switch (_TypeOfWindow)
            {
                case "Web":
                    switch (_TypeOfOperation)
                    {
                        case "EnterText":
                        case "EnterLanguageText":
                        case "WriteData":
                        case "EnterRandomNumber":
                            status = WebEnterData(_TypeControl, _ControlKeyword, _ControlKeywordValue, _DataRefferencekeyword, _StepId, _Description, _TypeOfOperation);
                            if(_TypeOfOperation== "WriteData")
                            {
                                genDetailedReport.Reports(_StepId, _Description, _TypeOfOperation, status, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, _DataRefferencekeyword);
                            }
                            else
                            {
                                genDetailedReport.Reports(_StepId, _Description, _TypeOfOperation, status, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            }
                            break;
                        case "ClickAction":
                        case "RightClickAction":
                            if (_TypeOfOperation == "RightClickAction")
                            {
                                status = WebClickOperation(_TypeControl, _ControlKeyword, _ControlKeywordValue, _StepId, _Description, _TypeOfOperation, MouseButtons.Right, codinates);
                            }
                            else
                            {
                                status = WebClickOperation(_TypeControl, _ControlKeyword, _ControlKeywordValue, _StepId, _Description, _TypeOfOperation, MouseButtons.Left, codinates);
                            }
                            genDetailedReport.Reports(_StepId, _Description, _TypeOfOperation, status, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "DblClickAction":
                            status = WebDoubleClickOperation(_TypeControl, _ControlKeyword, _ControlKeywordValue, _StepId, _Description, _TypeOfOperation, MouseButtons.Left, codinates);
                            genDetailedReport.Reports(_StepId, _Description, _TypeOfOperation, status, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "Clear":
                            status = WebClearControlData(_TypeControl, _ControlKeyword, _ControlKeywordValue, _StepId, _Description, _TypeOfOperation);
                            genDetailedReport.Reports(_StepId, _Description, _TypeOfOperation, status, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "Click":
                        case "RightClick":
                            if (_TypeOfOperation == "Click")
                            {
                                status = WebClickControl(_TypeControl, _ControlKeyword, _ControlKeywordValue, _StepId, _Description, _TypeOfOperation, MouseButtons.Left);
                            }
                            else if (_TypeOfOperation == "RightClick")
                            {
                                status = WebClickControl(_TypeControl, _ControlKeyword, _ControlKeywordValue, _StepId, _Description, _TypeOfOperation, MouseButtons.Right);
                            }
                            genDetailedReport.Reports(_StepId, _Description, _TypeOfOperation, status, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "Wait":
                        case "WaitPropertySet":
                            if(_TypeOfOperation== "WaitPropertySet")
                            {
                                status = WebWaitForControl(_TypeControl, _ControlKeyword, _ControlKeywordValue, _StepId, _Description, _TypeOfOperation, _DataRefferencekeyword);
                            }
                            else if(_TypeOfOperation == "Wait")
                            {
                                status = WebWaitForControl(_TypeControl, _ControlKeyword, _ControlKeywordValue, _StepId, _Description, _TypeOfOperation,"");
                            }
                            genDetailedReport.Reports(_StepId, _Description, _TypeOfOperation, status, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "WaitPageLoad":
                            status = PageLoadWait(_TypeControl, _ControlKeyword, _ControlKeywordValue, _StepId, _Description, _TypeOfOperation);
                            genDetailedReport.Reports(_StepId, _Description, _TypeOfOperation, status, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "ClickLanguageText":
                            status = WebClickControlText(_TypeControl, _ControlKeyword, _ControlKeywordValue, _StepId, _Description, _TypeOfOperation);
                            genDetailedReport.Reports(_StepId, _Description, _TypeOfOperation, status, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "PointClick":
                        case "PointRightClick":
                            if (_TypeOfOperation == "PointClick")
                            {
                                status = WebMouseMoveClick(_TypeControl, _ControlKeyword, _ControlKeywordValue, _StepId, _Description, _TypeOfOperation, MouseButtons.Left, codinates);
                            }
                            else if (_TypeOfOperation == "PointRightClick")
                            {
                                status = WebMouseMoveClick(_TypeControl, _ControlKeyword, _ControlKeywordValue, _StepId, _Description, _TypeOfOperation, MouseButtons.Right, codinates);
                            }
                            genDetailedReport.Reports(_StepId, _Description, _TypeOfOperation, status, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "Check":
                            status = WebSelectCheckBox(_TypeControl, _ControlKeyword, _ControlKeywordValue, _DataRefferencekeyword, _StepId, _Description, _TypeOfOperation);
                            genDetailedReport.Reports(_StepId, _Description, _TypeOfOperation, status, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "ScriptClick":
                            status = ScriptClick(_TypeOfOperation, _ControlKeyword, _ControlKeywordValue, _StepId, _Description);
                            genDetailedReport.Reports(_StepId, _Description, _TypeOfOperation, status, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        default:
                            Operation.FailerReason = "No " + _TypeOfOperation + " keyword found. Please check the keyword.";
                            throw new NoSuchOperationFound("No " + _TypeOfOperation + " keyword found. Please check the keyword.");
                    }
                    break;
                case "Window":
                    switch (_TypeOfOperation)
                    {
                        case "EnterText":
                        case "WriteData":
                        case "EnterRandomNumber":
                            WindowEnterData(_TypeControl, _ControlKeyword, _ControlKeywordValue, _DataRefferencekeyword);
                            if (_TypeOfOperation == "WriteData")
                            {
                                genDetailedReport.Reports(_StepId, _Description, _TypeOfOperation, status, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, _DataRefferencekeyword);
                            }
                            else
                            {
                                genDetailedReport.Reports(_StepId, _Description, _TypeOfOperation, status, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            }
                            break;
                        case "Click":
                            WindowClickControl(_TypeControl, _ControlKeyword, _ControlKeywordValue);
                            genDetailedReport.Reports(_StepId, _Description, _TypeOfOperation, status, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "Wait":
                            WindowWaitControl(_TypeControl, _ControlKeyword, _ControlKeywordValue);
                            genDetailedReport.Reports(_StepId, _Description, _TypeOfOperation, status, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "Clear":
                            status = WindowClearControl(_TypeControl, _ControlKeyword, _ControlKeywordValue, _StepId, _Description, _TypeOfOperation);
                            genDetailedReport.Reports(_StepId, _Description, _TypeOfOperation, status, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        case "PointClick":
                        case "PointRightClick":
                            if (_TypeOfOperation == "PointClick")
                            {
                                status = WindowMouseMoveClick(_TypeControl, _ControlKeyword, _ControlKeywordValue, _StepId, _Description, _TypeOfOperation, MouseButtons.Left, codinates);
                            }
                            else if (_TypeOfOperation == "PointRightClick")
                            {
                                status = WindowMouseMoveClick(_TypeControl, _ControlKeyword, _ControlKeywordValue, _StepId, _Description, _TypeOfOperation, MouseButtons.Right, codinates);
                            }
                            genDetailedReport.Reports(_StepId, _Description, _TypeOfOperation, status, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                            break;
                        default:
                            Operation.FailerReason = "No " + _TypeOfOperation + " keyword found. Please check the keyword.";
                            throw new NoSuchOperationFound("No Control Type " + _TypeOfOperation + " please check the Type of Operation.");
                    }
                    break;
                default:
                    Operation.FailerReason = "No " + _TypeOfWindow + " window type found, Please check the window type.";
                    throw new NoSuchWindowTypeFound("No " + _TypeOfWindow + " window type found, Please check the window type.");
            }
        }

        #region Web
        /// <summary>
        /// This used to enter the data to the control.based on the values given
        /// </summary>
        /// <param name="TypeControl">Specify for which Type of control you want to perform this action.</param>
        /// <param name="ControlKeyword">Property that you used to identify the control</param>
        /// <param name="ControlKeywordValue">What is the value of the property you specified</param>
        /// <param name="DataRefferencekeyword">Reference to the data you want to enter in the control.</param>
        /// <param name="steps">Step Number.</param>
        /// <param name="description">Description about the step.</param>
        /// <param name="Operationkeyword">This function will be triggered using Keyword: "EnterText".</param>
        /// <returns>Return the true or false value.</returns>
        public bool WebEnterData(string TypeControl, string ControlKeyword, string ControlKeywordValue, string DataRefferencekeyword, string StepId, string Description, string Operationkeyword)
        {
            HtmlControl genericsControl = (HtmlControl)Activator.CreateInstance(typeof(HtmlControl), new object[] { ParentWindow });
            Keyboard.SendKeysDelay = 0;
            Point p = new Point();
            try
            {
                switch (TypeControl)
                {
                    case "HtmlEdit":
                        switch (ControlKeyword)
                        {
                            case "Class":
                                genericsControl.SearchProperties[HtmlEdit.PropertyNames.Class] = ControlKeywordValue;
                                Keyboard.SendKeys(genericsControl, DataRefferencekeyword);
                                break;
                            case "IDVisible":
                                genericsControl.SearchProperties[HtmlEdit.PropertyNames.Id] = ControlKeywordValue;
                                bool d = genericsControl.TryGetClickablePoint(out p);
                                if (d)
                                {
                                    Keyboard.SendKeys(genericsControl, DataRefferencekeyword);
                                }
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlEdit.PropertyNames.Id] = ControlKeywordValue;
                                Keyboard.SendKeys(genericsControl, DataRefferencekeyword);
                                break;
                        }
                        return true;
                    case "HtmlComboBox":
                        switch (ControlKeyword)
                        {
                            case "ID":
                                genericsControl.SearchProperties[HtmlComboBox.PropertyNames.Id] = ControlKeywordValue;
                                Keyboard.SendKeys(genericsControl, DataRefferencekeyword);
                                break;
                            case "IDData":
                                string[] IdData = ExcelDataTable.ReadData(1, ControlKeywordValue).Split('+');
                                if (IdData.Count() > 1)
                                {
                                    if (IdData[1].StartsWith("Rec_"))
                                    {
                                        ControlKeywordValue = IdData[0].Replace("&", Operation.recordedData[IdData[1].Replace("Rec_", string.Empty)]);
                                    }
                                    else
                                    {
                                        ControlKeywordValue = IdData[0].Replace("&", ExcelDataTable.ReadData(1, IdData[1]));
                                    }
                                }
                                else {
                                    throw new NullReferenceException("Reference Data not found after + Symbol");
                                }
                                genericsControl.SearchProperties[HtmlHyperlink.PropertyNames.Id] = ControlKeywordValue;
                                Keyboard.SendKeys(genericsControl, DataRefferencekeyword);
                                break;
                            case "IDVisible":
                                genericsControl.SearchProperties[HtmlComboBox.PropertyNames.Id] = ControlKeywordValue;
                                bool d = genericsControl.TryGetClickablePoint(out p);
                                if (d)
                                {
                                    Keyboard.SendKeys(genericsControl, DataRefferencekeyword);
                                }
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlComboBox.PropertyNames.Id] = ControlKeywordValue;
                                Keyboard.SendKeys(genericsControl, DataRefferencekeyword);
                                break;
                        }
                        return true;
                    case "HtmlTextArea":
                        switch (ControlKeyword)
                        {
                            case "ID":
                                genericsControl.SearchProperties[HtmlTextArea.PropertyNames.Id] = ControlKeywordValue;
                                Keyboard.SendKeys(genericsControl, DataRefferencekeyword);
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlTextArea.PropertyNames.Id] = ControlKeywordValue;
                                Keyboard.SendKeys(genericsControl, DataRefferencekeyword);
                                break;
                        }
                        return true;
                    default:
                        Operation.FailerReason = TypeControl + " Control Type not configured for entering data. Please check the Control Type.";
                        genDetailedReport.Reports(StepId, Description, Operationkeyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                        genDetailedReport.FileCorreptionCheck();
                        throw new NoSuchControlTypeFound(TypeControl + " Control Type not configured for entering data. Please check the Control Type.");
                }
            }
            catch (UITestControlNotFoundException e)
            {
                genDetailedReport.Reports(StepId, Description, Operationkeyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                Operation.FailerReason = TypeControl + " with " + ControlKeywordValue + " as " + ControlKeyword + " not found in page";
                genDetailedReport.FileCorreptionCheck();
                TakeScreenshot(ControlKeywordValue + " as " + ControlKeyword + " not found");
                Assert.Fail(Operation.FailerReason);
                return false;
            }
            catch (FailedToPerformActionOnBlockedControlException e)
            {
                genericsControl.DrawHighlight();
                TakeScreenshot("ControlCannotBeClicked");
                genDetailedReport.Reports(StepId, Description, Operationkeyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                Operation.FailerReason = TypeControl + " with " + ControlKeywordValue + " as " + ControlKeyword + " is blocked by another control";
                genDetailedReport.FileCorreptionCheck();
                Assert.Fail(Operation.FailerReason);
                return false;
            }
        }

        ///<summary>
        ///<param name="name">name of failed screen shot</param>
        /// </summary>
        public void TakeScreenshot(string name)
        {
            string path = ConfigurationManager.AppSettings["ScreenShot"];
            path = path + @"\" + LoginOperatrion.ProjectName + "";
            Directory.CreateDirectory(path);
            try
            {
                Image image = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);
                using (Graphics g = Graphics.FromImage(image))
                {
                    g.CopyFromScreen(0, 0, 0, 0, Screen.PrimaryScreen.Bounds.Size);
                    image.Save(path + "\\" + name + ".jpeg", ImageFormat.Jpeg);
                    Operation.ErrorScreenPath = path + "\\" + name + ".jpeg";
                    image.Dispose();
                }
            }
            catch (Exception v) { }
        }






        /// <summary>
        /// This used to click the control.based on the values given
        /// </summary>
        /// <param name="TypeControl">Specify for which Type of control you want to perform this action.</param>
        /// <param name="ControlKeyword">Property that you used to identify the control</param>
        /// <param name="ControlKeywordValue">What is the value of the property you specified</param>
        /// <param name="steps">Step Number.</param>
        /// <param name="description">Description about the step.</param>
        /// <param name="Operationkeyword">This function will be triggered using Keyword: "Click"</param>
        /// <returns>Return the true or false value.</returns>
        public bool WebClickControl(string TypeControl, string ControlKeyword, string ControlKeywordValue, string StepId, string Description, string Operationkeyword, MouseButtons Button)
        {
            HtmlControl genericsControl = (HtmlControl)Activator.CreateInstance(typeof(HtmlControl), new object[] { ParentWindow });
            Point p = new Point();
            try
            {
                switch (TypeControl)
                {
                    case "HtmlInputButton":
                        switch (ControlKeyword)
                        {
                            case "InnerText":
                                genericsControl.SearchProperties[HtmlInputButton.PropertyNames.InnerText] = ControlKeywordValue;
                                Mouse.Click(genericsControl, Button);
                                break;
                            case "FriendlyName":
                                genericsControl.SearchProperties[HtmlInputButton.PropertyNames.FriendlyName] = ControlKeywordValue;
                                Mouse.Click(genericsControl, Button);
                                break;
                            case "IDData":
                                string[] IdData = ExcelDataTable.ReadData(1, ControlKeywordValue).Split('+');
                                if (IdData.Count() > 1)
                                {
                                    if (IdData[1].StartsWith("Rec_"))
                                    {
                                        ControlKeywordValue = IdData[0].Replace("&", Operation.recordedData[IdData[1].Replace("Rec_", string.Empty)]);
                                    }
                                    else
                                    {
                                        ControlKeywordValue = IdData[0].Replace("&", ExcelDataTable.ReadData(1, IdData[1]));
                                    }
                                }
                                else {
                                    throw new NullReferenceException("Reference Data not found after + Symbol");
                                }
                                genericsControl.SearchProperties[HtmlInputButton.PropertyNames.Id] = ControlKeywordValue;
                                Mouse.Click(genericsControl, Button);
                                break;
                            case "IDVisible":
                                genericsControl.SearchProperties[HtmlInputButton.PropertyNames.Id] = ControlKeywordValue;
                                bool d = genericsControl.TryGetClickablePoint(out p);
                                if (d)
                                {
                                    Mouse.Click(genericsControl, Button);
                                }
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlInputButton.PropertyNames.Id] = ControlKeywordValue;
                                Mouse.Click(genericsControl, Button);

                                break;
                        }
                        return true;
                    case "HtmlComboBox":
                        switch (ControlKeyword)
                        {
                            case "ID":
                                genericsControl.SearchProperties[HtmlComboBox.PropertyNames.Id] = ControlKeywordValue;
                                Mouse.Click(genericsControl, Button);
                                break;
                            case "IDData":
                                string[] IdData = ExcelDataTable.ReadData(1, ControlKeywordValue).Split('+');
                                if (IdData.Count() > 1)
                                {
                                    if (IdData[1].StartsWith("Rec_"))
                                    {
                                        ControlKeywordValue = IdData[0].Replace("&", Operation.recordedData[IdData[1].Replace("Rec_", string.Empty)]);
                                    }
                                    if (IdData[1].Contains('='))
                                    {
                                        WriteAndReadData.DataFilePath = ConfigurationManager.AppSettings["ExcelDataFile"] + "\\DynamicData" + "\\" + LoginOperatrion.batchforReport + ".xls";
                                        ExcelDataTable.PopulateRecordData(WriteAndReadData.DataFilePath);
                                        ControlKeywordValue = IdData[0].Replace("&", getDataFromDynamicExcel(ControlKeywordValue));
                                    }
                                    else
                                    {
                                        ControlKeywordValue = IdData[0].Replace("&", ExcelDataTable.ReadData(1, IdData[1]));
                                    }
                                }
                                else {
                                    throw new NullReferenceException("Reference Data not found after + Symbol");
                                }
                                genericsControl.SearchProperties[HtmlComboBox.PropertyNames.Id] = ControlKeywordValue;
                                Mouse.Click(genericsControl, Button);
                                break;
                            case "IDVisible":
                                genericsControl.SearchProperties[HtmlComboBox.PropertyNames.Id] = ControlKeywordValue;
                                bool d = genericsControl.TryGetClickablePoint(out p);
                                if (d)
                                {
                                    Mouse.Click(genericsControl, Button);
                                }
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlComboBox.PropertyNames.Id] = ControlKeywordValue;
                                Mouse.Click(genericsControl, Button);
                                break;
                        }
                        return true;
                    case "HtmlSpan":
                        switch (ControlKeyword)
                        {
                            case "InnerText":
                                genericsControl.SearchProperties[HtmlSpan.PropertyNames.InnerText] = ExcelDataTable.ReadData(1, ControlKeywordValue);
                                Mouse.Click(genericsControl, Button);
                                break;
                            case "IDClickDisable":
                                Thread.Sleep(200);
                                genericsControl.SearchProperties[HtmlCheckBox.PropertyNames.Id] = ControlKeywordValue;
                                string d = genericsControl.GetProperty("checked").ToString();
                                if (d == "False")
                                {
                                    ScriptClick("Click", "ID", ControlKeywordValue, StepId, Description);
                                }
                                break;
                            case "IDClickEnable":
                                Thread.Sleep(200);
                                genericsControl.SearchProperties[HtmlCheckBox.PropertyNames.Id] = ControlKeywordValue;
                                d = genericsControl.GetProperty("checked").ToString();
                                if (d == "True")
                                {
                                    ScriptClick("Click", "ID", ControlKeywordValue, StepId, Description);
                                }
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlSpan.PropertyNames.Id] = ControlKeywordValue;
                                Mouse.Click(genericsControl, Button);
                                break;
                        }
                        return true;
                    case "HtmlFileInput":
                        switch (ControlKeyword)
                        {
                            case "ID":
                                genericsControl.SearchProperties[HtmlFileInput.PropertyNames.Id] = ControlKeywordValue;
                                Mouse.Click(genericsControl, Button);
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlFileInput.PropertyNames.Id] = ControlKeywordValue;
                                Mouse.Click(genericsControl, Button);
                                break;
                        }
                        return true;
                    case "HtmlCell":
                        switch (ControlKeyword)
                        {
                            case "InnerText":
                                if (ControlKeywordValue.Contains('='))
                                {
                                    WriteAndReadData.DataFilePath = ConfigurationManager.AppSettings["ExcelDataFile"] + "\\DynamicData" + "\\" + Operation.Batch + ".xls";
                                    ExcelDataTable.PopulateRecordData(WriteAndReadData.DataFilePath);
                                    genericsControl.SearchProperties[HtmlHyperlink.PropertyNames.InnerText] = getDataFromDynamicExcel(ControlKeywordValue);
                                }
                                else
                                {
                                    genericsControl.SearchProperties[HtmlHyperlink.PropertyNames.InnerText] = ExcelDataTable.ReadData(1, ControlKeywordValue);
                                }
                                Mouse.Click(genericsControl, Button);
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlCell.PropertyNames.Id] = ControlKeywordValue;
                                Mouse.Click(genericsControl, Button);
                                break;
                        }
                        return true;
                    case "HtmlEdit":
                        switch (ControlKeyword)
                        {
                            case "Class":
                                genericsControl.SearchProperties[HtmlEdit.PropertyNames.Class] = ControlKeywordValue;
                                Mouse.Click(genericsControl, Button);
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlEdit.PropertyNames.Id] = ControlKeywordValue;
                                Mouse.Click(genericsControl, Button);
                                break;
                        }
                        return true;
                    case "HtmlCustom":
                        switch (ControlKeyword)
                        {
                            case "InnerText":
                                genericsControl.SearchProperties[HtmlCustom.PropertyNames.InnerText] = ExcelDataTable.ReadData(1, ControlKeywordValue);
                                Mouse.Click(genericsControl, Button);
                                break;
                            case "Class":
                                genericsControl.SearchProperties[HtmlCustom.PropertyNames.Class] = ExcelDataTable.ReadData(1, ControlKeywordValue);
                                Mouse.Click(genericsControl, Button);
                                break;
                            case "IDData":
                                string[] IdData = ExcelDataTable.ReadData(1, ControlKeywordValue).Split('+');
                                if (IdData.Count() > 1)
                                {
                                    if (IdData[1].StartsWith("Rec_"))
                                    {
                                        ControlKeywordValue = IdData[0].Replace("&", Operation.recordedData[IdData[1].Replace("Rec_", string.Empty)]);
                                    }
                                    else if (ControlKeywordValue.Contains('='))
                                    {
                                        WriteAndReadData.DataFilePath = ConfigurationManager.AppSettings["ExcelDataFile"] + "\\DynamicData" + "\\" + Operation.Batch + ".xls";
                                        ExcelDataTable.PopulateRecordData(WriteAndReadData.DataFilePath);
                                        ControlKeywordValue = IdData[0].Replace("&", getDataFromDynamicExcel(ControlKeywordValue));
                                    }
                                    else
                                    {
                                        ControlKeywordValue = IdData[0].Replace("&", ExcelDataTable.ReadData(1, IdData[1]));
                                    }
                                }
                                else {
                                    throw new NullReferenceException("Reference Data not found after + Symbol");
                                }
                                genericsControl.SearchProperties[HtmlCustom.PropertyNames.Id] = ControlKeywordValue;
                                Mouse.Click(genericsControl, Button);
                                break;
                            case "IDVisible":
                                genericsControl.SearchProperties[HtmlCustom.PropertyNames.Id] = ControlKeywordValue;
                                bool d = genericsControl.TryGetClickablePoint(out p);
                                if (d)
                                {
                                    Mouse.Click(genericsControl, Button);
                                }
                                break;
                            case "TextVisible":
                                genericsControl.SearchProperties[HtmlCustom.PropertyNames.InnerText] = ExcelDataTable.ReadData(1, ControlKeywordValue);
                                d = genericsControl.TryGetClickablePoint(out p);
                                if (d)
                                {
                                    Mouse.Click(genericsControl, Button);
                                }
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlCustom.PropertyNames.Id] = ControlKeywordValue;
                                Mouse.Click(genericsControl, Button);
                                break;
                        }
                        return true;
                    case "HtmlList":
                        switch (ControlKeyword)
                        {
                            case "InnerText":
                                genericsControl.SearchProperties[HtmlList.PropertyNames.InnerText] = ExcelDataTable.ReadData(1, ControlKeywordValue);
                                Mouse.Click(genericsControl, Button);
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlList.PropertyNames.Id] = ControlKeywordValue;
                                Mouse.Click(genericsControl, Button);
                                break;
                        }
                        return true;
                    case "HtmlListItem":
                        switch (ControlKeyword)
                        {
                            case "InnerText":
                                genericsControl.SearchProperties[HtmlListItem.PropertyNames.InnerText] = ExcelDataTable.ReadData(1, ControlKeywordValue);
                                Mouse.Click(genericsControl, Button);
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlListItem.PropertyNames.Id] = ControlKeywordValue;
                                Mouse.Click(genericsControl, Button);
                                break;
                        }
                        return true;
                    case "HtmlHyperlink":
                        switch (ControlKeyword)
                        {
                            case "InnerText":
                                if (ControlKeywordValue.Contains('='))
                                {
                                    WriteAndReadData.DataFilePath = ConfigurationManager.AppSettings["ExcelDataFile"] + "\\DynamicData" + "\\" + LoginOperatrion.batchforReport + ".xls";
                                    ExcelDataTable.PopulateRecordData(WriteAndReadData.DataFilePath);
                                    genericsControl.SearchProperties[HtmlHyperlink.PropertyNames.InnerText] = getDataFromDynamicExcel(ControlKeywordValue);
                                }
                                else
                                {
                                    genericsControl.SearchProperties[HtmlHyperlink.PropertyNames.InnerText] = ExcelDataTable.ReadData(1, ControlKeywordValue);
                                }
                                Mouse.Click(genericsControl, Button);
                                break;
                            case "FriendlyName":
                                genericsControl.SearchProperties[HtmlHyperlink.PropertyNames.FriendlyName] = ExcelDataTable.ReadData(1, ControlKeywordValue);
                                Mouse.Click(genericsControl, Button);
                                break;
                            case "IDVisible":
                                genericsControl.SearchProperties[HtmlHyperlink.PropertyNames.Id] = ControlKeywordValue;
                                bool d = genericsControl.TryGetClickablePoint(out p);
                                if (d)
                                {
                                    Mouse.Click(genericsControl, Button);
                                }
                                break;
                            case "IDData":
                                string[] IdData = ExcelDataTable.ReadData(1, ControlKeywordValue).Split('+');
                                if (IdData.Count() > 1)
                                {
                                    if (IdData[1].StartsWith("Rec_"))
                                    {
                                        ControlKeywordValue = IdData[0].Replace("&", Operation.recordedData[IdData[1].Replace("Rec_", string.Empty)]);
                                    }
                                    else
                                    {
                                        ControlKeywordValue = IdData[0].Replace("&", ExcelDataTable.ReadData(1, IdData[1]));
                                    }
                                }
                                else {
                                    throw new NullReferenceException("Reference Data not found after + Symbol");
                                }
                                genericsControl.SearchProperties[HtmlHyperlink.PropertyNames.Id] = ControlKeywordValue;
                                Mouse.Click(genericsControl, Button);
                                break;
                            case "Class":
                                genericsControl.SearchProperties[HtmlHyperlink.PropertyNames.Class] = ControlKeywordValue;
                                Mouse.Click(genericsControl, Button);
                                break;
                            case "AbsolutePath":
                                genericsControl.SearchProperties[HtmlHyperlink.PropertyNames.AbsolutePath] = ControlKeywordValue;
                                d = genericsControl.TryGetClickablePoint(out p);
                                if (d)
                                {
                                    Mouse.Click(genericsControl, Button);
                                }
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlHyperlink.PropertyNames.Id] = ControlKeywordValue;
                                Mouse.Click(genericsControl, Button);
                                break;
                        }
                        return true;
                    case "HtmlButton":
                        switch (ControlKeyword)
                        {
                            case "InnerText":
                                genericsControl.SearchProperties[HtmlButton.PropertyNames.InnerText] = ExcelDataTable.ReadData(1, ControlKeywordValue);
                                Mouse.Click(genericsControl, Button);
                                break;
                            case "TextVisible":
                                genericsControl.SearchProperties[HtmlButton.PropertyNames.InnerText] = ExcelDataTable.ReadData(1, ControlKeywordValue);
                                bool d = genericsControl.TryGetClickablePoint(out p);
                                if (d)
                                {
                                    Mouse.Click(genericsControl, Button);
                                }
                                break;
                            case "IDVisible":
                                genericsControl.SearchProperties[HtmlButton.PropertyNames.Id] = ControlKeywordValue;
                                d = genericsControl.TryGetClickablePoint(out p);
                                if (d)
                                {
                                    Mouse.Click(genericsControl, Button);
                                }
                                break;
                            case "IDData":
                                string[] IdData = ExcelDataTable.ReadData(1, ControlKeywordValue).Split('+');
                                if (IdData.Count() > 1)
                                {
                                    if (IdData[1].StartsWith("Rec_"))
                                    {
                                        ControlKeywordValue = IdData[0].Replace("&", Operation.recordedData[IdData[1].Replace("Rec_", string.Empty)]);
                                    }
                                    else
                                    {
                                        ControlKeywordValue = IdData[0].Replace("&", ExcelDataTable.ReadData(1, IdData[1]));
                                    }
                                }
                                else {
                                    throw new NullReferenceException("Reference Data not found after + Symbol");
                                }
                                genericsControl.SearchProperties[HtmlButton.PropertyNames.Id] = ControlKeywordValue;
                                Mouse.Click(genericsControl, Button);
                                break;
                            case "Class":
                                genericsControl.SearchProperties[HtmlButton.PropertyNames.Class] = ControlKeywordValue;
                                Mouse.Click(genericsControl, Button);
                                break;
                            case "ClassVisible":
                                genericsControl.SearchProperties[HtmlButton.PropertyNames.Class] = ControlKeywordValue;
                                d = genericsControl.TryGetClickablePoint(out p);
                                if (d)
                                {
                                    Mouse.Click(genericsControl, Button);
                                }
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlButton.PropertyNames.Id] = ControlKeywordValue;
                                Mouse.Click(genericsControl, Button);
                                break;
                        }
                        return true;
                    case "HtmlCheckBox":
                        switch (ControlKeyword)
                        {
                            case "InnerText":
                                genericsControl.SearchProperties[HtmlCheckBox.PropertyNames.InnerText] = ExcelDataTable.ReadData(1, ControlKeywordValue);
                                Mouse.Click(genericsControl, Button);
                                break;
                            case "Name":
                                genericsControl.SearchProperties[HtmlCheckBox.PropertyNames.Name] = ExcelDataTable.ReadData(1, ControlKeywordValue);
                                Mouse.Click(genericsControl, Button);
                                break;
                            case "Value":
                                genericsControl.SearchProperties[HtmlCheckBox.PropertyNames.Value] = ControlKeywordValue;
                                Mouse.Click(genericsControl, Button);
                                break;
                            case "Label":
                                genericsControl.SearchProperties[HtmlCheckBox.PropertyNames.LabeledBy] = ExcelDataTable.ReadData(1, ControlKeywordValue);
                                Mouse.Click(genericsControl, Button);
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlCheckBox.PropertyNames.Id] = ControlKeywordValue;
                                Mouse.Click(genericsControl, Button);
                                break;
                        }
                        return true;
                    case "HtmlRadioButton":
                        switch (ControlKeyword)
                        {
                            case "InnerText":
                                genericsControl.SearchProperties[HtmlRadioButton.PropertyNames.InnerText] = ExcelDataTable.ReadData(1, ControlKeywordValue);
                                Mouse.Click(genericsControl, Button);
                                break;
                            case "Label":
                                genericsControl.SearchProperties[HtmlRadioButton.PropertyNames.LabeledBy] = ExcelDataTable.ReadData(1, ControlKeywordValue);
                                Mouse.Click(genericsControl, Button);
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlRadioButton.PropertyNames.Id] = ControlKeywordValue;
                                Mouse.Click(genericsControl, Button);
                                break;
                        }
                        return true;
                    case "HtmlLabel":
                        switch (ControlKeyword)
                        {
                            case "InnerText":
                                if (ControlKeywordValue.Contains('='))
                                {
                                    WriteAndReadData.DataFilePath = ConfigurationManager.AppSettings["ExcelDataFile"] + "\\DynamicData" + "\\" + LoginOperatrion.batchforReport + ".xls";
                                    ExcelDataTable.PopulateRecordData(WriteAndReadData.DataFilePath);
                                    genericsControl.SearchProperties[HtmlLabel.PropertyNames.InnerText] = getDataFromDynamicExcel(ControlKeywordValue);
                                }
                                else
                                {
                                    genericsControl.SearchProperties[HtmlLabel.PropertyNames.InnerText] = ExcelDataTable.ReadData(1, ControlKeywordValue);
                                }
                                Mouse.Click(genericsControl, Button);
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlLabel.PropertyNames.Id] = ControlKeywordValue;
                                Mouse.Click(genericsControl, Button);
                                break;
                        }
                        return true;
                    case "HtmlDiv":
                        switch (ControlKeyword)
                        {
                            case "InnerText":
                                genericsControl.SearchProperties[HtmlDiv.PropertyNames.InnerText] = ExcelDataTable.ReadData(1, ControlKeywordValue);
                                Mouse.Click(genericsControl, Button);
                                break;
                            case "FriendlyName":
                                genericsControl.SearchProperties[HtmlDiv.PropertyNames.FriendlyName] = ExcelDataTable.ReadData(1, ControlKeywordValue);
                                Mouse.Click(genericsControl, Button);
                                break;
                            case "DisplayText":
                                genericsControl.SearchProperties[HtmlDiv.PropertyNames.DisplayText] = ExcelDataTable.ReadData(1, ControlKeywordValue);
                                Mouse.Click(genericsControl, Button);
                                break;
                            case "IDVisible":
                                genericsControl.SearchProperties[HtmlHyperlink.PropertyNames.Id] = ControlKeywordValue;
                                bool d = genericsControl.TryGetClickablePoint(out p);
                                if (d)
                                {
                                    Mouse.Click(genericsControl, Button);
                                }
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlDiv.PropertyNames.Id] = ControlKeywordValue;
                                Mouse.Click(genericsControl, Button);
                                break;
                        }
                        return true;
                    case "HtmlImage":
                        switch (ControlKeyword)
                        {
                            case "ID":
                                genericsControl.SearchProperties[HtmlImage.PropertyNames.InnerText] = ControlKeywordValue;
                                Mouse.Click(genericsControl, Button);
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlImage.PropertyNames.Id] = ControlKeywordValue;
                                Mouse.Click(genericsControl, Button);
                                break;
                        }
                        return true;
                    case "HtmlHeaderCell":
                        switch (ControlKeyword)
                        {
                            case "ID":
                                genericsControl.SearchProperties[HtmlHeaderCell.PropertyNames.Id] = ControlKeywordValue;
                                Mouse.Click(genericsControl, Button);
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlHeaderCell.PropertyNames.Id] = ControlKeywordValue;
                                Mouse.Click(genericsControl, Button);
                                break;
                        }
                        return true;
                        Playback.PlaybackSettings.WaitForReadyLevel = WaitForReadyLevel.UIThreadOnly;
                        return true;
                    default:
                        Operation.FailerReason = TypeControl + " Control Type not configured for Clicking. Please check the Control Type.";
                        genDetailedReport.Reports(StepId, Description, Operationkeyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                        genDetailedReport.FileCorreptionCheck();
                        throw new NoSuchControlTypeFound(TypeControl + " Control Type not configured for Clicking. Please check the Control Type.");
                }
            }
            catch (UITestControlNotFoundException e)
            {
                genDetailedReport.Reports(StepId, Description, Operationkeyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                Operation.FailerReason = TypeControl + " with " + ControlKeywordValue + " as " + ControlKeyword + " not found in page";
                genDetailedReport.FileCorreptionCheck();
                TakeScreenshot(ControlKeywordValue + " as " + ControlKeyword + " not found");
                Assert.Fail(Operation.FailerReason);
                return false;
            }
            catch (FailedToPerformActionOnBlockedControlException e)
            {
                genericsControl.DrawHighlight();
                TakeScreenshot(ControlKeywordValue + " as " + ControlKeyword + " not found");
                genDetailedReport.Reports(StepId, Description, Operationkeyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                Operation.FailerReason = TypeControl + " with " + ControlKeywordValue + " as " + ControlKeyword + " is blocked by another control";
                genDetailedReport.FileCorreptionCheck();
                Assert.Fail(Operation.FailerReason);
                return false;
            }
            catch (FailedToPerformActionOnHiddenControlException ex)
            {
                genDetailedReport.Reports(StepId, Description, Operationkeyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                Operation.FailerReason = TypeControl + " with " + ControlKeywordValue + " as " + ControlKeyword + " not found in page - Expected : " + ControlKeywordValue + " as " + ControlKeyword + " Should be displyed ";
                genDetailedReport.FileCorreptionCheck();
                TakeScreenshot(ControlKeywordValue + " as " + ControlKeyword + " not found");
                Assert.Fail(Operation.FailerReason);
                return false;
            }
        }

        /// <summary>
        /// This used to click the control.that we can't directly click, so you can specify the another control nearby and move the pointer using coordinates and click the control.
        /// </summary>
        /// <param name="TypeControl">Specify for which Type of control you want to perform this action.</param>
        /// <param name="ControlKeyword">Property that you used to identify the control</param>
        /// <param name="ControlKeywordValue">What is the value of the property you specified</param>
        /// <param name="steps">Step Number.</param>
        /// <param name="description">Description about the step.</param>
        /// <param name="Operationkeyword">This function will be triggered using Keyword: "ClickEdit" and "ClickDelete"</param>
        /// <param name="codinates">specify the x and y coordinates to move to a control.</param>
        /// <returns>Return the true or false value.</returns>
        public bool WebClickOperation(string TypeControl, string ControlKeyword, string ControlKeywordValue, string StepId, string Description, string Operationkeyword, MouseButtons Button, params int[] codinates)
        {
            HtmlControl genericsControl = (HtmlControl)Activator.CreateInstance(typeof(HtmlControl), new object[] { ParentWindow });
            try
            {
                switch (TypeControl)
                {
                    case "HtmlSpan":
                        switch (ControlKeyword)
                        {
                            case "InnerText":
                                genericsControl.SearchProperties[HtmlSpan.PropertyNames.InnerText] = ControlKeywordValue;
                                int X = genericsControl.BoundingRectangle.X;
                                int Y = genericsControl.BoundingRectangle.Y;
                                auto.MouseMove(X + codinates[0], Y + codinates[1], 1000);
                                Mouse.Click(Button);
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlSpan.PropertyNames.Id] = ControlKeywordValue;
                                X = genericsControl.BoundingRectangle.X;
                                Y = genericsControl.BoundingRectangle.Y;
                                auto.MouseMove(X + codinates[0], Y + codinates[1], 1000);
                                Mouse.Click(Button);
                                break;
                        }
                        return true;
                    case "HtmlLabel":
                        switch (ControlKeyword)
                        {
                            case "InnerText":
                                genericsControl.SearchProperties[HtmlLabel.PropertyNames.InnerText] = ControlKeywordValue;
                                int X = genericsControl.BoundingRectangle.X;
                                int Y = genericsControl.BoundingRectangle.Y;
                                auto.MouseMove(X + codinates[0], Y + codinates[1], 1000);
                                Mouse.Click(Button);
                                break;
                            case "FriendlyName":
                                genericsControl.SearchProperties[HtmlLabel.PropertyNames.FriendlyName] = ControlKeywordValue;
                                X = genericsControl.BoundingRectangle.X;
                                Y = genericsControl.BoundingRectangle.Y;
                                auto.MouseMove(X + codinates[0], Y + codinates[1], 1000);
                                Mouse.Click(Button);
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlLabel.PropertyNames.Id] = ControlKeywordValue;
                                X = genericsControl.BoundingRectangle.X;
                                Y = genericsControl.BoundingRectangle.Y;
                                auto.MouseMove(X + codinates[0], Y + codinates[1], 1000);
                                Mouse.Click(Button);
                                break;
                        }
                        return true;
                    case "HtmlRow":
                        switch (ControlKeyword)
                        {
                            case "InnerText":
                                genericsControl.SearchProperties[HtmlRow.PropertyNames.InnerText] = ControlKeywordValue;
                                int X = genericsControl.BoundingRectangle.X;
                                int Y = genericsControl.BoundingRectangle.Y;
                                auto.MouseMove(X + codinates[0], Y + codinates[1], 1000);
                                Mouse.Click(Button);
                                break;
                            case "SearchText":
                                genericsControl.SearchProperties[HtmlRow.PropertyNames.InnerText] = ControlKeywordValue;
                                X = genericsControl.BoundingRectangle.X;
                                Y = genericsControl.BoundingRectangle.Y;
                                auto.MouseMove(X + codinates[0], Y + codinates[1], 1000);
                                Mouse.Click(Button);
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlRow.PropertyNames.Id] = ControlKeywordValue;
                                X = genericsControl.BoundingRectangle.X;
                                Y = genericsControl.BoundingRectangle.Y;
                                auto.MouseMove(X + codinates[0], Y + codinates[1], 1000);
                                Mouse.Click(Button);
                                break;
                        }
                        return true;
                    case "HtmlEdit":
                        switch (ControlKeyword)
                        {
                            case "InnerText":
                                genericsControl.SearchProperties[HtmlEdit.PropertyNames.InnerText] = ControlKeywordValue;
                                int X = genericsControl.BoundingRectangle.X;
                                int Y = genericsControl.BoundingRectangle.Y;
                                auto.MouseMove(X + codinates[0], Y + codinates[1], 1000);
                                Mouse.Click(Button);
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlEdit.PropertyNames.Id] = ControlKeywordValue;
                                X = genericsControl.BoundingRectangle.X;
                                Y = genericsControl.BoundingRectangle.Y;
                                auto.MouseMove(X + codinates[0], Y + codinates[1], 1000);
                                Mouse.Click(Button);
                                break;
                        }
                        return true;
                    case "HtmlCustom":
                        switch (ControlKeyword)
                        {
                            case "InnerText":
                                genericsControl.SearchProperties[HtmlCustom.PropertyNames.InnerText] = ControlKeywordValue;
                                int X = genericsControl.BoundingRectangle.X;
                                int Y = genericsControl.BoundingRectangle.Y;
                                auto.MouseMove(X + codinates[0], Y + codinates[1], 1000);
                                Mouse.Click(Button);
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlCustom.PropertyNames.Id] = ControlKeywordValue;
                                X = genericsControl.BoundingRectangle.X;
                                Y = genericsControl.BoundingRectangle.Y;
                                auto.MouseMove(X + codinates[0], Y + codinates[1], 1000);
                                Mouse.Click(Button);
                                break;
                        }
                        return true;
                    case "HtmlListItem":
                        switch (ControlKeyword)
                        {
                            case "InnerText":
                                genericsControl.SearchProperties[HtmlListItem.PropertyNames.InnerText] = ControlKeywordValue;
                                int X = genericsControl.BoundingRectangle.X;
                                int Y = genericsControl.BoundingRectangle.Y;
                                auto.MouseMove(X + codinates[0], Y + codinates[1], 1000);
                                Mouse.Click(Button);
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlListItem.PropertyNames.Id] = ControlKeywordValue;
                                X = genericsControl.BoundingRectangle.X;
                                Y = genericsControl.BoundingRectangle.Y;
                                auto.MouseMove(X + codinates[0], Y + codinates[1], 1000);
                                Mouse.Click(Button);
                                break;
                        }
                        return true;
                    case "HtmlHyperlink":
                        switch (ControlKeyword)
                        {
                            case "InnerText":
                                genericsControl.SearchProperties[HtmlHyperlink.PropertyNames.InnerText] = ControlKeywordValue;
                                int X = genericsControl.BoundingRectangle.X;
                                int Y = genericsControl.BoundingRectangle.Y;
                                auto.MouseMove(X + codinates[0], Y + codinates[1], 1000);
                                Mouse.Click(Button);
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlHyperlink.PropertyNames.Id] = ControlKeywordValue;
                                X = genericsControl.BoundingRectangle.X;
                                Y = genericsControl.BoundingRectangle.Y;
                                auto.MouseMove(X + codinates[0], Y + codinates[1], 1000);
                                Mouse.Click(Button);
                                break;
                        }
                        return true;
                    case "HtmlButton":
                        switch (ControlKeyword)
                        {
                            case "InnerText":
                                genericsControl.SearchProperties[HtmlButton.PropertyNames.InnerText] = ControlKeywordValue;
                                int X = genericsControl.BoundingRectangle.X;
                                int Y = genericsControl.BoundingRectangle.Y;
                                auto.MouseMove(X + codinates[0], Y + codinates[1], 1000);
                                Mouse.Click(Button);
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlButton.PropertyNames.Id] = ControlKeywordValue;
                                X = genericsControl.BoundingRectangle.X;
                                Y = genericsControl.BoundingRectangle.Y;
                                auto.MouseMove(X + codinates[0], Y + codinates[1], 1000);
                                Mouse.Click(Button);
                                break;
                        }
                        return true;
                    default:
                        Operation.FailerReason = TypeControl + " Control Type not configured for Click. Please check the Control Type.";
                        genDetailedReport.Reports(StepId, Description, Operationkeyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                        genDetailedReport.FileCorreptionCheck();
                        throw new NoSuchControlTypeFound(TypeControl + " Control Type not configured for Click. Please check the Control Type.");
                }
            }
            catch (UITestControlNotFoundException e)
            {
                if (ControlKeyword != "SearchText")
                {
                    genDetailedReport.Reports(StepId, Description, Operationkeyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                    Operation.FailerReason = TypeControl + " with " + ControlKeywordValue + " as " + ControlKeyword + " not found in page";
                    genDetailedReport.FileCorreptionCheck();
                    TakeScreenshot(ControlKeywordValue + " as " + ControlKeyword + " not found");
                    Assert.Fail(Operation.FailerReason);
                    return false;
                }
                else
                {
                    genDetailedReport.Reports(StepId, Description, "SearchDataOperation", false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                    Operation.FailerReason = " ";
                    Assert.Fail(Operation.FailerReason);
                    return false;
                }
            }
        }

        /// <summary>
        /// This used to click the control.that we can't directly click, so you can specify the another control nearby and move the pointer using coordinates and click the control.
        /// </summary>
        /// <param name="TypeControl">Specify for which Type of control you want to perform this action.</param>
        /// <param name="ControlKeyword">Property that you used to identify the control</param>
        /// <param name="ControlKeywordValue">What is the value of the property you specified</param>
        /// <param name="steps">Step Number.</param>
        /// <param name="description">Description about the step.</param>
        /// <param name="Operationkeyword">This function will be triggered using Keyword: "DClickEdit" and "DClickDelete"</param>
        /// <param name="codinates">specify the x and y coordinates to move to a control.</param>
        /// <returns>Return the true or false value.</returns>
        public bool WebDoubleClickOperation(string TypeControl, string ControlKeyword, string ControlKeywordValue, string StepId, string Description, string Operationkeyword, MouseButtons Button, params int[] codinates)
        {
            HtmlControl genericsControl = (HtmlControl)Activator.CreateInstance(typeof(HtmlControl), new object[] { ParentWindow });
            try
            {
                switch (TypeControl)
                {
                    case "HtmlHyperlink":
                        switch (ControlKeyword)
                        {
                            case "InnerText":
                                genericsControl.SearchProperties[HtmlHyperlink.PropertyNames.InnerText] = ControlKeywordValue;
                                int X = genericsControl.BoundingRectangle.X;
                                int Y = genericsControl.BoundingRectangle.Y;
                                auto.MouseMove(X + codinates[0], Y + codinates[1], 1000);
                                Mouse.DoubleClick(Button);
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlHyperlink.PropertyNames.Id] = ControlKeywordValue;
                                X = genericsControl.BoundingRectangle.X;
                                Y = genericsControl.BoundingRectangle.Y;
                                auto.MouseMove(X + codinates[0], Y + codinates[1], 1000);
                                Mouse.DoubleClick(Button);
                                break;
                        }
                        return true;
                    default:
                        Operation.FailerReason = TypeControl + " Control Type not configured for Double Click. Please check the Control Type.";
                        genDetailedReport.Reports(StepId, Description, Operationkeyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                        genDetailedReport.FileCorreptionCheck();
                        throw new NoSuchControlTypeFound(TypeControl + " Control Type not configured for Double Click. Please check the Control Type.");
                }
            }
            catch (UITestControlNotFoundException e)
            {
                genDetailedReport.Reports(StepId, Description, "SearchDataOperation", false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                Operation.FailerReason = TypeControl + " with " + ControlKeywordValue + " as " + ControlKeyword + " not found in page";
                genDetailedReport.FileCorreptionCheck();
                TakeScreenshot(ControlKeywordValue + " as " + ControlKeyword + " not found");
                Assert.Fail(Operation.FailerReason);
                return false;
            }
        }

        /// <summary>
        /// This will get the data of the control and return the data and validate from the resource file.
        /// </summary>
        /// <param name="TypeControl">Specify for which Type of control you want to perform this action.</param>
        /// <param name="ControlKeyword">Property that you used to identify the control</param>
        /// <param name="ControlKeywordValue">What is the value of the property you specified</param>
        /// <param name="ScreenshotName">Here screen shot name will come.</param>
        /// <param name="AssertionMessage">message show when test execution fail</param>
        /// <param name="steps">Step Number.</param>
        /// <param name="description">Description about the step.</param>
        /// <param name="Operationkeyword">This function will be triggered using Keyword: "Validate", "ValidateNotEqual", "RecordData"</param>
        /// <returns>return the specified data.</returns>
        public string WebGetControlData(string TypeControl, string ControlKeyword, string ControlKeywordValue, string ScreenshotName, string AssertionMessage, string StepId, string Description, string Operationkeyword)
        {
            try
            {
                HtmlControl genericsControl = (HtmlControl)Activator.CreateInstance(typeof(HtmlControl), new object[] { ParentWindow });
                Point p = new Point();
                Thread.Sleep(2000);

                switch (TypeControl)
                {
                    case "HtmlSpan":
                        switch (ControlKeyword)
                        {
                            case "ID":
                                genericsControl.SearchProperties[HtmlSpan.PropertyNames.Id] = ControlKeywordValue;
                                return genericsControl.GetProperty("innerHTML").ToString();
                            case "IDData":
                                string[] IdData = ExcelDataTable.ReadData(1, ControlKeywordValue).Split('+');
                                if (IdData.Count() > 1)
                                {
                                    if (IdData[1].StartsWith("Rec_"))
                                    {
                                        ControlKeywordValue = IdData[0].Replace("&", Operation.recordedData[IdData[1].Replace("Rec_", string.Empty)]);
                                    }
                                    if (IdData[1].Contains('='))
                                    {
                                        ExcelDataTable.PopulateRecordData(WriteAndReadData.DataFilePath);
                                        ControlKeywordValue = IdData[0].Replace("&", getDataFromDynamicExcel(IdData[1]));
                                    }
                                    else
                                    {
                                        ControlKeywordValue = IdData[0].Replace("&", ExcelDataTable.ReadData(1, IdData[1]));
                                    }
                                }
                                else {
                                    throw new NullReferenceException("Reference Data not found after + Symbol");
                                }
                                genericsControl.SearchProperties[HtmlSpan.PropertyNames.Id] = ControlKeywordValue;
                                return genericsControl.GetProperty("innerHTML").ToString();
                            case "Class":
                                genericsControl.SearchProperties[HtmlSpan.PropertyNames.Class] = ControlKeywordValue;
                                return genericsControl.GetProperty("innerHTML").ToString();
                            default:
                                return null;
                        }
                    case "HtmlInputButton":
                        switch (ControlKeyword)
                        {
                            case "ID":
                                genericsControl.SearchProperties[HtmlInputButton.PropertyNames.Id] = ControlKeywordValue;
                                return genericsControl.GetProperty("ValueAttribute").ToString();
                            case "TagInstance":
                                genericsControl.SearchProperties[HtmlInputButton.PropertyNames.TagInstance] = ControlKeywordValue;
                                return genericsControl.GetProperty("Class").ToString();
                            default:
                                return null;
                        }
                    case "HtmlButton":
                        switch (ControlKeyword)
                        {
                            case "IDEnabled":
                                genericsControl.SearchProperties[HtmlButton.PropertyNames.Id] = ControlKeywordValue;
                                return genericsControl.GetProperty("Enabled").ToString();
                            case "ClassEnabled":
                                genericsControl.SearchProperties[HtmlButton.PropertyNames.Class] = ControlKeywordValue;
                                return genericsControl.GetProperty("Enabled").ToString();
                            default:
                                return null;
                        }
                    case "HtmlEdit":
                        switch (ControlKeyword)
                        {
                            case "ID":
                                genericsControl.SearchProperties[HtmlEdit.PropertyNames.Id] = ControlKeywordValue;
                                return genericsControl.GetProperty("ValueAttribute").ToString();
                            case "Class":
                                genericsControl.SearchProperties[HtmlEdit.PropertyNames.Class] = ControlKeywordValue;
                                return genericsControl.GetProperty("ValueAttribute").ToString();
                            default:
                                return null;
                        }
                    case "HtmlCustom":
                        switch (ControlKeyword)
                        {
                            case "ID":
                                genericsControl.SearchProperties[HtmlCustom.PropertyNames.Id] = ControlKeywordValue;
                                return genericsControl.GetProperty("innerHTML").ToString();
                            case "IDData":
                                string[] IdData = ExcelDataTable.ReadData(1, ControlKeywordValue).Split('+');
                                if (IdData.Count() > 1)
                                {
                                    if (IdData[1].StartsWith("Rec_"))
                                    {
                                        ControlKeywordValue = IdData[0].Replace("&", Operation.recordedData[IdData[1].Replace("Rec_", string.Empty)]);
                                    }
                                    if (IdData[1].Contains('='))
                                    {
                                        ExcelDataTable.PopulateRecordData(WriteAndReadData.DataFilePath);
                                        ControlKeywordValue = IdData[0].Replace("&", getDataFromDynamicExcel(IdData[1]));
                                    }
                                    else
                                    {
                                        ControlKeywordValue = IdData[0].Replace("&", ExcelDataTable.ReadData(1, IdData[1]));
                                    }
                                }
                                else {
                                    throw new NullReferenceException("Reference Data not found after + Symbol");
                                }
                                genericsControl.SearchProperties[HtmlCustom.PropertyNames.Id] = ControlKeywordValue;
                                return genericsControl.GetProperty("innerHTML").ToString();
                            case "IDClass":
                                genericsControl.SearchProperties[HtmlCustom.PropertyNames.Id] = ControlKeywordValue;
                                return genericsControl.GetProperty("class").ToString();
                            case "Class":
                                genericsControl.SearchProperties[HtmlCustom.PropertyNames.Class] = ControlKeywordValue;
                                return genericsControl.GetProperty("innerText").ToString();
                            default:
                                return null;
                        }
                    case "HtmlTable":
                        switch (ControlKeyword)
                        {
                            case "ID":
                                genericsControl.SearchProperties[HtmlTable.PropertyNames.Id] = ControlKeywordValue;
                                return genericsControl.GetProperty("innerHTML").ToString();
                            case "GetTableData":
                                return ScriptGetExcuite(Operationkeyword, ControlKeyword, ControlKeywordValue, StepId, Description, AssertionMessage);
                            case "IFrameTblData":
                                return ScriptGetExcuite(Operationkeyword, ControlKeyword, ControlKeywordValue, StepId, Description, AssertionMessage);
                            case "IDExists":
                                genericsControl.SearchProperties[HtmlTable.PropertyNames.Id] = ControlKeywordValue;
                                return genericsControl.WaitForControlExist().ToString();
                            default:
                                return null;
                        }
                    case "HtmlLabel":
                        switch (ControlKeyword)
                        {
                            case "Class":
                                genericsControl.SearchProperties[HtmlLabel.PropertyNames.Class] = ControlKeywordValue;
                                return genericsControl.GetProperty("innerHTML").ToString();
                            case "IDData":
                                string[] IdData = ExcelDataTable.ReadData(1, ControlKeywordValue).Split('+');
                                if (IdData.Count() > 1)
                                {
                                    if (IdData[1].StartsWith("Rec_"))
                                    {
                                        ControlKeywordValue = IdData[0].Replace("&", Operation.recordedData[IdData[1].Replace("Rec_", string.Empty)]);
                                    }
                                    else
                                    {
                                        ControlKeywordValue = IdData[0].Replace("&", ExcelDataTable.ReadData(1, IdData[1]));
                                    }
                                }
                                else {
                                    throw new NullReferenceException("Reference Data not found after + Symbol");
                                }
                                genericsControl.SearchProperties[HtmlLabel.PropertyNames.Id] = ControlKeywordValue;
                                return genericsControl.GetProperty("innerHTML").ToString();
                            default:
                                return null;
                        }
                    case "HtmlDiv":
                        switch (ControlKeyword)
                        {
                            case "ID":
                                genericsControl.SearchProperties[HtmlDiv.PropertyNames.Id] = ControlKeywordValue;
                                return genericsControl.GetProperty("innerHTML").ToString();
                            case "ClassVisible":
                                genericsControl.SearchProperties[HtmlDiv.PropertyNames.Class] = ControlKeywordValue;
                                bool d = genericsControl.TryGetClickablePoint(out p);
                                if (d == true)
                                {
                                    return genericsControl.GetProperty("innerHTML").ToString();
                                }
                                else
                                {
                                    return string.Empty;
                                }
                            case "IDData":
                                string[] IdData = ExcelDataTable.ReadData(1, ControlKeywordValue).Split('+');
                                if (IdData.Count() > 1)
                                {
                                    if (IdData[1].StartsWith("Rec_"))
                                    {
                                        ControlKeywordValue = IdData[0].Replace("&", Operation.recordedData[IdData[1].Replace("Rec_", string.Empty)]);
                                    }
                                    if (IdData[1].Contains('='))
                                    {
                                        ExcelDataTable.PopulateRecordData(WriteAndReadData.DataFilePath);
                                        ControlKeywordValue = IdData[0].Replace("&", getDataFromDynamicExcel(IdData[1]));
                                    }
                                    else
                                    {
                                        ControlKeywordValue = IdData[0].Replace("&", ExcelDataTable.ReadData(1, IdData[1]));
                                    }
                                }
                                else {
                                    throw new NullReferenceException("Reference Data not found after + Symbol");
                                }
                                genericsControl.SearchProperties[HtmlDiv.PropertyNames.Id] = ControlKeywordValue;
                                return genericsControl.GetProperty("innerHTML").ToString();
                            case "Class":
                                genericsControl.SearchProperties[HtmlDiv.PropertyNames.Class] = ControlKeywordValue;
                                return genericsControl.GetProperty("innerHTML").ToString();
                            
                            default:
                                return null;
                        }
                    case "HtmlImage":
                        switch (ControlKeyword)
                        {
                            case "ID":
                                genericsControl.SearchProperties[HtmlImage.PropertyNames.Id] = ControlKeywordValue;
                                return genericsControl.GetProperty("src").ToString();
                            case "IDData":
                                string[] IdData = ExcelDataTable.ReadData(1, ControlKeywordValue).Split('+');
                                if (IdData.Count() > 1)
                                {
                                    if (IdData[1].StartsWith("Rec_"))
                                    {
                                        ControlKeywordValue = IdData[0].Replace("&", Operation.recordedData[IdData[1].Replace("Rec_", string.Empty)]);
                                    }
                                    if (IdData[1].Contains('='))
                                    {
                                        ExcelDataTable.PopulateRecordData(WriteAndReadData.DataFilePath);
                                        ControlKeywordValue = IdData[0].Replace("&", getDataFromDynamicExcel(IdData[1]));
                                    }
                                    else
                                    {
                                        ControlKeywordValue = IdData[0].Replace("&", ExcelDataTable.ReadData(1, IdData[1]));
                                    }
                                }
                                else {
                                    throw new NullReferenceException("Reference Data not found after + Symbol");
                                }
                                genericsControl.SearchProperties[HtmlImage.PropertyNames.Id] = ControlKeywordValue;
                                return genericsControl.GetProperty("src").ToString();
                            case "IDVisible":
                                genericsControl.SearchProperties[HtmlImage.PropertyNames.Id] = ControlKeywordValue;
                                if (genericsControl.Exists)
                                { return genericsControl.GetProperty("src").ToString(); }
                                else
                                { return "Control is not visible"; }
                            case "ClassVisible":
                                genericsControl.SearchProperties[HtmlImage.PropertyNames.Class] = ControlKeywordValue;
                                var im = genericsControl.BoundingRectangle;
                                bool result = im.X != 0 ? true : false;
                                return result.ToString();
                            case "TagInstance":
                                genericsControl.SearchProperties[HtmlImage.PropertyNames.TagInstance] = ControlKeywordValue;
                                return genericsControl.GetProperty("src").ToString();
                            default:
                                return null;
                        }
                    case "HtmlListItem":
                        switch (ControlKeyword)
                        {
                            case "ID":
                                genericsControl.SearchProperties[HtmlListItem.PropertyNames.Id] = ControlKeywordValue;
                                return genericsControl.GetProperty("class").ToString();
                            default:
                                return null;
                        }
                    case "HtmlComboBox":
                        switch (ControlKeyword)
                        {
                            case "ID":
                                genericsControl.SearchProperties[HtmlComboBox.PropertyNames.Id] = ControlKeywordValue;
                                return genericsControl.InnerText;
                            case "IDSelected":
                                genericsControl.SearchProperties[HtmlComboBox.PropertyNames.Id] = ControlKeywordValue;
                                return genericsControl.GetProperty("SelectedItem").ToString();
                            default:
                                return null;
                        }
                    case "HtmlCell":
                        switch (ControlKeyword)
                        {
                            case "ID":
                                genericsControl.SearchProperties[HtmlCell.PropertyNames.Id] = ControlKeywordValue;
                                return genericsControl.InnerText;
                            case "TagInstance":
                                genericsControl.SearchProperties[HtmlCell.PropertyNames.TagInstance] = ControlKeywordValue;
                                return genericsControl.InnerText;
                            default:
                                return null;
                        }
                    case "HtmlHyperlink":
                        switch (ControlKeyword)
                        {
                            case "ID":
                                genericsControl.SearchProperties[HtmlHyperlink.PropertyNames.Id] = ControlKeywordValue;
                                return genericsControl.InnerText;
                            case "IDVisible":
                                genericsControl.SearchProperties[HtmlHyperlink.PropertyNames.Id] = ControlKeywordValue;
                                if (genericsControl.Exists)
                                { return genericsControl.InnerText; }
                                else
                                { return string.Empty; }
                            default:
                                return null;
                        }
                    case "HtmlCheckBox":
                        switch (ControlKeyword)
                        {
                            case "Label":
                                genericsControl.SearchProperties[HtmlCheckBox.PropertyNames.LabeledBy] = ExcelDataTable.ReadData(1, ControlKeywordValue);
                                return genericsControl.GetProperty("Checked").ToString();
                            case "ID":
                                genericsControl.SearchProperties[HtmlCheckBox.PropertyNames.Id] = ControlKeywordValue;
                                return genericsControl.GetProperty("Checked").ToString();
                            default:
                                return null;
                        }
                    default:
                        Operation.FailerReason = TypeControl + " Control Type not configured to Fetch data from control. Please check the Control Type.";
                        genDetailedReport.Reports(StepId, Description, Operationkeyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                        genDetailedReport.FileCorreptionCheck();
                        throw new NoSuchControlTypeFound(TypeControl + " Control Type not configured to Fetch data from control. Please check the Control Type.");
                }
            }
            catch (UITestControlNotFoundException e)
            {
                Operation.FailerReason = e.BasicMessage + AssertionMessage;
                string path = ConfigurationManager.AppSettings["ScreenShot"];
                path = path + @"\" + LoginOperatrion.ProjectName + "";
                Directory.CreateDirectory(path);
                try
                {
                    Image image = window.CaptureImage();
                    image.Save(path + "\\" + ScreenshotName + ".jpeg", ImageFormat.Jpeg);
                    Operation.ErrorScreenPath = path + "\\" + ScreenshotName + ".jpeg";
                    image.Dispose();
                }
                catch (Exception v) { }
                genDetailedReport.Reports(StepId, Description, Operationkeyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genDetailedReport.FileCorreptionCheck();
                Assert.Fail(e.BasicMessage + AssertionMessage);
                return null;
            }
            catch (Exception exp)
            {
                Operation.FailerReason = "Control not visible " + AssertionMessage;
                string path = ConfigurationManager.AppSettings["ScreenShot"];
                path = path + @"\" + LoginOperatrion.ProjectName + "";
                Directory.CreateDirectory(path);
                try
                {
                    Image image = window.CaptureImage();
                    image.Save(path + "\\" + ScreenshotName + ".jpeg", ImageFormat.Jpeg);
                    Operation.ErrorScreenPath = path + "\\" + ScreenshotName + ".jpeg";
                    image.Dispose();
                }
                catch (Exception v) { }
                genDetailedReport.Reports(StepId, Description, Operationkeyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genDetailedReport.FileCorreptionCheck();
                Assert.Fail("Control not visible " + AssertionMessage);
                return null;
            }

        }

        /// <summary>
        /// Wait for the control ready. 
        /// </summary>
        /// <param name="TypeControl">Specify for which Type of control you want to perform this action.</param>
        /// <param name="ControlKeyword">Property that you used to identify the control</param>
        /// <param name="ControlKeywordValue">What is the value of the property you specified</param>
        /// <param name="steps">Step Number.</param>
        /// <param name="description">Description about the step.</param>
        /// <param name="Operationkeyword">This function will be triggered using Keyword: "Wait" </param>
        /// <returns>return true or false value</returns>
        public bool WebWaitForControl(string TypeControl, string ControlKeyword, string ControlKeywordValue, string StepId, string Description, string Operationkeyword, string DataRefferencekeyword)
        {
            HtmlControl genericsControl = (HtmlControl)Activator.CreateInstance(typeof(HtmlControl), new object[] { ParentWindow });
            Thread.Sleep(1000);
            try
            {
                switch (TypeControl)
                {
                    case "HtmlCustom":
                        switch (ControlKeyword)
                        {
                            case "ID":
                                genericsControl.SearchProperties[HtmlCustom.PropertyNames.Id] = ControlKeywordValue;
                                genericsControl.WaitForControlExist();
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlCustom.PropertyNames.Id] = ControlKeywordValue;
                                genericsControl.WaitForControlExist();
                                break;
                        }
                        return true;
                    case "HtmlDiv":
                        switch (ControlKeyword)
                        {
                            case "ID":
                                genericsControl.SearchProperties[HtmlDiv.PropertyNames.Id] = ControlKeywordValue;
                                genericsControl.WaitForControlExist();
                                break;
                            case "Class":
                                genericsControl.SearchProperties[HtmlDiv.PropertyNames.Class] = ControlKeywordValue;
                                genericsControl.WaitForControlExist();
                                break;
                            case "IsNotEmpty":
                                genericsControl.SearchProperties[HtmlDiv.PropertyNames.Id] = ControlKeywordValue;
                                genericsControl.WaitForControlPropertyNotEqual("InnerText", string.Empty);
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlDiv.PropertyNames.Id] = ControlKeywordValue;
                                genericsControl.WaitForControlExist();
                                break;
                        }
                        return true;
                    case "HtmlSpan":
                        switch (ControlKeyword)
                        {
                            case "IDPropertyEqual":
                                genericsControl.SearchProperties[HtmlSpan.PropertyNames.Id] = ControlKeywordValue;
                                string property = DataRefferencekeyword.Split('=')[0];
                                string propertyValue = DataRefferencekeyword.Split('=')[1];
                                if (propertyValue.Contains(':'))
                                {
                                    CommonLanguageTemplateReader.Message(Operation.lang, propertyValue);
                                    propertyValue = languageResource.Msg_GetTemplateMessage;
                                }
                                else
                                {
                                    propertyValue = ExcelDataTable.ReadData(1, propertyValue);
                                }
                                genericsControl.WaitForControlPropertyEqual(property,propertyValue);
                                break;
                            case "IDData":
                                string[] IdData = ExcelDataTable.ReadData(1, ControlKeywordValue).Split('+');
                                if (IdData.Count() > 1)
                                {
                                    if (IdData[1].StartsWith("Rec_"))
                                    {
                                        ControlKeywordValue = IdData[0].Replace("&", Operation.recordedData[IdData[1].Replace("Rec_", string.Empty)]);
                                    }
                                    else
                                    {
                                        ControlKeywordValue = IdData[0].Replace("&", ExcelDataTable.ReadData(1, IdData[1]));
                                    }
                                }
                                else {
                                    throw new NullReferenceException("Reference Data not found after + Symbol");
                                }
                                genericsControl.SearchProperties[HtmlSpan.PropertyNames.Id] = ControlKeywordValue;
                                genericsControl.WaitForControlExist();
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlSpan.PropertyNames.Id] = ControlKeywordValue;
                                genericsControl.WaitForControlExist();
                                break;
                        }
                        return true;
                    case "HtmlEdit":
                        switch (ControlKeyword)
                        {
                            case "IsNotEmpty":
                                genericsControl.SearchProperties[HtmlEdit.PropertyNames.Id] = ControlKeywordValue;
                                genericsControl.WaitForControlPropertyNotEqual("ValueAttribute", string.Empty);
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlEdit.PropertyNames.Id] = ControlKeywordValue;
                                genericsControl.WaitForControlExist();
                                break;
                        }
                        return true;
                    case "HtmlButton":
                        switch (ControlKeyword)
                        {
                            case "ID":
                                genericsControl.SearchProperties[HtmlEdit.PropertyNames.Id] = ControlKeywordValue;
                                genericsControl.WaitForControlExist();
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlEdit.PropertyNames.Id] = ControlKeywordValue;
                                genericsControl.WaitForControlExist();
                                break;
                        }
                        return true;
                    case "HtmlComboBox":
                        switch (ControlKeyword)
                        {
                            case "ID":
                                genericsControl.SearchProperties[HtmlComboBox.PropertyNames.Id] = ControlKeywordValue;
                                genericsControl.WaitForControlExist();
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlComboBox.PropertyNames.Id] = ControlKeywordValue;
                                genericsControl.WaitForControlExist();
                                break;
                        }
                        return true;
                    case "HtmlTable":
                        switch (ControlKeyword)
                        {
                            case "ID":
                                genericsControl.SearchProperties[HtmlTable.PropertyNames.Id] = ControlKeywordValue;
                                genericsControl.WaitForControlReady();
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlTable.PropertyNames.Id] = ControlKeywordValue;
                                genericsControl.WaitForControlReady();
                                break;
                        }
                        return true;
                    case "HtmlInputButton":
                        switch (ControlKeyword)
                        {
                            case "ID":
                                genericsControl.SearchProperties[HtmlInputButton.PropertyNames.Id] = ControlKeywordValue;
                                genericsControl.WaitForControlExist();
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlInputButton.PropertyNames.Id] = ControlKeywordValue;
                                genericsControl.WaitForControlExist();
                                break;
                        }
                        return true;
                    case "HtmlTextArea":
                        switch (ControlKeyword)
                        {
                            case "ID":
                                genericsControl.SearchProperties[HtmlTextArea.PropertyNames.Id] = ControlKeywordValue;
                                genericsControl.WaitForControlExist();
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlTextArea.PropertyNames.Id] = ControlKeywordValue;
                                genericsControl.WaitForControlExist();
                                break;
                        }
                        return true;
                    case "HtmlHyperlink":
                        switch (ControlKeyword)
                        {
                            case "InnerText":
                                genericsControl.SearchProperties[HtmlHyperlink.PropertyNames.InnerText] = ControlKeywordValue;
                                genericsControl.WaitForControlExist();
                                break;
                            case "Class":
                                genericsControl.SearchProperties[HtmlHyperlink.PropertyNames.Class] = ControlKeywordValue;
                                genericsControl.WaitForControlExist();
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlHyperlink.PropertyNames.Id] = ControlKeywordValue;
                                genericsControl.WaitForControlExist();
                                break;
                        }
                        return true;
                    default:
                        Operation.FailerReason = TypeControl + " Control Type not configured wait for control. Please check the Control Type.";
                        genDetailedReport.Reports(StepId, Description, Operationkeyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                        throw new NoSuchControlTypeFound(TypeControl + " Control Type not configured wait for control. Please check the Control Type.");
                }
            }
            catch (Exception e)
            {
                genDetailedReport.Reports(StepId, Description, Operationkeyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                return false;
            }
        }

        /// <summary>
        /// Clear the data in the control 
        /// </summary>
        /// <param name="TypeControl">Specify for which Type of control you want to perform this action.</param>
        /// <param name="ControlKeyword">Property that you used to identify the control</param>
        /// <param name="ControlKeywordValue">What is the value of the property you specified</param>
        /// <param name="steps">Step Number.</param>
        /// <param name="description">Description about the step.</param>
        /// <param name="Operationkeyword">This function will be triggered using Keyword: "Clear" </param>
        /// <returns></returns>
        public bool WebClearControlData(string TypeControl, string ControlKeyword, string ControlKeywordValue, string StepId, string Description, string Operationkeyword)
        {
            HtmlControl genericsControl = (HtmlControl)Activator.CreateInstance(typeof(HtmlControl), new object[] { ParentWindow });
            try
            {
                switch (TypeControl)
                {
                    case "HtmlEdit":
                        switch (ControlKeyword)
                        {
                            case "ID":
                                genericsControl.SearchProperties[HtmlEdit.PropertyNames.Id] = ControlKeywordValue;
                                Mouse.Click(genericsControl);
                                auto.Send("{BACKSPACE 20}");
                                auto.Send("{DEL 20}");
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlEdit.PropertyNames.Id] = ControlKeywordValue;
                                Mouse.Click(genericsControl);
                                auto.Send("{BACKSPACE 20}");
                                auto.Send("{DEL 20}");
                                break;
                        }
                        return true;
                    //break;
                    case "HtmlTextArea":
                        switch (ControlKeyword)
                        {
                            case "ID":
                                genericsControl.SearchProperties[HtmlEdit.PropertyNames.Id] = ControlKeywordValue;
                                Mouse.Click(genericsControl);
                                auto.Send("{BACKSPACE 30}");
                                auto.Send("{DEL 30}");
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlEdit.PropertyNames.Id] = ControlKeywordValue;
                                Mouse.Click(genericsControl);
                                auto.Send("{BACKSPACE 30}");
                                auto.Send("{DEL 30}");
                                break;
                        }
                        return true;
                    //break;
                    default:
                        Operation.FailerReason = TypeControl + " Control Type not configured to clear data. Please check the Control Type.";
                        genDetailedReport.Reports(StepId, Description, Operationkeyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                        genDetailedReport.FileCorreptionCheck();
                        throw new NoSuchControlTypeFound(TypeControl + " Control Type not configured to clear data. Please check the Control Type.");
                }
            }
            catch (UITestControlNotFoundException e)
            {
                genDetailedReport.Reports(StepId, Description, Operationkeyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                Operation.FailerReason = TypeControl + " with " + ControlKeywordValue + " as " + ControlKeyword + " not found in page";
                genDetailedReport.FileCorreptionCheck();
                TakeScreenshot(ControlKeywordValue + " as " + ControlKeyword + " not found");
                Assert.Fail(Operation.FailerReason);
                return false;
            }
            catch (FailedToPerformActionOnBlockedControlException e)
            {
                genericsControl.DrawHighlight();
                TakeScreenshot("ControlCannotBeClicked");
                genDetailedReport.Reports(StepId, Description, Operationkeyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                Operation.FailerReason = TypeControl + " with " + ControlKeywordValue + " as " + ControlKeyword + " is blocked by another control";
                genDetailedReport.FileCorreptionCheck();
                Assert.Fail(Operation.FailerReason);
                return false;
            }

        }

        /// <summary>
        /// This is used to sleep execution for some time.
        /// </summary>
        /// <param name="data"></param>
        public void GotoSleep(string data)
        {
            switch (data)
            {
                case "Maximum":
                    Thread.Sleep(LoginOperatrion.max);
                    break;
                case "Medium":
                    Thread.Sleep(LoginOperatrion.mid);
                    break;
                case "Minimum":
                    Thread.Sleep(LoginOperatrion.min);
                    break;
                default:
                    int Time = Convert.ToInt32(data.Split(':')[0]);
                    Thread.Sleep(Time*1000);
                    break;
            }

        }

        /// <summary>
        /// Click the control based on the text. the controls doesn't have id value, if the text change according to the 
        /// </summary>
        /// <param name="TypeControl">Specify for which Type of control you want to perform this action.</param>
        /// <param name="ControlKeyword">Property that you used to identify the control</param>
        /// <param name="ControlKeywordValue">What is the value of the property you specified</param>
        /// <param name="steps">Step Number.</param>
        /// <param name="description">Description about the step.</param>
        /// <param name="Operationkeyword">This function will be triggered using Keyword: "ClickLanguageText" </param>
        /// <returns>return true or false value</returns>
        public bool WebClickControlText(string TypeControl, string ControlKeyword, string ControlKeywordValue, string StepId, string Description, string Operationkeyword)
        {
            HtmlControl genericsControl = (HtmlControl)Activator.CreateInstance(typeof(HtmlControl), new object[] { ParentWindow });
            Thread.Sleep(2000);
            Point p = new Point();
            try
            {
                switch (TypeControl)
                {
                    case "HtmlButton":
                        switch (ControlKeyword)
                        {
                            case "ClickTextAvailable":
                                genericsControl.SearchProperties[HtmlButton.PropertyNames.InnerText] = ControlKeywordValue;
                                if (genericsControl.Exists)
                                {
                                    Mouse.Click(genericsControl);
                                }
                                break;
                            case "ClickFNameAvailable":
                                genericsControl.SearchProperties[HtmlButton.PropertyNames.FriendlyName] = ControlKeywordValue;
                                if (genericsControl.Exists)
                                {
                                    Mouse.Click(genericsControl);
                                }
                                break;
                            case "FriendlyName":
                                genericsControl.SearchProperties[HtmlButton.PropertyNames.FriendlyName] = ControlKeywordValue;
                                Mouse.Click(genericsControl);
                                break;
                            case "DisplayText":
                                genericsControl.SearchProperties[HtmlButton.PropertyNames.DisplayText] = ControlKeywordValue;
                                Mouse.Click(genericsControl);
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlButton.PropertyNames.Id] = ControlKeywordValue;
                                Mouse.Click(genericsControl);
                                break;
                        }
                        return true;
                    case "HtmlCustom":
                        switch (ControlKeyword)
                        {
                            case "ClickTextAvailable":
                                genericsControl.SearchProperties[HtmlCustom.PropertyNames.InnerText] = ControlKeywordValue;
                                if (genericsControl.Exists)
                                {
                                    Mouse.Click(genericsControl);
                                }
                                break;
                            case "InnerText":
                                genericsControl.SearchProperties[HtmlCustom.PropertyNames.InnerText] = ControlKeywordValue;
                                Mouse.Click(genericsControl);
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlCustom.PropertyNames.Id] = ControlKeywordValue;
                                Mouse.Click(genericsControl);
                                break;
                        }
                        return true;
                    case "HtmlHyperlink":
                        switch (ControlKeyword)
                        {
                            case "ClickTextAvailable":
                                genericsControl.SearchProperties[HtmlHyperlink.PropertyNames.InnerText] = ControlKeywordValue;
                                if (genericsControl.Exists)
                                {
                                    auto.Send("{TAB}");
                                    auto.Send("{ENTER}");
                                }
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlHyperlink.PropertyNames.InnerText] = ControlKeywordValue;
                                Mouse.Click(genericsControl);
                                break;
                        }
                        return true;
                    case "HtmlDiv":
                        switch (ControlKeyword)
                        {
                            case "InnerText":
                                genericsControl.SearchProperties[HtmlDiv.PropertyNames.InnerText] = ControlKeywordValue;
                                Mouse.Click(genericsControl);
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlDiv.PropertyNames.InnerText] = ControlKeywordValue;
                                Mouse.Click(genericsControl);
                                break;
                        }
                        return true;
                    case "HtmlLabel":
                        switch (ControlKeyword)
                        {
                            case "InnerText":
                                genericsControl.SearchProperties[HtmlLabel.PropertyNames.InnerText] = ControlKeywordValue;
                                Mouse.Click(genericsControl);
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlLabel.PropertyNames.InnerText] = ControlKeywordValue;
                                Mouse.Click(genericsControl);
                                break;
                        }
                        return true;
                    default:
                        Operation.FailerReason = TypeControl + " Control Type not configured to Click. Please check the Control Type.";
                        genDetailedReport.Reports(StepId, Description, Operationkeyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                        genDetailedReport.FileCorreptionCheck();
                        throw new NoSuchControlTypeFound(TypeControl + " Control Type not configured to Click. Please check the Control Type.");
                }
            }
            catch (UITestControlNotFoundException)
            {
                genDetailedReport.Reports(StepId, Description, Operationkeyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                Operation.FailerReason = TypeControl + " with " + ControlKeywordValue + " as " + ControlKeyword + " not found in page";
                genDetailedReport.FileCorreptionCheck();
                TakeScreenshot(ControlKeywordValue + " as " + ControlKeyword + " not found");
                Assert.Fail(Operation.FailerReason);
                return false;
            }
            catch (FailedToPerformActionOnBlockedControlException e)
            {
                genericsControl.DrawHighlight();
                TakeScreenshot("ControlCanontbeClicked");
                genDetailedReport.Reports(StepId, Description, Operationkeyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                Operation.FailerReason = TypeControl + " with " + ControlKeywordValue + " as " + ControlKeyword + " is blocked by another control";
                genDetailedReport.FileCorreptionCheck();
                Assert.Fail(Operation.FailerReason);
                return false;
            }
        }

        /// <summary>
        /// This used to click the control.that we can't directly click, so you can specify the another control nearby and move the pointer using coordinates and click the control.
        /// </summary>
        /// <param name="TypeControl">Specify for which Type of control you want to perform this action.</param>
        /// <param name="ControlKeyword">Property that you used to identify the control</param>
        /// <param name="ControlKeywordValue">What is the value of the property you specified</param>
        /// <param name="steps">Step Number.</param>
        /// <param name="description">Description about the step.</param>
        /// <param name="Operationkeyword">This function will be triggered using Keyword: "MouseMove" </param>
        /// <param name="codinates">specify the x and y coordinates to move to a control</param>
        /// <returns></returns>
        public bool WebMouseMoveClick(string TypeControl, string ControlKeyword, string ControlKeywordValue, string StepId, string Description, string Operationkeyword, MouseButtons Button, params int[] codinates)
        {
            HtmlControl genericsControl = (HtmlControl)Activator.CreateInstance(typeof(HtmlControl), new object[] { ParentWindow });
            Point p = new Point();
            try
            {
                switch (TypeControl)
                {
                    case "HtmlCustom":
                        switch (ControlKeyword)
                        {
                            case "InnerText":
                                genericsControl.SearchProperties[HtmlCustom.PropertyNames.InnerText] = ExcelDataTable.ReadData(1, ControlKeywordValue);
                                int X = genericsControl.BoundingRectangle.X;
                                int Y = genericsControl.BoundingRectangle.Y;
                                auto.MouseMove(X + codinates[0], Y + codinates[1], 1000);
                                Mouse.Click(Button);
                                break;
                            case "TextVisible":
                                genericsControl.SearchProperties[HtmlCustom.PropertyNames.InnerText] = ExcelDataTable.ReadData(1, ControlKeywordValue);
                                bool d = genericsControl.TryGetClickablePoint(out p);
                                if (d)
                                {
                                    X = genericsControl.BoundingRectangle.X;
                                    Y = genericsControl.BoundingRectangle.Y;
                                    auto.MouseMove(X + codinates[0], Y + codinates[1], 1000);
                                    Mouse.Click(Button);
                                }
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlCustom.PropertyNames.Id] = ControlKeywordValue;
                                X = genericsControl.BoundingRectangle.X;
                                Y = genericsControl.BoundingRectangle.Y;
                                auto.MouseMove(X + codinates[0], Y + codinates[1], 1000);
                                Mouse.Click(Button);
                                break;
                        }
                        return true;
                    case "HtmlDiv":
                        switch (ControlKeyword)
                        {
                            case "InnerText":
                                genericsControl.SearchProperties[HtmlDiv.PropertyNames.InnerText] = ControlKeywordValue;
                                int X = genericsControl.BoundingRectangle.X;
                                int Y = genericsControl.BoundingRectangle.Y;
                                auto.MouseMove(X + codinates[0], Y + codinates[1], 1000);
                                Mouse.Click(Button);
                                break;
                            case "FriendlyName":
                                genericsControl.SearchProperties[HtmlDiv.PropertyNames.FriendlyName] = ControlKeywordValue;
                                X = genericsControl.BoundingRectangle.X;
                                Y = genericsControl.BoundingRectangle.Y;
                                auto.MouseMove(X + codinates[0], Y + codinates[1], 1000);
                                Mouse.Click(Button);
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlDiv.PropertyNames.Id] = ControlKeywordValue;
                                X = genericsControl.BoundingRectangle.X;
                                Y = genericsControl.BoundingRectangle.Y;
                                auto.MouseMove(X + codinates[0], Y + codinates[1], 1000);
                                Mouse.Click(Button);
                                break;
                        }
                        return true;
                    case "HtmlButton":
                        switch (ControlKeyword)
                        {
                            case "FriendlyName":
                                genericsControl.SearchProperties[HtmlButton.PropertyNames.FriendlyName] = ControlKeywordValue;
                                int X = genericsControl.BoundingRectangle.X;
                                int Y = genericsControl.BoundingRectangle.Y;
                                auto.MouseMove(X + codinates[0], Y + codinates[1], 1000);
                                Mouse.Click(Button);
                                break;
                            case "InnerText":
                                genericsControl.SearchProperties[HtmlButton.PropertyNames.InnerText] = ControlKeywordValue;
                                X = genericsControl.BoundingRectangle.X;
                                Y = genericsControl.BoundingRectangle.Y;
                                auto.MouseMove(X + codinates[0], Y + codinates[1], 1000);
                                Mouse.Click(Button);
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlButton.PropertyNames.Id] = ControlKeywordValue;
                                X = genericsControl.BoundingRectangle.X;
                                Y = genericsControl.BoundingRectangle.Y;
                                auto.MouseMove(X + codinates[0], Y + codinates[1], 1000);
                                Mouse.Click(Button);
                                break;
                        }
                        return true;
                    case "HtmlHyperlink":
                        switch (ControlKeyword)
                        {
                            case "InnerText":
                                genericsControl.SearchProperties[HtmlHyperlink.PropertyNames.InnerText] = ControlKeywordValue;
                                int X = genericsControl.BoundingRectangle.X;
                                int Y = genericsControl.BoundingRectangle.Y;
                                auto.MouseMove(X + codinates[0], Y + codinates[1], 1000);
                                Mouse.Click(Button);
                                break;
                            case "TextVisible":
                                genericsControl.SearchProperties[HtmlHyperlink.PropertyNames.InnerText] = ExcelDataTable.ReadData(1, ControlKeywordValue);
                                bool d = genericsControl.TryGetClickablePoint(out p);
                                if (d)
                                {
                                    X = genericsControl.BoundingRectangle.X;
                                    Y = genericsControl.BoundingRectangle.Y;
                                    auto.MouseMove(X + codinates[0], Y + codinates[1], 1000);
                                    Mouse.Click(Button);
                                }
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlHyperlink.PropertyNames.Id] = ControlKeywordValue;
                                X = genericsControl.BoundingRectangle.X;
                                Y = genericsControl.BoundingRectangle.Y;
                                auto.MouseMove(X + codinates[0], Y + codinates[1], 1000);
                                Mouse.Click(Button);
                                break;
                        }
                        return true;
                    case "HtmlHeaderCell":
                        switch (ControlKeyword)
                        {
                            case "InnerText":
                                genericsControl.SearchProperties[HtmlHeaderCell.PropertyNames.InnerText] = ControlKeywordValue;
                                int X = genericsControl.BoundingRectangle.X;
                                int Y = genericsControl.BoundingRectangle.Y;
                                auto.MouseMove(X + codinates[0], Y + codinates[1], 1000);
                                Mouse.Click(Button);
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlHeaderCell.PropertyNames.Id] = ControlKeywordValue;
                                X = genericsControl.BoundingRectangle.X;
                                Y = genericsControl.BoundingRectangle.Y;
                                auto.MouseMove(X + codinates[0], Y + codinates[1], 1000);
                                Mouse.Click(Button);
                                break;
                        }
                        return true;
                    case "HtmlRow":
                        switch (ControlKeyword)
                        {
                            case "InnerText":
                                genericsControl.SearchProperties[HtmlRow.PropertyNames.InnerText] = ControlKeywordValue;
                                int X = genericsControl.BoundingRectangle.X;
                                int Y = genericsControl.BoundingRectangle.Y;
                                auto.MouseMove(X + codinates[0], Y + codinates[1], 1000);
                                Mouse.Click(Button);
                                break;
                            case "LanguageInnerText":
                                CommonLanguageTemplateReader.Message(Operation.lang, ControlKeywordValue);
                                genericsControl.SearchProperties[HtmlRow.PropertyNames.InnerText] = languageResource.Msg_GetTemplateMessage;
                                X = genericsControl.BoundingRectangle.X;
                                Y = genericsControl.BoundingRectangle.Y;
                                auto.MouseMove(X + codinates[0], Y + codinates[1], 1000);
                                Mouse.Click(Button);
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlRow.PropertyNames.Id] = ControlKeywordValue;
                                X = genericsControl.BoundingRectangle.X;
                                Y = genericsControl.BoundingRectangle.Y;
                                auto.MouseMove(X + codinates[0], Y + codinates[1], 1000);
                                Mouse.Click(Button);
                                break;
                        }
                        return true;
                    case "HtmlLabel":
                        switch (ControlKeyword)
                        {
                            case "InnerText":
                                genericsControl.SearchProperties[HtmlLabel.PropertyNames.InnerText] = ControlKeywordValue;
                                int X = genericsControl.BoundingRectangle.X;
                                int Y = genericsControl.BoundingRectangle.Y;
                                auto.MouseMove(X + codinates[0], Y + codinates[1], 1000);
                                Mouse.Click(Button);
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlLabel.PropertyNames.Id] = ControlKeywordValue;
                                X = genericsControl.BoundingRectangle.X;
                                Y = genericsControl.BoundingRectangle.Y;
                                auto.MouseMove(X + codinates[0], Y + codinates[1], 1000);
                                Mouse.Click(Button);
                                break;
                        }
                        return true;
                    case "HtmlSpan":
                        switch (ControlKeyword)
                        {
                            case "InnerText":
                                genericsControl.SearchProperties[HtmlSpan.PropertyNames.InnerText] = ControlKeywordValue;
                                int X = genericsControl.BoundingRectangle.X;
                                int Y = genericsControl.BoundingRectangle.Y;
                                auto.MouseMove(X + codinates[0], Y + codinates[1], 1000);
                                Mouse.Click(Button);
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlSpan.PropertyNames.Id] = ControlKeywordValue;
                                X = genericsControl.BoundingRectangle.X;
                                Y = genericsControl.BoundingRectangle.Y;
                                auto.MouseMove(X + codinates[0], Y + codinates[1], 1000);
                                Mouse.Click(Button);
                                break;
                        }
                        return true;
                    case "HtmlImage":
                        switch (ControlKeyword)
                        {
                            case "IDData":
                                string[] IdData = ExcelDataTable.ReadData(1, ControlKeywordValue).Split('+');
                                if (IdData.Count() > 1)
                                {
                                    if (IdData[1].StartsWith("Rec_"))
                                    {
                                        ControlKeywordValue = IdData[0].Replace("&", Operation.recordedData[IdData[1].Replace("Rec_", string.Empty)]);
                                    }
                                    if (IdData[1].Contains('='))
                                    {
                                        ExcelDataTable.PopulateRecordData(WriteAndReadData.DataFilePath);
                                        ControlKeywordValue = IdData[0].Replace("&", getDataFromDynamicExcel(IdData[1]));
                                    }
                                    else
                                    {
                                        ControlKeywordValue = IdData[0].Replace("&", ExcelDataTable.ReadData(1, IdData[1]));
                                    }
                                }
                                else {
                                    throw new NullReferenceException("Reference Data not found after + Symbol");
                                }
                                genericsControl.SearchProperties[HtmlImage.PropertyNames.Id] = ControlKeywordValue;
                                int X = genericsControl.BoundingRectangle.X;
                                int Y = genericsControl.BoundingRectangle.Y;
                                auto.MouseMove(X + codinates[0], Y + codinates[1], 1000);
                                Mouse.Click(Button);
                                break;
                        }
                        return true;
                    default:
                        Operation.FailerReason = TypeControl + " Control Type not configured for Mouse Move. Please check the Control Type.";
                        genDetailedReport.Reports(StepId, Description, Operationkeyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                        genDetailedReport.FileCorreptionCheck();
                        throw new NoSuchControlTypeFound(TypeControl + " Control Type not configured for Mouse Move. Please check the Control Type.");
                }
            }
            catch (UITestControlNotFoundException e)
            {
                genDetailedReport.Reports(StepId, Description, Operationkeyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                Operation.FailerReason = TypeControl + " with " + ControlKeywordValue + " as " + ControlKeyword + " not found in page";
                genDetailedReport.FileCorreptionCheck();
                TakeScreenshot(ControlKeywordValue + " as " + ControlKeyword + " not found");
                Assert.Fail(Operation.FailerReason);
                return false;
            }

        }

        /// <summary>
        /// Set the value to a control using script.
        /// </summary>
        /// <param name="TypeControl">Specify for which Type of control you want to perform this action.</param>
        /// <param name="ControlKeyword">Property that you used to identify the control</param>
        /// <param name="ControlKeywordValue">What is the value of the property you specified</param>
        /// <param name="DataRefferencekeyword">Reference to the data that you want to enter in to the control.</param>
        /// <param name="steps">Step Number.</param>
        /// <param name="description">Description about the step.</param>
        public void ScriptExcuite(string TypeOfOperation, string ControlKeyword, string ControlKeywordValue, string DataRefferencekeyword, string StepId, string Description)
        {
            BrowserWindow windows = new BrowserWindow(); 
            try
            {
                switch (TypeOfOperation)
                {
                    case "SetValue":
                        switch (ControlKeyword)
                        {
                            case "ID":
                                windows.ExecuteScript("document.getElementById('" + ControlKeywordValue + "').value='" + DataRefferencekeyword + "'");
                                break;
                            case "Attribute":
                                windows.ExecuteScript("document.getElementById('" + ControlKeywordValue + "').setAttribute('" + DataRefferencekeyword.Split('=')[0] + "','" + DataRefferencekeyword.Split('=')[1] + "');");
                                break;
                            case "SelectItem":
                                windows.ExecuteScript(@"$('#" + ControlKeywordValue + "').val($('#" + ControlKeywordValue + " option:eq(" + DataRefferencekeyword + ")').val())");
                                break;
                            default:
                                break;
                        }
                        break;
                    default:
                        break;
                }
            }
            catch (Exception e)
            {
                Operation.FailerReason = "Could not find control with specified ID value Please check the Control Keyword Value.";
                genDetailedReport.Reports(StepId, Description, TypeOfOperation, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genDetailedReport.FileCorreptionCheck();
                TakeScreenshot(ControlKeywordValue + " as " + ControlKeyword + " not found");
                throw new NoSuchControlTypeFound(ControlKeywordValue + " control can not find Please check the Control Keyword Value.");
            }
        }

        /// <summary>
        /// Set the attribute of the control using script.
        /// </summary>
        /// <param name="TypeControl">Specify for which Type of control you want to perform this action.</param>
        /// <param name="ControlKeyword">Property and value of that property that you separated by ':'</param>
        /// <param name="ControlKeywordValue">What is the value of the property you specified</param>
        /// <param name="DataRefferencekeyword">Reference to the data that you want to enter in to the control.</param>
        /// <param name="steps">Step Number.</param>
        /// <param name="description">Description about the step.</param>
        public void ScriptSetAttribute(string TypeOfOperation, string ControlKeyword, string ControlKeywordValue, string DataRefferencekeyword, string StepId, string Description)
        {
            BrowserWindow windows = new BrowserWindow();
            try
            {
                switch (TypeOfOperation)
                {
                    case "SetAttribute":
                        switch (ControlKeyword)
                        {
                            case "ID":
                                windows.ExecuteScript("document.getElementById('" + ControlKeywordValue + "').setAttribute('" + DataRefferencekeyword.Split('=')[0] + "','" + DataRefferencekeyword.Split('=')[1] + "');");
                                break;
                            default:
                                break;
                        }
                        break;
                    default:
                        break;
                }
            }
            catch (Exception e)
            {
                Operation.FailerReason = "Could not find control with specified ID value Please check the Control Keyword Value. Attribute Could not be set for this control";
                genDetailedReport.Reports(StepId, Description, TypeOfOperation, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genDetailedReport.FileCorreptionCheck();
                TakeScreenshot(ControlKeywordValue + " as " + ControlKeyword + " not found");
                throw new NoSuchControlTypeFound(ControlKeywordValue + " control can not find Please check the Control Keyword Value. Attribute Could not be set for this control");
            }
        }


        /// <summary>
        /// This used to excite the script, to get the record data.
        /// </summary>
        /// <param name="TypeOfOperation">This function will be triggered using Keyword: ""
        /// "GetValueMatch"
        /// "GetValueNoMatch"
        /// "ValidateEmptyValueEqual"
        /// "ValidateEmptyValueNotEqual"
        /// "ScrollToControl"
        /// "ValidateControlCount"
        /// "ValidatePartialPickSatus"
        /// "IsReadOnly"
        /// "IsChecked"
        /// "GetDeviceIDValue" : This will get the Device ID value in the Device->Device Information after search for the device that we added.
        /// </param>
        /// <param name="ControlKeyword">Property that you used to identify the control</param>
        /// <param name="ControlKeywordValue">What is the value of the property you specified</param>
        /// <param name="StepId">Step Number.</param>
        /// <param name="Description">Description about the step.</param>
        /// <returns>return specified value</returns>
        public string ScriptGetExcuite(string TypeOfOperation, string ControlKeyword, string ControlKeywordValue, string StepId, string Description, string AssertionMsg, params string[] Typeofcontrol)
        {
            BrowserWindow windoww = new BrowserWindow();
            try
            {
                string[] IdData = ExcelDataTable.ReadData(1, ControlKeywordValue).Split('+');
                if (IdData.Count() > 1)
                {
                    if (IdData[1].StartsWith("Rec_"))
                    {
                        ControlKeywordValue = IdData[0].Replace("&", Operation.recordedData[IdData[1].Replace("Rec_", string.Empty)]);
                    }
                    if (IdData[1].Contains(":"))
                    {
                        WriteAndReadData.DataFilePath = ConfigurationManager.AppSettings["ExcelDataFile"] + "\\DynamicData" + "\\" + Operation.Batch + ".xls";
                        ExcelDataTable.PopulateRecordData(WriteAndReadData.DataFilePath);
                        ControlKeywordValue = IdData[0].Replace("&", getDataFromDynamicExcel(IdData[1]));
                    }
                    else
                    {
                        ControlKeywordValue = IdData[0].Replace("&", ExcelDataTable.ReadData(1, IdData[1]));
                    }
                }

                if (Typeofcontrol.Length != 0 && Typeofcontrol[0]!="")
                {
                        string value = WebGetControlData(Typeofcontrol[0].ToString(), ControlKeyword, ControlKeywordValue, "", "", StepId, Description, TypeOfOperation).Trim();
                        return value;
                }
                else
                {
                    switch (TypeOfOperation)
                    {
                        case "GetValueMatch":
                        case "GetValueNoMatch":
                        case "GetValueContains":
                        case "ValueNotContains":
                        case "ValidateEmptyValueEqual":
                        case "ValidateEmptyValueNotEqual":
                        case "ScrollToControl":
                        case "ValidateControlCount":
                        case "VallidateNoControlCount":
                        case "IsReadOnly":
                        case "IsNotReadOnly":
                        case "ConditionEquals":
                        case "ConditionNotEquals":
                        case "IsChecked":
                        case "ReadData":
                            switch (ControlKeyword)
                            {
                                case "Class":
                                    string data = windoww.ExecuteScript("var data = document.getElementsByClassName('" + ControlKeywordValue + "')[0].innerText; return data;").ToString();
                                    return data;
                                case "ID":
                                    data = windoww.ExecuteScript("var data = document.getElementById('" + ControlKeywordValue + "').getAttribute('value'); return data;").ToString();
                                    return data;
                                case "IDValue":
                                    data = windoww.ExecuteScript("var data = document.getElementById('" + ControlKeywordValue + "').value; return data;").ToString();
                                    return data;
                                case "IDText":
                                    data = windoww.ExecuteScript("var data = document.getElementById('" + ControlKeywordValue + "').innerHTML; return data;").ToString();
                                    return data;
                                case "IDTextValue":
                                    data = windoww.ExecuteScript("var data = document.getElementById('" + ControlKeywordValue + "').innerText; return data;").ToString();
                                    return data;
                                case "IDStyle":
                                    data = windoww.ExecuteScript("var data = document.getElementById('" + ControlKeywordValue + "').getAttribute('style'); return data;").ToString();
                                    return data;
                                case "IDClass":
                                    data = windoww.ExecuteScript("var data = document.getElementById('" + ControlKeywordValue + "').getAttribute('class'); return data;").ToString();
                                    return data;
                                case "IDScroll":
                                    windoww.ExecuteScript("var data = document.getElementById('" + ControlKeywordValue + "').scrollIntoView(true);");
                                    return "true";
                                case "IDElementsCount":
                                    data = windoww.ExecuteScript("var data = document.getElementById('" + ControlKeywordValue + "').childElementCount; return data;").ToString();
                                    return data;
                                case "IDReadOnly":
                                    data = windoww.ExecuteScript("var data = document.getElementById('" + ControlKeywordValue + "').readOnly; return data;").ToString();
                                    return data;
                                case "IsEditable":
                                case "IsDisabledOnly":
                                    data = windoww.ExecuteScript("var data = document.getElementById('" + ControlKeywordValue + "').disabled; return data;").ToString();
                                    return data;
                                case "TRStyle":
                                    string[] value = ControlKeywordValue.Split(':');
                                    data = windoww.ExecuteScript("var data = document.getElementById('" + value[0] + "').getElementsByTagName('tbody')[0].getElementsByTagName('tr')['" + value[1] + "'].getAttribute('style');return data;").ToString();
                                    return data;
                                case "IDChecked":
                                    data = windoww.ExecuteScript("var data = document.getElementById('" + ControlKeywordValue + "').checked; return data;").ToString();
                                    return data;
                                case "SelectItem":
                                    data = windoww.ExecuteScript("var data = document.getElementById('" + ControlKeywordValue + "');var item = data.options[data.selectedIndex].text; return item;").ToString();
                                    return data;
                                case "TagInstance":
                                    string[] tagdata = ControlKeywordValue.Split('=');
                                    data = windoww.ExecuteScript("var data = document.getElementById('" + tagdata[0] + "').getElementsByTagName('" + tagdata[1].Split(':')[0] + "')[" + tagdata[1].Split(':')[1] + "].innerText; return data;").ToString().Trim();
                                    return data;
                                case "ImgTagInstance":
                                    data = windoww.ExecuteScript("var data = document.getElementById('" + ControlKeywordValue + "').getElementsByTagName('img')[0].getAttribute('src'); return data;").ToString();
                                    return data;
                                case "GetTransferReqStatus":
                                    data = windoww.ExecuteScript("var data = document.getElementById('" + ControlKeywordValue + "').getElementsByTagName('img')[1].getAttribute('src'); return data;").ToString();
                                    return data;
                                case "GetTableData":
                                    data = windoww.ExecuteScript("var data = document.getElementById('" + ControlKeywordValue.Split('|')[0] + "').getElementsByTagName('td')['"+ ControlKeywordValue.Split('|')[1] + "'].innerText; return data;").ToString().Trim();
                                    return data;
                                    //top get style of control using class name 
                                case "ClassStyle":
                                    data = window.ExecuteScript("var data = document.getElementsByClassName('" + ControlKeywordValue + "')[0].getAttribute('style'); return data;").ToString();
                                    return data;
                                case "GetTblControl":
                                    // Gets the class value of the first control tag inside the table specify table id value and tag name of the control also occurrence of the tag name and attribute you want to get separated with '|' symbol
                                    data = window.ExecuteScript("var data = document.getElementById('" + ControlKeywordValue.Split('|')[0] + "').getElementsByTagName('" + ControlKeywordValue.Split('|')[1] + "')[" + ControlKeywordValue.Split('|')[2] + "].getAttribute('" + ControlKeywordValue.Split('|')[3] + "'); return data;").ToString();
                                    return data;
                                case "IFrameTblData":
                                    string[] str = ControlKeywordValue.Split('|');
                                    if (str.Length > 3)
                                    {
                                        //to get title  of pass icon using iframe id and table id along with td index value and title attribute
                                        data = window.ExecuteScript("var data=window.frames['" + ControlKeywordValue.Split('|')[0] + "'].document.getElementById('" + ControlKeywordValue.Split('|')[1] + "').getElementsByTagName('td')['" + ControlKeywordValue.Split('|')[2] + "'].getAttribute('" + ControlKeywordValue.Split('|')[3] + "'); return data;").ToString();
                                        return data;
                                    }
                                    else
                                    {
                                        //to get iframe table data using iframe id and table id along with td index value
                                        data = window.ExecuteScript("var data=window.frames['" + ControlKeywordValue.Split('|')[0] + "'].document.getElementById('" + ControlKeywordValue.Split('|')[1] + "').getElementsByTagName('td')['" + ControlKeywordValue.Split('|')[2] + "'].innerText; return data;").ToString();
                                        return data;
                                    }
                                default:
                                    return null;
                            }
                        default:
                            return null;
                    }
                }
            }
             catch (Exception e)
            {
                Operation.FailerReason = ControlKeywordValue + " control can not find Please check the Control Keyword Value," + AssertionMsg;
                genDetailedReport.Reports(StepId, Description, TypeOfOperation, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genDetailedReport.FileCorreptionCheck();
                TakeScreenshot(ControlKeywordValue + " as " + ControlKeyword + " not found");
                throw new NoSuchControlTypeFound(ControlKeywordValue + " control can not find Please check the Control Keyword Value.");
            }
        }

        /// <summary>
        /// Used to click the control using script
        /// </summary>
        /// <param name="TypeOfOperation">This function will be triggered using Keyword: "ScriptClick"</param>
        /// <param name="ControlKeyword">Property that you used to identify the control</param>
        /// <param name="ControlKeywordValue">What is the value of the property you specified</param>
        /// <param name="steps">Step Number.</param>
        /// <param name="description">Description about the step.</param>
        /// <returns>return true or false value.</returns>
        public bool ScriptClick(string TypeOfOperation, string ControlKeyword, string ControlKeywordValue, string StepId, string Description)
        {
            BrowserWindow br = new BrowserWindow();
            try
            {
                switch (ControlKeyword)
                {
                    case "ID":
                        br.ExecuteScript("var data = document.getElementById('" + ControlKeywordValue + "').click();");
                        return true;
                    case "IDData":
                        string[] IdData = ExcelDataTable.ReadData(1, ControlKeywordValue).Split('+');
                        if (IdData.Count() > 1)
                        {
                            if (IdData[1].StartsWith("Rec_"))
                            {
                                ControlKeywordValue = IdData[0].Replace("&", Operation.recordedData[IdData[1].Replace("Rec_", string.Empty)]);
                            }
                            else if (ControlKeywordValue.Contains('='))
                            {
                                WriteAndReadData.DataFilePath = ConfigurationManager.AppSettings["ExcelDataFile"] + "\\DynamicData" + "\\" + Operation.Batch + ".xls";
                                ExcelDataTable.PopulateRecordData(WriteAndReadData.DataFilePath);
                                ControlKeywordValue = IdData[0].Replace("&" , getDataFromDynamicExcel(ControlKeywordValue));
                            }
                            else
                            {
                                ControlKeywordValue = IdData[0].Replace("&", ExcelDataTable.ReadData(1, IdData[1]));
                            }
                        }
                        else {
                            throw new NullReferenceException("Reference Data not found after + Symbol");
                        }
                        br.ExecuteScript("var data = document.getElementById('" + ControlKeywordValue + "').click();");
                        return true;
                    case "CheckingExist":
                        var data = br.ExecuteScript("var data = document.getElementById('" + ControlKeywordValue + "').childElementCount; return data;").ToString();
                        if (data != "1")
                        {
                            br.ExecuteScript("var data = document.getElementById('" + ControlKeywordValue + "').click();");
                            Thread.Sleep(1000);
                        }
                        return true;
                    case "CheckRemove":
                        data = br.ExecuteScript("var data = document.getElementById('" + ControlKeywordValue + "').childElementCount; return data;").ToString();
                        if (data == "1")
                        {
                            br.ExecuteScript("var data = document.getElementById('" + ControlKeywordValue + "').click();");
                            Thread.Sleep(1000);
                        }
                        return true;
                    case "Function":
                        br.ExecuteScript("" + ControlKeywordValue + "");
                        return true;
                    default:
                        return false;
                }
            }
            catch (Exception e)
            {
                Operation.FailerReason = "Could not find control with specified ID value Please check the Control Keyword Value.";
                genDetailedReport.Reports(StepId, Description, TypeOfOperation, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genDetailedReport.FileCorreptionCheck();
                TakeScreenshot(ControlKeywordValue + " as " + ControlKeyword + " not found");
                throw new NoSuchControlTypeFound(ControlKeywordValue + " control can not find Please check the Control Keyword Value.");
            }
        }

        /// <summary>
        /// This will validate the control is displaying or not or available or not.
        /// </summary>
        /// <param name="TypeControl">Specify for which Type of control you want to perform this action.</param>
        /// <param name="ControlKeyword">Property that you used to identify the control</param>
        /// <param name="ControlKeywordValue">What is the value of the property you specified</param>
        /// <param name="DataRefferencekeyword">Reference to the data that you want to enter in to the control.</param>
        /// <param name="steps">Step Number.</param>
        /// <param name="description">Description about the step.</param>
        /// <param name="Operationkeyword">This function will be triggered using Keyword: "ValidateControlAvailable"</param>
        /// <returns>return true or false value.</returns>
        public bool ControlAvailability(string TypeControl, string ControlKeyword, string ControlKeywordValue, string StepId, string Description, string Operationkeyword)
        {
            HtmlControl genericsControl = (HtmlControl)Activator.CreateInstance(typeof(HtmlControl), new object[] { ParentWindow });
            Point p = new Point();
            Thread.Sleep(2000);
            try
            {
                switch (TypeControl)
                {
                    case "HtmlDiv":
                        switch (ControlKeyword)
                        {
                            case "IDVisible":
                                if (ControlKeywordValue.Contains('='))
                                {
                                    string _data = getDataFromDynamicExcel(ControlKeywordValue);
                                    genericsControl.SearchProperties[HtmlDiv.PropertyNames.Id] = _data;
                                }
                                else
                                {
                                    genericsControl.SearchProperties[HtmlDiv.PropertyNames.Id] = ControlKeywordValue;
                                }
                                bool d = genericsControl.TryGetClickablePoint(out p);
                                return d;
                            case "ClassVisible":
                                genericsControl.SearchProperties[HtmlDiv.PropertyNames.Class] = ControlKeywordValue;
                                d = genericsControl.TryGetClickablePoint(out p);
                                return d;
                            case "InnerText":
                                genericsControl.SearchProperties[HtmlDiv.PropertyNames.InnerText] = ExcelDataTable.ReadData(1, ControlKeywordValue);
                                if (genericsControl.Exists)
                                { return true; }
                                else
                                { return false; }
                            case "LanguageInnerText":
                                CommonLanguageTemplateReader.Message(Operation.lang, ControlKeywordValue);
                                genericsControl.SearchProperties[HtmlDiv.PropertyNames.InnerText] = languageResource.Msg_GetTemplateMessage;
                                if (genericsControl.Exists)
                                { return true; }
                                else
                                { return false; }
                            default:
                                if (ControlKeywordValue.Contains('='))
                                {
                                    string _data = getDataFromDynamicExcel(ControlKeywordValue);
                                    genericsControl.SearchProperties[HtmlDiv.PropertyNames.InnerText] = _data;
                                }
                                else
                                {
                                    genericsControl.SearchProperties[HtmlDiv.PropertyNames.InnerText] = ControlKeywordValue;
                                }
                                if (genericsControl.Exists)
                                { return true; }
                                else
                                { return false; }
                        }
                    case "HtmlHyperlink":
                        switch (ControlKeyword)
                        {
                            case "IDVisible":
                                genericsControl.SearchProperties[HtmlHyperlink.PropertyNames.Id] = ControlKeywordValue;
                                bool d = genericsControl.TryGetClickablePoint(out p);
                                return d;
                            case "InnerText":
                                genericsControl.SearchProperties[HtmlHyperlink.PropertyNames.InnerText] = ExcelDataTable.ReadData(1, ControlKeywordValue);
                                if (genericsControl.Exists)
                                { return true; }
                                else
                                { return false; }
                            default:
                                genericsControl.SearchProperties[HtmlHyperlink.PropertyNames.Id] = ControlKeywordValue;
                                if (genericsControl.Exists)
                                { return true; }
                                else
                                { return false; }
                        }
                    case "HtmlCell":
                        switch (ControlKeyword)
                        {
                            case "InnerText":
                                if(ControlKeywordValue.Contains('='))
                                {
                                    string _data = getDataFromDynamicExcel(ControlKeywordValue);
                                    genericsControl.SearchProperties[HtmlCell.PropertyNames.InnerText] = _data;
                                }
                                else
                                {
                                    genericsControl.SearchProperties[HtmlCell.PropertyNames.InnerText] = ExcelDataTable.ReadData(1, ControlKeywordValue);
                                }
                                if (genericsControl.Exists)
                                { return true; }
                                else
                                { return false; }
                            case "LanguageInnerText":
                                CommonLanguageTemplateReader.Message(Operation.lang, ControlKeywordValue);
                                genericsControl.SearchProperties[HtmlCell.PropertyNames.InnerText] = languageResource.Msg_GetTemplateMessage;
                                if (genericsControl.Exists)
                                { return true; }
                                else
                                { return false; }
                            default:
                                genericsControl.SearchProperties[HtmlCell.PropertyNames.Id] = ControlKeywordValue;
                                if (genericsControl.Exists)
                                { return true; }
                                else
                                { return false; }
                        }
                    case "HtmlCheckBox":
                        switch (ControlKeyword)
                        {
                            case "IDVisible":
                                genericsControl.SearchProperties[HtmlCheckBox.PropertyNames.Id] = ControlKeywordValue;
                                bool d = genericsControl.TryGetClickablePoint(out p);
                                return d;
                            default:
                                genericsControl.SearchProperties[HtmlCheckBox.PropertyNames.Id] = ControlKeywordValue;
                                d = genericsControl.TryGetClickablePoint(out p);
                                return d;
                        }
                    case "HtmlButton":
                        switch (ControlKeyword)
                        {
                            case "IDVisible":
                                genericsControl.SearchProperties[HtmlButton.PropertyNames.Id] = ControlKeywordValue;
                                bool d = genericsControl.TryGetClickablePoint(out p);
                                return d;
                            default:
                                genericsControl.SearchProperties[HtmlButton.PropertyNames.Id] = ControlKeywordValue;
                                if (genericsControl.Exists)
                                { return true; }
                                else
                                { return false; }
                        }
                    case "HtmlSpan":
                        switch (ControlKeyword)
                        {
                            case "IDVisible":
                                genericsControl.SearchProperties[HtmlSpan.PropertyNames.Id] = ControlKeywordValue;
                                bool d = genericsControl.TryGetClickablePoint(out p);
                                return d;
                            case "InnerText":
                                if (ControlKeywordValue.Contains('='))
                                {
                                    string _data = getDataFromDynamicExcel(ControlKeywordValue);
                                    genericsControl.SearchProperties[HtmlSpan.PropertyNames.InnerText] = _data;
                                }
                                else
                                {
                                    genericsControl.SearchProperties[HtmlSpan.PropertyNames.InnerText] = ExcelDataTable.ReadData(1, ControlKeywordValue);
                                }
                                if (genericsControl.Exists)
                                { return true; }
                                else
                                { return false; }
                            case "LanguageInnerText":
                                CommonLanguageTemplateReader.Message(Operation.lang, ControlKeywordValue);
                                genericsControl.SearchProperties[HtmlCell.PropertyNames.InnerText] = languageResource.Msg_GetTemplateMessage;
                                if (genericsControl.Exists)
                                { return true; }
                                else
                                { return false; }
                            default:
                                genericsControl.SearchProperties[HtmlSpan.PropertyNames.Id] = ControlKeywordValue;
                                if (genericsControl.Exists)
                                { return true; }
                                else
                                { return false; }
                        }
                    case "HtmlInputButton":
                        switch (ControlKeyword)
                        {
                            case "IDVisible":
                                genericsControl.SearchProperties[HtmlInputButton.PropertyNames.Id] = ControlKeywordValue;
                                bool d = genericsControl.TryGetClickablePoint(out p);
                                return d;
                            case "ClassVisible":
                                genericsControl.SearchProperties[HtmlInputButton.PropertyNames.Class] = ControlKeywordValue;
                                d = genericsControl.TryGetClickablePoint(out p);
                                return d;
                            default:
                                genericsControl.SearchProperties[HtmlInputButton.PropertyNames.Id] = ControlKeywordValue;
                                if (genericsControl.Exists)
                                { return true; }
                                else
                                { return false; }
                        }
                    case "HtmlEdit":
                        switch (ControlKeyword)
                        {
                            case "ValueAttribute":
                                if (ControlKeywordValue.Contains('='))
                                {
                                    string _data = getDataFromDynamicExcel(ControlKeywordValue);
                                    genericsControl.SearchProperties[HtmlSpan.PropertyNames.ValueAttribute] = _data;
                                }
                                else
                                {
                                    genericsControl.SearchProperties[HtmlSpan.PropertyNames.InnerText] = ExcelDataTable.ReadData(1, ControlKeywordValue);
                                }
                                if (genericsControl.Exists)
                                { return true; }
                                else
                                { return false; }
                            default:
                                genericsControl.SearchProperties[HtmlInputButton.PropertyNames.Id] = ControlKeywordValue;
                                if (genericsControl.Exists)
                                { return true; }
                                else
                                { return false; }
                        }
                    case "HtmlLabel":
                        switch (ControlKeyword)
                        {
                            case "InnerText":
                                if (ControlKeywordValue.Contains('='))
                                {
                                    string _data = getDataFromDynamicExcel(ControlKeywordValue);
                                    genericsControl.SearchProperties[HtmlLabel.PropertyNames.InnerText] = _data;
                                }
                                else
                                {
                                    genericsControl.SearchProperties[HtmlLabel.PropertyNames.InnerText] = ExcelDataTable.ReadData(1, ControlKeywordValue);
                                }
                                if (genericsControl.Exists)
                                { return true; }
                                else
                                { return false; }
                            default:
                                genericsControl.SearchProperties[HtmlLabel.PropertyNames.Id] = ControlKeywordValue;
                                if (genericsControl.Exists)
                                { return true; }
                                else
                                { return false; }
                        }
                    case "HtmlComboBox":
                        switch (ControlKeyword)
                        {
                            case "ID":
                                genericsControl.SearchProperties[HtmlComboBox.PropertyNames.Id] = ControlKeywordValue;
                                if (genericsControl.Exists)
                                { return true; }
                                else
                                { return false; }
                            default:
                                genericsControl.SearchProperties[HtmlComboBox.PropertyNames.Id] = ControlKeywordValue;
                                if (genericsControl.Exists)
                                { return true; }
                                else
                                { return false; }
                        }
                    case "HtmlTable":
                        switch (ControlKeyword)
                        {
                            case "IDVisible":
                                genericsControl.SearchProperties[HtmlTable.PropertyNames.Id] = ControlKeywordValue;
                                bool d = genericsControl.TryGetClickablePoint(out p);
                                return d;
                            default:
                                genericsControl.SearchProperties[HtmlTable.PropertyNames.Id] = ControlKeywordValue;
                                if (genericsControl.Exists)
                                { return true; }
                                else
                                { return false; }
                        }
                    case "HtmlCustom":
                        switch (ControlKeyword)
                        {
                            case "Class":
                                genericsControl.SearchProperties[HtmlCustom.PropertyNames.Class] = ControlKeywordValue;
                                if (genericsControl.Exists)
                                { return true; }
                                else
                                { return false; }
                            case "LanguageInnerText":
                                CommonLanguageTemplateReader.Message(Operation.lang, ControlKeywordValue);
                                genericsControl.SearchProperties[HtmlCustom.PropertyNames.InnerText] = languageResource.Msg_GetTemplateMessage;
                                if (genericsControl.Exists)
                                { return true; }
                                else
                                { return false; }
                            case "IDVisible":
                                genericsControl.SearchProperties[HtmlCustom.PropertyNames.Id] = ControlKeywordValue;
                                bool d = genericsControl.TryGetClickablePoint(out p);
                                return d;
                            case "InnerText":
                                if (ControlKeywordValue.Contains('='))
                                {
                                    string _data = getDataFromDynamicExcel(ControlKeywordValue);
                                    genericsControl.SearchProperties[HtmlCustom.PropertyNames.InnerText] = _data;
                                }
                                else
                                {
                                    genericsControl.SearchProperties[HtmlCustom.PropertyNames.InnerText] = ExcelDataTable.ReadData(1, ControlKeywordValue);
                                }
                                if (genericsControl.Exists)
                                { return true; }
                                else
                                { return false; }
                            case "IDData":
                                string[] IdData = ExcelDataTable.ReadData(1, ControlKeywordValue).Split('+');
                                if (IdData.Count() > 1)
                                {
                                    if (IdData[1].StartsWith("Rec_"))
                                    {
                                        ControlKeywordValue = IdData[0].Replace("&", Operation.recordedData[IdData[1].Replace("Rec_", string.Empty)]);
                                    }
                                    if (IdData[1].Contains('='))
                                    {
                                        WriteAndReadData.DataFilePath = ConfigurationManager.AppSettings["ExcelDataFile"] + "\\DynamicData" + "\\" + Operation.Batch + ".xls";
                                        ExcelDataTable.PopulateRecordData(WriteAndReadData.DataFilePath);
                                        ControlKeywordValue = IdData[0].Replace("&", getDataFromDynamicExcel(ControlKeywordValue));
                                    }
                                    else
                                    {
                                        ControlKeywordValue = IdData[0].Replace("&", ExcelDataTable.ReadData(1, IdData[1]));
                                    }
                                }
                                else {
                                    throw new NullReferenceException("Reference Data not found after + Symbol");
                                }
                                genericsControl.SearchProperties[HtmlCustom.PropertyNames.Id] = ControlKeywordValue;
                                if (genericsControl.Exists)
                                { return true; }
                                else
                                { return false; }
                            default:
                                genericsControl.SearchProperties[HtmlCustom.PropertyNames.Id] = ControlKeywordValue;
                                if (genericsControl.Exists)
                                { return true; }
                                else
                                { return false; }
                        }
                    case "HtmlImage":
                        switch (ControlKeyword)
                        {
                            case "ImagePath":
                                CommonLanguageTemplateReader.Message(Operation.lang, ControlKeywordValue);
                                genericsControl.SearchProperties[HtmlImage.PropertyNames.Src] = languageResource.Msg_GetTemplateMessage;
                                var im = genericsControl.BoundingRectangle;
                                bool result = im.X != 0 ? true : false; 
                                return result;
                           case "IDData":
                                string[] IdData = ExcelDataTable.ReadData(1, ControlKeywordValue).Split('+');
                                if (IdData.Count() > 1)
                                {
                                    if (IdData[1].StartsWith("Rec_"))
                                    {
                                        ControlKeywordValue = IdData[0].Replace("&", Operation.recordedData[IdData[1].Replace("Rec_", string.Empty)]);
                                    }
                                    if (IdData[1].Contains('='))
                                    {
                                        WriteAndReadData.DataFilePath = ConfigurationManager.AppSettings["ExcelDataFile"] + "\\DynamicData" + "\\" + LoginOperatrion.batchforReport + ".xls";
                                        ExcelDataTable.PopulateRecordData(WriteAndReadData.DataFilePath);
                                        ControlKeywordValue = IdData[0].Replace("&", getDataFromDynamicExcel(ControlKeywordValue));
                                    }
                                    else
                                    {
                                        ControlKeywordValue = IdData[0].Replace("&", ExcelDataTable.ReadData(1, IdData[1]));
                                    }
                                }
                                else {
                                    throw new NullReferenceException("Reference Data not found after + Symbol");
                                }
                                genericsControl.SearchProperties[HtmlImage.PropertyNames.Id] = ControlKeywordValue;
                                bool d = genericsControl.TryGetClickablePoint(out p);
                                return d;
                            default:
                                genericsControl.SearchProperties[HtmlImage.PropertyNames.Id] = ControlKeywordValue;
                                if (genericsControl.Exists)
                                { return true; }
                                else
                                { return false; }
                        }
                    default:
                        Operation.FailerReason = TypeControl + " Control Type not configured for entering data. Please check the Control Type.";
                        genDetailedReport.Reports(StepId, Description, Operationkeyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                        genDetailedReport.FileCorreptionCheck();
                        throw new NoSuchControlTypeFound(TypeControl + " Control Type not configured for entering data. Please check the Control Type.");
                }
            }
            catch (Exception e)
            {
                return false;
            }
        }

        /// <summary>
        /// This will check and Unchecked the check box based on the option given in the test data options like "Yes" and "No"
        /// </summary>
        /// <param name="TypeControl">Specify for which Type of control you want to perform this action.</param>
        /// <param name="ControlKeyword">Property that you used to identify the control</param>
        /// <param name="ControlKeywordValue">What is the value of the property you specified</param>
        /// <param name="DataRefferencekeyword">Reference to the option that you want to check the check box or not</param>
        /// <param name="steps">Step Number.</param>
        /// <param name="description">Description about the step.</param>
        /// <param name="Operationkeyword">This function will be triggered using Keyword: "Check"</param>
        /// <returns>return true or false value.</returns>
        public bool WebSelectCheckBox(string TypeControl, string ControlKeyword, string ControlKeywordValue, string DataRefferencekeyword, string StepId, string Description, string Operationkeyword)
        {
            HtmlControl genericsControl = (HtmlControl)Activator.CreateInstance(typeof(HtmlControl), new object[] { ParentWindow });
            Point p = new Point();
            Thread.Sleep(2000);
            bool d;
            try
            {
                switch (TypeControl)
                {
                    case "HtmlCheckBox":
                        switch (ControlKeyword)
                        {
                            case "Check":
                                genericsControl.SearchProperties[HtmlCheckBox.PropertyNames.Id] = ControlKeywordValue;
                                string status = genericsControl.GetProperty("Checked").ToString();
                                if (status == "False")
                                {
                                    Mouse.Click(genericsControl);
                                }
                                break;
                            case "UnCheck":
                                genericsControl.SearchProperties[HtmlCheckBox.PropertyNames.Id] = ControlKeywordValue;
                                status = genericsControl.GetProperty("Checked").ToString();
                                if (status == "True")
                                {
                                    Mouse.Click(genericsControl);
                                }
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlCheckBox.PropertyNames.Id] = ControlKeywordValue;
                                status = genericsControl.GetProperty("Checked").ToString();
                                if (DataRefferencekeyword == "Yes" && status == "False")
                                {
                                    Mouse.Click(genericsControl);
                                }
                                else if (DataRefferencekeyword == "No" && status == "True")
                                {
                                    Mouse.Click(genericsControl);
                                }
                                break;
                        }
                        return true;
                    case "HtmlComboBox":
                        switch (ControlKeyword)
                        {
                            case "Check":
                                genericsControl.SearchProperties[HtmlComboBox.PropertyNames.Id] = ControlKeywordValue;
                                string status = genericsControl.GetProperty("Selected").ToString();
                                if (status == "False")
                                {
                                    Mouse.Click(genericsControl);
                                }
                                break;
                            case "UnCheck":
                                genericsControl.SearchProperties[HtmlComboBox.PropertyNames.Id] = ControlKeywordValue;
                                status = genericsControl.GetProperty("Selected").ToString();
                                if (status == "True")
                                {
                                    Mouse.Click(genericsControl);
                                }
                                break;
                            case "IDVisible":
                                genericsControl.SearchProperties[HtmlComboBox.PropertyNames.Id] = ControlKeywordValue;
                                status = genericsControl.GetProperty("Selected").ToString();
                                d = genericsControl.TryGetClickablePoint(out p);
                                if (d)
                                {
                                    if (DataRefferencekeyword == "Yes" && status == "False")
                                    {
                                        Mouse.Click(genericsControl);
                                    }
                                    else if (DataRefferencekeyword == "No" && status == "True")
                                    {
                                        Mouse.Click(genericsControl);
                                    }
                                }
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlComboBox.PropertyNames.Id] = ControlKeywordValue;
                                status = genericsControl.GetProperty("Selected").ToString();
                                if (DataRefferencekeyword == "Yes" && status == "False")
                                {
                                    Mouse.Click(genericsControl);
                                }
                                else if (DataRefferencekeyword == "No" && status == "True")
                                {
                                    Mouse.Click(genericsControl);
                                }
                                break;
                        }
                        return true;
                    default:
                        Operation.FailerReason = TypeControl + " Control Type not configured for Check Control. Please check the Control Type.";
                        genDetailedReport.Reports(StepId, Description, Operationkeyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                        genDetailedReport.FileCorreptionCheck();
                        throw new NoSuchControlTypeFound(TypeControl + " Control Type not configured for Check Control. Please check the Control Type.");
                }
            }
            catch (UITestControlNotFoundException e)
            {
                genDetailedReport.Reports(StepId, Description, Operationkeyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                Operation.FailerReason = TypeControl + " with " + ControlKeywordValue + " as " + ControlKeyword + " not found in page";
                genDetailedReport.FileCorreptionCheck();
                TakeScreenshot(ControlKeywordValue + " as " + ControlKeyword + " not found");
                Assert.Fail(Operation.FailerReason);
                return false;
            }
            catch (FailedToPerformActionOnBlockedControlException e)
            {
                genericsControl.DrawHighlight();
                TakeScreenshot("ControlCannotbeClicked");
                genDetailedReport.Reports(StepId, Description, Operationkeyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                Operation.FailerReason = TypeControl + " with " + ControlKeywordValue + " as " + ControlKeyword + " is blocked by another control";
                genDetailedReport.FileCorreptionCheck();
                Assert.Fail(Operation.FailerReason);
                return false;
            }
        }
        #endregion

        #region PageLoadWait
        /// <summary>
        /// 
        /// </summary>
        /// <param name="TypeControl">Specify for which Type of control you want to wait for </param>
        /// <param name="ControlKeyword">Property that you used to identify the control</param>
        /// <param name="ControlKeywordValue">What is the value of the property you specified</param>
        /// <param name="steps">Step Number.</param>
        /// <param name="description">Description about the step.</param>
        /// <param name="Operationkeyword">This function will be triggered using Keyword: "WaitPageLoad"</param>
        /// <returns>return true or false value</returns>
        public bool PageLoadWait(string TypeControl, string ControlKeyword, string ControlKeywordValue, string StepId, string Description, string Operationkeyword)
        {
            HtmlControl genericsControl = (HtmlControl)Activator.CreateInstance(typeof(HtmlControl), new object[] { ParentWindow });
            Playback.PlaybackSettings.WaitForReadyLevel = WaitForReadyLevel.UIThreadOnly;
            Playback.PlaybackSettings.WaitForReadyTimeout = 20000;
            Thread.Sleep(2000);
            try
            {
                switch (TypeControl)
                {
                    case "HtmlTable":
                        switch (ControlKeyword)
                        {
                            case "ID":
                                genericsControl.SearchProperties[HtmlTable.PropertyNames.Id] = ControlKeywordValue;
                                genericsControl.WaitForControlReady();
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlTable.PropertyNames.Id] = ControlKeywordValue;
                                genericsControl.WaitForControlReady();
                                break;
                        }
                        return true;
                    case "HtmlDiv":
                        switch (ControlKeyword)
                        {
                            case "ID":
                                genericsControl.SearchProperties[HtmlDiv.PropertyNames.Class] = ControlKeywordValue;
                                genericsControl.WaitForControlReady();
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlDiv.PropertyNames.Id] = ControlKeywordValue;
                                genericsControl.WaitForControlReady();
                                break;
                        }
                        return true;
                    case "HtmlHyperlink":
                        switch (ControlKeyword)
                        {
                            case "ID":
                                genericsControl.SearchProperties[HtmlHyperlink.PropertyNames.Id] = ControlKeywordValue;
                                genericsControl.WaitForControlEnabled();
                                break;
                            case "Class":
                                genericsControl.SearchProperties[HtmlHyperlink.PropertyNames.Class] = ControlKeywordValue;
                                genericsControl.WaitForControlExist();
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlHyperlink.PropertyNames.Id] = ControlKeywordValue;
                                genericsControl.WaitForControlReady();
                                break;
                        }
                        return true;
                    case "HtmlSpan":
                        switch (ControlKeyword)
                        {
                            case "ID":
                                genericsControl.SearchProperties[HtmlSpan.PropertyNames.Id] = ControlKeywordValue;
                                genericsControl.WaitForControlReady();
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlSpan.PropertyNames.Id] = ControlKeywordValue;
                                genericsControl.WaitForControlReady();
                                break;
                        }
                        return true;
                    case "HtmlComboBox":
                        switch (ControlKeyword)
                        {
                            case "ID":
                                genericsControl.SearchProperties[HtmlComboBox.PropertyNames.Id] = ControlKeywordValue;
                                genericsControl.WaitForControlReady();
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlComboBox.PropertyNames.Id] = ControlKeywordValue;
                                genericsControl.WaitForControlReady();
                                break;
                        }
                        return true;
                    case "HtmlButton":
                        switch (ControlKeyword)
                        {
                            case "ID":
                                genericsControl.SearchProperties[HtmlButton.PropertyNames.Id] = ControlKeywordValue;
                                genericsControl.WaitForControlReady();
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlButton.PropertyNames.Id] = ControlKeywordValue;
                                genericsControl.WaitForControlReady();
                                break;
                        }
                        return true;
                    case "HtmlCustom":
                        switch (ControlKeyword)
                        {
                            case "ID":
                                genericsControl.SearchProperties[HtmlCustom.PropertyNames.Id] = ControlKeywordValue;
                                genericsControl.WaitForControlReady();
                                break;
                            default:
                                genericsControl.SearchProperties[HtmlCustom.PropertyNames.Id] = ControlKeywordValue;
                                genericsControl.WaitForControlReady();
                                break;
                        }
                        return true;
                    default:
                        Operation.FailerReason = TypeControl + " Control Type not configured wait for control. Please check the Control Type.";
                        genDetailedReport.Reports(StepId, Description, Operationkeyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                        throw new NoSuchControlTypeFound(TypeControl + " Control Type not configured wait for control. Please check the Control Type.");
                }
            }
            catch (Exception e)
            {
                genDetailedReport.Reports(StepId, Description, Operationkeyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                return false;
            }
        }
        #endregion

        #region GetHoverTableData 
        /// <summary> 
        /// This used to execute script, to get data present in hover table . 
        /// </summary> 
        /// <param name="_DataRefferencekeyword">Data reference keyword value which you want to match</param> 
        /// <param name="Typeofcontrol">Type of control like HtmlDiv etc.</param> 
        /// <param name="TypeOfOperation">Operation which you want to perform</param> 
        /// <param name="ControlKeyword">Property that you used to identify the control</param> 
        /// <param name="ControlKeywordValue">What is the value of the property you specified</param> 
        /// <param name="StepId">Step Number.</param> 
        /// <param name="Description">Description about the step.</param> 
        /// <returns>return specified value</returns> 
        public string GetHoverTableData(string TypeOfOperation, string ControlKeyword, string ControlKeywordValue, string StepId, string Description, string _DataRefferencekeyword, string Typeofcontrol)
        {
            BrowserWindow window = new BrowserWindow();
            try
            {
                string unit = string.Empty, unit2 = string.Empty, UOM = string.Empty, symbol = string.Empty, add = string.Empty, str = string.Empty, result = string.Empty;
                if (_DataRefferencekeyword.Contains(':'))
                {
                    unit = _DataRefferencekeyword.Split(':')[0];
                    unit2 = _DataRefferencekeyword.Split(':')[1];
                    UOM = ExcelDataTable.ReadData(1, unit);
                    symbol = ExcelDataTable.ReadData(1, unit2);
                    add = UOM + symbol;
                }
                switch (Typeofcontrol)
                {
                    case "HtmlDiv":
                        switch (ControlKeyword)
                        {
                            case "MultiRows":
                                string v = window.ExecuteScript("var data=document.getElementById('" + ControlKeywordValue + "').innerText; return data;").ToString();
                                int count = v.IndexOf(add);
                                if (count > 0)
                                {
                                    result = v.Substring(count, add.Length);
                                    if (add.ToString() == result.ToString())
                                    {
                                        return "1";
                                    }
                                    else { return "-1"; }
                                }
                                else
                                {
                                    return "-1";
                                }
                                break;
                            case "SingleRow":
                                string data = window.ExecuteScript("var data=document.getElementById('" + ControlKeywordValue + "').innerText; return data;").ToString();
                                int counts = data.LastIndexOf("\n") + 1;
                                if (counts > 0)
                                {
                                    result = data.Substring(counts, data.Length - counts);
                                    if (add.ToString() == result.ToString())
                                    {
                                        return "1";
                                    }
                                    else { return "-1"; }
                                }
                                else
                                {
                                    return "-1";
                                }
                                break;
                            default:
                                return null;
                        }
                    default:
                        return null;
                }
            }
            catch (Exception e)
            {
                Operation.FailerReason = "Could not find control with specified ID value Please check the Control Keyword Value.";
                genDetailedReport.Reports(StepId, Description, TypeOfOperation, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genDetailedReport.FileCorreptionCheck();
                
                throw new NoSuchControlTypeFound(ControlKeywordValue + " control can not find Please check the Control Keyword Value.");
            }
        }
        #endregion
        #region BrowserTabAvailable
        /// <summary>
        /// if browser tab with specified tab name available it will return true value and close the tab
        /// </summary>
        /// <param name="ControlKeyValue">Browser Window Name Specify</param>
        /// <returns>return the true or false value</returns>
        public bool BrowserTabValidation(string ControlKeyValue)
        {
            BrowserWindow win = new BrowserWindow();
            WinTabList tablist = new WinTabList(win);
            WinTabPage newTab = new WinTabPage(tablist);
            newTab.SearchProperties.Add(WinTabPage.PropertyNames.Name, ControlKeyValue, PropertyExpressionOperator.Contains);
            bool Status = newTab.Exists;
            if (Status)
            {
                WinButton cloase = new WinButton(newTab);
                cloase.SearchProperties.Add(WinButton.PropertyNames.Name, "Close", PropertyExpressionOperator.Contains);
                Mouse.Click(cloase);
                return Status;
            }
            else
            {
                return Status;
            }

        }
        #endregion

        #region Get Data based on the reference data from the data file.
        /// <summary>
        /// This will get the data from the dynamic created excel file. 
        /// </summary>
        /// <param name="ControlKeyword">Property of the control that you do operation.</param>
        /// <param name="TypeOfControl">Type of the control that you want to do operation.</param>
        /// <param name="TypeOfWindow">Type of the window or the technology. like web or window</param>
        /// <param name="ControlKeywordValue">Value of the control keyword value.</param>
        /// <param name="DataReferenceKeyword">Reference to the data that you want to enter or do the operation</param>
        /// <param name="StepNumber">Step Number</param>
        /// <param name="Description">Description for the step</param>
        /// <param name="keyword">"WriteData"</param>
        public void GetAndWriteDataforReff(string ControlKeyword, string TypeOfControl, string TypeOfWindow, string ControlKeywordValue, string DataReferenceKeyword, string StepNumber, string Description, string keyword)
        {
            //First read the recorded excel file from the location 
            string[] RefResources = { };
            string[] readResources = { };
            string ser = WriteAndReadData.DataFilePath;
            ExcelDataTable.PopulateRecordData(WriteAndReadData.DataFilePath);
            string _data = string.Empty;
            //Store=Store1,Material=Material1:Required Quantity+Less Quantity
            //Material=Material1:Batch Number+Less Quantity
            if (DataReferenceKeyword.Contains('+') || DataReferenceKeyword.Contains('-'))
            {
                if (DataReferenceKeyword.Contains('+'))
                {
                    readResources = DataReferenceKeyword.Split('+');
                    _data = getDataFromDynamicExcel(readResources[0]);
                    _data = calculatedata(_data + '+' + ExcelDataTable.ReadData(1, readResources[1])).ToString();
                }
                else if (DataReferenceKeyword.Contains('-'))
                {
                    readResources = DataReferenceKeyword.Split('-');
                    _data = getDataFromDynamicExcel(readResources[0]);
                    _data = calculatedata(_data + '-' + ExcelDataTable.ReadData(1, readResources[1])).ToString();
                }
            }
            else
            {
                _data = getDataFromDynamicExcel(DataReferenceKeyword);
              
            }
            OperationStart(keyword, TypeOfControl, ControlKeyword, ControlKeywordValue, _data, TypeOfWindow, StepNumber, Description);
        }
        /// <summary>
        /// Getting data from dynamic created Excels using a reference value
        /// </summary>
        /// <param name="DataReferenceKeyword">Specify the reference data with heading. "reference data heading=reference data:Required data heading"</param>
        /// <returns>returns required data that matches with the reference data </returns>
        public string getDataFromDynamicExcel(string DataReferenceKeyword)
        {
            //Store=Store1,Material=Material1:Batch Number
            //Material=Material1:Batch Number

            //Store=Store1,Material=Rec_Material:Batch Number
            //Material=Rec_Material:Batch Number

            string store = string.Empty;
            string material = string.Empty;
            string _exldata = string.Empty;
            string order = string.Empty;
            string reference = string.Empty;
            string [] readResources = DataReferenceKeyword.Split(':');
            if (readResources.Count() > 1)
            {
                if (DataReferenceKeyword.Contains("Store") && DataReferenceKeyword.Contains("Material"))
                {
                    string[] RDReferences = readResources[0].Split(',');
                    foreach (var reff in RDReferences)
                    {
                        if (reff.Contains("Store"))
                        {
                            if (reff.Split('=')[1].Contains("Rec_"))
                            {
                                store = Operation.recordedData[reff.Split('=')[1].Replace("Rec_", string.Empty)];
                            }
                            else
                            {
                                store = ExcelDataTable.ReadData(1, reff.Split('=')[1]);
                            }
                            
                        }
                        else if (reff.Contains("Material"))
                        {
                            if (reff.Split('=')[1].Contains("Rec_"))
                            {
                                material = Operation.recordedData[reff.Split('=')[1].Replace("Rec_", string.Empty)];
                            }
                            else
                            {
                                material = ExcelDataTable.ReadData(1, reff.Split('=')[1]);
                            }
                        }
                        else
                        {

                        }
                    }
                    _exldata = ExcelDataTable.ReadRefferenceInfoByMaterialWithStore(store, material, readResources[1]);
                }
                else
                {
                    if (readResources[0].Contains("Order"))
                    {
                        if (readResources[0].Split('=')[1].Contains("Rec_"))
                        {
                            order = Operation.recordedData[readResources[0].Split('=')[1].Replace("Rec_", string.Empty)];
                        }
                        else
                        {
                            order = ExcelDataTable.ReadData(1, readResources[0].Split('=')[1].ToString());
                        }
                        _exldata = ExcelDataTable.ReadRefferenceInfoByOrder(order, readResources[1].ToString());
                    }
                    else if (readResources[0].Contains("Material"))
                    {
                        if (readResources[0].Split('=')[1].Contains("Rec_"))
                        {
                            material = Operation.recordedData[readResources[0].Split('=')[1].Replace("Rec_", string.Empty)];
                        }
                        else
                        {
                            material = ExcelDataTable.ReadData(1, readResources[0].Split('=')[1].ToString());
                        }
                        _exldata = ExcelDataTable.ReadRefferenceInfoByMaterial(material, readResources[1].ToString());
                    }
                    else if (readResources[0].Contains("Store"))
                    {
                        if (readResources[0].Split('=')[1].Contains("Rec_"))
                        {
                            store = Operation.recordedData[readResources[0].Split('=')[1].Replace("Rec_", string.Empty)];
                        }
                        else
                        {
                            store = ExcelDataTable.ReadData(1, readResources[0].Split('=')[1].ToString());
                        }
                        _exldata = ExcelDataTable.ReadRefferenceInfo(store, readResources[1].ToString());
                    }
                    else if (readResources[0].Contains("Reference"))
                    {
                        if (readResources[0].Split('=')[1].Contains("Rec_"))
                        {
                            reference = Operation.recordedData[readResources[0].Split('=')[1].Replace("Rec_", string.Empty)];
                        }
                        else
                        {
                            reference = ExcelDataTable.ReadData(1, readResources[0].Split('=')[1].ToString());
                        }
                        _exldata = ExcelDataTable.ReadRefferenceInfo(reference, readResources[1].ToString());
                    }
                    else
                    {
                        throw new Exception("Reference Resource Not Specified Please Provide Proper Reference");
                    }
                }
            }
            return _exldata;
        }
        #endregion

        #region Window
        public void WindowEnterData(string TypeControl, string ControlKeyword, string ControlKeywordValue, string DataRefferencekeyword)
        {
            WinControl WGenericControl = (WinControl)Activator.CreateInstance(typeof(WinControl), new object[] { mTestApplicationWindow });
            switch (TypeControl)
            {
                case "WinEdit":
                    switch (ControlKeyword)
                    {
                        case "FilePath":
                            UITestControl FilePathData = TopParrentwindow();
                            FilePathData.SearchProperties[WinControl.PropertyNames.Name] = ControlKeywordValue;
                            Thread.Sleep(LoginOperatrion.max);
                            Thread.Sleep(1000);
                            Keyboard.SendKeys(DataRefferencekeyword);
                            Thread.Sleep(LoginOperatrion.mid);
                            auto.Send("{ENTER}");
                            Playback.PlaybackSettings.WaitForReadyLevel = WaitForReadyLevel.Disabled;
                            break;
                        case "ControlName":
                            WGenericControl.SearchProperties[WinControl.PropertyNames.ControlName] = ControlKeywordValue;
                            Keyboard.SendKeys(WGenericControl, DataRefferencekeyword);
                            break;
                        default:
                            WGenericControl.SearchProperties[WinControl.PropertyNames.Name] = ControlKeywordValue;
                            Keyboard.SendKeys(WGenericControl, DataRefferencekeyword);
                            break;
                    }
                    break;
                case "WinComboBox":
                    switch (ControlKeyword)
                    {
                        case "ControlName":
                            WGenericControl.SearchProperties[WinComboBox.PropertyNames.ControlName] = ControlKeywordValue;
                            Keyboard.SendKeys(WGenericControl, DataRefferencekeyword);
                            break;
                        default:
                            WGenericControl.SearchProperties[WinDateTimePicker.PropertyNames.Name] = ControlKeywordValue;
                            Keyboard.SendKeys(WGenericControl, DataRefferencekeyword);
                            break;
                    }
                    break;
                default:
                    Operation.FailerReason = TypeControl + " Control Type not configured for Enter data. Please check the Control Type.";
                    throw new NoSuchControlTypeFound(TypeControl + " Control Type not configured for Enter data. Please check the Control Type.");
            }
        }
        public void WindowClickControl(string TypeControl, string ControlKeyword, string ControlKeywordValue)
        {
            try
            {
                WinControl WGenericControl = (WinControl)Activator.CreateInstance(typeof(WinControl), new object[] { mTestApplicationWindow });
                switch (TypeControl)
                {
                    case "WinSplitButton":
                        switch (ControlKeyword)
                        {
                            case "Name":
                                WGenericControl.SearchProperties[WinControl.PropertyNames.Name] = ControlKeywordValue;
                                Mouse.Click(WGenericControl);
                                break;
                            default:
                                WGenericControl.SearchProperties[WinControl.PropertyNames.Name] = ControlKeywordValue;
                                Mouse.Click(WGenericControl);
                                break;
                        }
                        break;
                    case "WinDateTimePicker":
                        switch (ControlKeyword)
                        {
                            case "ControlName":
                                WGenericControl.SearchProperties[WinDateTimePicker.PropertyNames.ControlName] = ControlKeywordValue;
                                try
                                {
                                    Mouse.Click(WGenericControl);
                                }
                                catch(Exception e)
                                {
                                    if (WGenericControl.HasFocus != true)
                                    {
                                        throw new Exception();
                                    }
                                }
                                break;
                            default:
                                WGenericControl.SearchProperties[WinDateTimePicker.PropertyNames.Name] = ControlKeywordValue;
                                Mouse.Click(WGenericControl);
                                break;
                        }
                        break;
                    case "WinButton":
                        switch (ControlKeyword)
                        {
                            case "FriendlyName":
                                WGenericControl.SearchProperties[WinButton.PropertyNames.FriendlyName] = ControlKeywordValue;
                                Mouse.Click(WGenericControl);
                                break;
                            case "ControlName":
                                WGenericControl.SearchProperties[WinButton.PropertyNames.ControlName] = ControlKeywordValue;
                                try
                                {
                                    Mouse.Click(WGenericControl);
                                }
                                catch(Exception e)
                                {
                                    if (WGenericControl.HasFocus == true)
                                    {
                                        auto.Send("{SPACE}");
                                    }
                                    else
                                    {
                                        throw new Exception();
                                    }
                                }
                                break;
                            default:
                                WGenericControl.SearchProperties[WinButton.PropertyNames.Name] = ControlKeywordValue;
                                try
                                {
                                    Mouse.Click(WGenericControl);
                                }
                                catch (Exception e)
                                {
                                    if (WGenericControl.HasFocus == true)
                                    {
                                        auto.Send("{SPACE}");
                                    }
                                    else
                                    {
                                        throw new Exception();
                                    }
                                }
                                break;
                        }
                        break;
                    case "WinMenuItem":
                        switch (ControlKeyword)
                        {
                            case "Name":
                                WGenericControl.SearchProperties[WinMenuItem.PropertyNames.ControlName] = ControlKeywordValue;
                                Mouse.Click(WGenericControl);
                                break;
                            case "FriendlyName":
                                WGenericControl.SearchProperties[WinMenuItem.PropertyNames.FriendlyName] = ControlKeywordValue;
                                Mouse.Click(WGenericControl);
                                break;
                            default:
                                WGenericControl.SearchProperties[WinMenuItem.PropertyNames.Name] = ControlKeywordValue;
                                Mouse.Click(WGenericControl);
                                break;
                        }
                        break;
                    case "WinMenu":
                        switch (ControlKeyword)
                        {
                            case "Name":
                                WGenericControl.SearchProperties[WinMenu.PropertyNames.Name] = ControlKeywordValue;
                                Mouse.Click(WGenericControl);
                                break;
                            case "FriendlyName":
                                WGenericControl.SearchProperties[WinMenu.PropertyNames.FriendlyName] = ControlKeywordValue;
                                Mouse.Click(WGenericControl);
                                break;
                            default:
                                WGenericControl.SearchProperties[WinMenu.PropertyNames.Name] = ControlKeywordValue;
                                Mouse.Click(WGenericControl);
                                break;
                        }
                        break;
                    case "WinMenuBar":
                        switch (ControlKeyword)
                        {
                            case "Name":
                                WGenericControl.SearchProperties[WinMenuBar.PropertyNames.Name] = ControlKeywordValue;
                                Mouse.Click(WGenericControl);
                                break;
                            default:
                                WGenericControl.SearchProperties[WinMenuBar.PropertyNames.Name] = ControlKeywordValue;
                                Mouse.Click(WGenericControl);
                                break;
                        }
                        break;
                    default:
                        Operation.FailerReason = TypeControl + " Control Type not configured for Click. Please check the Control Type.";
                        throw new NoSuchControlTypeFound(TypeControl + " Control Type not configured for Click. Please check the Control Type.");
                }
            }
            catch (UITestControlNotFoundException e)
            {
                UITestControl testApp = TopParrentWinWindow();
                string path = ConfigurationManager.AppSettings["ScreenShot"];
                path = path + @"\\" + LoginOperatrion.ProjectName + "";
                Directory.CreateDirectory(path);
                try
                {
                    Image image = testApp.CaptureImage();
                    image.Save(path + "\\ControlNotFound.jpeg", ImageFormat.Jpeg);
                    Operation.ErrorScreenPath = path + "\\ControlNotFound.jpeg";
                    image.Dispose();
                }
                catch (Exception v) { }
                Assert.Fail("Control Not Found");
            }
        }
        public void WindowWaitControl(string TypeControl, string ControlKeyword, string ControlKeywordValue)
        {
            WinControl WGenericControl = (WinControl)Activator.CreateInstance(typeof(WinControl), new object[] { mTestApplicationWindow });
            switch (TypeControl)
            {
                case "WinToolBar":
                    switch (ControlKeyword)
                    {
                        case "Name":
                            WGenericControl.SearchProperties[WinControl.PropertyNames.Name] = ControlKeywordValue;
                            WGenericControl.WaitForControlExist();
                            break;
                        default:
                            WGenericControl.SearchProperties[WinControl.PropertyNames.Name] = ControlKeywordValue;
                            WGenericControl.WaitForControlExist();
                            break;
                    }
                    break;
                case "WinButton":
                    switch (ControlKeyword)
                    {
                        case "ControlName":
                            WGenericControl.SearchProperties[WinButton.PropertyNames.ControlName] = ControlKeywordValue;
                            WGenericControl.WaitForControlExist();
                            break;
                        default:
                            WGenericControl.SearchProperties[WinButton.PropertyNames.Name] = ControlKeywordValue;
                            WGenericControl.WaitForControlExist();
                            break;
                    }
                    break;
                case "Window":
                    switch (ControlKeyword)
                    {
                        case "Class":
                            WGenericControl.SearchProperties[WinWindow.PropertyNames.ClassName] = ControlKeywordValue;
                            WGenericControl.WaitForControlReady(2000);
                            break;
                        case "ClassDisappear":
                            WGenericControl.SearchProperties[WinWindow.PropertyNames.ClassName] = ControlKeywordValue;
                            WGenericControl.WaitForControlNotExist(2000);
                            break;
                        default:
                            WGenericControl.SearchProperties[WinWindow.PropertyNames.Name] = ControlKeywordValue;
                            WGenericControl.WaitForControlReady();
                            break;
                    }
                    break;
                default:
                    Operation.FailerReason = TypeControl + " Control Type not configured for Click. Please check the Control Type.";
                    throw new NoSuchControlTypeFound(TypeControl + " Control Type not configured for Click. Please check the Control Type.");
            }

        }
        public string WindowGetControlData(string TypeControl, string ControlKeyword, string ControlKeywordValue, string ScreenshotName, string AssertionMessage, string StepId, string Description, string Operationkeyword)
        {
            string dd = string.Empty;
            try
            {
                WinControl WGenericControl = (WinControl)Activator.CreateInstance(typeof(WinControl), new object[] { mTestApplicationWindow });
                switch (TypeControl)
                {
                    case "WinDateTimePicker":
                        switch (ControlKeyword)
                        {
                            case "ControlName":
                                WGenericControl.SearchProperties[WinDateTimePicker.PropertyNames.ControlName] = ControlKeywordValue.Split('+')[0];
                                var df = WGenericControl.GetChildren();
                                DateTime date = (DateTime)WGenericControl.GetChildren()[3].GetProperty("DateTime");
                                if (ControlKeywordValue.Contains("+Sec"))
                                {
                                    dd = date.ToString(ExcelDataTable.ReadData(1, "DateFormat").ToString() + " hh:mm:ss tt");
                                }
                                else
                                {
                                    dd = date.ToString(ExcelDataTable.ReadData(1, "DateFormat").ToString() + " hh:mm tt");
                                }
                                return dd;
                            default:
                                return null;
                        }
                    default:
                        Operation.FailerReason = TypeControl + " Control Type not configured to Fetch data from control. Please check the Control Type.";
                        genDetailedReport.Reports(StepId, Description, Operationkeyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                        genDetailedReport.FileCorreptionCheck();
                        throw new NoSuchControlTypeFound(TypeControl + " Control Type not configured to Fetch data from control. Please check the Control Type.");
                }
            }
            catch (UITestControlNotFoundException e)
            {
                Operation.FailerReason = e.BasicMessage + AssertionMessage;
                string path = ConfigurationManager.AppSettings["ScreenShot"];
                path = path + @"\" + LoginOperatrion.ProjectName + "";
                Directory.CreateDirectory(path);
                try
                {
                    Image image = window.CaptureImage();
                    image.Save(path + "\\" + ScreenshotName + ".jpeg", ImageFormat.Jpeg);
                    Operation.ErrorScreenPath = path + "\\" + ScreenshotName + ".jpeg";
                    image.Dispose();
                }
                catch (Exception v) { }
                genDetailedReport.Reports(StepId, Description, Operationkeyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                genDetailedReport.FileCorreptionCheck();
                Assert.Fail(e.BasicMessage + AssertionMessage);
                return null;
            }
        }
        public bool WindowMouseMoveClick(string TypeControl, string ControlKeyword, string ControlKeywordValue, string StepId, string Description, string Operationkeyword, MouseButtons Button, params int[] codinates)
        {
            WinControl WGenericControl = (WinControl)Activator.CreateInstance(typeof(WinControl), new object[] { mTestApplicationWindow });
            try
            {
                switch (TypeControl)
                {
                    case "WinDateTimePicker":
                        switch (ControlKeyword)
                        {
                            case "ControlName":
                                WGenericControl.SearchProperties[WinDateTimePicker.PropertyNames.ControlName] = ExcelDataTable.ReadData(1, ControlKeywordValue);
                                int X = WGenericControl.BoundingRectangle.X;
                                int Y = WGenericControl.BoundingRectangle.Y;
                                auto.MouseMove(X + codinates[0], Y + codinates[1], 1000);
                                Mouse.Click(Button);
                                break;
                            default:
                                WGenericControl.SearchProperties[WinDateTimePicker.PropertyNames.Name] = ControlKeywordValue;
                                X = WGenericControl.BoundingRectangle.X;
                                Y = WGenericControl.BoundingRectangle.Y;
                                auto.MouseMove(X + codinates[0], Y + codinates[1], 1000);
                                Mouse.Click(Button);
                                break;
                        }
                        return true;
                    case "WinEdit":
                        switch (ControlKeyword)
                        {
                            case "ControlName":
                                WGenericControl.SearchProperties[WinEdit.PropertyNames.ControlName] = ExcelDataTable.ReadData(1, ControlKeywordValue);
                                int X = WGenericControl.BoundingRectangle.X;
                                int Y = WGenericControl.BoundingRectangle.Y;
                                auto.MouseMove(X + codinates[0], Y + codinates[1], 1000);
                                Mouse.Click(Button);
                                break;
                            default:
                                WGenericControl.SearchProperties[WinEdit.PropertyNames.Name] = ControlKeywordValue;
                                X = WGenericControl.BoundingRectangle.X;
                                Y = WGenericControl.BoundingRectangle.Y;
                                auto.MouseMove(X + codinates[0], Y + codinates[1], 1000);
                                Mouse.Click(Button);
                                break;
                        }
                        return true;
                    case "WinCheckBox":
                        switch (ControlKeyword)
                        {
                            case "ControlName":
                                WGenericControl.SearchProperties[WinCheckBox.PropertyNames.ControlName] = ExcelDataTable.ReadData(1, ControlKeywordValue);
                                int X = WGenericControl.BoundingRectangle.X;
                                int Y = WGenericControl.BoundingRectangle.Y;
                                auto.MouseMove(X + codinates[0], Y + codinates[1], 1000);
                                Mouse.Click(Button);
                                break;
                            default:
                                WGenericControl.SearchProperties[WinCheckBox.PropertyNames.Name] = ControlKeywordValue;
                                X = WGenericControl.BoundingRectangle.X;
                                Y = WGenericControl.BoundingRectangle.Y;
                                auto.MouseMove(X + codinates[0], Y + codinates[1], 1000);
                                Mouse.Click(Button);
                                break;
                        }
                        return true;
                    case "WinButton":
                        switch (ControlKeyword)
                        {
                            case "ControlName":
                                WGenericControl.SearchProperties[WinButton.PropertyNames.ControlName] = ExcelDataTable.ReadData(1, ControlKeywordValue);
                                int X = WGenericControl.BoundingRectangle.X;
                                int Y = WGenericControl.BoundingRectangle.Y;
                                auto.MouseMove(X + codinates[0], Y + codinates[1], 1000);
                                Mouse.Click(Button);
                                break;
                            default:
                                WGenericControl.SearchProperties[WinButton.PropertyNames.Name] = ControlKeywordValue;
                                X = WGenericControl.BoundingRectangle.X;
                                Y = WGenericControl.BoundingRectangle.Y;
                                auto.MouseMove(X + codinates[0], Y + codinates[1], 1000);
                                Mouse.Click(Button);
                                break;
                        }
                        return true;
                    default:
                        Operation.FailerReason = TypeControl + " Control Type not configured for Mouse Move in Test App. Please check the Control Type.";
                        genDetailedReport.Reports(StepId, Description, Operationkeyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                        genDetailedReport.FileCorreptionCheck();
                        throw new NoSuchControlTypeFound(TypeControl + " Control Type not configured for Mouse Move In Test App. Please check the Control Type.");
                }
            }
            catch (UITestControlNotFoundException e)
            {
                genDetailedReport.Reports(StepId, Description, Operationkeyword, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                Operation.FailerReason = TypeControl + " with " + ControlKeywordValue + " as " + ControlKeyword + " not found in page";
                genDetailedReport.FileCorreptionCheck();
                Assert.Fail(Operation.FailerReason);
                return false;
            }

        }
        /// <summary>
        /// This is used to click on the menu and sub menu of the test application
        /// </summary>
        /// <param name="TypeControl">Specify 'Name' as your Control Type</param>
        /// <param name="ControlKeyword">"WindowMenuClick" will used as the Key word to click the menu and sub menu of that menu</param>
        /// <param name="ControlKeywordValue">Here specify the Menu and Sub menu  for that menu to be clicked</param>
        public void WindowMenuClick(string TypeControl, string ControlKeyword, string ControlKeywordValue)
        {
            UITestControl testApp = TopParrentWinWindow();
            try
            {
                WinMenuItem menu = new WinMenuItem(testApp);
                menu.SearchProperties[UITestControl.PropertyNames.Name] = ControlKeywordValue.Split(':')[0].ToString();
                Mouse.Click(menu);

                WinMenuItem Submenu = new WinMenuItem(menu);
                Submenu.SearchProperties[UITestControl.PropertyNames.Name] = ControlKeywordValue.Split(':')[1].ToString();
                Mouse.Click(Submenu);
            }
            catch(Exception e)
            {
                string path = ConfigurationManager.AppSettings["ScreenShot"];
                path = path + @"\\" + LoginOperatrion.ProjectName + "";
                Directory.CreateDirectory(path);
                try
                {
                    Image image = testApp.CaptureImage();
                    image.Save(path + "\\ControlNotFound.jpeg", ImageFormat.Jpeg);
                    Operation.ErrorScreenPath = path + "\\ControlNotFound.jpeg";
                    image.Dispose();
                }
                catch (Exception v) { }
                Assert.Fail("Control Not Found");
            }
            
        }
        public bool WindowClearControl(string TypeControl, string ControlKeyword, string ControlKeywordValue, string Step , string Description ,string typeofoperation  )
        {
            try
            {
                WinControl WGenericControl = (WinControl)Activator.CreateInstance(typeof(WinControl), new object[] { mTestApplicationWindow });
                switch (TypeControl)
                {
                    case "WinEdit":
                        switch (ControlKeyword)
                        {
                            case "ControlName":
                                WGenericControl.SearchProperties[WinEdit.PropertyNames.ControlName] = ControlKeywordValue;
                                Mouse.Click(WGenericControl);
                                Thread.Sleep(1000);
                                auto.Send("{BACKSPACE 20}");
                                auto.Send("{DEL 20}");
                                break;
                            default:
                                WGenericControl.SearchProperties[WinEdit.PropertyNames.Name] = ControlKeywordValue;
                                Mouse.Click(WGenericControl);
                                Thread.Sleep(1000);
                                auto.Send("{BACKSPACE 20}");
                                auto.Send("{DEL 20}");
                                break;
                        }
                        return true;
                    default:
                        Operation.FailerReason = TypeControl + " Control Type not configured to clear data in Windows application. Please check the Control Type.";
                        genDetailedReport.Reports(Step, Description, typeofoperation, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                        genDetailedReport.FileCorreptionCheck();
                        throw new NoSuchControlTypeFound(TypeControl + " Control Type not configured to clear data in Windows application. Please check the Control Type.");
                }
            }
            catch (UITestControlNotFoundException e)
            {
                genDetailedReport.Reports(Step, Description, typeofoperation, false, LoginOperatrion.batchforReport, LoginOperatrion.DetaildReportStatus, "");
                Operation.FailerReason = TypeControl + " with " + ControlKeywordValue + " as " + ControlKeyword + " not found in Test App";
                genDetailedReport.FileCorreptionCheck();
                Assert.Fail(Operation.FailerReason);
                return false;
            }
        }
        #endregion

        #region basic Web
        private BrowserWindow mParentwindow;
        public BrowserWindow ParentWindow
        {
            get
            {
                if (this.mParentwindow == null)
                {
                    this.mParentwindow = TopParrentwindow();
                }
                return this.mParentwindow;
            }
        }
        public BrowserWindow TopParrentwindow()
        {
            BrowserWindow widow = new BrowserWindow();
            widow.SearchProperties[UITestControl.PropertyNames.ClassName] = BrowserWindow.CurrentBrowser.ToString();
            return widow;
        }
        /// <summary>
        /// This will calculate two data given in the data reference. and return the result to the called function according to the symbol
        /// addition and subtraction is allowed.
        /// </summary>
        /// <param name="datatoCalculate">Data reference of the two values separated by the operation like addition and subtraction.</param>
        /// <returns></returns>
        public int calculatedata(string datatoCAlculate)
        {
            int result = 0;
            if (datatoCAlculate.Contains('+'))
            {
                result = Convert.ToInt32(datatoCAlculate.Split('+')[0]) + Convert.ToInt32(datatoCAlculate.Split('+')[1]);
            }
            else if (datatoCAlculate.Contains('-'))
            {
                result = Convert.ToInt32(datatoCAlculate.Split('-')[0]) - Convert.ToInt32(datatoCAlculate.Split('-')[1]);
            }
            else if (datatoCAlculate.Contains('*'))
            {
                result = Convert.ToInt32(datatoCAlculate.Split('*')[0]) * Convert.ToInt32(datatoCAlculate.Split('*')[1]);
            }
            else
            {
                throw new Exception("No Calculation Symbol Not Found..Calculation is not possible");
            }
            return result;
        }
        #endregion


        #region  Basic Window
        private WinWindow mTestApplicationWindow;

        public WinWindow WinParentWindow
        {
            get
            {
                if (this.mTestApplicationWindow == null)
                {
                    this.mTestApplicationWindow = TopParrentWinWindow();
                }
                return this.mTestApplicationWindow;
            }
        }

        public WinWindow TopParrentWinWindow()
        {
            WinWindow WinWin = new WinWindow();
            WinWin.TechnologyName = "MSAA";
            WinWin.SearchProperties.Add(new PropertyExpression(WinWindow.PropertyNames.ClassName,"WindowsForms10.Window",PropertyExpressionOperator.Contains));
            return WinWin;
        }
        #endregion
    }
}
