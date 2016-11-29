using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using CanYouLib.ExcelLib;
using CanYouLib.ExcelLib.Utility;
using GalaSoft.MvvmLight.Command;
using LjcDesktopApp.Models;
using Microsoft.Win32;

namespace LjcDesktopApp.ViewModel
{
    public class AutoPlanViewModel
    {
        private IList<TaskModel> _taskModels;
        private string _sourceFileName;

        /// <summary>
        /// 导入计划Excel
        /// </summary>
        public RelayCommand ImportPlanCommand
        {
            get
            {
                return new RelayCommand(() =>
               {
                   var importExcel = new ImportExcel();
                   var dialog =
                       new Microsoft.Win32.OpenFileDialog { Filter = "excel|*.xls" };
                   if (dialog.ShowDialog() == true)
                   {
                       _sourceFileName = dialog.FileName;
                       var ds = importExcel.ImportDataSet(_sourceFileName, false, 1, 0);
                       var dt = ds.Tables[0];
                       _taskModels = ModelConverter<TaskModel>.ConvertToModel(dt);
                       CalSchedule(_taskModels);
                   }
               });
            }
        }

        /// <summary>
        /// 导出计划Excel
        /// </summary>
        public RelayCommand ExportPlanCommand
        {
            get
            {
                return new RelayCommand(() =>
                {
                    var sheetName = "Sheet1";
                    var exportExcel = new ExportExcel();

                    string rootPath = AppDomain.CurrentDomain.BaseDirectory;
                    var fileId = Path.GetFileNameWithoutExtension(_sourceFileName) + DateTime.Now.ToString("yyyyMMddHHmmss");//避免文件重复
                    var tempFilePath = rootPath + "\\" + fileId + ".xls";//存放临时文件的路径
                    exportExcel.CreateExcel(sheetName, 1);
                    var docuSum = new DocumentSummary()
                    {
                        ApplicationName = "AutoSchedule",
                        Author = "ljc",
                        //FirstRow = 0
                    };

                    var fileData = exportExcel.ExportData(_taskModels, sheetName, rootPath, docuSum);
                    var file = new FileStream(tempFilePath, FileMode.Create);
                    file.Write(fileData, 0, fileData.Length - 1);
                    file.Close();

                    #region 执行导出逻辑
                    using (ExcelExpert report = new ExcelExpert(tempFilePath))
                    {
                        SaveFileDialog dialog = new SaveFileDialog
                        {
                            Filter = "文档|*.xls;*.xlsx",
                            FileName = fileId,
                            RestoreDirectory = true
                        };
                        if (dialog.ShowDialog() == true)
                        {
                            report.Save(dialog.FileName);

                            if (MessageBox.Show("下载成功！是否打开文件？") == MessageBoxResult.OK)
                            {
                                System.Diagnostics.Process.Start(dialog.FileName);
                            }
                            File.Delete(tempFilePath);//删除临时文件
                        }
                        else
                        {
                            File.Delete(tempFilePath);//删除临时文件
                        }
                    }
                    #endregion
                });
            }
        }

        private void CalSchedule(IList<TaskModel> list)
        {
            var members = new List<string>();
            var memberStrList = list.Select(p => p.TaskMember).Distinct().ToList();
            foreach (var memberStr in memberStrList)
            {
                if (memberStr.Contains("、"))
                {
                    var subMembers = memberStr.Split('、').ToList();
                    foreach (var subMember in subMembers)
                    {
                        if (!members.Contains(subMember))
                        {
                            members.Add(subMember);
                        }
                    }
                }
                else if (!members.Contains(memberStr))
                {
                    members.Add(memberStr);
                }
            }

            foreach (var taskModel in list)
            {
                taskModel.PlanEndTime = null;
            }
            foreach (var member in members)
            {
                var subList = list.Where(p => p.TaskMember.Contains(member)).ToList();
                if (subList.Count == 0 || subList.First().PlanStartTime == null)
                {
                    continue;
                }

                string dateFormat = "yyyy/MM/dd HH:mm:ss";
                //string dateFormat = "yyyy/MM/dd";
                var firstStartTime = DateTime.Parse(subList.First().PlanStartTime);
                DateTime lastEndTime = firstStartTime;
                foreach (var taskModel in subList)
                {
                    taskModel.HolidayRemark = null;//先清空

                    var startTime = lastEndTime;
                    taskModel.PlanStartTime = startTime.ToString(dateFormat);
                    var spentDays = double.Parse(taskModel.PlanSpentDays);


                    if (startTime.AddDays(spentDays).Hour != 0 && startTime.AddDays(spentDays).Date.ToString(dateFormat) == "2016/12/08"
                          || startTime.AddDays(spentDays).Date.ToString(dateFormat) == "2016/12/09"
                          || startTime.AddDays(spentDays).Date.ToString(dateFormat) == "2016/12/10"
                            || startTime.AddDays(spentDays).Date.ToString(dateFormat) == "2016/12/11")
                    //搬家2天+周末2天
                    {
                        spentDays += 4;
                        taskModel.HolidayRemark = "搬家2天+周末2天";
                    }
                    else if (startTime.AddDays(spentDays).Hour != 0 && startTime.AddDays(spentDays).Date.ToString(dateFormat) == "2016/12/31"
                          || startTime.AddDays(spentDays).Date.ToString(dateFormat) == "2017/01/01"
                          || startTime.AddDays(spentDays).Date.ToString(dateFormat) == "2017/01/02")
                    //元旦放假3天
                    {
                        spentDays += 3;
                        taskModel.HolidayRemark = "元旦放假3天";
                    }
                    else if (startTime.AddDays(spentDays).Hour != 0 && startTime.AddDays(spentDays).DayOfWeek == DayOfWeek.Saturday
                           || startTime.AddDays(spentDays).DayOfWeek == DayOfWeek.Sunday)
                    //周末两天休假
                    {
                        spentDays += 2;
                        taskModel.HolidayRemark = "周末两天休假";
                    }

                    var endTime = startTime.AddDays(spentDays);

                    var endDateStr = endTime.Hour == 0 ? endTime.AddDays(-1).ToString(dateFormat) : endTime.ToString(dateFormat);
                    if (taskModel.PlanEndTime == null)
                    {
                        taskModel.PlanEndTime = endDateStr;
                    }
                    else if (taskModel.PlanEndTime != endDateStr)
                    //多人任务但排的计划时间不一致时，分别显示
                    {
                        taskModel.PlanEndTime += "\n" + member + ":" + endDateStr;
                    }

                    if (startTime.AddDays(spentDays).Hour == 0 && startTime.AddDays(spentDays).Date.ToString(dateFormat) == "2016/12/08")
                    //搬家2天+周末2天
                    {
                        lastEndTime = endTime.AddDays(4);
                        taskModel.HolidayRemark = endDateStr + "之后是搬家2天+周末2天";
                    }
                    else if (startTime.AddDays(spentDays).Hour == 0 && startTime.AddDays(spentDays).Date.ToString(dateFormat) == "2016/12/31")
                    //元旦放假3天
                    {
                        lastEndTime = endTime.AddDays(3);
                        taskModel.HolidayRemark = endDateStr + "之后是元旦3天假期";
                    }
                    else if (endTime.Hour == 0 && endTime.DayOfWeek == DayOfWeek.Saturday)
                    //周末两天休假
                    {
                        lastEndTime = endTime.AddDays(2);
                        taskModel.HolidayRemark = endDateStr + "之后是周末2天假期";
                    }
                    else
                    {
                        lastEndTime = endTime;
                    }
                }
            }

        }

    }
}
