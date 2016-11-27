using CanYouLib.ExcelLib;
using GalaSoft.MvvmLight.Command;

namespace LjcDesktopApp.ViewModel
{
    public class AutoPlanViewModel
    {
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
                       var sourceFileName = dialog.FileName;
                       var ds = importExcel.ImportDataSet(sourceFileName, false, 1, 1);
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
                });
            }
        }
    }
}
