using Microsoft.Win32;
using Prism.Commands;
using Prism.Mvvm;
using QuickOutputCert.Models;
using QuickOutputCert.Services;
using System.Collections.Generic;
using System.IO;
using System.Windows;

namespace QuickOutputCert.ViewModels
{
    public class MainViewModel : BindableBase
    {
        #region Private Properties
        private List<DataModel> ObcDataModel = new List<DataModel>();
        #endregion
        #region Public Propertis
        public string FileName { get; set; }
        #endregion
        #region Command
        public DelegateCommand InputCommand { get; }
        public DelegateCommand OutputCommand { get; }
        #endregion
        public MainViewModel()
        {
            InputCommand = new DelegateCommand(() =>
            {
                OpenFileDialog openFileDialog = new OpenFileDialog
                {
                    Filter = "Excel File(*.xlsx)|*.xlsx",
                    RestoreDirectory = true
                };
                if (openFileDialog.ShowDialog().Value)
                {
                    FileName = Path.GetFileNameWithoutExtension(openFileDialog.FileName);
                    ObcDataModel = ExcelDataSort.GetData(openFileDialog.FileName);
                }
            });
            OutputCommand = new DelegateCommand(() =>
            {
                if (ObcDataModel.Count < 1)
                    MessageBox.Show("資料不能為空");
                else
                    FileName.OutputReportDetail(ObcDataModel);
            });
        }
    }
}
