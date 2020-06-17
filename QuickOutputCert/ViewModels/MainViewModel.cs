using Microsoft.Win32;
using Prism.Commands;
using Prism.Mvvm;
using QuickOutputCert.Models;
using QuickOutputCert.Services;
using System.Collections.Generic;
using System.IO;

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
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;
                if (openFileDialog.ShowDialog().Value)
                {
                    FileName = Path.GetFileNameWithoutExtension(openFileDialog.FileName);
                    ObcDataModel = ExcelDataSort.GetData(openFileDialog.FileName);
                }
            });
            OutputCommand = new DelegateCommand(() =>
            {
                FileName.OutputReportDetail(ObcDataModel);
            });
        }
    }
}
