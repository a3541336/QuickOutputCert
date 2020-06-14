using Prism.Mvvm;
using Prism.Commands;
using QuickOutputCert.Models;
using System.Collections.ObjectModel;
using Microsoft.Win32;
using System.Windows.Interop;
using System.Windows;
using System.IO;
using QuickOutputCert.Services;
using System.Windows.Documents;
using System.Collections.Generic;

namespace QuickOutputCert.ViewModels
{
    public class MainViewModel: BindableBase
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
                //openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;
                if (openFileDialog.ShowDialog().Value)
                {
                    FileName = Path.GetFileName(openFileDialog.FileName);
                    ObcDataModel = ExcelDataSort.GetData(openFileDialog.FileName);
                }
            });
            OutputCommand = new DelegateCommand(() =>
            {
                OutputWordData.OutputReportDetail(ObcDataModel);
            });
        }
    }
}
 