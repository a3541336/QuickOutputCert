using Prism.Mvvm;
using Prism.Commands;
using QuickOutputCert.Models;
using System.Collections.ObjectModel;
using Microsoft.Win32;

namespace QuickOutputCert.ViewModels
{
    public class MainViewModel
    {
        #region Private Properties
        private ObservableCollection<DataModel> ObcDataModel = new ObservableCollection<DataModel>();
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
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;
                openFileDialog.ShowDialog();
            });
        }
    }
}
