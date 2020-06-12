using Prism.Ioc;
using System;
using System.Windows;

namespace QuickOutputCert
{
    /// <summary>
    /// App.xaml 的互動邏輯
    /// </summary>
    public partial class App
    {
        protected override void RegisterTypes(IContainerRegistry containerRegistry)
        {  
        }

        protected override Window CreateShell()
        {
            return Container.Resolve<MainWindow>();
        }
    }
}
