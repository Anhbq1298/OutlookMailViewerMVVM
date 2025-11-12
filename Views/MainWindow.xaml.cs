using System.Windows;
using OutlookMailViewerMVVM.ViewModels;

namespace OutlookMailViewerMVVM.Views
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            DataContext = new MainViewModel();
        }
    }
}
