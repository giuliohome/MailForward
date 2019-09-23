using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace MailForward
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private OutlookHelper outlookHelper;
        public MainWindow()
        {
            InitializeComponent();
            outlookHelper = new OutlookHelper();
            DataContext = outlookHelper;
        }

        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            await outlookHelper.ReadConfig();
        }


        private async void BtnForwardFolder_Click(object sender, RoutedEventArgs e)
        {
            await outlookHelper.ReadConfig();
            await outlookHelper.ForwardItems();
        }

        private void BtnExit_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private async void BtnSelectFolder_Click(object sender, RoutedEventArgs e)
        {
            await outlookHelper.ReadConfig();
            await outlookHelper.SelectFolder();
            outlookHelper.DisplayFolder();
        }

        private async void BtnDisplayFolder_Click(object sender, RoutedEventArgs e)
        {
            await outlookHelper.ReadConfig();
            outlookHelper.DisplayFolder();
        }

        private async void BtnSettings_Click(object sender, RoutedEventArgs e)
        {
            await outlookHelper.ReadConfig();
            var settings = new Settings
            {
                DataContext = outlookHelper
            };
            settings.ShowDialog();
            await outlookHelper.SaveConfig();
        }
    }
}
