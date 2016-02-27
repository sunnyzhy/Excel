using Microsoft.Win32;
using System.Windows;
using System;
using System.Windows.Media;
using G.Excel.Logic;
using System.Windows.Input;
using System.Windows.Controls;
using System.Windows.Media.Imaging;

namespace G.Excel.Project
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        private SolidColorBrush _red;
        private SolidColorBrush _green;

        public MainWindow()
        {
            InitializeComponent();

            this.DataContext = new
            {
                ImageSource = new BitmapImage(new Uri("pack://application:,,,/G.Excel.Project;component/excel.ico", UriKind.Absolute)),
                Title = "Excel数据表统计"
            };
            _red = new SolidColorBrush(Color.FromArgb(255, 255, 0, 0));
            _green = new SolidColorBrush(Color.FromArgb(255, 0, 128, 0));
        }

        private void Window_MouseLeftButtonDown_1(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (e.ButtonState == MouseButtonState.Pressed)
            {
                this.DragMove();
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            ShowMessage(string.Empty);
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Multiselect = false;
            ofd.Filter = "Excel Files (*.xls, *.xlsx)|*.xls;*.xlsx";
            if (ofd.ShowDialog() != true)
            {
                return;
            }
            this.txtFilePath.Text = ofd.FileName;
        }

        private void btnExecute_Click_1(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(this.txtFilePath.Text.Trim()))
            {
                this.txtMsg.Foreground = _red;
                ShowMessage("请先选择Excel文件");
                return;
            }
            this.txtMsg.Foreground = _green;

            AnalysisLib.CreateInstance().SetButtonEnableEvent += SetButtonEnable;
            AnalysisLib.CreateInstance().ShowMessageEvent += ShowMessage;
            AnalysisLib.CreateInstance().AnalysisData(this.txtFilePath.Text.Trim());
        }

        /// <summary>
        /// 设置按钮是否可用
        /// </summary>
        /// <param name="value"></param>
        private void SetButtonEnable(bool value)
        {
            this.btnExecute.Dispatcher.Invoke(new Action(() => this.btnExecute.IsEnabled = value));
        }

        /// <summary>
        /// 显示提示信息
        /// </summary>
        /// <param name="msg"></param>
        private void ShowMessage(string msg)
        {
            this.txtMsg.Dispatcher.Invoke(new Action(() => this.txtMsg.Text = msg));
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            Environment.Exit(0);
        }


    }
}
