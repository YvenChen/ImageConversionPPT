using System.IO;
using System.Windows;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace ImageConversionPPT
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }


        /// <summary>
        /// 选择图片文件夹路径
        /// </summary>
        private void BtnPath_Click(object sender, RoutedEventArgs e)
        {
            CommonOpenFileDialog dlg = new();
            dlg.IsFolderPicker = true;
            //dlg.InitialDirectory = Environment.CurrentDirectory;

            if (dlg.ShowDialog() == CommonFileDialogResult.Ok)
            {
                TebPath.Text = dlg.FileName;
            }
        }


        /// <summary>
        /// Start
        /// </summary>
        private async void BtnStart_Click(object sender, RoutedEventArgs e)
        {
            DisableButton();

            string[] dirs = Directory.GetDirectories(TebPath.Text, "*", SearchOption.AllDirectories);

            foreach (string dirPath in dirs)
            {
                ICP.ImagePath = dirPath;
                bool result = await ICP.MainProcess();
                if (!result)
                {
                    MessageBox.Show("执行失败，请稍后再试。", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
            }
            
            MessageBox.Show("执行完成！", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            EnableButton();
        }


        /// <summary>
        /// 禁用按钮
        /// </summary>
        private void DisableButton()
        {
            BtnStart.IsEnabled = false;
            BtnPath.IsEnabled = false;
        }


        /// <summary>
        /// 启用按钮
        /// </summary>
        private void EnableButton()
        {
            BtnStart.IsEnabled = true;
            BtnPath.IsEnabled = true;
        }
    }
}
