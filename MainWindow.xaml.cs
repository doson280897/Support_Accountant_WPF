using System;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Threading;
using MessageBox = System.Windows.MessageBox;

namespace Support_Accountant
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private DispatcherTimer curTimeTimer;

        public MainWindow()
        {
            InitializeComponent();
            InitializeTimer();
        }

        private void TabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        #region Timer Management
        private void InitializeTimer()
        {
            curTimeTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(1)
            };
            curTimeTimer.Tick += CurTimeTimer_Tick;
            curTimeTimer.Start();
        }

        private void CurTimeTimer_Tick(object sender, EventArgs e)
        {
            label_CurTime.Content = DateTime.Now.ToString("ddd dd/MM/yyyy - HH:mm:ss", CultureInfo.InvariantCulture);
        }
        #endregion

        #region Menu Management
        private void Help_About_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Support Accountant v1.0\nDeveloped by SonDNT\n© 2025", "About Support Accountant", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        #endregion

        private void btn_Browse_Rename_Click(object sender, RoutedEventArgs e)
        {
            using (var folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "Select XML/PDF Folder";
                folderDialog.ShowNewFolderButton = false;

                // In WPF, use ShowDialog() with a Window handle
                var result = folderDialog.ShowDialog();

                if (result == System.Windows.Forms.DialogResult.OK)
                {
                    txtBox_Browse_Rename.Text = folderDialog.SelectedPath;
                }
            }
        }
    }
}
