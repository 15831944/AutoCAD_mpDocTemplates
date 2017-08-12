using System.ComponentModel;
using System.Globalization;
using System.Windows;
using mpSettings;
using ModPlus;

namespace mpDocTemplates
{
    /// <summary>
    /// Логика взаимодействия для ExportProgressDialog.xaml
    /// </summary>
    public partial class ExportProgressDialog
    {
        readonly BackgroundWorker _backgroundWorker = new BackgroundWorker
        {
            WorkerSupportsCancellation = true,
            WorkerReportsProgress = true
        };

        public ExportProgressDialog(string whyWeAreWaiting, DoWorkEventHandler work)
        {
            InitializeComponent();
            MpWindowHelpers.OnWindowStartUp(
                this,
                MpSettings.GetValue("Settings", "MainSet", "Theme"),
                MpSettings.GetValue("Settings", "MainSet", "AccentColor"),
                MpSettings.GetValue("Settings", "MainSet", "BordersType")
                );
            this.Title.Text = whyWeAreWaiting; // Show in title bar
            _backgroundWorker.DoWork += work; // Event handler to be called in context of new thread.
            _backgroundWorker.ProgressChanged += backgroundWorker_ProgressChanged;
            _backgroundWorker.RunWorkerCompleted += backgroundWorker_RunWorkerCompleted;
        }

        void backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            this.Close();
        }

        void backgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            // Если значение процентов = 0, значит не показываем текст с процентами
            this.TbProgress.Visibility = e.ProgressPercentage == -1 ? Visibility.Collapsed : Visibility.Visible;
            // Процент
            this.ProgressBar.Value = e.ProgressPercentage;
            // Процент в виде текста
            this.TbProgress.Text = e.ProgressPercentage.ToString(CultureInfo.InvariantCulture) + "%";
            // Что сейчас делаем
            this.TbCurrentWorkLabel.Text = e.UserState as string;
        }

        private void BtCancel_OnClick(object sender, RoutedEventArgs e)
        {
            this.TbCurrentWorkLabel.Text = "Отмена...";
            _backgroundWorker.CancelAsync(); // Tell worker to abort.
            this.BtCancel.IsEnabled = false;
        }

        private void ExportProgressDialog_OnLoaded(object sender, RoutedEventArgs e)
        {
            _backgroundWorker.RunWorkerAsync();
        }
    }
}
