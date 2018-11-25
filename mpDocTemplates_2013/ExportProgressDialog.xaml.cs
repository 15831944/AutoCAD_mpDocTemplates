using System.ComponentModel;
using System.Globalization;
using System.Windows;

namespace mpDocTemplates
{
    public partial class ExportProgressDialog
    {
        private const string LangItem = "mpDocTemplates";
        readonly BackgroundWorker _backgroundWorker = new BackgroundWorker
        {
            WorkerSupportsCancellation = true,
            WorkerReportsProgress = true
        };

        public ExportProgressDialog(string whyWeAreWaiting, DoWorkEventHandler work)
        {
            InitializeComponent();
            Title.Text = whyWeAreWaiting; // Show in title bar
            _backgroundWorker.DoWork += work; // Event handler to be called in context of new thread.
            _backgroundWorker.ProgressChanged += backgroundWorker_ProgressChanged;
            _backgroundWorker.RunWorkerCompleted += backgroundWorker_RunWorkerCompleted;
        }

        void backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            Close();
        }

        void backgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            // Если значение процентов = 0, значит не показываем текст с процентами
            TbProgress.Visibility = e.ProgressPercentage == -1 ? Visibility.Collapsed : Visibility.Visible;
            // Процент
            ProgressBar.Value = e.ProgressPercentage;
            // Процент в виде текста
            TbProgress.Text = e.ProgressPercentage.ToString(CultureInfo.InvariantCulture) + "%";
            // Что сейчас делаем
            TbCurrentWorkLabel.Text = e.UserState as string;
        }

        private void BtCancel_OnClick(object sender, RoutedEventArgs e)
        {
            TbCurrentWorkLabel.Text = ModPlusAPI.Language.GetItem(LangItem, "h18") + "...";
            _backgroundWorker.CancelAsync(); // Tell worker to abort.
            BtCancel.IsEnabled = false;
        }

        private void ExportProgressDialog_OnLoaded(object sender, RoutedEventArgs e)
        {
            _backgroundWorker.RunWorkerAsync();
        }
    }
}
