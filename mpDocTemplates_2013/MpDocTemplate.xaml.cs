namespace mpDocTemplates
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.ComponentModel;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using System.Threading;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Input;
    using Autodesk.AutoCAD.DatabaseServices;
    using Autodesk.AutoCAD.Internal;
    using ModPlusAPI;
    using ModPlusAPI.IO.Office.Word;
    using ModPlusAPI.IO.Office.Word.FindAndReplace;
    using ModPlusAPI.Windows;
    using AcApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;
    using Visibility = System.Windows.Visibility;

    public partial class MpDocTemplate
    {
        private const string LangItem = "mpDocTemplates";

        private ObservableCollection<TemplateItem> _kapDocs;
        private ObservableCollection<TemplateItem> _linDocs;
        private List<TextBox> _textBoxes;
        private List<string> _toReplace;

        private readonly List<string> _fieldNames = new List<string>
        {
            "NormKontrol",
            "zG2",
            "Author",
            "ChiefEngineer",
            "GIP",
            "zG1",
            "zG9"
        };

        public MpDocTemplate()
        {
            InitializeComponent();
            Title = ModPlusAPI.Language.GetItem(LangItem, "h1");
        }

        #region windows standard

        private void MpDocTemplate_OnMouseEnter(object sender, MouseEventArgs e)
        {
            Focus();
        }

        private void MpDocTemplate_OnMouseLeave(object sender, MouseEventArgs e)
        {
            Utils.SetFocusToDwgView();
        }
        
        #endregion

        // Окно загрузилось
        private void MpDocTemplate_OnLoaded(object sender, RoutedEventArgs e)
        {
            try
            {
                _textBoxes = new List<TextBox>
                {
                    TbController,
                    TbDescription,
                    TbEmployer,
                    TbEngineer,
                    TbGIP,
                    TbNumProj,
                    TbOrganization,
                    TbResolution,
                    TbCustomer
                };

                // Данные из файла настроек
                foreach (var tb in _textBoxes)
                {
                    tb.Text = UserConfigFile.GetValue(LangItem, tb.Name);
                }

                // Создаем новые коллекции, заполняем и биндим их
                _kapDocs = new ObservableCollection<TemplateItem>();
                TemplateData.FillKapItems(_kapDocs);
                LbKap.ItemsSource = _kapDocs;
                _linDocs = new ObservableCollection<TemplateItem>();
                TemplateData.FillLinItems(_linDocs);
                LbLin.ItemsSource = _linDocs;
            }
            catch (System.Exception exception)
            {
                ExceptionBox.Show(exception);
            }
        }

        // Окно закрылось
        private void MpDocTemplate_OnClosed(object sender, EventArgs e)
        {
            try
            {
                // ReSharper disable once InvertIf
                if (_textBoxes != null)
                {
                    foreach (var tb in _textBoxes)
                    {
                        UserConfigFile.SetValue(LangItem, tb.Name, tb.Text, false);
                    }

                    UserConfigFile.SaveConfigFile();
                }
            }
            catch (System.Exception exception)
            {
                ExceptionBox.Show(exception);
            }
        }

        // Взять из полей
        private void BtGetFromFields_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                var dbsi = AcApp.DocumentManager.MdiActiveDocument.Database.SummaryInfo;
                var dbsib = new DatabaseSummaryInfoBuilder(dbsi);
                for (var i = 0; i < _fieldNames.Count; i++)
                {
                    if (dbsib.CustomPropertyTable.Contains(_fieldNames[i]))
                        _textBoxes[i].Text = dbsib.CustomPropertyTable[_fieldNames[i]].ToString();
                }
            }
            catch (System.Exception exception)
            {
                ExceptionBox.Show(exception);
            }
        }

        private System.Exception _error;
        private bool _hasError;
        private List<string> _fileToDelete;

        // Создать
        private void BtCreate_OnClick(object sender, RoutedEventArgs e)
        {
            _error = new System.Exception(); // Ошибка, возможная при асинхронной работе
            _fileToDelete = new List<string>(); // Список файлов на удаление
            if (!_kapDocs.Any(x => x.Create) & !_linDocs.Any(x => x.Create))
            {
                ModPlusAPI.Windows.MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "h19"));
                return;
            }

            // Значения в текстовых полях в порядке поиска и замены
            _toReplace = new List<string>
            {
                TbDescription.Text,
                TbGIP.Text, TbEngineer.Text,
                TbEmployer.Text.Split(' ').GetValue(0).ToString(),
                DateTime.Now.Year.ToString(CultureInfo.InvariantCulture) + " г.",
                TbNumProj.Text, TbController.Text.Split(' ').GetValue(0).ToString(),
                TbOrganization.Text, TbResolution.Text, TbCustomer.Text
            };

            // Запускаем ProgressDialog
            var dialogProgress = new ExportProgressDialog(
                ModPlusAPI.Language.GetItem(LangItem, "h20"),
                CreateTemplates)
            {
                Topmost = true,
                BtCancel = { Visibility = Visibility.Visible }
            };
            dialogProgress.ShowDialog();

            // Если была ошибка
            if (_hasError & _error != null)
                ExceptionBox.Show(_error);

            // Удаляем файлы
            if (_fileToDelete.Any())
            {
                foreach (var file in _fileToDelete)
                {
                    try
                    {
                        if (File.Exists(file))
                            File.Delete(file);
                    }
                    catch
                    {
                        // ignored
                    }
                }
            }
        }

        private void CreateTemplates(object sender, DoWorkEventArgs e)
        {
            var worker = sender as BackgroundWorker;

            // Список для поиска
            var toFind = new List<string>
            {
                "<Description>",
                "<GIP>", "<Engineer>",
                "<Employer>", "<Year>",
                "<NumProj>", "<Controller>","<Organization>",
                "<Resolution>", "<Customer>"
            };
            var assembly = Assembly.GetExecutingAssembly();
            worker?.ReportProgress(0, ModPlusAPI.Language.GetItem(LangItem, "h21"));
            var wordAutomation = new WordAutomation();

            // Запускаем Word
            wordAutomation.CreateWordApplication();

            // Проходим по объектам kap
            foreach (var kapDoc in _kapDocs.Where(x => x.Create))
            {
                if (worker != null && worker.CancellationPending)
                {
                    wordAutomation.CloseWordApp();
                    break;
                }

                worker?.ReportProgress(0, ModPlusAPI.Language.GetItem(LangItem, "h22") + ": " + ModPlusAPI.Language.GetItem(LangItem, "h15") + ": " + kapDoc.Name);

                // Временный файл
                var tmp = Path.GetTempFileName();
                var templateFullPath = Path.ChangeExtension(tmp, ".docx");

                // Имя внедренного ресурса
                var resourceName = "mpDocTemplates.Resources.Kap." + kapDoc.Name + ".docx";

                // Читаем ресурс в поток и сохраняем как временный файл
                using (Stream stream = assembly.GetManifestResourceStream(resourceName))
                {
                    SaveStreamToFile(templateFullPath, stream);
                }

                try
                {
                    using (FlatDocument flatDocument = new FlatDocument(templateFullPath))
                    {
                        // Помещаем в список на удаление
                        _fileToDelete?.Add(tmp);
                        _fileToDelete?.Add(templateFullPath);
                        for (var i = 0; i < toFind.Count; i++)
                        {
                            if (worker != null && worker.CancellationPending)
                            {
                                wordAutomation.CloseWordApp();
                                break;
                            }

                            worker?.ReportProgress(
                                Convert.ToInt32((decimal)i / toFind.Count * 100),
                                ModPlusAPI.Language.GetItem(LangItem, "h22") + ": " + ModPlusAPI.Language.GetItem(LangItem, "h15") + ": " + kapDoc.Name);
                            flatDocument.FindAndReplace(toFind[i], _toReplace[i]);
                            Thread.Sleep(50);
                        }
                    }

                    // Создаем документ, используя временный файл
                    worker?.ReportProgress(0, 
                        ModPlusAPI.Language.GetItem(LangItem, "h23") + ": " + ModPlusAPI.Language.GetItem(LangItem, "h15") + ": " + kapDoc.Name);
                    wordAutomation.CreateWordDoc(templateFullPath, true);
                }
                catch (System.Exception exception)
                {
                    // Если словили ошибку, то закрываем ворд
                    wordAutomation.CloseWordApp();
                    _error = exception;
                    _hasError = true;
                }
            }

            // Проходим по объектам Лин
            foreach (var linDoc in _linDocs.Where(x => x.Create))
            {
                if (worker != null && worker.CancellationPending)
                {
                    wordAutomation.CloseWordApp();
                    break;
                }

                worker?.ReportProgress(0,
                    ModPlusAPI.Language.GetItem(LangItem, "h22") + ": " + ModPlusAPI.Language.GetItem(LangItem, "h16") + ": " + linDoc.Name);

                // Временный файл
                var tmp = Path.GetTempFileName();
                var templateFullPath = Path.ChangeExtension(tmp, ".docx");

                // Имя внедренного ресурса
                var resourceName = "mpDocTemplates.Resources.Lin." + linDoc.Name + ".docx";

                // Читаем ресурс в поток и сохраняем как временный файл
                using (var stream = assembly.GetManifestResourceStream(resourceName))
                {
                    SaveStreamToFile(templateFullPath, stream);
                }

                try
                {
                    using (FlatDocument flatDocument = new FlatDocument(templateFullPath))
                    {
                        // Помещаем в список на удаление
                        _fileToDelete?.Add(tmp);
                        _fileToDelete?.Add(templateFullPath);
                        for (var i = 0; i < toFind.Count; i++)
                        {
                            if (worker != null && worker.CancellationPending)
                            {
                                wordAutomation.CloseWordApp();
                                break;
                            }

                            worker?.ReportProgress(
                                Convert.ToInt32((decimal)i / toFind.Count * 100),
                                ModPlusAPI.Language.GetItem(LangItem, "h22") + ": " + ModPlusAPI.Language.GetItem(LangItem, "h16") + ": " + linDoc.Name);
                            flatDocument.FindAndReplace(toFind[i], _toReplace[i]);
                            Thread.Sleep(50);
                        }
                    }

                    // Создаем документ, используя временный файл
                    worker?.ReportProgress(0,
                        ModPlusAPI.Language.GetItem(LangItem, "h23") + ": " + ModPlusAPI.Language.GetItem(LangItem, "h16") + ": " + linDoc.Name);
                    wordAutomation.CreateWordDoc(templateFullPath, true);
                }
                catch (System.Exception exception)
                {
                    // Если словили ошибку, то закрываем ворд
                    wordAutomation.CloseWordApp();
                    _error = exception;
                    _hasError = true;
                }
            }

            // Делаем word видимым
            wordAutomation.MakeWordAppVisible();
        }

        private static void SaveStreamToFile(string fileFullPath, Stream stream)
        {
            if (stream.Length == 0)
                return;

            // Create a FileStream object to write a stream to a file
            using (var fileStream = File.Create(fileFullPath, (int)stream.Length))
            {
                // Fill the bytes[] array with the stream data
                var bytesInStream = new byte[stream.Length];
                stream.Read(bytesInStream, 0, bytesInStream.Length);

                // Use FileStream object to write to the specified file
                fileStream.Write(bytesInStream, 0, bytesInStream.Length);
            }
        }
    }
}
