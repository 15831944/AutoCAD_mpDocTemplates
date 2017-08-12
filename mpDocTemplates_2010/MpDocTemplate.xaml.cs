#if ac2010
using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;
#elif ac2013
using AcApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;
#endif
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Internal;
using Autodesk.AutoCAD.Runtime;
using mpMsg;
using mpSettings;
using ModPlus;
using System.Globalization;
using Visibility = System.Windows.Visibility;
using System.IO;
using System.Reflection;
using System.Linq;

namespace mpDocTemplates
{
    /// <summary>
    /// Логика взаимодействия для MpDocTemplate.xaml
    /// </summary>
    public partial class MpDocTemplate
    {
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
            MpWindowHelpers.OnWindowStartUp(
                this,
                MpSettings.GetValue("Settings", "MainSet", "Theme"),
                MpSettings.GetValue("Settings", "MainSet", "AccentColor"),
                MpSettings.GetValue("Settings", "MainSet", "BordersType")
                );
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

        private void MpDocTemplate_OnPreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape) Close();
        }

        private void MpDocTemplate_OnMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            DragMove();
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
                    TbOrganization
                };
                // Данные из файла настроек
                foreach (var tb in _textBoxes)
                {
                    tb.Text = MpSettings.GetValue("Settings", "mpDocTemplates", tb.Name);
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
                MpExWin.Show(exception);
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
                        MpSettings.SetValue("Settings", "mpDocTemplates", tb.Name, tb.Text, false);
                    }
                    MpSettings.SaveFile();
                }
            }
            catch (System.Exception exception)
            {
                MpExWin.Show(exception);
            }
        }
        // Взять из полей
        private void BtGetFromFields_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                var dbsi = AcApp.DocumentManager.MdiActiveDocument.Database.SummaryInfo;
                var dbsib = new DatabaseSummaryInfoBuilder(dbsi);
                for (var i = 0; i < _textBoxes.Count; i++)
                {
                    if (dbsib.CustomPropertyTable.Contains(_fieldNames[i]))
                        _textBoxes[i].Text = dbsib.CustomPropertyTable[_fieldNames[i]].ToString();
                }
            }
            catch (System.Exception exception)
            {
                MpExWin.Show(exception);
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
                MpMsgWin.Show("Вы не указали ни одного шаблона для создания");
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
                TbOrganization.Text
            };
            // Запускаем ProgressDialog
            var dialogProgress = new ExportProgressDialog(
                "Создание шаблонов",
                CreateTemplates)
            {
                Topmost = true,
                BtCancel = { Visibility = Visibility.Visible }
            };
            dialogProgress.ShowDialog();
            // Если была ошибка
            if(_hasError & _error != null)
                MpExWin.Show(_error);
            // Удаляем файлы
            if(_fileToDelete.Any())
                foreach (var file in _fileToDelete)
                {
                    try
                    {
                        if(File.Exists(file))
                            File.Delete(file);
                    }
                    catch
                    {
                        //ignored
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
                "<NumProj>", "<Controller>","<Organization>"
            };
            var assembly = Assembly.GetExecutingAssembly();
            var wordAutomation = new WordAutomation();
            // Запускаем Word
            wordAutomation.CreateWordApplication();
            // Проходим по объектам kap
            foreach (var kapDoc in _kapDocs.Where(x => x.Create))
            {
                if (worker != null && worker.CancellationPending) { wordAutomation.CloseWordApp(); break;}
                worker?.ReportProgress(0, "Создание шаблона: Объекты кап. стр-ва: " + kapDoc.Name);
                // Временный файл
                var tmp = Path.GetTempFileName();
                var templateFullPath = Path.ChangeExtension(tmp, ".docx");
                // Имя внедренного ресурса
                var resourceName = "mpDocTemplates.Resources.Kap." + kapDoc.Name + ".docx";
                // Читаем ресурс в поток и сохраняем как временный файл
                using (var stream = assembly.GetManifestResourceStream(resourceName))
                {
                    SaveStreamToFile(templateFullPath, stream);
                }
                try
                {
                    // Создаем документ, используя временный файл
                    wordAutomation.CreateWordDoc(templateFullPath, false);
                    // Помещяем в список на удаление
                    _fileToDelete?.Add(tmp);
                    _fileToDelete?.Add(templateFullPath);
                    for (var i = 0; i < toFind.Count; i++)
                    {
                        if (worker != null && worker.CancellationPending) { wordAutomation.CloseWordApp(); break; }
                        worker?.ReportProgress(Convert.ToInt32(((decimal)i / toFind.Count) * 100),
                            "Создание шаблона: Объекты кап. стр-ва: " + kapDoc.Name);
                        wordAutomation.FindReplace(toFind[i], _toReplace[i]);
                    }
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
                if (worker != null && worker.CancellationPending) { wordAutomation.CloseWordApp(); break; }
                worker?.ReportProgress(0, "Создание шаблона: Линейные объекты: " + linDoc.Name);
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
                    // Создаем документ, используя временный файл
                    wordAutomation.CreateWordDoc(templateFullPath, false);
                    // Помещяем в список на удаление
                    _fileToDelete?.Add(tmp);
                    _fileToDelete?.Add(templateFullPath);
                    for (var i = 0; i < toFind.Count; i++)
                    {
                        if (worker != null && worker.CancellationPending) { wordAutomation.CloseWordApp(); break; }
                        worker?.ReportProgress(Convert.ToInt32(((decimal)i / toFind.Count) * 100),
                            "Создание шаблона: Линейные объекты: " + linDoc.Name);
                        wordAutomation.FindReplace(toFind[i], _toReplace[i]);
                    }
                }
                catch (System.Exception exception)
                {
                    // Если словили ошибку, то закрываем ворд
                    wordAutomation.CloseWordApp();
                    _error = exception;
                    _hasError = true;
                }
            }
            wordAutomation.MakeWordAppVisible();
        }
        private static void SaveStreamToFile(string fileFullPath, Stream stream)
        {
            if (stream.Length == 0) return;

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
    /// <summary>
    /// Класс описывает один элемент в списке шаблонов
    /// Имеет имя, описание и параметр Создавать или нет
    /// </summary>
    class TemplateItem
    {
        /// <summary>
        /// Имя шаблона
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// Создавать или нет
        /// </summary>
        public bool Create { get; set; }
        /// <summary>
        /// Описание шаблона
        /// </summary>
        public string ToolTip { get; set; }
    }
    /// <summary>
    /// Постоянные значения
    /// </summary>
    static class TemplateData
    {
        public static void FillKapItems(ObservableCollection<TemplateItem> col)
        {
            for (var i = 0; i < KapNames.Count; i++)
            {
                col.Add(
                    new TemplateItem
                    {
                        Name = KapNames[i],
                        ToolTip = KapToolTips[i],
                        Create = false
                    });
            }
        }

        public static void FillLinItems(ObservableCollection<TemplateItem> col)
        {
            for (var i = 0; i < LinNames.Count; i++)
            {
                col.Add(new TemplateItem
                {
                    Name = LinNames[i],
                    ToolTip = LinToolTips[i],
                    Create = false
                });
            }
        }
        static readonly List<string> KapNames = new List<string>
        {
            "ПЗ","ПЗУ","АР","КР","ЭОМ","В","К","ОВ","СС","ГС","ТХ","ПОС","ПОД","ООС","ПБ","ОДИ","ЭЭ","СМ"
        };
        static readonly List<string> KapToolTips = new List<string>
        {
            "Пояснительная записка",
            "Схема планировочной организации земельного участка",
            "Архитектурные решения",
            "Конструктивные и объемно-планировочные решения",
            "Система электроснабжения",
            "Система водоснабжения",
            "Система водоотведения",
            "Отопление, вентиляция и кондиционирование воздуха, тепловые сети",
            "Сети связи",
            "Система газоснабжения",
            "Технологические решения",
            "Проект организации строительства",
            "Проект организации работ по сносу или демонтажу&#x0a;объектов капиатльного строительства",
            "Перечень мероприятий по охране окружающей среды",
            "Мероприятия по обеспечению пожарной безопасности",
            "Мероприятия по обеспечению доступа инвалидов",
            "Перечень мероприятий по обеспечению соблюдения требований&#x0a;энергетической эффективности и требований оснащенности зданий, строений,&#x0a;сооружений приборами учета используемых энергетических ресурсов",
            "Смета на строительство объектов капитального строительства"
        };
        static readonly List<string> LinNames = new List<string>
        {
            "ПЗ","ППО","ТКР","ИЛО","ПОС","ПОД","ООС","ПБ","СМ"
        };
        static readonly List<string> LinToolTips = new List<string>
        {
            "Пояснительная записка",
            "Проект полосы отвода",
            "Технологические и конструктивные решения линейного&#x0a;объекта. Искусственные сооружения",
            "Здания, строения и сооружения, входящие в&#x0a;инфраструктуру линейного объекта",
            "Проект организации строительства",
            "Проект организации работ по сносу (демонтажу) линейного объекта",
            "Мероприятия по охране окружающей среды",
            "Мероприятия по обеспечению пожарной безопасности",
            "Смета на строительство"
        };
    }
    /// <summary>
    /// Запуск функции в автокаде
    /// </summary>
    public class AcadFunction
    {
        MpDocTemplate _mpDocTemplate;

        [CommandMethod("ModPlus", "mpDocTemplates", CommandFlags.Modal)]
        public void StartFunction()
        {
            if (_mpDocTemplate == null)
            {
                _mpDocTemplate = new MpDocTemplate();
                _mpDocTemplate.Closed += window_Closed;
            }
            if (_mpDocTemplate.IsLoaded)
                _mpDocTemplate.Activate();
            else
                AcApp.ShowModalWindow(
                    AcApp.MainWindow.Handle, _mpDocTemplate);
        }
        void window_Closed(object sender, EventArgs e)
        {
            _mpDocTemplate = null;
        }
    }
}
