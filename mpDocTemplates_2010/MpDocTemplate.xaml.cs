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
using System.Globalization;
using Visibility = System.Windows.Visibility;
using System.IO;
using System.Reflection;
using System.Linq;
using System.Threading;
using ModPlusAPI;
using ModPlusAPI.IO.Office.Word;
using ModPlusAPI.IO.Office.Word.FindAndReplace;
using ModPlusAPI.Windows;

namespace mpDocTemplates
{
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

        private void MpDocTemplate_OnPreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape) Close();
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
                    tb.Text = UserConfigFile.GetValue(UserConfigFile.ConfigFileZone.Settings, "mpDocTemplates", tb.Name);
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
                        UserConfigFile.SetValue(UserConfigFile.ConfigFileZone.Settings, "mpDocTemplates", tb.Name, tb.Text, false);
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
                foreach (var file in _fileToDelete)
                {
                    try
                    {
                        if (File.Exists(file))
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
                        // Помещяем в список на удаление
                        _fileToDelete?.Add(tmp);
                        _fileToDelete?.Add(templateFullPath);
                        for (var i = 0; i < toFind.Count; i++)
                        {
                            if (worker != null && worker.CancellationPending)
                            {
                                wordAutomation.CloseWordApp();
                                break;
                            }
                            worker?.ReportProgress(Convert.ToInt32(((decimal)i / toFind.Count) * 100),
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
                        // Помещяем в список на удаление
                        _fileToDelete?.Add(tmp);
                        _fileToDelete?.Add(templateFullPath);
                        for (var i = 0; i < toFind.Count; i++)
                        {
                            if (worker != null && worker.CancellationPending)
                            {
                                wordAutomation.CloseWordApp();
                                break;
                            }
                            worker?.ReportProgress(Convert.ToInt32(((decimal)i / toFind.Count) * 100),
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
    internal class TemplateItem
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
    /// <summary>Постоянные значения</summary>
    internal static class TemplateData
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
            "ПЗ","ПЗУ","АР","КР","ЭОМ","В","К","ОВ","СС","ГС","ТХ","ПОС","ПОД","ООС","ПБ","ОДИ","ТБЭ","СМ","ЭЭ"
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
            "Проект организации работ по сносу или демонтажу" + Environment.NewLine +
            "объектов капиатльного строительства",
            "Перечень мероприятий по охране окружающей среды",
            "Мероприятия по обеспечению пожарной безопасности",
            "Мероприятия по обеспечению доступа инвалидов",
            "Требования к обеспечению безопасной эксплуатации объектов капитального строительства",
            "Смета на строительство объектов капитального строительства",
            "Перечень мероприятий по обеспечению соблюдения требований"+ Environment.NewLine +
            "энергетической эффективности и требований оснащенности зданий, строений,"+ Environment.NewLine +
            "сооружений приборами учета используемых энергетических ресурсов"
        };
        static readonly List<string> LinNames = new List<string>
        {
            "ПЗ","ППО",
            "ТКР (автомобильные дороги)",
            "ТКР (железные дороги)",
            "ТКР (метрополитен)",
            "ТКР (линии связи)",
            "ТКР (магистральные трубопроводы)",
            "ИЛО","ПОС","ПОД","ООС","ПБ","СМ"
        };
        static readonly List<string> LinToolTips = new List<string>
        {
            "Пояснительная записка",
            "Проект полосы отвода",
            "Технологические и конструктивные решения линейного"+ Environment.NewLine +
            "объекта. Искусственные сооружения",
            "Технологические и конструктивные решения линейного"+ Environment.NewLine +
            "объекта. Искусственные сооружения",
            "Технологические и конструктивные решения линейного"+ Environment.NewLine +
            "объекта. Искусственные сооружения",
            "Технологические и конструктивные решения линейного"+ Environment.NewLine +
            "объекта. Искусственные сооружения",
            "Технологические и конструктивные решения линейного"+ Environment.NewLine +
            "объекта. Искусственные сооружения",
            "Здания, строения и сооружения, входящие в"+ Environment.NewLine +
            "инфраструктуру линейного объекта",
            "Проект организации строительства",
            "Проект организации работ по сносу (демонтажу) линейного объекта",
            "Мероприятия по охране окружающей среды",
            "Мероприятия по обеспечению пожарной безопасности",
            "Смета на строительство"
        };
    }
    /// <summary>Запуск функции в автокаде</summary>
    public class AcadFunction
    {
        MpDocTemplate _mpDocTemplate;

        [CommandMethod("ModPlus", "mpDocTemplates", CommandFlags.Modal)]
        public void StartFunction()
        {
            Statistic.SendCommandStarting(new Interface());

            if (_mpDocTemplate == null)
            {
                _mpDocTemplate = new MpDocTemplate();
                _mpDocTemplate.Closed += Window_Closed;
            }
            if (_mpDocTemplate.IsLoaded)
                _mpDocTemplate.Activate();
            else
                AcApp.ShowModalWindow(
                    AcApp.MainWindow.Handle, _mpDocTemplate);
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            _mpDocTemplate = null;
        }
    }
}
