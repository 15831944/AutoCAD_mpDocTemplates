namespace mpDocTemplates
{
    using System;
    using Autodesk.AutoCAD.ApplicationServices.Core;
    using Autodesk.AutoCAD.Runtime;
    using ModPlusAPI;

    /// <summary>
    /// Запуск функции в автокаде
    /// </summary>
    public class AcadFunction
    {
        MpDocTemplate _mpDocTemplate;

        [CommandMethod("ModPlus", "mpDocTemplates", CommandFlags.Modal)]
        public void StartFunction()
        {
            Statistic.SendCommandStarting(new ModPlusConnector());

            if (_mpDocTemplate == null)
            {
                _mpDocTemplate = new MpDocTemplate();
                _mpDocTemplate.Closed += Window_Closed;
            }

            if (_mpDocTemplate.IsLoaded)
            {
                _mpDocTemplate.Activate();
            }
            else
            {
                Application.ShowModalWindow(
                    Application.MainWindow.Handle, _mpDocTemplate);
            }
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            _mpDocTemplate = null;
        }
    }
}