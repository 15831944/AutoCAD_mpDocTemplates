namespace mpDocTemplates
{
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
}