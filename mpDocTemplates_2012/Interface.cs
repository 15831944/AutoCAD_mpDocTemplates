using mpPInterface;

namespace mpDocTemplates
{
    public class Interface : IPluginInterface
    {
        public string Name => "mpDocTemplates";
        public string AvailCad => "2012";
        public string LName => "Шаблоны";
        public string Description => "Функция позволяет создавать шаблоны текстовой части раздела согласно &quot;Постановление РФ №87&quot; с титульным листом согласно ГОСТ Р 21.1101-2013";
        public string Author => "Пекшев Александр aka Modis";
        public string Price => "0";
    }
}
