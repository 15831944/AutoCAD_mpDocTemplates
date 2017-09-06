using System;
using System.Collections.Generic;
using ModPlusAPI.Interfaces;

namespace mpDocTemplates
{
    public class Interface : IModPlusFunctionInterface
    {
        public SupportedProduct SupportedProduct => SupportedProduct.AutoCAD;
        public string Name => "mpDocTemplates";
        public string AvailProductExternalVersion => "2010";
        public string FullClassName => string.Empty;
        public string AppFullClassName => string.Empty;
        public Guid AddInId => Guid.Empty;
        public string LName => "Шаблоны. Стадия П";
        public string Description => "Функция позволяет создавать шаблоны текстовой части раздела согласно \"Постановление РФ №87\" с титульным листом согласно ГОСТ Р 21.1101-2013";
        public string Author => "Пекшев Александр aka Modis";
        public string Price => "0";
        public bool CanAddToRibbon => true;
        public string FullDescription => string.Empty;
        public string ToolTipHelpImage => string.Empty;
        public List<string> SubFunctionsNames => new List<string>();
        public List<string> SubFunctionsLames => new List<string>();
        public List<string> SubDescriptions => new List<string>();
        public List<string> SubFullDescriptions => new List<string>();
        public List<string> SubHelpImages => new List<string>();
        public List<string> SubClassNames => new List<string>();
    }
}
