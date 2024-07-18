using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;
using TFlex.Reporting.CAD.MacroGenerator.Macros;
using TFlex.DOCs.Model.References;
using TFlex.DOCs.Model.References.Documents;
using TFlex.DOCs.Model.References.Nomenclature;
using System.Linq.Expressions;

using TFlex.Reporting.CAD.MacroGenerator.ObjectModel;

using TFlex.PdmFramework.Remarks;


public class Macro : ReportBaseMacroProvider
{
    public class Cell
    {
        public Guid guid { get; set; }
        public string value { get; set; }
        public string tag { get; set; }
        public Cell(string tag, Guid guid)
        {
            this.tag = tag;
            this.guid = guid;
            this.value = "";
        }
    }
    public class Record
    { 
        private Dictionary<string, Cell> cells; 
        public bool isHeader = false;
        public Record(bool _isHeader = false)
        {
                cells = new Dictionary<string, Cell>()
                {
                    { "Формат", new Cell("Ф", new Guid("42a5dd37-e537-46ab-88d2-97060ca46c1c")) },
                    { "Зона", new Cell("З", new Guid("1367dda9-7850-4c15-b123-636ab692034c")) },
                    { "Позиция", new Cell("П", new Guid("ab34ef56-6c68-4e23-a532-dead399b2f2e")) },
                    { "Обозначение", new Cell("Обозначение", new Guid("ae35e329-15b4-4281-ad2b-0e8659ad2bfb")) },
                    { "Наименование", new Cell("Наименование", new Guid("45e0d244-55f3-4091-869c-fcf0bb643765")) },
                    { "Кол", new Cell("Кол", new Guid("3f5fc6c8-d1bf-4c3d-b7ff-f3e636603818")) },
                    { "Примечание", new Cell("Прим", new Guid("a3d509de-a28f-4719-936b-fb2da0ca72ce")) },
                    { "Раздел спецификации", new Cell("Раздел", new Guid("7e2425f7-15ea-4921-be03-b60db93fbe28")) },
                };
            this.isHeader = _isHeader;
        }
        public Cell this[string cellName]
        {
            get { return cells[cellName]; }
        }
    }
    public class SectionList
    {

        private Dictionary<string, List<ReferenceObject>> sections;
            
        public SectionList()
        {
            sections = new Dictionary<string, List<ReferenceObject>>()
            {
                { "Documentation",  new List<ReferenceObject>() },
                { "Complexes",  new List<ReferenceObject>() },
                { "Assemblies",  new List<ReferenceObject>() },
                { "Parts",  new List<ReferenceObject>() },
                { "StandartParts",  new List<ReferenceObject>() },
                { "MiscParts",  new List<ReferenceObject>() },
                { "Materials",  new List<ReferenceObject>() },
                { "Kits",  new List<ReferenceObject>() },
            };

        }
        public List<ReferenceObject> this[string sectionName]
        { 
            get { return sections[sectionName]; }
        }
    }
    public Dictionary<string, Guid> Guids = new Dictionary<string, Guid>()
        {
            {"Формат",  new Guid("42a5dd37-e537-46ab-88d2-97060ca46c1c")},
            {"Зона",  new Guid("1367dda9-7850-4c15-b123-636ab692034c")},
            {"Позиция",  new Guid("ab34ef56-6c68-4e23-a532-dead399b2f2e")},
            {"Обозначение",  new Guid("ae35e329-15b4-4281-ad2b-0e8659ad2bfb")},
            {"Наименование",  new Guid("45e0d244-55f3-4091-869c-fcf0bb643765")},
            {"Количество",  new Guid("3f5fc6c8-d1bf-4c3d-b7ff-f3e636603818")},
            {"Примечание",  new Guid("a3d509de-a28f-4719-936b-fb2da0ca72ce")},
            {"Раздел спецификации",  new Guid("7e2425f7-15ea-4921-be03-b60db93fbe28")},
            {"ID", new Guid("647ad6c4-a23c-403e-bab1-ac57e696f396")},
            {"Типоразмер", new Guid("adb931cf-7978-40f4-a4d7-4398c3e147de")},
            {"Типкрепежа", new Guid("f868c90a-ded0-4e47-b1c3-6683e8048f81")},
            //ГОСТ 1759.0
            {"Класс точности", new Guid("")},
            {"Исполнение", new Guid("")},
            {"D резьбы", new Guid("")},
            {"Шаг резьбы", new Guid("")},
            {"Напр. резьбы", new Guid("")},
            {"Поле допуска резьбы", new Guid("")},
            {"Длина изделия", new Guid("")},
            {"Класс прочности", new Guid("")},
            {"Указание стали", new Guid("")},
            {"Марка материала", new Guid("")},
            {"Вид покрытия", new Guid("")},
            {"Толщина покрытия", new Guid("")},
        };
    private static Dictionary<String, String> SingularPluralMap = new Dictionary<string, string>()
        {
            {"Болт", "Болты"},
            {"Винт", "Винты"},
            {"Гайка", "Гайки"},
            {"Заклепка", "Заклепки"},
            {"Подшипник", "Подшипники"},
            {"Фланец", "Фланцы"},
            {"Шайба", "Шайбы"},
            {"Шпилька", "Шпильки"},
            {"Гвоздь", "Гвозди"},
            {"Кольцо", "Кольца"},
            {"Соединение", "Соединения"},
            {"Крышка", "Крышки"},
            {"Манжет", "Манжеты"},
            {"Профиль", "Профили"},
            {"Труба", "Трубы"},
            {"Скоба", "Скобы"},
            {"Двутавр", "Двутавры"},
            {"Швеллер", "Швеллеры"},
            {"Шплинт", "Шплинты"},
            {"Шпонка", "Шпонки"},
            {"Штифт", "Штифты"},
            {"Шуруп", "Шурупы"},
        };

    public Macro(ReportGenerationBaseMacroContext context)        : base(context)    {    }
    public override void Run()
    {
        System.Diagnostics.Debugger.Launch();
        System.Diagnostics.Debugger.Break();
        List<Record> records = new List<Record>();
        NomenclatureObject current_obj = (ReferenceObject)CurrentObject as NomenclatureObject;
        List<NomenclatureObject> versions = current_obj.GetVersions();
        NomenclatureObject base_version = versions[0]; 
        var allPart = versions.SelectMany(v => v.Children.OfType<NomenclatureObject>()).ToList();
        var variable_part = new List<NomenclatureObject>();
        
        /*--------------------------------TODO-------------------------------------  
        Не забыть переписать эту хрень на лямбду !  =°,,°=
         -------------------------------------------------------------------------*/
        foreach (NomenclatureObject obj in allPart) 
        {
            List<double> quantity = new List<double>();
            foreach (var v in versions)
            {
                try {
                    var i = obj.GetParentLink(v).Id;
                    quantity.Add( (double)obj.GetParentLink(v)[Guids["Количество"]].Value);
                }
                catch (NullReferenceException) 
                {
                    variable_part.Add(obj);
                    quantity.Clear();
                }
            }
            if (quantity.Distinct().Count() >1)  { variable_part.Add(obj); }
        }

        var constantPart = allPart.Except(variable_part).ToList();
        filling_by_list(constantPart, base_version, records);
        //Переменные данные для исполнений
        records.Add(new Record());
        Record dummy = new Record();
        dummy["Обозначение"].value = "Переменные данные";
        dummy["Наименование"].value = "для исполнений";
        records.Add(dummy); 
        foreach (var v in versions)
        {
            Record header = new Record(true);
            header["Наименование"].value = v.Denotation;
            records.Add(header);
            filling_by_list(variable_part, v, records);
        }
    }
    public void filling_by_list(List<NomenclatureObject> section, NomenclatureObject parent, List<Record> rl)
    {
        SectionList sl = new SectionList();
        
        foreach (NomenclatureObject obj in section.Where(o => o.GetParentLink(parent) != null))
        {
            var child = obj.GetParentLink(parent).ChildObject;
            switch (obj.GetParentLink(parent)[Guids["Раздел спецификации"]].Value.ToString())
            {
                case "Документация": sl["Documentation"].Add(child); break;
                case "Комплексы": sl["Complexes"].Add(child); break;
                case "Сборочные единицы": sl["Assemblies"].Add(child); break;
                case "Детали": sl["Parts"].Add(child); break;
                case "Стандартные изделия": sl["StandartParts"].Add(child); break;
                case "Прочие изделия": sl["MiscParts"].Add(child); break;
                case "Материалы": sl["Materials"].Add(child); break;
                case "Комплекты": sl["Kits"].Add(child); break;
            }

        }
            
        rl.AddRange( MakeList( sl["Documentation"].OrderBy(obj => obj[Guids["Обозначение"]].Value).ToList(), parent, "Документация"));
        rl.AddRange(MakeList( sl["Complexes"].OrderBy(obj => obj[Guids["Обозначение"]].Value).ToList(), parent, "Комплексы"));
        rl.AddRange(MakeList( sl["Assemblies"].OrderBy(obj => obj[Guids["Обозначение"]].Value).ToList(), parent, "Сборочные единицы"));
        rl.AddRange(MakeList( sl["Parts"].OrderBy(obj => obj[Guids["Обозначение"]].Value).ToList(), parent, "Детали"));

        rl.AddRange(MakeListSti(sl["StandartParts"], parent, "Стандартные изделия"));
        rl.AddRange(MakeList( sl["MiscParts"], parent, "Прочие изделия"));
        rl.AddRange(MakeList( sl["Materials"], parent, "Материалы"));
        rl.AddRange(MakeList( sl["Kits"], parent, "Комплекты")); 
    }
    public void print_test(List<Record> list)
    {
        Переменная["$Наименование"] = ТекущийОбъект.Параметр["Наименование"];
        Переменная["$Обозначение"] = ТекущийОбъект.Параметр["Обозначение"];
        foreach (var cl in list)
        {
            var текст = Текст["Текст1"];
            var шаблон = текст["Наименование"];
            var строка = текст.Таблица.ДобавитьСтроку(шаблон);
            if(cl.isHeader)
            {
                шаблон = текст["Раздел"];
                строка = текст.Таблица.ДобавитьСтроку(шаблон);
                строка["Раздел"].Текст = cl["Наименование"].value;
                continue;
            }
            var cellNames = new List<string> { "Формат", "Зона", "Позиция", "Обозначение", "Наименование", "Кол", "Примечание" };
            foreach (var cellName in cellNames)
            {                           
                строка[cl[cellName].tag].Текст = cl[cellName].value; 
            }
        }
    }
    public List<Record> MakeListSti(List<ReferenceObject> list, NomenclatureObject parent, string name_sec)
    {
        if (list == null || list.Count == 0)
        {
            return new List<Record>();
        }
        var records = new List<Record>();
        Record header = new Record(true);
        header["Наименование"].value = name_sec;
        records.Add(header);
        var sti = list.GroupBy(obj => obj[Guids["Стандарт"]].Value.ToString())
         .Select(g => new
         {
             GroupName = g.Key,
             Count = g.Count(),
             PartList = g.Select(p => p)
         });
        foreach (var group in sti)
        {
            if (group.Count > 1)
            {
                string NamePart = group.PartList.FirstOrDefault()[Guids["Наименвоание изделия"]].GetString();
                string PluralName = SingularPluralMap.ContainsKey(NamePart) ? SingularPluralMap[NamePart] : NamePart;
                Record groupHeader = new Record();
                groupHeader["Наименование"].value = PluralName + " " + group.GroupName; ;
                records.Add(groupHeader);

            }
            foreach (var child in group.PartList.Distinct())
            {
                Record item = new Record();
                try { item["Позиция"].value = child.GetParentLink(parent)[item["Позиция"].guid].Value.ToString(); }
                catch (NullReferenceException) { }
                try
                {
                    string varName = MakeNameSti(child);
                    if (group.Count > 1)
                    {
                        item["Наименование"].value = varName.Length > 1 ? varName : child[item["Наименование"].guid].GetString();
                    }
                    else item["Наименование"].value = varName.Length > 1 ?
                            child[Guids["Наименование изделия"]].GetString() + varName + child[Guids["Стандарт"]].GetString() :
                            child[item["Наименование"].guid].GetString();
                }
                catch (NullReferenceException) { }
                try { item["Кол"].value = parent.GetChildLinks(child).Sum(link => (double)link[item["Кол"].guid].Value).ToString(); }
                catch (NullReferenceException) { }
                try { item["Примечание"].value = parent.GetChildLink(child)[item["Примечание"].guid].Value.ToString(); }
                catch (NullReferenceException) { }
                records.Add(item);
                }
           }
           return records;
    }
        
    private string MakeNameSti(ReferenceObject child)
    {
        string str = string.Empty;
        string precisionClass = string.IsNullOrEmpty(child[Guids["Класс точности"]].ToString()) ? "": child[Guids["Класс точности"]].ToString();
        string variant = string.IsNullOrEmpty(child[Guids["Исполнение"]].ToString()) ? "" : child[Guids["Исполнение"]].ToString();
        string diameter = string.IsNullOrEmpty(child[Guids["D резьбы"]].ToString()) ? "" : "М" + child[Guids["D резьбы"]].ToString();
        string pitch = child[Guids["Шаг резьбы"]].GetDouble() == 0 ? "" : "х" + child[Guids["D резьбы"]].ToString();
        string leftHanded = string.IsNullOrEmpty(child[Guids["Напр. резьбы"]].ToString()) ? "" : "-"+child[Guids["Напр. резьбы"]].ToString();
        string threadTol = string.IsNullOrEmpty(child[Guids["Поле допуска резьбы"]].ToString()) ? "" : "-" + child[Guids["Поле допуска резьбы"]].ToString();
        string length = string.IsNullOrEmpty(child[Guids["Длина изделия"]].ToString()) ? "" : "-" + child[Guids["Длина изделия"]].ToString();
        string strength = string.IsNullOrEmpty(child[Guids["Класс прочности"]].ToString()) ? "" : "." + child[Guids["Класс прочности"]].ToString();
        string steelType = string.IsNullOrEmpty(child[Guids["Указание стали"]].ToString()) ? "" : "." + child[Guids["Указание стали"]].ToString();
        string material = string.IsNullOrEmpty(child[Guids["Марка материала"]].ToString()) ? "" : "." + child[Guids["Марка материала"]].ToString();
        string coatType = string.IsNullOrEmpty(child[Guids["Вид покрытия"]].ToString()) ? "" : "." + child[Guids["Вид покрытия"]].ToString();
        string coatTickness = child[Guids["Толщина покрытия"]].GetInt32() == 0 ? "" : child[Guids["Толщина покрытия"]].ToString();

        str = precisionClass + variant + diameter +pitch + leftHanded + threadTol + steelType + material + coatType + coatTickness;
        return str;
    }

    public List<Record> MakeList(List<ReferenceObject> list, NomenclatureObject v, string name_sec)
    {
        if (list == null || list.Count == 0)
        {
            return new List<Record>();
        }
        var records = new List<Record>();
        foreach (var child in list.Distinct())
        {
            Record item = new Record();
            try { item["Формат"].value = child[item["Формат"].guid].Value.ToString() ?? ""; }
            catch (NullReferenceException) { }
            try { item["Зона"].value = child.GetParentLink(v)[item["Зона"].guid].Value.ToString(); }
            catch (NullReferenceException) { }
            try { item["Позиция"].value = child.GetParentLink(v)[item["Позиция"].guid].Value.ToString(); }
            catch (NullReferenceException) { }
            try { item["Обозначение"].value = child[item["Обозначение"].guid].Value.ToString(); }
            catch (NullReferenceException) { }
            try { item["Наименование"].value = child[item["Наименование"].guid].Value.ToString(); }
            catch (NullReferenceException) { }
            try { item["Кол"].value = v.GetChildLinks(child).Sum(link => (double)link[item["Кол"].guid].Value).ToString(); }
            catch (NullReferenceException) { }
            try { item["Примечание"].value = v.GetChildLink(child)[item["Примечание"].guid].Value.ToString(); }                             
            catch (NullReferenceException) { }
            records.Add(item);
        }
        return records;
    }

}