using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Xml;
using System.Xml.Linq;
using System.IO;
using Kompas6API5;
using KompasAPI7;
using KAPITypes;
using Kompas6Constants;

namespace Steps.NET
{
    public partial class Form1 : Form
    {
        private KompasObject kompas;
        private IApplication appl;
        //private IDocuments doc_add;
        //private IKompasDocument2D doc;
        private IKompasDocument docum;
        private ILayoutSheets sheets;
        private ILayoutSheet sheet;
        private ISheetFormat format;
        private ksDocument2D _ksDocument2D;
        private ksStamp _ksStamp;
        private ksTextItemParam _ksTextItemParam;
        private ksTextLineParam _ksTextLineParam;
        private IKompasAPIObject _IKompasAPIObject;
        private IDocuments _IDocuments;
        private ksDynamicArray _ksDynamicArray;
        //private ksParagraphParam _ksParagraphParam;
        //private ksTextParam _ksTextParam;
        private TextItemFont _TextItemFont;
        DynamicArray _DynamicArray;

        public string decimal_number;
        public string primary_applicability;
        public static string name;
        public string head_piece;
        public string active_path;
        public string active_name;
        int LayoutStyleNumber_count = 0;
        string global_path;
        string global_path_to_xmlfile;
        string path;
        string litera;
        string date;

        public Form1()
        {
            InitializeComponent();
            var location = System.Reflection.Assembly.GetExecutingAssembly().Location;
            global_path = Path.GetDirectoryName(location);
            global_path_to_xmlfile = global_path + "/xmlfile.xml";
        }
        private void button5_Click(object sender, EventArgs e)
        {
            kompas = (KompasObject)Marshal.GetActiveObject("KOMPAS.Application.5");
            _ksDocument2D = (ksDocument2D)kompas.ActiveDocument2D();
            if (_ksDocument2D != null)
            {
                if (textBox1.Text != "" & textBox2.Text != "" & textBox3.Text != "" & textBox4.Text != "")
                {
                    try
                    {
                        //capture_active_title_block(); 
                        if (capture_active_title_block() == 2) //возвращает 1 - отсутствует децимальный номер, 2 - штатная работа проргаммы
                        {
                            get_active_path();

                            if (System.IO.File.Exists(active_path + " ВН.cdw"))
                            {
                                kompas.ksMessage("Данная ведомость уже существует");
                            }
                            else
                            {
                                create_new_list();
                                filling_tables();
                                filling_tables2();
                                filling_tables3();
                                this.Close();
                            }
                        }
                    }
                    catch
                    {
                        kompas = (KompasObject)Marshal.GetActiveObject("KOMPAS.Application.5");
                        kompas.ksMessage("Видимо какая-то ошибка в оформлении чертежа.          Вероятные ошибки: -Отсутствие децимального номера;   -Библиотека оформлений не подключена; -Чертеж не сохранен; -Слишком короткое Имя в поле textbox.");
                    }
                }
                else
                    kompas.ksMessage("Не все текстовые поля заполнены.");
            }
            else
            { 
                kompas.ksMessage("Перед использованием библиотеки откройте нужный чертеж!");
                this.Close();
            }
        }       
        public int capture_active_title_block()
        {
            int a;
            decimal_number = null;
            ksDynamicArray items;

            kompas = (KompasObject)Marshal.GetActiveObject("KOMPAS.Application.5");
            _ksTextLineParam = (ksTextLineParam)kompas.GetParamStruct(29);
            _ksTextItemParam = (ksTextItemParam)kompas.GetParamStruct(31);

            _ksDocument2D = (ksDocument2D)kompas.ActiveDocument2D();
            _ksStamp = (ksStamp)_ksDocument2D.GetStampEx(1);
            _ksStamp.ksOpenStamp();
           
            #region получаем децимальный номер
            _ksStamp.ksColumnNumber(2); //номер ячейки основной надписи
            _ksDynamicArray = (ksDynamicArray)_ksStamp.ksGetStampColumnText(0);

            if (_ksDynamicArray.ksGetArrayCount() == 0)
            { 
                kompas.ksMessage("В документе отсутствует децимальный номер!");
                a = 1;
                return a;
            }
            
            for (int i = 0; i < _ksDynamicArray.ksGetArrayCount(); i++)
            {
                _ksDynamicArray.ksGetArrayItem(i, _ksTextLineParam);
                items = (ksDynamicArray)_ksTextLineParam.GetTextItemArr();
                decimal_number = "";
                for (int j = 0; j < items.ksGetArrayCount(); j++)
                {
                    items.ksGetArrayItem(j, _ksTextItemParam);
                    decimal_number += _ksTextItemParam.s;
                }
            }
            #endregion

            #region получаем первичную применяемость

            _ksStamp.ksColumnNumber(25); //номер ячейки основной надписи
            _ksDynamicArray.ksClearArray();
            _ksDynamicArray = (ksDynamicArray)_ksStamp.ksGetStampColumnText(0);

            for (int i = 0; i < _ksDynamicArray.ksGetArrayCount(); i++)
            {
                _ksDynamicArray.ksGetArrayItem(i, _ksTextLineParam);
                items = (ksDynamicArray)_ksTextLineParam.GetTextItemArr();
                primary_applicability = "";
                for (int j = 0; j < items.ksGetArrayCount(); j++)
                {
                    items.ksGetArrayItem(j, _ksTextItemParam);
                    primary_applicability += _ksTextItemParam.s;
                }
            }
            if (primary_applicability.Length == 0)
                kompas.ksMessage("В документе отсутствует первичная применяемость!");
            #endregion

            #region получаем наименование
            _ksStamp.ksColumnNumber(1); //номер ячейки основной надписи          
            _ksDynamicArray = (ksDynamicArray)_ksStamp.ksGetStampColumnText(0);

            if (_ksDynamicArray.ksGetArrayCount() == 0)
                kompas.ksMessage("В документе отсутствует наименование!");

            for (int i = 0; i < _ksDynamicArray.ksGetArrayCount(); i++)
            {
                _ksDynamicArray.ksGetArrayItem(i, _ksTextLineParam);
                items = (ksDynamicArray)_ksTextLineParam.GetTextItemArr();
                name = "";
                for (int j = 0; j < items.ksGetArrayCount(); j++)
                {
                    items.ksGetArrayItem(j, _ksTextItemParam);                    
                    name += _ksTextItemParam.s;
                }
            }
            #endregion

            #region получаем головное изделие

            _ksStamp.ksColumnNumber(1000); //номер ячейки основной надписи
            _ksDynamicArray = (ksDynamicArray)_ksStamp.ksGetStampColumnText(0);

            for (int i = 0; i < _ksDynamicArray.ksGetArrayCount(); i++)
            {
                _ksDynamicArray.ksGetArrayItem(i, _ksTextLineParam);
                items = (ksDynamicArray)_ksTextLineParam.GetTextItemArr();
                head_piece = "";
                for (int j = 0; j < items.ksGetArrayCount(); j++)
                {
                    items.ksGetArrayItem(j, _ksTextItemParam);
                    head_piece += _ksTextItemParam.s;
                }
            }
            if (head_piece.Length == 0)
                kompas.ksMessage("В документе отсутствует головное изделие!");
            #endregion

            #region получаем литеру
            _ksStamp.ksColumnNumber(41); //номер ячейки основной надписи
            _ksDynamicArray = (ksDynamicArray)_ksStamp.ksGetStampColumnText(0);

            for (int i = 0; i < _ksDynamicArray.ksGetArrayCount(); i++)
            {
                _ksDynamicArray.ksGetArrayItem(i, _ksTextLineParam);
                items = (ksDynamicArray)_ksTextLineParam.GetTextItemArr();
                litera = "";
                for (int j = 0; j < items.ksGetArrayCount(); j++)
                {
                    items.ksGetArrayItem(j, _ksTextItemParam);
                    litera += _ksTextItemParam.s;
                }
            }
            if (litera.Length == 0)
                kompas.ksMessage("В документе отсутствует литера!");
            #endregion

            #region получаем дату из графы разработчика
            _ksStamp.ksColumnNumber(130); //номер ячейки основной надписи          
            _ksDynamicArray = (ksDynamicArray)_ksStamp.ksGetStampColumnText(0);

            for (int i = 0; i < _ksDynamicArray.ksGetArrayCount(); i++)
            {
                _ksDynamicArray.ksGetArrayItem(i, _ksTextLineParam);
                items = (ksDynamicArray)_ksTextLineParam.GetTextItemArr();
                date = "";
                for (int j = 0; j < items.ksGetArrayCount(); j++)
                {
                    items.ksGetArrayItem(j, _ksTextItemParam);
                    date += _ksTextItemParam.s;
                }
            }
            if (date.Length == 0)
                kompas.ksMessage("Дата в графе 'Разработал' отсутствует!");
            #endregion

            _ksStamp.ksCloseStamp();

            #region Получаем путь исходного активного документа
            _IKompasAPIObject = (IKompasAPIObject)Marshal.GetActiveObject("KOMPAS.Application.7");
            appl = (IApplication)_IKompasAPIObject.Application;
            docum = (IKompasDocument)appl.ActiveDocument;
            string path = docum.PathName;
            _ksDocument2D.ksSaveDocument(path);
            #endregion
            a = 2;
            return a;
        }
        public void get_active_path()
        {
            _IKompasAPIObject = (IKompasAPIObject)Marshal.GetActiveObject("KOMPAS.Application.7");
            appl = (IApplication)_IKompasAPIObject.Application;
            docum = (IKompasDocument)appl.ActiveDocument;

            sheets = (ILayoutSheets)docum.LayoutSheets;
            sheet = (ILayoutSheet)sheets.ItemByNumber[1];
            sheet.SheetType = 0;

            if (sheet.LayoutStyleNumber == 111)
                LayoutStyleNumber_count = 2;
            else
                if (sheet.LayoutStyleNumber == 1)
                LayoutStyleNumber_count = 3;
            else
                kompas.ksMessage("Выбран неверный номер оформления для документа");

            path = docum.Path;
            active_path = docum.PathName;
            active_path = active_path.Remove(active_path.Length - 4);
            active_name = docum.Name;
        }
        public void create_new_list()
        {
            _IKompasAPIObject = (IKompasAPIObject)Marshal.GetActiveObject("KOMPAS.Application.7");
            appl = (IApplication)_IKompasAPIObject.Application;
            appl.Visible = true;

            _IDocuments = (IDocuments)appl.Documents;
            _IDocuments.Add((DocumentTypeEnum)1, true);

            docum = (IKompasDocument)appl.ActiveDocument;
            sheets = (ILayoutSheets)docum.LayoutSheets;
            sheet = (ILayoutSheet)sheets.ItemByNumber[1];
            sheet.SheetType = 0;
            //sheet.LayoutLibraryFileName = "C:\\Users\\smirnovgv\\Desktop\\Radioavionika.LYT";

            if (LayoutStyleNumber_count == 2)
                sheet.LayoutStyleNumber = 2;
            else
            if (LayoutStyleNumber_count == 3)
                sheet.LayoutStyleNumber = 3;
            else
                kompas.ksMessage("Выбран неверный номер оформления для документа");

            format = (ISheetFormat)sheet.Format;
            format.Format = (ksDocumentFormatEnum)4;
            format.VerticalOrientation = true;
            sheet.Update();

            docum = (IKompasDocument)appl.ActiveDocument;
            sheets = (ILayoutSheets)docum.LayoutSheets;
            sheet = (ILayoutSheet)sheets.Add();
            sheet.SheetType = 0;
            //sheet.LayoutLibraryFileName = "C:\\Users\\smirnovgv\\Desktop\\Radioavionika.LYT";
            sheet.LayoutStyleNumber = 7;
            format = (ISheetFormat)sheet.Format;
            format.Format = (ksDocumentFormatEnum)4;
            format.VerticalOrientation = true;
            sheet.Update();

            docum = (IKompasDocument)appl.ActiveDocument;
            sheets = (ILayoutSheets)docum.LayoutSheets;
            sheet = (ILayoutSheet)sheets.Add();
            sheet.SheetType = 0;
            //sheet.LayoutLibraryFileName = "C:\\Users\\smirnovgv\\Desktop\\Radioavionika.LYT";

            if (LayoutStyleNumber_count == 2)
                sheet.LayoutStyleNumber = 5;
            else 
            if (LayoutStyleNumber_count == 3)
                sheet.LayoutStyleNumber = 6;
            else
                kompas.ksMessage("Выбран неверный номер оформления для документа");

            format = (ISheetFormat)sheet.Format;
            format.Format = (ksDocumentFormatEnum)4;
            format.VerticalOrientation = true;
            sheet.Update();
            //docum.Close(0);
        }
        public void filling_tables()
        {
            string _name;

            kompas = (KompasObject)Marshal.GetActiveObject("KOMPAS.Application.5");
            _ksDocument2D = (ksDocument2D)kompas.ActiveDocument2D();

            _ksStamp = (ksStamp)_ksDocument2D.GetStampEx(1);
            _ksStamp.ksOpenStamp();

            _ksStamp.ksColumnNumber(1000);
            _ksTextItemParam = (ksTextItemParam)kompas.GetParamStruct(31); //ko_TextItemParam = 31
            _ksTextItemParam.s = "";
            _ksTextItemParam.s = head_piece;
            _ksStamp.ksTextLine(_ksTextItemParam);

            _ksStamp.ksColumnNumber(1001); //номер ячейки основной надписи
            _ksTextItemParam = (ksTextItemParam)kompas.GetParamStruct(31); //ko_TextItemParam = 31
            _ksTextItemParam.s = "";
            _ksTextItemParam.s = decimal_number + " Д56-ЛУ";
            _ksDocument2D.ksTextLine(_ksTextItemParam);

            _ksStamp.ksColumnNumber(1002); //номер ячейки основной надписи
            ksDynamicArray items;
            ksDynamicArray lines;

            items = (ksDynamicArray)kompas.GetDynamicArray(4);
            lines = (ksDynamicArray)kompas.GetDynamicArray(3);

            //_ksParagraphParam = (ksParagraphParam)kompas.GetParamStruct(27);
            //_ksParagraphParam.Init();
            //_ksParagraphParam.x = 100;
            //_ksParagraphParam.y = 100;

            //_ksTextParam = (ksTextParam)kompas.GetParamStruct(28);
            //_ksTextParam.SetParagraphParam(_ksParagraphParam);
            //_ksTextParam.SetTextLineArr(lines);
            //_ksDocument2D.ksParagraph(_ksParagraphParam);

            _ksTextLineParam = (ksTextLineParam)kompas.GetParamStruct(29);
            _ksTextItemParam.s = "";
            _ksTextItemParam.s = decimal_number + " Д56";
            _ksDocument2D.ksTextLine(_ksTextItemParam);

            _TextItemFont = (TextItemFont)_ksTextItemParam.GetItemFont();
            _TextItemFont.SetBitVectorValue(0x1000, true);

            string _decimal_number = decimal_number.Remove(0, 5);
            _ksTextItemParam.s = "";
            _ksTextItemParam.s = "(" + _decimal_number + "_V01.zip)";
            _ksDocument2D.ksTextLine(_ksTextItemParam);

            _ksStamp.ksColumnNumber(110); //Разраб.
            _ksTextItemParam = (ksTextItemParam)kompas.GetParamStruct(31); //ko_TextItemParam = 31
            _name = textBox1.Text;
            _name = _name.Remove(0, 5);
            _ksTextItemParam.s = "";
            _ksTextItemParam.s = _name;
            _ksDocument2D.ksTextLine(_ksTextItemParam);

            _ksStamp.ksColumnNumber(111); //Пров.
            _ksTextItemParam = (ksTextItemParam)kompas.GetParamStruct(31); //ko_TextItemParam = 31
            _name = textBox2.Text;
            _name = _name.Remove(0, 5);
            _ksTextItemParam.s = "";
            _ksTextItemParam.s = _name;
            _ksDocument2D.ksTextLine(_ksTextItemParam);

            _ksStamp.ksColumnNumber(114); //Н.контр.
            _ksTextItemParam = (ksTextItemParam)kompas.GetParamStruct(31); //ko_TextItemParam = 31
            _name = textBox3.Text;
            _name = _name.Remove(0, 5);
            _ksTextItemParam.s = "";
            _ksTextItemParam.s = _name;
            _ksDocument2D.ksTextLine(_ksTextItemParam);

            _ksStamp.ksColumnNumber(115); //Утв.
            _ksTextItemParam = (ksTextItemParam)kompas.GetParamStruct(31); //ko_TextItemParam = 31
            _name = textBox4.Text;
            _name = _name.Remove(0, 5);
            _ksTextItemParam.s = "";
            _ksTextItemParam.s = _name;
            _ksDocument2D.ksTextLine(_ksTextItemParam);

            _ksStamp.ksColumnNumber(1057); //децимальный номер
            _ksTextItemParam = (ksTextItemParam)kompas.GetParamStruct(31); //ko_TextItemParam = 31
            _ksTextItemParam.s = "";
            _ksTextItemParam.s = decimal_number + " ВН"; 
            _ksStamp.ksTextLine(_ksTextItemParam);

            _ksStamp.ksColumnNumber(25); //первичная применяемость
            _ksTextItemParam = (ksTextItemParam)kompas.GetParamStruct(31); //ko_TextItemParam = 31
            _ksTextItemParam.s = "";
            _ksTextItemParam.s = primary_applicability; //запись децимального номера в переменную
            _ksStamp.ksTextLine(_ksTextItemParam);

            _ksStamp.ksColumnNumber(1056); //литера
            if (litera.Length == 1)
            {
                _ksTextLineParam = (ksTextLineParam)kompas.GetParamStruct(29);
                _TextItemFont.height = 5;
                _ksTextItemParam.s = litera;
                _ksDocument2D.ksTextLine(_ksTextItemParam);
            }
            else
            {
                if (litera.Length == 2)
                {
                    _ksTextLineParam = (ksTextLineParam)kompas.GetParamStruct(29);
                    _TextItemFont.height = 5;
                    _ksTextItemParam.s = litera[0].ToString();
                    _ksDocument2D.ksTextLine(_ksTextItemParam);

                    _TextItemFont = (TextItemFont)_ksTextItemParam.GetItemFont();
                    _TextItemFont.SetBitVectorValue(0x5, true);

                    _ksTextLineParam = (ksTextLineParam)kompas.GetParamStruct(29);
                    _TextItemFont.height = 2;
                    _ksTextItemParam.s = litera[1].ToString();
                    _ksDocument2D.ksTextLine(_ksTextItemParam);
                }
                else kompas.ksMessage("Ошибка в присвоении(построении) литеры! Проверьте исходный файл.");
            }                       

            #region записываем многострочное наименование         
            _ksStamp.ksColumnNumber(1058); //многострочное наименование
            _ksTextItemParam = (ksTextItemParam)kompas.GetParamStruct(31); //ko_TextItemParam = 31
            _TextItemFont = (TextItemFont)_ksTextItemParam.GetItemFont();
            _TextItemFont.height = 7;
            _ksTextItemParam.s = "";
            _ksTextItemParam.s = name; 
            _ksStamp.ksTextLine(_ksTextItemParam);

            _TextItemFont = (TextItemFont)_ksTextItemParam.GetItemFont();
            _TextItemFont.SetBitVectorValue(0x1000, true);
            _TextItemFont.height = 5;
            _ksTextItemParam.s = "";
            _ksTextItemParam.s = "Ведомость документов на";
            _ksDocument2D.ksTextLine(_ksTextItemParam);

            _TextItemFont = (TextItemFont)_ksTextItemParam.GetItemFont();
            _TextItemFont.SetBitVectorValue(0x1000, true);
            _TextItemFont.height = 5;
            _ksTextItemParam.s = "";
            _ksTextItemParam.s = "носителях данных";
            _ksDocument2D.ksTextLine(_ksTextItemParam);
            #endregion

            _ksStamp.ksCloseStamp();
        }

        public void filling_tables2()
        {
            kompas = (KompasObject)Marshal.GetActiveObject("KOMPAS.Application.5");
            _ksDocument2D = (ksDocument2D)kompas.ActiveDocument2D();

            _ksStamp = (ksStamp)_ksDocument2D.GetStampEx(2);
            _ksStamp.ksOpenStamp();

            _ksStamp.ksColumnNumber(25); //Первичная применяемость
            _ksTextItemParam = (ksTextItemParam)kompas.GetParamStruct(31); //ko_TextItemParam = 31
            _TextItemFont.height = 5;
            _ksTextItemParam.s = "";
            _ksTextItemParam.s = decimal_number + " ВН";
            _ksDocument2D.ksTextLine(_ksTextItemParam);

            _ksStamp.ksColumnNumber(1000);
            _ksTextItemParam = (ksTextItemParam)kompas.GetParamStruct(31); //ko_TextItemParam = 31
            _ksTextItemParam.s = "";
            _ksTextItemParam.s = head_piece;
            _ksStamp.ksTextLine(_ksTextItemParam);

           

            _ksStamp.ksColumnNumber(1050);
            _ksTextItemParam = (ksTextItemParam)kompas.GetParamStruct(31); //ko_TextItemParam = 31
            _TextItemFont.height = 3.5;
            _ksTextItemParam.s = "";
            _ksTextItemParam.s = name;
            _ksStamp.ksTextLine(_ksTextItemParam);

            _ksStamp.ksColumnNumber(1051);
            _ksTextItemParam = (ksTextItemParam)kompas.GetParamStruct(31); //ko_TextItemParam = 31
            _ksTextItemParam.s = "";
            _ksTextItemParam.s = decimal_number + " Д56-ЛУ";
            _ksStamp.ksTextLine(_ksTextItemParam);

            #region литера
            _ksStamp.ksColumnNumber(1056); //литера
            if (litera.Length == 1)
            {
                _ksTextLineParam = (ksTextLineParam)kompas.GetParamStruct(29);
                _TextItemFont.height = 5;
                _ksTextItemParam.s = "Лит. " + litera;
                _ksDocument2D.ksTextLine(_ksTextItemParam);
            }
            else
            {
                if (litera.Length == 2)
                {
                    _ksTextLineParam = (ksTextLineParam)kompas.GetParamStruct(29);
                    _TextItemFont.height = 5;
                    _ksTextItemParam.s = "Лит. " + litera[0].ToString();
                    _ksDocument2D.ksTextLine(_ksTextItemParam);

                    _TextItemFont = (TextItemFont)_ksTextItemParam.GetItemFont();
                    _TextItemFont.SetBitVectorValue(0x5, true);

                    _ksTextLineParam = (ksTextLineParam)kompas.GetParamStruct(29);
                    _TextItemFont.height = 2;
                    _ksTextItemParam.s = litera[1].ToString();
                    _ksDocument2D.ksTextLine(_ksTextItemParam);
                }                
            }
            #endregion
            #region заполнение листа утверждений фамилиями
            _ksStamp.ksColumnNumber(1052);
            _ksTextItemParam = (ksTextItemParam)kompas.GetParamStruct(31); //ko_TextItemParam = 31
            _ksTextItemParam.s = textBox4.Text;
            _ksStamp.ksTextLine(_ksTextItemParam);

            _ksStamp.ksColumnNumber(1053);
            _ksTextItemParam = (ksTextItemParam)kompas.GetParamStruct(31); //ko_TextItemParam = 31
            _ksTextItemParam.s = "";
            _ksTextItemParam.s = textBox2.Text;
            _ksStamp.ksTextLine(_ksTextItemParam);

            _ksStamp.ksColumnNumber(1054);
            _ksTextItemParam = (ksTextItemParam)kompas.GetParamStruct(31); //ko_TextItemParam = 31
            _ksTextItemParam.s = "";
            _ksTextItemParam.s = textBox1.Text;
            _ksStamp.ksTextLine(_ksTextItemParam);

            _ksStamp.ksColumnNumber(1055);
            _ksTextItemParam = (ksTextItemParam)kompas.GetParamStruct(31); //ko_TextItemParam = 31
            _ksTextItemParam.s = "";
            _ksTextItemParam.s = textBox3.Text;
            _ksStamp.ksTextLine(_ksTextItemParam);
            #endregion

            _ksStamp.ksCloseStamp();
        }
        public void filling_tables3()
        {
            string _name;

            kompas = (KompasObject)Marshal.GetActiveObject("KOMPAS.Application.5");
            _ksDocument2D = (ksDocument2D)kompas.ActiveDocument2D();

            _ksStamp = (ksStamp)_ksDocument2D.GetStampEx(3);
            _ksStamp.ksOpenStamp();

            _ksStamp.ksColumnNumber(110); //Разраб.
            _ksTextItemParam = (ksTextItemParam)kompas.GetParamStruct(31); //ko_TextItemParam = 31
            _name = textBox1.Text;
            _name = _name.Remove(0, 5);
            _ksTextItemParam.s = "";
            _ksTextItemParam.s = _name;
            _ksDocument2D.ksTextLine(_ksTextItemParam);

            _ksStamp.ksColumnNumber(111); //Пров.
            _ksTextItemParam = (ksTextItemParam)kompas.GetParamStruct(31); //ko_TextItemParam = 31
            _name = textBox2.Text;
            _name = _name.Remove(0, 5);
            _ksTextItemParam.s = "";
            _ksTextItemParam.s = _name;
            _ksDocument2D.ksTextLine(_ksTextItemParam);

            _ksStamp.ksColumnNumber(114); //Н.контр.
            _ksTextItemParam = (ksTextItemParam)kompas.GetParamStruct(31); //ko_TextItemParam = 31
            _name = textBox3.Text;
            _name = _name.Remove(0, 5);
            _ksTextItemParam.s = "";
            _ksTextItemParam.s = _name;
            _ksDocument2D.ksTextLine(_ksTextItemParam);

            _ksStamp.ksColumnNumber(115); //Утв.
            _ksTextItemParam = (ksTextItemParam)kompas.GetParamStruct(31); //ko_TextItemParam = 31
            _name = textBox4.Text;
            _name = _name.Remove(0, 5);
            _ksTextItemParam.s = "";
            _ksTextItemParam.s = _name;
            _ksDocument2D.ksTextLine(_ksTextItemParam);

            _ksStamp.ksColumnNumber(1057); //Децимальный номер
            _ksTextItemParam = (ksTextItemParam)kompas.GetParamStruct(31); //ko_TextItemParam = 31
            _ksTextItemParam.s = "";
            _ksTextItemParam.s = decimal_number + " Д56";
            _ksDocument2D.ksTextLine(_ksTextItemParam);

            _ksStamp.ksColumnNumber(25); //Первичная применяемость
            _ksTextItemParam = (ksTextItemParam)kompas.GetParamStruct(31); //ko_TextItemParam = 31
            _ksTextItemParam.s = "";
            _ksTextItemParam.s = decimal_number + " ВН";
            _ksDocument2D.ksTextLine(_ksTextItemParam);

            _ksStamp.ksColumnNumber(1000); //Головное изделие
            _ksTextItemParam = (ksTextItemParam)kompas.GetParamStruct(31); //ko_TextItemParam = 31
            _ksTextItemParam.s = "";
            _ksTextItemParam.s = head_piece;
            _ksStamp.ksTextLine(_ksTextItemParam);

            _ksStamp.ksColumnNumber(1058); //Наименование
            _ksTextItemParam = (ksTextItemParam)kompas.GetParamStruct(31); //ko_TextItemParam = 31
            _TextItemFont = (TextItemFont)_ksTextItemParam.GetItemFont();
            _TextItemFont.height = 7;
            _ksTextItemParam.s = "";
            _ksTextItemParam.s = name;
            _ksDocument2D.ksTextLine(_ksTextItemParam);

            _TextItemFont = (TextItemFont)_ksTextItemParam.GetItemFont();
            _TextItemFont.SetBitVectorValue(0x1000, true);
            _TextItemFont.height = 5;
            _ksTextItemParam.s = "";
            _ksTextItemParam.s = "Ведомость НДЗ";
            _ksDocument2D.ksTextLine(_ksTextItemParam);

            _ksStamp.ksColumnNumber(1000); //Головное изделие
            _ksTextItemParam = (ksTextItemParam)kompas.GetParamStruct(31); //ko_TextItemParam = 31
            _ksTextItemParam.s = "";
            _ksTextItemParam.s = head_piece; //запись децимального номера в переменную
            _ksStamp.ksTextLine(_ksTextItemParam);

            _ksStamp.ksColumnNumber(1003);//Идентификатор тома
            _ksTextItemParam = (ksTextItemParam)kompas.GetParamStruct(31); //ko_TextItemParam = 31
            string number = decimal_number;
            number = number.Remove(0, 5);
            number = number.Replace(".", "_");
            _ksTextItemParam.s = "";
            _ksTextItemParam.s = number;
            _ksStamp.ksTextLine(_ksTextItemParam);

            _ksStamp.ksColumnNumber(1006);//Дата создания НДЗ
            _ksTextItemParam = (ksTextItemParam)kompas.GetParamStruct(31); //ko_TextItemParam = 31
            _ksTextItemParam.s = "";
            _ksTextItemParam.s = date;
            _ksStamp.ksTextLine(_ksTextItemParam);

            _ksStamp.ksColumnNumber(1007); //Децимальный номер
            _ksTextItemParam = (ksTextItemParam)kompas.GetParamStruct(31); //ko_TextItemParam = 31
            _ksTextItemParam.s = "";
            _ksTextItemParam.s = decimal_number + " Д56";
            _ksDocument2D.ksTextLine(_ksTextItemParam);

            string _decimal_number = decimal_number.Remove(0, 5);
            _ksStamp.ksColumnNumber(1008);
            _ksTextItemParam = (ksTextItemParam)kompas.GetParamStruct(31); //ko_TextItemParam = 31
            _ksTextItemParam.s = "";
            _ksTextItemParam.s = _decimal_number + "_V01.zip";
            _ksDocument2D.ksTextLine(_ksTextItemParam);

            _ksStamp.ksColumnNumber(1009);
            _ksTextItemParam = (ksTextItemParam)kompas.GetParamStruct(31); //ko_TextItemParam = 31
            _ksTextItemParam.s = "";
            _ksTextItemParam.s = date;
            _ksStamp.ksTextLine(_ksTextItemParam);

            #region литера
            _ksStamp.ksColumnNumber(1056); //литера
            if (litera.Length == 1)
            {
                _ksTextLineParam = (ksTextLineParam)kompas.GetParamStruct(29);
                _TextItemFont.height = 5;
                _ksTextItemParam.s = litera;
                _ksDocument2D.ksTextLine(_ksTextItemParam);
            }
            else
            {
                if (litera.Length == 2)
                {
                    _ksTextLineParam = (ksTextLineParam)kompas.GetParamStruct(29);
                    _TextItemFont.height = 5;
                    _ksTextItemParam.s = litera[0].ToString();
                    _ksDocument2D.ksTextLine(_ksTextItemParam);

                    _TextItemFont = (TextItemFont)_ksTextItemParam.GetItemFont();
                    _TextItemFont.SetBitVectorValue(0x5, true);

                    _ksTextLineParam = (ksTextLineParam)kompas.GetParamStruct(29);
                    _TextItemFont.height = 2;
                    _ksTextItemParam.s = litera[1].ToString();
                    _ksDocument2D.ksTextLine(_ksTextItemParam);
                }
            }
            #endregion

            _ksStamp.ksCloseStamp();

            _IKompasAPIObject = (IKompasAPIObject)Marshal.GetActiveObject("KOMPAS.Application.7");
            appl = (IApplication)_IKompasAPIObject.Application;
            docum = (IKompasDocument)appl.ActiveDocument;

            docum.SaveAs(path + decimal_number + " ВН - " + name + ".cdw");
            docum.Close(0);
        }
        public string current_data_time(int n) //n = 1 - год; n = 2 - дата.месяц.год
        {
            string _date = Convert.ToString(DateTime.Today);
            char[] ar = _date.ToCharArray();
            int i;
            string new_date = null;
            if (n == 1)
            {
                for (i = 6; i < _date.Length - 8; i++)
                    new_date = new_date + Convert.ToString(ar[i]);
                new_date += "г.";
                return new_date;
            }
            if (n == 2)
            {
                for (i = 0; i < _date.Length - 8; i++)
                    if (i != _date.Length - 12 && i != _date.Length - 11)
                        new_date = new_date + Convert.ToString(ar[i]);
                return new_date;
            }
            return new_date;
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            FileInfo MyFile = new FileInfo(global_path_to_xmlfile);
            if (MyFile.Exists == false)
            {
                FileStream fs = MyFile.Create();
                fs.Close();
                XDocument xdoc = new XDocument();
                XElement xmlfile = new XElement("xmlfile");

                xdoc.Add(xmlfile);
                xdoc.Save(global_path_to_xmlfile);

                XmlDocument xDoc = new XmlDocument();
                xDoc.Load(global_path_to_xmlfile);
                XmlElement xRoot = xDoc.DocumentElement;

                XmlElement list = xDoc.CreateElement("list");

                XmlElement elem1 = xDoc.CreateElement("Разраб.");
                XmlElement elem2 = xDoc.CreateElement("Пров.");
                XmlElement elem3 = xDoc.CreateElement("Н.контр.");
                XmlElement elem4 = xDoc.CreateElement("Утв.");

                XmlText elem1_text = xDoc.CreateTextNode(textBox1.Text);
                XmlText elem2_text = xDoc.CreateTextNode(textBox2.Text);
                XmlText elem3_text = xDoc.CreateTextNode(textBox3.Text);
                XmlText elem4_text = xDoc.CreateTextNode(textBox4.Text);
                
                xRoot.AppendChild(list);

                list.AppendChild(elem1);
                list.AppendChild(elem2);
                list.AppendChild(elem3);
                list.AppendChild(elem4);

                elem1.AppendChild(elem1_text);
                elem2.AppendChild(elem2_text);
                elem3.AppendChild(elem3_text);
                elem4.AppendChild(elem4_text);

                xDoc.Save(global_path_to_xmlfile);
            }
            else
            {
                XmlDocument xDoc = new XmlDocument();
                xDoc.Load(global_path_to_xmlfile);
                XmlElement xRoot = xDoc.DocumentElement;

                int i = 0;
                TextBox[] textboxes = new TextBox[4]
                { textBox1, textBox2, textBox3, textBox4 };

                foreach (XmlNode xnode in xRoot.ChildNodes)
                {
                    foreach (XmlNode xnode2 in xnode.ChildNodes)
                    {
                        textboxes[i].Text = xnode2.InnerText;
                        i++;
                    }
                }
            }

        }
        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "" || textBox2.Text == "" || textBox3.Text == "" || textBox4.Text == "")
            {
                kompas = (KompasObject)Marshal.GetActiveObject("KOMPAS.Application.5");
                kompas.ksMessage("Не все текстовые поля заполнены.");
                return;
            } 
            else
            button1.Enabled = false;

            XmlDocument xDoc = new XmlDocument();
            xDoc.Load(global_path_to_xmlfile);
            XmlElement xRoot = xDoc.DocumentElement;

            int i=0;
            string[] textboxes = new string[4] 
            { textBox1.Text, textBox2.Text, textBox3.Text, textBox4.Text };

            foreach (XmlNode xnode in xRoot.ChildNodes)
            {
                foreach (XmlNode xnode2 in xnode.ChildNodes)
                {
                    xnode2.InnerText = textboxes[i];
                    i++;
                }
            }
            xDoc.Save(global_path_to_xmlfile);
        }
        #region событие изменения состояния кнопки
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            button1.Enabled = true;
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            button1.Enabled = true;
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            button1.Enabled = true;
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            button1.Enabled = true;
        }
        #endregion
    }
}