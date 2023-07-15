using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Cells;
using System.Text.RegularExpressions;
using System.Xml.Serialization;
using System.Xml;
using Excel = Microsoft.Office.Interop;
namespace Final
{
    class Program
    {
        static void Main(string[] args)
        {
            //XMLHealer -- проверять xml файл на недопустимые, к примеру, 16-тиричные символы
            //Transliterator -- Словарь для транслитерации русских фамилий обратно

            int A = 100;
            while (A != 5)
            {
                Console.WriteLine("Какое слияние произвести:\n" +
                    "1)Слияние внутри одного документа( для файлов расширением .xlsx)\n 2)Слияние внутри одного документа одинаковых по структуре файлов" +
                    " (для файлов с расширением .xlsx\n 3)Слияние разных документов .xlsx \n4) Слияние разных документов, xml & xlsx\n5)Выход");
                A = Convert.ToInt32(Console.ReadLine());
                switch (A)
                {
                    case 1:
                        {
                            //Программа по слиянию листов в одном доке xlsx , далее перевод в xml в конце с Save();
                            //ЛИСТЫ ВСЕ ПО СТРУКТУРЫ И КОЛИЧЕСТВО СТРОК ЗАГОЛОВКОВ И ТИПОВ ДАННЫХ В НИХ РАЗНЫЕ, ИХ ОЧЕНЬ КРУТО
                            //ДРУГ С ДРУГОМ СЛИВАТЬ


                            Console.WriteLine("Введите путь для открытия файла");
                            string Path = Console.ReadLine();
                            Workbook workbook = new Workbook(Path);
                            Console.WriteLine("Выберите лист, в который перенести:");
                            for (int i = 0; i < workbook.Worksheets.Count; i++) //Вывод списка листов
                            {
                                Console.Write(i + 1); Console.Write(" | ");
                                Console.WriteLine(workbook.Worksheets[i].Name);
                            }



                            int NumberSheetTo = Convert.ToInt32(Console.ReadLine());    //Задание листа, из которого перенос
                            NumberSheetTo--;
                            if (NumberSheetTo < 0 || NumberSheetTo >= workbook.Worksheets.Count)
                            {
                                Console.WriteLine("Такого листа не существует");
                            }



                            Worksheet WSTo = workbook.Worksheets[NumberSheetTo];
                            Console.WriteLine(WSTo.Name);
                            Console.WriteLine("Выберите лист, который перенести:");



                            int NumberSheetFrom = Convert.ToInt32(Console.ReadLine());  //Задание листа, в который перенос
                            NumberSheetFrom--;
                            if (NumberSheetFrom < 0 || NumberSheetFrom >= workbook.Worksheets.Count)
                            {
                                Console.WriteLine("Такого листа не существует");
                            }



                            Worksheet WSFrom = workbook.Worksheets[NumberSheetFrom];
                            Console.WriteLine(WSFrom.Name);


                            Console.Write("Сколько строк заголовков в переносимом листе "); Console.WriteLine(WSFrom.Name); Console.WriteLine("?");
                            int Zagolovki = Convert.ToInt32(Console.ReadLine());
                            int ZagSave = Zagolovki;

                            Console.WriteLine("сколько строк заголовков в листе, куда перенос"); Console.WriteLine(WSTo.Name);
                            int ZagolovkiTo = Convert.ToInt32(Console.ReadLine());
                            int ZagToSave = ZagolovkiTo - 1;
                            int ZagToSaveTo2 = Zagolovki;
                            int CounterRowsTo = WSTo.Cells.Rows.Count;  //Количество строк в листе куда перенос
                            int CounterRowsFrom = WSFrom.Cells.Rows.Count - Zagolovki; //Количество строк в листе откуда перенос

                            for (int j = 0; j < WSFrom.Cells.Columns.Count; j++)
                            {
                                for (int m = 0; m < WSTo.Cells.Columns.Count; m++)
                                {
                                    string ColName = Transliterator.ColumnParserToString[m];
                                    Console.Write(ColName); Console.Write("|===|");
                                    Console.WriteLine(WSTo.Cells[ZagToSave - 1, m].Value);
                                }
                                //
                                Console.WriteLine("В какую колонну из этих перенести ");// Console.Write(WSTo.Name);Console.WriteLine();

                                string columnato = Console.ReadLine();
                                int columnatoint = Transliterator.ColumnParser[columnato];
                                //
                                for (int n = 0; n < WSFrom.Cells.Columns.Count; n++)
                                {
                                    string ColName = Transliterator.ColumnParserToString[n];
                                    Console.Write(ColName); Console.Write("|===|");
                                    Console.WriteLine(WSFrom.Cells[ZagSave - 2, n].Value);
                                }
                                Console.WriteLine("Из какой колонны перенести ");// Console.Write(WSFrom.Name);Console.WriteLine();
                                string columnafrom = Console.ReadLine();
                                Zagolovki = ZagSave;// Обнуляем, откуда плыть по колонне после нижнего внутр. цикла
                                int columnafromint = Transliterator.ColumnParser[columnafrom];
                                //
                                for (int i = CounterRowsTo; i < CounterRowsTo + CounterRowsFrom; i++)
                                {
                                    WSTo.Cells[i, columnatoint].Value = WSFrom.Cells[Zagolovki, columnafromint].Value;
                                    Zagolovki++;
                                }
                            }
                            /////////////////////////////////////РАБОТАЕТ//////////////////////////
                            /////Листы соединены, идём далее, проверяем колонну с авторами на транслит
                            Console.WriteLine("В какой колонне указаны авторы?");
                            string AuthorColumn = Console.ReadLine();
                            int AuthorColumnINT = Transliterator.ColumnParser[AuthorColumn];
                            Console.WriteLine("сколько строк заголовков в листе"); Console.WriteLine(WSTo.Name);
                            Console.WriteLine("Начинаю перевод:");

                            for (int i = ZagSave; i < WSTo.Cells.Rows.Count; i++)
                            {
                                var F = WSTo.Cells[i, AuthorColumnINT].Value;
                                string f = F.ToString();
                                foreach (var name in Transliterator.Urusification)
                                {
                                    f = Regex.Replace(f, name.Key, name.Value, RegexOptions.IgnoreCase);
                                }
                                //}
                                WSTo.Cells[i, AuthorColumnINT].Value = f;
                                Console.WriteLine(WSTo.Cells[i, AuthorColumnINT].Value);
                            }
                            Console.WriteLine("Перевод успешно завершён");





                            /////Фамилии переведены, идём далее, проверяем строки на совпадения, по институтам, факультетам,
                            /////названиям и авторам
                            Console.WriteLine("По каким колоннам сравнить строки (через пробел) ?");
                            string Kolonni = Console.ReadLine();
                            string[] Columnes = Kolonni.Split(' ');
                            int[] ArrayColumnes = new int[Columnes.Length];
                            for (int i = 0; i < ArrayColumnes.Length; i++)
                            {
                                string F = Columnes[i].ToString();
                                foreach (var elem in Transliterator.ColumnParser)
                                {
                                    F = F.Replace(elem.Key, elem.Value.ToString());
                                }
                            ArrayColumnes[i] = Convert.ToInt32(F);
                            //Console.WriteLine(ArrayColumnes[i]);
                            }
                            for (int chetchik = 0; chetchik < ArrayColumnes.Length; chetchik++) 
                            {
                                int CurrentColumn = ArrayColumnes[chetchik];
                                for (int i = ZagToSaveTo2; i < WSTo.Cells.Rows.Count; i++)
                                {
                                    for (int j = ZagToSaveTo2; j < WSTo.Cells.Rows.Count; j++)
                                    {
                                        if (WSTo.Cells[i, CurrentColumn].Value == WSTo.Cells[j, CurrentColumn].Value)
                                        {
                                            WSTo.Cells.DeleteRow(j);
                                        }
                                    }
                                } 
                            }
                            ///// Сохранение
                            Console.WriteLine("Введите путь для сохранения");
                            Console.WriteLine("Введите формат xml, чтобы сохранить в этом формате");
                            string PathToSave = Console.ReadLine();
                            workbook.Save(PathToSave);
                            /////
                            break;
                        }



                    case 2:
                        {
                            Console.WriteLine("Введите путь для открытия файла");
                            string Path = Console.ReadLine();
                            Workbook workbook = new Workbook(Path);
                            Console.WriteLine("Выберите лист, в который перенести:");
                            for (int i = 0; i < workbook.Worksheets.Count; i++) //Вывод списка листов
                            {
                                Console.Write(i + 1); Console.Write(" | ");
                                Console.WriteLine(workbook.Worksheets[i].Name);
                            }



                            int NumberSheetTo = Convert.ToInt32(Console.ReadLine());    //Задание листа, из которого перенос
                            NumberSheetTo--;
                            if (NumberSheetTo < 0 || NumberSheetTo >= workbook.Worksheets.Count)
                            {
                                Console.WriteLine("Такого листа не существует");
                            }
                            


                            Worksheet WSTo = workbook.Worksheets[NumberSheetTo];
                            Console.WriteLine(WSTo.Name);
                            Console.WriteLine("Выберите лист, который перенести:");



                            int NumberSheetFrom = Convert.ToInt32(Console.ReadLine());  //Задание листа, в который перенос
                            NumberSheetFrom--;
                            if (NumberSheetFrom < 0 || NumberSheetFrom >= workbook.Worksheets.Count)
                            {
                                Console.WriteLine("Такого листа не существует");
                            }
                            Worksheet WSFrom = workbook.Worksheets[NumberSheetFrom];
                            Console.WriteLine(WSFrom.Name);

                            Console.WriteLine("Сколько строк занимают заголовки?");
                            int Zagolovki = Convert.ToInt32(Console.ReadLine());
                            int CounterRowsTo = WSTo.Cells.Rows.Count;
                            int CounterRowsFrom = WSFrom.Cells.Rows.Count;
                            int CounterColumnsTo = WSTo.Cells.Columns.Count;
                            int CounterColumnsFrom = WSFrom.Cells.Columns.Count;
                            int CountColumns;
                            if (CounterColumnsTo > CounterColumnsFrom)
                            {
                                CountColumns = CounterColumnsTo;
                            }
                            else if (CounterColumnsTo < CounterColumnsFrom)
                            {
                                CountColumns = CounterColumnsFrom;
                            }
                            else if (CounterColumnsFrom == CounterColumnsTo)
                            {
                                CountColumns = (CounterColumnsFrom + CounterColumnsTo) / 2;
                            }
                            int a = Zagolovki;
                            for (int i = 0; i < CounterRowsTo + CounterRowsFrom; i++)
                            {
                                for (int j = 0; j < CounterColumnsTo; j++)
                                {
                                    WSTo.Cells[i + CounterRowsTo, j].Value = WSFrom.Cells[a, j].Value;

                                }
                                a++;
                            }
                            Console.WriteLine("Листы объединены");
                            Console.WriteLine("В какой колонне указаны авторы?");
                            string AuthorColumn = Console.ReadLine();
                            int AuthorColumnINT = Transliterator.ColumnParser[AuthorColumn];
                            Console.WriteLine("сколько строк заголовков в листе"); Console.WriteLine(WSTo.Name);
                            Console.WriteLine("Начинаю перевод:");

                            for (int i = Zagolovki + 1; i < WSTo.Cells.Rows.Count; i++)
                            {
                                var F = WSTo.Cells[i, AuthorColumnINT].Value;
                                string f = F.ToString();
                                foreach (var name in Transliterator.Urusification)
                                {
                                    f = Regex.Replace(f, name.Key, name.Value, RegexOptions.IgnoreCase);
                                }
                                //}
                                WSTo.Cells[i, AuthorColumnINT].Value = f;
                                Console.WriteLine(WSTo.Cells[i, AuthorColumnINT].Value);
                            }
                            Console.WriteLine("Перевод успешно завершён");
                            int del = 0;
                            for (int chetchik = 0; chetchik < CounterColumnsTo; chetchik++)
                            {
                                int CurrentColumn = CounterColumnsTo;
                                for (int i = Zagolovki + 1; i < WSTo.Cells.Rows.Count; i++)
                                {
                                    for (int j = Zagolovki + 2; j < WSTo.Cells.Rows.Count; j++)
                                    {
                                        if (WSTo.Cells[i, CurrentColumn].Value == WSTo.Cells[j, CurrentColumn].Value)
                                        {
                                            del++;
                                        }
                                        if (del == CounterColumnsTo) 
                                        {
                                            WSTo.Cells.DeleteRow(j);
                                        }
                                    }
                                }
                            }
                            Console.WriteLine("Листы слиты");
                            Console.WriteLine("Введите путь для сохранения");
                            Console.WriteLine("Введите формат xml, чтобы сохранить в этом формате");
                            string PathToSave = Console.ReadLine();
                            workbook.Save(PathToSave);
                            break;
                        }
                    case 3:
                        {
                            
                                



                                Console.WriteLine("Введите путь для открытия 1 книги (куда перенести)");
                                string Path = Console.ReadLine();
                                Console.WriteLine("Введите путь для открытия 2 книги (откуда перенести)");
                                string Path2 = Console.ReadLine();




                                Workbook workbook = new Workbook(Path);
                                Workbook workbook2 = new Workbook(Path2);
                                Console.WriteLine("Выберите лист, в который перенести:");
                                for (int i = 0; i < workbook.Worksheets.Count; i++) //Вывод списка листов
                                {
                                    Console.Write(i + 1); Console.Write(" | ");
                                    Console.WriteLine(workbook.Worksheets[i].Name);
                                }



                                int NumberSheetTo = Convert.ToInt32(Console.ReadLine());    //Задание листа, из которого перенос
                                NumberSheetTo--;
                                if (NumberSheetTo < 0 || NumberSheetTo >= workbook.Worksheets.Count)
                                {
                                    Console.WriteLine("Такого листа не существует");
                                }



                                Worksheet WSTo = workbook2.Worksheets[NumberSheetTo];
                                Console.WriteLine(WSTo.Name);
                                Console.WriteLine("Выберите лист, который перенести:");



                                int NumberSheetFrom = Convert.ToInt32(Console.ReadLine());  //Задание листа, в который перенос
                                NumberSheetFrom--;
                                if (NumberSheetFrom < 0 || NumberSheetFrom >= workbook.Worksheets.Count)
                                {
                                    Console.WriteLine("Такого листа не существует");
                                }



                                Worksheet WSFrom = workbook2.Worksheets[NumberSheetFrom];
                                Console.WriteLine(WSFrom.Name);


                                Console.Write("Сколько строк заголовков в переносимом листе "); Console.WriteLine(WSFrom.Name); Console.WriteLine("?");
                                int Zagolovki = Convert.ToInt32(Console.ReadLine());
                                int ZagSave = Zagolovki;

                                Console.WriteLine("сколько строк заголовков в листе, куда перенос"); Console.WriteLine(WSTo.Name);
                                int ZagolovkiTo = Convert.ToInt32(Console.ReadLine());
                                int ZagToSave = ZagolovkiTo - 1;
                                int ZagToSaveTo2 = Zagolovki;
                                int CounterRowsTo = WSTo.Cells.Rows.Count;  //Количество строк в листе куда перенос
                                int CounterRowsFrom = WSFrom.Cells.Rows.Count - Zagolovki; //Количество строк в листе откуда перенос

                                for (int j = 0; j < WSFrom.Cells.Columns.Count; j++)
                                {
                                    for (int m = 0; m < WSTo.Cells.Columns.Count; m++)
                                    {
                                        string ColName = Transliterator.ColumnParserToString[m];
                                        Console.Write(ColName); Console.Write("|===|");
                                        Console.WriteLine(WSTo.Cells[ZagToSave - 1, m].Value);
                                    }
                                    //
                                    Console.WriteLine("В какую колонну из этих перенести ");// Console.Write(WSTo.Name);Console.WriteLine();

                                    string columnato = Console.ReadLine();
                                    int columnatoint = Transliterator.ColumnParser[columnato];
                                    //
                                    for (int n = 0; n < WSFrom.Cells.Columns.Count; n++)
                                    {
                                        string ColName = Transliterator.ColumnParserToString[n];
                                        Console.Write(ColName); Console.Write("|===|");
                                        Console.WriteLine(WSFrom.Cells[ZagSave - 2, n].Value);
                                    }
                                    Console.WriteLine("Из какой колонны перенести ");// Console.Write(WSFrom.Name);Console.WriteLine();
                                    string columnafrom = Console.ReadLine();
                                    Zagolovki = ZagSave;// Обнуляем, откуда плыть по колонне после нижнего внутр. цикла
                                    int columnafromint = Transliterator.ColumnParser[columnafrom];
                                    //
                                    for (int i = CounterRowsTo; i < CounterRowsTo + CounterRowsFrom; i++)
                                    {
                                        WSTo.Cells[i, columnatoint].Value = WSFrom.Cells[Zagolovki, columnafromint].Value;
                                        Zagolovki++;
                                    }
                                }
                                /////////////////////////////////////РАБОТАЕТ//////////////////////////
                                /////Листы соединены, идём далее, проверяем колонну с авторами на транслит
                                Console.WriteLine("В какой колонне указаны авторы?");
                                string AuthorColumn = Console.ReadLine();
                                int AuthorColumnINT = Transliterator.ColumnParser[AuthorColumn];
                                Console.WriteLine("сколько строк заголовков в листе"); Console.WriteLine(WSTo.Name);
                                Console.WriteLine("Начинаю перевод:");

                                for (int i = ZagSave; i < WSTo.Cells.Rows.Count; i++)
                                {
                                    var F = WSTo.Cells[i, AuthorColumnINT].Value;
                                    string f = F.ToString();
                                    foreach (var name in Transliterator.Urusification)
                                    {
                                        f = Regex.Replace(f, name.Key, name.Value, RegexOptions.IgnoreCase);
                                    }
                                    //}
                                    WSTo.Cells[i, AuthorColumnINT].Value = f;
                                    Console.WriteLine(WSTo.Cells[i, AuthorColumnINT].Value);
                                }
                                Console.WriteLine("Перевод успешно завершён");





                                /////Фамилии переведены, идём далее, проверяем строки на совпадения, по институтам, факультетам,
                                /////названиям и авторам
                                Console.WriteLine("По каким колоннам сравнить строки (через пробел) ");
                                string Kolonni = Console.ReadLine();
                                string[] Columnes = Kolonni.Split(' ');
                                int[] ArrayColumnes = new int[Columnes.Length];
                                for (int i = 0; i < ArrayColumnes.Length; i++)
                                {
                                    string F = Columnes[i].ToString();
                                    foreach (var elem in Transliterator.ColumnParser)
                                    {
                                        F = F.Replace(elem.Key, elem.Value.ToString());
                                    }
                                    ArrayColumnes[i] = Convert.ToInt32(F);
                                    //Console.WriteLine(ArrayColumnes[i]);
                                }
                                for (int chetchik = 0; chetchik < ArrayColumnes.Length; chetchik++)
                                {
                                    int CurrentColumn = ArrayColumnes[chetchik];
                                    for (int i = ZagToSaveTo2; i < WSTo.Cells.Rows.Count; i++)
                                    {
                                        for (int j = ZagToSaveTo2; j < WSTo.Cells.Rows.Count; j++)
                                        {
                                            if (WSTo.Cells[i, CurrentColumn].Value == WSTo.Cells[j, CurrentColumn].Value)
                                            {
                                                WSTo.Cells.DeleteRow(j);
                                            }
                                        }
                                    }
                                }
                                ///// Сохранение
                                
                                //=====ЛИБО: workbook1.Combine(workbook2); И после этого уже проверки=====
                                Console.WriteLine("Введите путь для сохранения");
                                Console.WriteLine("Введите формат xml, чтобы сохранить в этом формате");
                                string PathToSave = Console.ReadLine();
                                workbook.Save(PathToSave);
                                /////
                            
                            break;
                        }
                    case 4:
                        {
                            Console.WriteLine("Введите путь к первому файлу");
                            string Path = Console.ReadLine();

                            string healed = XMLHealer.FixerReader(Path);//Таблетка для xml-я
                            XMLHealer.FixerWriter(healed, Path);
                            var workbook = new Workbook(Path);

                            Worksheet WSTo = workbook.Worksheets[0];

                            //Console.WriteLine("Введите путь ко второму файлу");
                           // string Path2 = Console.ReadLine();

                            //string healed2 = XMLHealer.FixerReader(Path2);//Таблетка для xml-я
                            //XMLHealer.FixerWriter(healed2, Path2);

                            //var workbook2 = new Workbook(Path2);
                            //XmlDocument xmlDoc = new XmlDocument();
                            //xmlDoc.Load(Path);
                            //XmlDocument xmlDoc2 = new XmlDocument();
                            //xmlDoc2.Load(Path2);



 
                            //workbook.Combine(workbook2);
                            //Console.WriteLine("Введите путь для сохранения объединенного файла xml");

                            //string Path3 = Console.ReadLine();
                            //workbook.Save(Path3);

                            Console.WriteLine("Сколько строк заголовков?");
                            int zagolovki = Convert.ToInt32(Console.ReadLine());

                            Console.WriteLine("В какой колонне содержится имя?");
                            string LastNameColumn = Console.ReadLine();//AW
                            int LastNameColumnINT = Transliterator.ColumnParser[LastNameColumn];//колонна lastname


                            //Проход по lastname для перевода
                            for (int i = zagolovki; i < WSTo.Cells.Rows.Count; i++)
                            {
                                var F = WSTo.Cells[i, LastNameColumnINT].Value;
                                string f = F.ToString();
                                foreach (var name in Transliterator.Urusification) 
                                {
                                    //WSTo.Cells[i, LastNameColumnINT].Value;
                                    f = Regex.Replace(f, name.Key, name.Value, RegexOptions.IgnoreCase);
                                }
                                WSTo.Cells[i, LastNameColumnINT].Value = f;
                                Console.WriteLine(WSTo.Cells[i, LastNameColumnINT].Value);
                            }
                            Console.WriteLine("Перевод успешно завершён");
                            LastNameColumnINT = LastNameColumnINT + 1;
                            //Проходка по соседней колонне, где инициалы, И ОНИ РАСПОЛОЖЕНЫ ОТДЕЛЬНО, КРУТОЙ XML ФАЙЛ СТАВЛЮ КЛАСС
                            for (int i = zagolovki; i < WSTo.Cells.Rows.Count; i++)
                            {
                                var F = WSTo.Cells[i, LastNameColumnINT].Value;
                                string f = F.ToString();
                                foreach (var name in Transliterator.Urusification)
                                {
                                    //WSTo.Cells[i, LastNameColumnINT].Value;
                                    f = Regex.Replace(f, name.Key, name.Value, RegexOptions.IgnoreCase);
                                }
                                WSTo.Cells[i, LastNameColumnINT].Value = f;
                                Console.WriteLine(WSTo.Cells[i, LastNameColumnINT].Value);
                            }
                            Console.WriteLine("Перевод успешно завершён");

                            //workbook.Worksheets[0].Cells.RemoveDuplicates();
                            //workbook.Save(Path3);

                            XmlDocument xmlDocument = new XmlDocument();
                            xmlDocument.Load(Console.ReadLine());

                            //XmlElement? xRoot = xmlDocument.DocumentElement; //получаем корневой элемент
                            //XmlNodeList elemList = xRoot.GetElementsByTagName(xRoot.InnerText);


                            //if (xRoot != null)
                            //{
                            //    // обход всех узлов в корневом элементе
                            //    foreach (XmlElement xnode in xRoot)
                            //    {
                            //        // получаем атрибут 
                            //        XmlNode? attr = xnode.Attributes.GetNamedItem(xRoot.Name);
                            //        Console.WriteLine(attr?.Value);
                            //        Console.WriteLine(xRoot.Name);
                            //        // обходим все дочерние узлы элемента 
                            //        foreach (XmlNode childnode in xnode.ChildNodes)
                            //        {
                            //            Console.Write("\t");
                            //            Console.WriteLine(childnode.Name);
                            //            if (childnode.HasChildNodes == true)
                            //            {
                            //                foreach (XmlNode grandchildnode in childnode.ChildNodes)
                            //                {
                            //                    Console.Write("\t\t");
                            //                    Console.WriteLine(grandchildnode.Name);
                            //                }
                            //            }
                            //        }
                            //    }

                            //}
                            xmlDocument.Save(Console.ReadLine());
                            break;
                        }
                    case 5:
                        {

                            break;
                        }
                }
            }
        }
    }
}