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
    static class XMLHealer
    {
        public static string CleanInvalidXmlChars(this string StrInput)//Метод уничтожения недопустимых для чтения 16-ричных символов
        {

            if (string.IsNullOrWhiteSpace(StrInput))
            {
                return StrInput;
            }

            string RegularExp = @"[^\x09\x0A\x0D\x20-\xD7FF\xE000-\xFFFD\x10000-x10FFFF]";
            return Regex.Replace(StrInput, RegularExp, String.Empty);
        }
        public static string FixerReader(string File)
        {
            string Stroka = "";
            try
            {

                StreamReader Reader = new StreamReader(File);
                Stroka = Reader.ReadLine();
                Reader.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return Stroka;
        }
        public static void FixerWriter(string Stroka, string NewFile)
        {

            try
            {
                StreamWriter Writer = new StreamWriter(NewFile);
                Writer.Write(Stroka);
                Writer.Close();
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
        }
    }
}
