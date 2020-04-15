using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SimpleCSVFormatChecker
{
    public class Program
    {
        static void Main(string[] args)
        {
            new Program().Check(args[0]);
        }

        private const string UTF_8_BOM = "EFBBBF";

        private void writeLine(string str, ConsoleColor color)
        {
            var originalColor = Console.ForegroundColor;
            Console.ForegroundColor = color;
            Console.WriteLine(str);
            Console.ForegroundColor = originalColor;
        }

        public void Check(string path)
        {
            var originalColor = Console.ForegroundColor;
            var list = new List<string>();
            var files = Directory.GetFiles(path, "*", SearchOption.AllDirectories);
            foreach (var file in files)
            {
                if (!isFileUnicode(file))
                {
                    writeLine($"请检查csv:{file}的格式", ConsoleColor.Red);
                    list.Add(Path.GetFileName(file));
                }
                else
                {
                    writeLine($"{file}", ConsoleColor.White);
                }
            }

            if (list.Count > 0)
            {
                writeLine("请检查以下非UTF8-BOM格式的CSV", ConsoleColor.Red);
                var phrase = list.Aggregate((partialPhrase, word) => $"{partialPhrase} {word}");
                writeLine(phrase, ConsoleColor.Red);
            }

            Console.ForegroundColor = originalColor;
        }

        private bool isFileUnicode(string filename)
        {
            bool ret = false;
            try
            {
                byte[] buff = new byte[3];
                using (FileStream fs = new FileStream(filename, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    fs.Read(buff, 0, 3);
                }

                string hx = "";
                foreach (byte letter in buff)
                {
                    hx += string.Format("{0:X2}", Convert.ToInt32(letter));
                    //Checking to see the first bytes matches with any of the defined Unicode BOM
                    //We only check for UTF8 and UTF16 here.
                    ret = UTF_8_BOM.Equals(hx);
                    if (ret)
                    {
                        break;
                    }
                }
            }
            catch (IOException)
            {
                //ignore any exception
            }
            return ret;
        }
    }
}
