using System;
using System.Linq;
using System.Diagnostics;

namespace FixDescriptions
{
    class Program
    {
        
        Filereader filereader;

        public Program()
        {
            filereader = new Filereader();
            string inPath = AppDomain.CurrentDomain.BaseDirectory + "in.xlsx";
            Console.WriteLine(inPath);
            Excel excel = new Excel(inPath);

            try
            {
                for (int i = 0; i < excel.getRows() && i < 10; i++)
                {
                    excel.WriteCell(i, 0, FixRow(excel.ReadCell(i, 0)));
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

        
            Console.WriteLine();
            Console.WriteLine("Finished! Used time: " + (DateTime.UtcNow - Process.GetCurrentProcess().StartTime.ToUniversalTime()) + "!");

            excel.Save();
            excel.Close();

        }


        static void Main()
        {
            new Program();

        }

        public static string Capitalize(string s)
        {
            if (s.Length != 0)
            {                
                s = s.ToLower();
                char[] a = s.ToCharArray();

                if (a.Contains('('))
                {
                    int index = Array.IndexOf(a, '(') + 1;
                    if (index < a.Length)
                        a[index] = char.ToUpper(a[index]);
                }

                a[0] = char.ToUpper(a[0]);
                return new string(a);
            }
            return "";
        }

        bool IsTag(string s)
        {
            if (s.Any(c => char.IsDigit(c)))
            {
                return true;
            }

            if (s.Contains('-'))
            {
                return true;
            }

            return false;
        }

        string FixRow(String s)
        {

            /*
             * Fiks 16 char bug 
             * Gjer alle ord capitalized
             * Gjer alle tag UPPERCASE
             * Gjer alle uppercase ord UPPERCASE
             * Gjer alle lowercase ord lowercase
             * 
             * UNIMPLEMENTED: Bytt ut ord.
             */
                     
            Console.WriteLine(s);

            //Fjern 16 bit ord fiksen
            s = s.Replace("- ", "");

            //Capitalize Alle Orda og Uppercase TAGS
            string[] words = s.Split(' ');
            for (int i = 0; i < words.Length; i++)
            {
                words[i] = FixWord(words[i]);

                //Uppercase exceptions
                foreach (string uppercase in filereader.uppercase)
                {
                    string fixedWord = words[i].Replace("(", " ").Replace(")", " ").Trim();
                    if (fixedWord.Equals(uppercase))                        
                    {
                        words[i] = words[i].Replace(uppercase, uppercase.ToUpper());
                    }
                }

                //lowercase exception
                foreach (string lowercase in filereader.lowercase)
                {
                    string fixedWord = words[i].Replace("(", " ").Replace(")", " ").Trim();
                    if (fixedWord.Equals(lowercase))
                    {
                        words[i] = words[i].Replace(lowercase, lowercase.ToLower());
                    }
                }


            }

            s = string.Join(" ", words);

            Console.WriteLine(s);
            return s;
        }

        string FixWord(string s)
        {
            if (IsTag(s))
            {
                return s.ToUpper();
            }
            else
            {
                return Capitalize(s);
            }

        }

    }
}
