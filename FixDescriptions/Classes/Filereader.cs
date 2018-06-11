using System;
using System.IO;

public class Filereader
{
    public string[] lowercase { get; set; }
    public string[] uppercase { get; set; }

    public Filereader()
    {
        LoadWords();
    }

    private void LoadWords()
    {
        try
        {
            lowercase = File.ReadAllLines(@"lowercasewords.list");
            uppercase = File.ReadAllLines(@"uppercasewords.list");

            for (int i = 0; i < lowercase.Length; i++)
            {
                lowercase[i] = FixDescriptions.Program.Capitalize(lowercase[i].ToLower());
            }

            for (int i = 0; i < lowercase.Length; i++)
            {
                uppercase[i] = FixDescriptions.Program.Capitalize(uppercase[i].ToLower());
            }

        }
        catch (Exception e)
        {
            File.Create(@"lowercasewords.list");
            File.Create(@"uppercasewords.list");
        }
    }
}
