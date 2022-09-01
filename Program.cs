using System;

namespace Import_Excel
{
    class Program
    {
        static void Main(string[] args)
        {
            string tablename = null;
            string filename = null;
            string databasename = "example.db";

            int i = 0;
            while (i < args.Length)
            {
                if (match("--file", args[i]))
                {
                    filename = args[i + 1];
                    i = i + 2;
                }
                else if (match("--table", args[i]))
                {
                    tablename = args[i + 1];
                    i = i + 2;
                }
                else if (match("--database", args[i]))
                {
                    databasename = args[i + 1];
                    i = i + 2;
                }
                else if (match("--verbose", args[i]))
                {
                    int index = 0;
                    foreach (string s in args)
                    {
                        Console.WriteLine(string.Format("argument {0}: {1}", index, s));
                        index = index + 1;
                    }
                    i = i + 1;
                }
                else
                {
                    i = i + 1;
                }
            }

            import_excel e = new import_excel();
            e.initialize(databasename, filename);
            e.import(true, tablename, null != tablename);
            e.cleanup();

        }

        static public bool match(string a, string b)
        {
            return a.Equals(b, StringComparison.CurrentCultureIgnoreCase);
        }
    }
}
