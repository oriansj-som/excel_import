using OfficeOpenXml;
using sqlite_Example;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Import_Excel
{
    class import_excel
    {
        public SQLiteDatabase mydatabase;
        private ExcelWorkbook book;

        public void initialize(string Database, string existingFile)
        {
            if (!File.Exists(existingFile))
            {
                Console.WriteLine("Unable to find " + existingFile);
                Environment.Exit(3);
            }

            FileInfo f = new FileInfo(existingFile);

            if (!File.Exists(Database))
            {
                Console.WriteLine("Unable to find " + Database);
                Environment.Exit(4);
            }

            // Connect to Database
            mydatabase = new SQLiteDatabase(Database);

            book = new ExcelPackage(f).Workbook;
        }

        public void import(bool emptyfail, string tablename, bool over_ride)
        {
            int i = 0;
            string[] tables = null;
            if (over_ride) tables = tablename.Split(',');
            foreach (ExcelWorksheet sheet in book.Worksheets)
            {
                if (over_ride) import_table(sheet, emptyfail, tables[i]);
                else import_table(sheet, emptyfail, null);
                i = i + 1;
            }
        }

        public void cleanup()
        {
            mydatabase.SQLiteDatabase_Close();
        }

        private void import_table(ExcelWorksheet ws, bool emptyfail, string table_over)
        {
            string table_name = ws.Name;
            if (null != table_over) table_name = table_over;
            List<string> col_names = new List<string>();
            int col = 1;
            int row = 1;
            string s = "";

            // Collect titles
            do
            {
                s = read_cell(row, col, ws);
                if(!match("", s)) col_names.Add(s);
                col = col + 1;
            } while (!match("", s));

            // Create table
            setup_table(table_name, col_names);

            row = row + 1;
            do
            {
                col = 1;
                Dictionary<string, string> tmp = new Dictionary<string, string>();
                foreach (string key in col_names)
                {
                    s = read_cell(row, col, ws);
                    if (!match("", s)) tmp.Add(key, s);
                    col = col + 1;
                }
                if(0 < tmp.Count)
                {
                    bool result = mydatabase.Insert(table_name, tmp);
                    if (!result)
                    {
                        Console.WriteLine("Failed to insert:");
                        Console.WriteLine(string.Join(Environment.NewLine, tmp));
                        if (emptyfail) Environment.Exit(11);
                    }
                }
                row = row + 1;
            } while (!match("", read_cell(row, 1, ws)));

        }

        private void setup_table(string tablename, List<string> table_cols)
        {
            mydatabase.ExecuteNonQuery(string.Format("drop table if exists '{0}';", tablename));
            string sql = string.Format("CREATE TABLE '{0}' ( ", tablename);
            foreach (string s in table_cols)
            {
                if ('`' == sql[sql.Length - 1]) sql = sql + ", ";
                sql = sql + "`" + s + "`";
            }
            sql = sql + ");";

            mydatabase.ExecuteNonQuery(sql);
        }

        private string read_cell(int row, int col, ExcelWorksheet currentWorksheet)
        {
            try
            {
                string s = currentWorksheet.Cells[row, col].Text;
                return s;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Environment.Exit(5);
            }

            // Not possible
            Console.WriteLine("WTF read_cell null??");
            return null;
        }

        private bool match(string a, string b)
        {
            return a.Equals(b, StringComparison.CurrentCultureIgnoreCase);
        }
    }
}
