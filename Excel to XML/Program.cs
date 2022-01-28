using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Data;
using ExcelDataReader;

namespace Excel_to_XML
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputFolderLocation = args[0];
                string outputFolderLocation = args[1];

                string[] files = Directory.GetFiles(inputFolderLocation);

                foreach (string f in files)
                {
                    string fileName = f;
                    string NewName = Path.GetFileNameWithoutExtension(fileName) + ".csv";

                    DataSet result;

                    result = ExcelToDataSet(fileName);
                    DataSetToCSV(result.Tables[0], NewName);

                    //Console.Read();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.Read();
            }           
        }

        public static DataSet ExcelToDataSet(string ExcelFileInput)
        {
            using (var stream = File.Open(ExcelFileInput, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    return reader.AsDataSet();
                }
            }
        }

        public static void DataSetToCSV(System.Data.DataTable dtDataTable, string strFilePath)
        {
            StreamWriter sw = new StreamWriter(strFilePath, false);
            //headers    
            for (int i = 0; i < dtDataTable.Columns.Count; i++)
            {
                sw.Write(dtDataTable.Columns[i]);
                if (i < dtDataTable.Columns.Count - 1)
                {
                    sw.Write(",");
                }
            }
            sw.Write(sw.NewLine);
            foreach (DataRow dr in dtDataTable.Rows)
            {
                for (int i = 0; i < dtDataTable.Columns.Count; i++)
                {
                    if (!Convert.IsDBNull(dr[i]))
                    {
                        string value = dr[i].ToString();
                        if (value.Contains(','))
                        {
                            value = String.Format("\"{0}\"", value);
                            sw.Write(value);
                        }
                        else
                        {
                            sw.Write(dr[i].ToString());
                        }
                    }
                    if (i < dtDataTable.Columns.Count - 1)
                    {
                        sw.Write(",");
                    }
                }
                sw.Write(sw.NewLine);
            }
            sw.Close();
        }
    }


}
