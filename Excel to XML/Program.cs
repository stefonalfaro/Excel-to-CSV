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
                    string fileLocation = f;
                    string fileName = Path.GetFileNameWithoutExtension(fileLocation);
                    string NewName = outputFolderLocation+"\\"+ fileName;

                    DataSet result = ExcelToDataSet(fileLocation);
                    DataSetToCSV(result.Tables[0], NewName);

                    System.Threading.Thread.Sleep(1000);
                    File.Delete(fileLocation);

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

            //headers 
            /*
            for (int i = 0; i < dtDataTable.Columns.Count; i++)
            {
                sw.Write(dtDataTable.Columns[i]);
                if (i < dtDataTable.Columns.Count - 1)
                {
                    sw.Write(",");
                }
            }
            sw.Write(sw.NewLine);*/

            int count = 1;
            foreach (DataRow dr in dtDataTable.Rows)
            {
                if (count != 1)
                {
                    StreamWriter sw = new StreamWriter(strFilePath + "_" + count + ".csv", false);
                    for (int i = 0; i < dtDataTable.Columns.Count; i++)
                    {
                        sw.Write(dtDataTable.Columns[i]);
                        if (i < dtDataTable.Columns.Count - 1)
                        {
                            sw.Write(",");
                        }
                    }
                    sw.Write(sw.NewLine);
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
                    sw.Close();
                }
                count++;
            }          
        }
    }


}
