using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel;

namespace xlsxToCsv
{
    public class ExcelConvert
    {
        DataSet result = new DataSet();

        public bool convert(string src, string tar)
        {
            getExcelData(src);
            return converToCSV(tar);
        }

        private void getExcelData(string file)
        {
            if (file.EndsWith(".xlsx"))
            {
                // Reading from a binary Excel file (format; *.xlsx)
                FileStream stream = File.Open(file, FileMode.Open, FileAccess.Read);
                IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                result = excelReader.AsDataSet();
                excelReader.Close();
            }

            if (file.EndsWith(".xls"))
            {
                // Reading from a binary Excel file ('97-2003 format; *.xls)
                FileStream stream = File.Open(file, FileMode.Open, FileAccess.Read);
                IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
                result = excelReader.AsDataSet();
                excelReader.Close();
            }

            List<string> items = new List<string>();
            for (int i = 0; i < result.Tables.Count; i++)
                items.Add(result.Tables[i].TableName.ToString());
        }

        private bool converToCSV(string toFilePath)
        {
            int index = 0;
            // sheets in excel file becomes tables in dataset
            // result.Tables[0].TableName.ToString(); // to get sheet name (table name)

            string a = "";
            int row_no = 0;
            if(result.Tables.Count == 0)
            {
                return false;
            }
            while (row_no < result.Tables[index].Rows.Count)
            {
                for (int i = 0; i < result.Tables[index].Columns.Count; i++)
                {
                    a += result.Tables[index].Rows[row_no][i].ToString() + ",";
                }
                row_no++;
                a += "\n";
            }
            string output = toFilePath;
            StreamWriter csv = new StreamWriter(@output, false);
            csv.Write(a);
            csv.Close();

            return true;
        }
    }
}
