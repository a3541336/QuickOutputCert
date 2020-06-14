using QuickOutputCert.Models;
using System.Reflection;
using ExcelDataReader;
using System.IO;
using System.Text;
using System.Data;
using System.Collections.Generic;

namespace QuickOutputCert.Services
{
    public static class ExcelDataSort
    {
        public static List<DataModel> GetData(string appPath)
        {
            DataSet ds = new DataSet();
            List<DataModel> dataModel = new List<DataModel>();
            //取得DataModel屬性陣列
            PropertyInfo[] prop = new DataModel().GetType().GetProperties();
            var extension = Path.GetExtension(appPath).ToLower();
            using (var stream = new FileStream(appPath, FileMode.Open, FileAccess.Read))
            {
                IExcelDataReader reader = null;
                if (extension == ".xls")
                    reader = ExcelReaderFactory.CreateBinaryReader(stream, new ExcelReaderConfiguration() { FallbackEncoding = Encoding.GetEncoding("Big5") });
                else if (extension == ".xlsx")
                    reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                using(reader)
                {
                    ds = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        UseColumnDataType = false,
                        ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                        {
                            //設定讀取資料時是否忽略標題
                            FilterRow = rowReader => rowReader.Depth > 1
                        }
                    }) ;
                    //取得第一個欄位
                    var table = ds.Tables[0];
                    foreach(DataRow row in table.Rows)
                    {
                        int i = 0;
                        DataModel dm = new DataModel();
                        foreach(var data in row.ItemArray)
                        {
                            prop[i].SetValue(dm, data.ToString());
                            i++;
                        }
                        dataModel.Add(dm);
                    }
                }
            }
            return dataModel;
        }
    }
}
