

using QuickOutputCert.Models;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using TemplateEngine.Docx;

namespace QuickOutputCert.Services
{
    public static class OutputWordData
    {
        public static void OutputReportDetail(this string fileName, List<DataModel> dataModels)
        {
            List<IContentItem> fields = new List<IContentItem>();
            TableContent tableContent = null;
            RepeatContent repeatContent = new RepeatContent("NewPage");
            //單號設定
            fields.Add(new FieldContent(("Name"), "123456"));
            int index = 0; //偵測當頁的設備數量，每頁最多10台設備
            string SealNoPrint = null;
            string LastSealNo = null;
            dataModels.ForEach((e) =>
            {
                #region 解析封識號並存入SealNoPrint內
                /*透過LastSealNo來存取前一個遇到的封識號，在做判斷是否為連續，若是連續則將連續號省略
                 *當不連續時(int.Parse(e.SealNo) - int.Parse(LastSealNo) != 1)，使用"-"符號將連續將最後一個連續封識號放入，並在增加","將目前不連續號放入
                 */
                if (SealNoPrint == null)
                {
                    SealNoPrint = e.SealNo;
                }
                else if (int.Parse(e.SealNo) - int.Parse(LastSealNo) != 1 | int.Parse(e.SealNo) == int.Parse(LastSealNo))
                {
                    SealNoPrint += $" - {LastSealNo}";
                    SealNoPrint += $", {e.SealNo}";
                }
                LastSealNo = e.SealNo;
                if (index == 0)
                {
                    fields = new List<IContentItem>();
                    tableContent = new TableContent("Row1");
                }
                #endregion
                tableContent.AddRow(
                    new FieldContent("Index", e.Index),
                    new FieldContent("ProdName", e.ProdName),
                    new FieldContent("DeviceModel", e.DeviceModel),
                    new FieldContent("Count", e.Count),
                    new FieldContent("Place", e.Place),
                    new FieldContent("ProdDate", e.ProdDate),
                    new FieldContent("Evaluation", e.Evaluation),
                    new FieldContent("Content1", $"({e.Manufacturer})"),
                    new FieldContent("Content2", $"({e.DeviceNo})")
                    );
                if (index == 9)
                {
                    repeatContent.AddItem(tableContent);
                }
                index = (index + 1) % 10;
            });
            //當設備最後一個時，將最後的封識號填上
            SealNoPrint += $" - {LastSealNo}";

            //當資料不足10筆時，用空白欄位補滿
            while (index < 10 && index != 0)
            {
                tableContent.AddRow(
                    new FieldContent("Index", ""),
                    new FieldContent("ProdName", ""),
                    new FieldContent("DeviceModel", ""),
                    new FieldContent("Count", ""),
                    new FieldContent("Place", ""),
                    new FieldContent("UseYear", ""),
                    new FieldContent("Evaluation", ""),
                    new FieldContent("Content1", ""),
                    new FieldContent("Content2", "")
                    );
                index++;
                if (index == 9)
                {
                    repeatContent.AddItem(tableContent);
                }
            }
            string[] arrStr = fileName.Split(' ');
            var fillContent = new Content();
            fillContent.Fields.Add(new FieldContent(("Name"), arrStr[0]));
            fillContent.Fields.Add(new FieldContent(("Port"), arrStr[1]));
            fillContent.Fields.Add(new FieldContent(("SealNo"), SealNoPrint));
            fillContent.Repeats.Add(repeatContent);
            File.Copy("Resources/核對單範本.docx", "機電核對單.docx", true);
            using (var outputDocument = new TemplateProcessor("機電核對單.docx").SetRemoveContentControls(true))
            {

                outputDocument.FillContent(fillContent);
                outputDocument.SaveChanges(); //儲存變更檔案
            }
            Process.Start("機電核對單.docx");
        }
    }
}
