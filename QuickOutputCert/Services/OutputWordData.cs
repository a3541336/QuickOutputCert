

using QuickOutputCert.Models;
using System.Collections.Generic;
using System.IO;
using TemplateEngine.Docx;

namespace QuickOutputCert.Services
{
    public static class OutputWordData
    {
        public static void OutputReportDetail(List<DataModel> dataModels)
        {
            List<IContentItem> fields = new List<IContentItem>();
            TableContent tableContent = null;
            RepeatContent repeatContent = new RepeatContent("NewPage");
            //單號設定
            fields.Add(new FieldContent(("Name"), "123456"));
            int index = 0; //偵測當頁的設備數量，每頁最多10台設備
            dataModels.ForEach((e) =>
            {
                if (index == 0)
                {
                    fields = new List<IContentItem>();
                    tableContent = new TableContent("Row1");
                }

                tableContent.AddRow(
                    new FieldContent("Index", e.Index),
                    new FieldContent("ProdName", e.ProdName),
                    new FieldContent("DeviceModel", e.DeviceModel),
                    new FieldContent("Count", e.Count),
                    new FieldContent("Place", e.Place),
                    new FieldContent("UseYear", e.UseYear),
                    new FieldContent("Evaluation", e.Evaluation),
                    new FieldContent("Content1", e.Manufacturer),
                    new FieldContent("Content2", e.DeviceNo)
                    );
                if (index == 9)
                {
                    repeatContent.AddItem(tableContent);
                }
                index = (index + 1) % 10;
            });
            while (index > 10)
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
            }
            var fillContent = new Content();
            fillContent.Fields.Add(new FieldContent(("Name"), "123456"));
            fillContent.Repeats.Add(repeatContent);
            File.Copy("Resources/核對單範本.docx", "機電核對單.docx");
            using (var outputDocument = new TemplateProcessor("機電核對單.docx").SetRemoveContentControls(true))
            {

                outputDocument.FillContent(fillContent);
                outputDocument.SaveChanges(); //儲存變更檔案
            }
        }
    }
}
