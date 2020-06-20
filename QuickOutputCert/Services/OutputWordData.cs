

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
            /*
             * tableContent1為針對核對單明細表使用
             * tableContent2為針對使用現場檢驗紀錄使用
             */
            TableContent tableContent1 = null;
            TableContent tableContent2 = new TableContent("Row1");
            RepeatContent repeatContent = new RepeatContent("NewPage");
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
                #endregion
                #region 針對核對單明細表的商業邏輯
                if (index == 0)
                {
                    tableContent1 = new TableContent("Row1");
                }
                tableContent1.AddRow(
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
                    repeatContent.AddItem(tableContent1);
                }
                #endregion
                #region 針對現場檢驗紀錄表的商業邏輯
                tableContent2.AddRow(
                    new FieldContent("Index", e.Index),
                    new FieldContent("Check1", e.Check1),
                    new FieldContent("Check2", e.Check2),
                    new FieldContent("ProdName", e.ProdName),
                    new FieldContent("Manufacturer", e.Manufacturer),
                    new FieldContent("DeviceModel", e.DeviceModel),
                    new FieldContent("DeviceNo", e.DeviceNo),
                    new FieldContent("ProdDate", e.ProdDate),
                    new FieldContent("Place", e.Place),
                    new FieldContent("DeviceCheck1", e.DeviceCheck1),
                    new FieldContent("DeviceCheck2", e.DeviceCheck2),
                    new FieldContent("DeviceCheck3", e.DeviceCheck3),
                    new FieldContent("DeviceCheck4", e.DeviceCheck4),
                    new FieldContent("DeviceCheck5", e.DeviceCheck5),
                    new FieldContent("DeviceCheck6", e.DeviceCheck6),
                    new FieldContent("DeviceCheck7", e.DeviceCheck7),
                    new FieldContent("DeviceCheck8", e.DeviceCheck8),
                    new FieldContent("DeviceCheck9", e.DeviceCheck9),
                    new FieldContent("DeviceCheck10", e.DeviceCheck10),
                    new FieldContent("DeviceCheck11", e.DeviceCheck11),
                    new FieldContent("DeviceCheck12", e.DeviceCheck12),
                    new FieldContent("Evaluation", e.Evaluation),
                    new FieldContent("Remark", e.Remark),
                    new FieldContent("SealNo", e.SealNo)
                    );
                #endregion
                index = (index + 1) % 10;
            });
            //當設備最後一個時，將最後的封識號填上
            SealNoPrint += $" - {LastSealNo}";

            //當資料不足10筆時，用空白欄位補滿
            while (index < 10 && index != 0)
            {
                tableContent1.AddRow(
                    new FieldContent("Index", ""),
                    new FieldContent("ProdName", ""),
                    new FieldContent("DeviceModel", ""),
                    new FieldContent("Count", ""),
                    new FieldContent("Place", ""),
                    new FieldContent("ProdDate", ""),
                    new FieldContent("Evaluation", ""),
                    new FieldContent("Content1", ""),
                    new FieldContent("Content2", "")
                    );
                tableContent2.AddRow(
                    new FieldContent("Index", ""),
                    new FieldContent("Check1", ""),
                    new FieldContent("Check2", ""),
                    new FieldContent("ProdName", ""),
                    new FieldContent("Manufacturer", ""),
                    new FieldContent("DeviceModel", ""),
                    new FieldContent("DeviceNo", ""),
                    new FieldContent("ProdDate", ""),
                    new FieldContent("Place", ""),
                    new FieldContent("DeviceCheck1", ""),
                    new FieldContent("DeviceCheck2", ""),
                    new FieldContent("DeviceCheck3", ""),
                    new FieldContent("DeviceCheck4", ""),
                    new FieldContent("DeviceCheck5", ""),
                    new FieldContent("DeviceCheck6", ""),
                    new FieldContent("DeviceCheck7", ""),
                    new FieldContent("DeviceCheck8", ""),
                    new FieldContent("DeviceCheck9", ""),
                    new FieldContent("DeviceCheck10", ""),
                    new FieldContent("DeviceCheck11", ""),
                    new FieldContent("DeviceCheck12", ""),
                    new FieldContent("Evaluation", ""),
                    new FieldContent("Remark", ""),
                    new FieldContent("SealNo", "")
                    );
                index++;
                if (index == 9)
                {
                    repeatContent.AddItem(tableContent1);
                }
            }
            /*
             * 檔案命名規則 證書編號 進口口岸 檢驗員
             */
            string[] arrStr = fileName.Split(' ');
            #region 核對單明細表資料匯入及檔案呼叫
            var fillContent = new Content();
            fillContent.Fields.Add(new FieldContent(("Name"), arrStr[0]));
            if (arrStr.Length > 2)
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
            #endregion
            #region 現場檢驗紀錄表資料匯入及檔案呼叫
            var fillContent2 = new Content();
            fillContent2.Fields.Add(new FieldContent(("Name"), arrStr[0]));
            if (arrStr.Length > 3)
                fillContent2.Fields.Add(new FieldContent(("Inspecter"), arrStr[2]));
            fillContent2.Tables.Add(tableContent2);
            File.Copy("Resources/進口舊機電設備裝運前現場檢驗記錄新版10台.docx", "進口舊機電設備裝運前現場檢驗記錄新版10台.docx", true);
            using (var outputDocument = new TemplateProcessor("進口舊機電設備裝運前現場檢驗記錄新版10台.docx").SetRemoveContentControls(true))
            {

                outputDocument.FillContent(fillContent2);
                outputDocument.SaveChanges(); //儲存變更檔案
            }
            Process.Start("進口舊機電設備裝運前現場檢驗記錄新版10台.docx");
            #endregion
        }
    }
}
