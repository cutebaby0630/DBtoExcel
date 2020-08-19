using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using SqlServerHelper.Core;
using SqlServerHelper;
using System.ComponentModel;
using LicenseContext = OfficeOpenXml.LicenseContext;
using System.Reflection;
using System.ComponentModel.DataAnnotations;

namespace DBtoExcel
{
    public static class ExcelExtensions
    {
        // SetQuickStyle，指定前景色/背景色/水平對齊
        public static void SetQuickStyle(this ExcelRange range,
            Color fontColor,
            Color bgColor = default(Color),
            ExcelHorizontalAlignment hAlign = ExcelHorizontalAlignment.Left)
        {
            range.Style.Font.Color.SetColor(fontColor);
            if (bgColor != default(Color))
            {
                range.Style.Fill.PatternType = ExcelFillStyle.Solid; // 一定要加這行..不然會報錯
                range.Style.Fill.BackgroundColor.SetColor(bgColor);
            }
            range.Style.HorizontalAlignment = hAlign;
        }

        //讓文字上有連結
        public static void SetHyperlink(this ExcelRange range, Uri uri)
        {
            range.Hyperlink = uri;
            range.Style.Font.UnderLine = true;
            range.Style.Font.Color.SetColor(Color.Blue);
        }
    }
    class Program
    {
        static void Main(string[] args)
        {
            string sql = @"
                            SELECT 'BLOMBlood' SourceTable,a.ChargeItemId,a.ChargeCode,a.ItemChineseName, a.CreateTime,a.ModifyTime ,b.BloodId ItemId,b.BloodCode ItemCode,b.BloodChineseName ItemName
                            FROM dbo.CHGMChargeItem a FULL OUTER JOIN dbo.BLOMBlood b ON a.ChargeItemId = b.BloodId 
                            WHERE a.OrderTypeCode = 'BLO' AND (a.ChargeItemId IS NULL OR b.BloodId IS NULL OR a.ChargeCode <> b.BloodCode)
                            UNION all
                            SELECT 'EXAMExamination' SourceTable,a.ChargeItemId,a.ChargeCode,a.ItemChineseName, a.CreateTime,a.ModifyTime ,b.ExaminationId ItemId,b.ExaminationCode ItemCode,b.ExamChineseName ItemName
                            FROM dbo.CHGMChargeItem a FULL OUTER JOIN dbo.EXAMExamination b ON a.ChargeItemId = b.ExaminationId 
                            WHERE a.OrderTypeCode = 'EXA' AND (a.ChargeItemId IS NULL OR b.ExaminationId IS NULL OR a.ChargeCode <> b.ExaminationCode)
                            UNION all
                            SELECT 'LABMLaboratory' SourceTable,a.ChargeItemId,a.ChargeCode,a.ItemChineseName, a.CreateTime,a.ModifyTime ,b.LaboratoryId ItemId,b.LaboratoryCode ItemCode,b.LaboratoryChineseName ItemName
                            FROM dbo.CHGMChargeItem a FULL OUTER JOIN dbo.LABMLaboratory b ON a.ChargeItemId = b.LaboratoryId 
                            WHERE a.OrderTypeCode = 'LAB' AND (a.ChargeItemId IS NULL OR b.LaboratoryId IS NULL OR a.ChargeCode <> b.LaboratoryCode)
                            UNION all
                            SELECT 'OPRMOperation' SourceTable,a.ChargeItemId,a.ChargeCode,a.ItemChineseName, a.CreateTime,a.ModifyTime ,b.OperationId ItemId,b.OperationCode ItemCode,b.OprChineseName ItemName
                            FROM dbo.CHGMChargeItem a FULL OUTER JOIN dbo.OPRMOperation b ON a.ChargeItemId = b.OperationId 
                            WHERE a.OrderTypeCode = 'OPR' AND (a.ChargeItemId IS NULL OR b.OperationId IS NULL OR a.ChargeCode <> b.OperationCode)
                            UNION all
                            SELECT 'PHRMMedication' SourceTable,a.ChargeItemId,a.ChargeCode,a.ItemChineseName, a.CreateTime,a.ModifyTime,b.MedicationId ItemId,b.MedicationCode ItemCode,b.GenericName ItemName
                            FROM dbo.CHGMChargeItem a FULL OUTER JOIN dbo.PHRMMedication b ON a.ChargeItemId = b.MedicationId 
                            WHERE a.OrderTypeCode = 'PHR' AND (a.ChargeItemId IS NULL OR b.MedicationId IS NULL OR a.ChargeCode <> b.MedicationCode)
                            UNION all
                            SELECT 'PEXMHealthCheckItem' SourceTable,a.ChargeItemId,a.ChargeCode,a.ItemChineseName, a.CreateTime,a.ModifyTime ,b.HealthCheckItemId  ItemId,b.HealthCheckItemCode ItemCode,b.ItemChineseName ItemName
                            FROM dbo.CHGMChargeItem a FULL OUTER JOIN dbo.PEXMHealthCheckItem b ON a.ChargeItemId = b.HealthCheckItemId 
                            WHERE a.OrderTypeCode = 'PEX' AND (a.ChargeItemId IS NULL OR b.HealthCheckItemId IS NULL OR a.ChargeCode <> b.HealthCheckItemCode)
                            UNION all
                            SELECT 'TREMTreatment' SourceTable,a.ChargeItemId,a.ChargeCode,a.ItemChineseName, a.CreateTime ,a.ModifyTime,b.TreatmentId ItemId,b.TreatmentCode ItemCode,b.TreatmentChineseName ItemName
                            FROM dbo.CHGMChargeItem a FULL OUTER JOIN dbo.TREMTreatment b ON a.ChargeItemId = b.TreatmentId 
                            WHERE a.OrderTypeCode = 'TRE' AND (a.ChargeItemId IS NULL OR b.TreatmentId IS NULL OR a.ChargeCode <> b.TreatmentCode)
                            ";
            IConfiguration config = new ConfigurationBuilder().AddJsonFile("appsetting.json", optional: true, reloadOnChange: true).Build();
            //取得連線字串
            string connString = config.GetConnectionString("DefaultConnection");
            //string connString = "Data Source=10.1.222.181;Initial Catalog={0};Integrated Security=False;User ID={1};Password={2};Pooling=True;MultipleActiveResultSets=True;Connect Timeout=120;Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite";
            SqlServerDBHelper sqlHelper = new SqlServerDBHelper(string.Format(connString, "HISDB", "msdba", "1qaz@wsx"));
            List<DBdata> migrationTableInfoList = sqlHelper.QueryAsync<DBdata>(sql).Result?.ToList();
            DataTable dt = sqlHelper.FillTableAsync(sql).Result;

            var excelname = new FileInfo(DateTime.Now.ToString("yyyyMMddhhmm") + ".xlsx");
            //ExcelPackage.LicenseContext = LicenseContext.Commercial;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var excel = new ExcelPackage(excelname))
            {
                excel.Workbook.Worksheets.Add("結果");
                ExcelWorksheet firstsheet = excel.Workbook.Worksheets[0];
                int rowIndex = 1;
                int colIndex = 1;
                //4.3.1塞資料到某一格
                firstsheet.Cells[rowIndex, colIndex++].Value = "SourceTable";
                firstsheet.Cells[rowIndex, colIndex++].Value = "ChargeItemId";
                firstsheet.Cells[rowIndex, colIndex++].Value = "ChargeCode";
                firstsheet.Cells[rowIndex, colIndex++].Value = "ItemChineseName";
                firstsheet.Cells[rowIndex, colIndex++].Value = "CreateTime";
                firstsheet.Cells[rowIndex, colIndex++].Value = "ModifyTime";
                firstsheet.Cells[rowIndex, colIndex++].Value = "ItemId";
                firstsheet.Cells[rowIndex, colIndex++].Value = "ItemCode";
                firstsheet.Cells[rowIndex, colIndex++].Value = "ItemName";
                //4.3.2 Cell Style
                firstsheet.Cells[rowIndex, 1, rowIndex, colIndex - 1]
                 .SetQuickStyle(Color.Black, Color.LightPink, ExcelHorizontalAlignment.Center);

                foreach (var v in migrationTableInfoList)
                {
                    rowIndex++;
                    colIndex = 1;
                    firstsheet.Cells[rowIndex, colIndex++].Value = v.SourceTable;
                    firstsheet.Cells[rowIndex, colIndex++].Value = v.ChargeItemId;
                    firstsheet.Cells[rowIndex, colIndex++].Value = v.ChargeCode;
                    firstsheet.Cells[rowIndex, colIndex++].Value = v.ItemChineseName;
                    firstsheet.Cells[rowIndex, colIndex].Value = v.CreateTime;
                    firstsheet.Cells[rowIndex, colIndex].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    firstsheet.Cells[rowIndex, colIndex++].Style.Numberformat.Format = "yyyy/MM/dd HH:mm:ss";
                    firstsheet.Cells[rowIndex, colIndex].Value = v.ModifyTime;
                    firstsheet.Cells[rowIndex, colIndex].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                    firstsheet.Cells[rowIndex, colIndex++].Style.Numberformat.Format = "yyyy/MM/dd HH:mm:ss";
                    firstsheet.Cells[rowIndex, colIndex++].Value = v.ItemId;
                    firstsheet.Cells[rowIndex, colIndex++].Value = v.ItemCode;
                    firstsheet.Cells[rowIndex, colIndex++].Value = v.ItemName;

                }
                //4.3.3 儲存格和字數相等
                int startColumn = firstsheet.Dimension.Start.Column;
                int endColumn = firstsheet.Dimension.End.Column;
                for (int count = startColumn; count <= endColumn; count++)
                {
                    firstsheet.Column(count).AutoFit();
                }
                Byte[] bin = excel.GetAsByteArray();
                File.WriteAllBytes(@"C:\Users\v-vyin\SchedulerDB_ExcelFile\" + excelname, bin);

            }

            var helper = new SMTPHelper("lovemath0630@gmail.com", "koormyktfbbacpmj", "smtp.gmail.com", 587, true, true); //寄出信email
            string subject = $"有關收標與子檔無法對應項目"; //信件主旨
            string body = $"Hi All, \r\n\r\n無法對應的項目如附件，\r\n\r\n再麻煩查收，感謝\r\n\r\n Best Regards, \r\n\r\n Vicky Yin";//信件內容
            string attachments = null;//附件
            var fileName = $@"C:\Users\v-vyin\SchedulerDB_ExcelFile\{excelname}";//附件位置
            if (File.Exists(fileName.ToString()))
            {
                attachments = fileName.ToString();
            }
            string toMailList = "Leon.Yen@microsoft.com;Sol.Lee@microsoft.com";//收件者
            string ccMailList = "NickyXu@dualred.onmicrosoft.com;v-vyin@microsoft.com";//CC收件者

            helper.SendMail(toMailList, ccMailList, null, subject, body, attachments);
        }
        public class DBdata
        {
            public string SourceTable { get; set; }
            public string ChargeItemId { get; set; }
            public string ChargeCode { get; set; }
            public string ItemChineseName { get; set; }
            public DateTime CreateTime { get; set; }
            public DateTime ModifyTime { get; set; }
            public string ItemId { get; set; }
            public string ItemCode { get; set; }
            public string ItemName { get; set; }
        }

    }
}
