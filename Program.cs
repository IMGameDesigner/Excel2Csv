using System;
using System.IO;
using System.Text;
using System.Threading;

namespace Excel2Csv
{
    class Program
    {
		static string xlsPath = @"D:\\excel2csv\xls_tmp\";
        static string csvPath = @"D:\\excel2csv\csv";

        static void Main(string[] args)
        {
            xlsPath = @ReadConfig.GetExcelPath();
            csvPath = @ReadConfig.GetCsvPath();
            //重要
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);//注册Nuget包System.Text.Encoding.CodePages中的编码到.NET Core

            DirectoryInfo di = new DirectoryInfo(xlsPath);
            var files = di.GetFiles("*.xlsx");

            foreach (var file in files)
            {
                new ExcelConvert().ConvertCsv(xlsPath + file.Name,csvPath,file.Name.Split('.')[0]);
            }
		}
    }
}
