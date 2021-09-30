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
            //重要
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);//注册Nuget包System.Text.Encoding.CodePages中的编码到.NET Core

            DirectoryInfo di = new DirectoryInfo(@"D:\\excel2csv\xls_tmp\");
            var files = di.GetFiles("*.xlsx");

            foreach (var file in files)
            {
                new ExcelConvert().ConvertCsv(xlsPath + file.Name,csvPath,file.Name.Split('.')[0]);
            }
		}
    }
}
