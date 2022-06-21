using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;

public class ReadConfig
{
    static string Read(string path)
    {
        FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read);
        int len = (int)fs.Length;
        byte[] bytes = new byte[len];
        int readedLength = fs.Read(bytes, 0, len);
        Console.WriteLine("读取了行数："+readedLength);
        string str= Encoding.UTF8.GetString(bytes, 0, len);
        fs.Close();
        return str;
    }

    public static string GetCsvPath()
    {
        return Read(".\\Config\\CsvPath.txt");
    }
    public static string GetExcelPath()
    {
        return Read(".\\Config\\ExcelPath.txt");
    }
}
