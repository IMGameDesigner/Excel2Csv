#### 工具概要

* 简单实用的高效Excel转Csv工具，支持Unicode编码

#### 使用方法：

1. 双击.csproj打开本C#项目，等待NuGet安装完毕
2. 确保xlsx里面只有一个sheet
3. xlsx文件放到指定目录，如D:\\excel2csv\xls_tmp\
4. 新建D:\\excel2csv\csv目录（csv生成地址）
5. 运行，生成csv
6. 如果把 “D:\\。。。” 改成 “.\” ，绝对路径变为相对exe的路径。“..\”为上层目录


#### 备注

* 单个xls文件可以用wps手动转csv
* 本仓库用了NuGet的第三方库
* 疑问，csv用text查看：每行字段间用逗号隔开，海贼是怎么在一个字符串字段里放逗号分隔符的
* 打印出来的内容怪怪的
