using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using System.IO;
using System;
using System.Collections.Generic;

internal class Program
{
    private static void Main(string[] args)
    {
        // Для .xlsx файла
        IWorkbook workbook;
        using (FileStream file = new FileStream(@"E:\obs-studio\HMF_24_dataset_sch1.xlsx", FileMode.Open, FileAccess.Read))
        {
            workbook = new XSSFWorkbook(file);
        }

        int count = 0, count1 = 0, count2 = 0, count3 = 0, count4 = 0, count5 = 0, count6 = 0, count7 = 0, count8 = 0, count9 = 0,
            count11 = 0, count12 = 0, count13 = 0, count14 = 0, count15 = 0, count16 = 0, count17 = 0, count18 = 0, count19 = 0;
        ISheet sheet = workbook.GetSheetAt(0);
        IRow row = sheet.GetRow(2);
        ICell cell = row.GetCell(5);
        var value = cell.StringCellValue;

        ISheet sheetType = workbook.GetSheetAt(0);
        IRow rowType = sheet.GetRow(2);
        ICell cellType = row.GetCell(3);
        var valueType = cellType.StringCellValue;

        ISheet sheetSits = workbook.GetSheetAt(0);
        IRow rowSits = sheet.GetRow(2);
        ICell cellSits = row.GetCell(8);
        var valueSits = cellSits.StringCellValue;

        for (int i = 2; i < 11859; i++)
        {
            sheet = workbook.GetSheetAt(0);
            row = sheet.GetRow(i);
            cell = row.GetCell(5);
            value = cell.StringCellValue;

            sheetType = workbook.GetSheetAt(0);
            rowType = sheet.GetRow(i);
            cellType = row.GetCell(3);
            valueType = cellType.StringCellValue;

            sheetSits = workbook.GetSheetAt(0);
            rowSits = sheet.GetRow(i);
            cellSits = row.GetCell(8);
            valueSits = cellSits.StringCellValue;




            if (value == "район Богородское")
            {
                Console.WriteLine(Convert.ToString(value) + "-" + Convert.ToString(i + 1) + Convert.ToString(valueType) + "-" + Convert.ToString(valueSits));
                count++;
                if (valueType == "бар")
                {
                    count11 += Convert.ToInt32(valueSits);
                    count1++;
                }
                else if (valueType == "буфет")
                {
                    count12 += Convert.ToInt32(valueSits);
                    count2++;
                }
                else if (valueType == "закусочная")
                {
                    count13 += Convert.ToInt32(valueSits);
                    count3++;
                }
                else if (valueType == "кафе")
                {
                    count14 += Convert.ToInt32(valueSits);
                    count4++;
                }
                else if (valueType == "кафетерий")
                {
                    count15 += Convert.ToInt32(valueSits);
                    count5++;
                }
                else if (valueType == "магазин (отдел кулинарии)")
                {
                    count16 += Convert.ToInt32(valueSits);
                    count6++;
                }
                else if (valueType == "предприятие быстрого обслуживания")
                {
                    count17 += Convert.ToInt32(valueSits);
                    count7++;
                }
                else if (valueType == "ресторан")
                {
                    count18 += Convert.ToInt32(valueSits);
                    count8++;
                }
                else if (valueType == "столовая")
                {
                    count19 += Convert.ToInt32(valueSits);
                    count9++;
                }
            }


        }
        for (int i = 11859; i < 20857; i++)
        {
            sheet = workbook.GetSheetAt(0);
            row = sheet.GetRow(i);
            cell = row.GetCell(5);
            value = cell.StringCellValue;

            sheetType = workbook.GetSheetAt(0);
            rowType = sheet.GetRow(i);
            cellType = row.GetCell(3);
            valueType = cellType.StringCellValue;

            sheetSits = workbook.GetSheetAt(0);
            rowSits = sheet.GetRow(i);
            cellSits = row.GetCell(8);
            valueSits = cellSits.StringCellValue;


            if (value == "район Богородское")
            {
                count++;
                Console.WriteLine(Convert.ToString(value) + "-" + Convert.ToString(i + 1) + Convert.ToString(valueType) + "-" + Convert.ToString(valueSits));
                if (valueType == "бар")
                {
                    count11 += Convert.ToInt32(valueSits);
                    count1++;
                }
                else if (valueType == "буфет")
                {
                    count12 += Convert.ToInt32(valueSits);
                    count2++;
                }
                else if (valueType == "закусочная")
                {
                    count13 += Convert.ToInt32(valueSits);
                    count3++;
                }
                else if (valueType == "кафе")
                {
                    count14 += Convert.ToInt32(valueSits);
                    count4++;
                }
                else if (valueType == "кафетерий")
                {
                    count15 += Convert.ToInt32(valueSits);
                    count5++;
                }
                else if (valueType == "магазин (отдел кулинарии)")
                {
                    count16 += Convert.ToInt32(valueSits);
                    count6++;
                }
                else if (valueType == "предприятие быстрого обслуживания")
                {
                    count17 += Convert.ToInt32(valueSits);
                    count7++;
                }
                else if (valueType == "ресторан")
                {
                    count18 += Convert.ToInt32(valueSits);
                    count8++;
                }
                else if (valueType == "столовая")
                {
                    count19 += Convert.ToInt32(valueSits);
                    count9++;
                }
            }
        }

        Console.WriteLine("район Богородское                | " + Convert.ToString(count));
        Console.WriteLine("бар                              | " + Convert.ToString(count1) + " - " + Convert.ToString(count11));
        Console.WriteLine("буфет                            | " + Convert.ToString(count2) + " - " + Convert.ToString(count12));
        Console.WriteLine("закусочная                       | " + Convert.ToString(count3) + " - " + Convert.ToString(count13));
        Console.WriteLine("кафе                             | " + Convert.ToString(count4) + " - " + Convert.ToString(count14));
        Console.WriteLine("кафетерий                        | " + Convert.ToString(count5) + " - " + Convert.ToString(count15));
        Console.WriteLine("магазин (отдел кулинарии)        | " + Convert.ToString(count6) + " - " + Convert.ToString(count16));
        Console.WriteLine("предприятие быстрого обслуживания| " + Convert.ToString(count7) + " - " + Convert.ToString(count17));
        Console.WriteLine("ресторан                         | " + Convert.ToString(count8) + " - " + Convert.ToString(count18));
        Console.WriteLine("столовая                         | " + Convert.ToString(count9) + " - " + Convert.ToString(count19));



        Console.ReadLine();
    }
}