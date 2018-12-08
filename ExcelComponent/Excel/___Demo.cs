namespace ExcelComponent.Excel
{
    partial class Demo
    {
        protected void Test()
        {

            XlsDocument xls = new XlsDocument();
            xls.FileName = "D:\\12333333.xls";
            string sheetName = "chc 实例";
            Worksheet sheet = xls.Workbook.Worksheets.Add(sheetName);//填加名为"chc 实例"的sheet页
            Cells cells = sheet.Cells;//Cells实例是sheet页中单元格（cell）集合
            //单元格1-base
            Cell cell = cells.Add(1, 2, "抗");//设定第一行，第二例单元格的值
            cell.HorizontalAlignment = HorizontalAlignments.Centered;//设定文字居中
            cell.Font.FontName = "方正舒体";//设定字体
            cell.Font.Height = 20 * 20;//设定字大小（字体大小是以 1/20 point 为单位的）
            cell.UseBorder = true;//使用边框
            cell.BottomLineStyle = 2;//设定边框底线为粗线
            cell.BottomLineColor = Colors.DarkRed;//设定颜色为暗红

            //cell的格式还可以定义在一个xf对象中
            CellFormat cellXF = xls.NewXF();//为xls生成一个XF实例（XF是cell格式对象）
            cellXF.HorizontalAlignment = HorizontalAlignments.Centered;//设定文字居中
            cellXF.Font.FontName = "方正舒体";//设定字体
            cellXF.Font.Height = 20 * 20;//设定字大小（字体大小是以 1/20 point 为单位的）
            cellXF.UseBorder = true;//使用边框
            cellXF.BottomLineStyle = 2;//设定边框底线为粗线
            cellXF.BottomLineColor = Colors.DarkRed;//设定颜色为暗红

            cell = cells.Add(2, 2, "震", cellXF);//以设定好的格式填加cell
 
            cellXF.Font.FontName = "仿宋_GB2312";
            cell = cells.Add(3, 2, "救", cellXF);//格式可以多次使用

            ColumnInfo colInfo = new ColumnInfo(xls, sheet);//生成列格式对象
            //设定colInfo格式的起作用的列为第2列到第5列(列格式为0-base)
            colInfo.ColumnIndexStart = 1;//起始列为第二列
            colInfo.ColumnIndexEnd = 5;//终止列为第六列
            colInfo.Width = 15 * 256;//列的宽度计量单位为 1/256 字符宽
            sheet.AddColumnInfo(colInfo);//把格式附加到sheet页上（注：AddColumnInfo方法有点小问题，不给把colInfo对象多次附给sheet页）
            colInfo.ColumnIndexEnd = 6;//可以更改列对象的值
            ColumnInfo colInfo2 = new ColumnInfo(xls, sheet);//通过新生成一个列格式对象，才到能设定其它列宽度
            colInfo2.ColumnIndexStart = 7;
            colInfo2.ColumnIndexEnd = 8;
            colInfo2.Width = 1 * 256;
            sheet.AddColumnInfo(colInfo2);

            MergeArea meaA = new MergeArea(1, 2, 3, 4);//一个合并单元格实例(合并第一行、第三例 到 第二行、第四例)
            sheet.AddMergeArea(meaA);//填加合并单元格
            cellXF.VerticalAlignment = VerticalAlignments.Centered;
            cellXF.Font.Height = 48 * 20;
            cellXF.Font.Bold = true;
            cellXF.Pattern = 3;//设定单元格填充风格。如果设定为0，则是纯色填充
            cellXF.PatternBackgroundColor = Colors.DarkRed;//填充的底色
            cellXF.PatternColor = Colors.DarkGreen;//设定填充线条的颜色
            cell = cells.Add(1, 3, "灾", cellXF);
            xls.Save();
        }
    }
}
