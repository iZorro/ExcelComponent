/*


using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
 

namespace com.Ole.excel
{
    /// <summary>
    /// 数据简单绑定
    /// </summary>
    public class DataBandExcel
    {
        /// <summary>
        /// 数据绑定
        /// </summary>
        /// <param name="title"></param>
        /// <param name="t"></param>
        public void Band<T>(Dictionary<string, string> title, IList<T> t,string filePath)
        {
            var entity = t[0];
            var tableName = entity.GetType().Name;
            XlsDocument xls = new XlsDocument();
            Worksheet sheet = xls.Workbook.Worksheets.Add(tableName);
            Cells cells = sheet.Cells;
            int key_index = 1; //列索引
            int line_index = 1; // 行数
            #region 设置标题
            foreach (var item in title.Keys)
            {
                Cell cell = cells.Add(line_index, key_index, item);
                key_index++;
            }
            #endregion

            line_index++;
            var _keyNames = entity.GetType().GetProperties();
            foreach (var souce in t)
            {
                key_index = 1;
                foreach (System.Reflection.PropertyInfo _key_ in _keyNames)
                {
                    if (title.Values.Contains(_key_.Name))
                    {
                        var xls_value = _key_.GetValue(souce, null); 
                        var code = XY.TypeToCode(_key_.PropertyType);

                        #region 不同类型设置

                    
                        switch (code)
                        {
                            case TypeCode.Boolean:
                                xls_value = xls_value.ToString();
                                break;
                            case TypeCode.Byte:
                                break;
                            case TypeCode.Char:
                                break;
                            case TypeCode.DBNull:
                                break;
                            case TypeCode.DateTime:
                                xls_value = XY.Convert.FormatDateTime(xls_value, DateTimeStyle.Style6);
                                break;
                            case TypeCode.Decimal:
                                break;
                            case TypeCode.Double:
                                break;
                            case TypeCode.Empty:
                                break;
                            case TypeCode.Int16:
                                break;
                            case TypeCode.Int32:
                                break;
                            case TypeCode.Int64:
                                break;
                            case TypeCode.Object:
                                break;
                            case TypeCode.SByte:
                                break;
                            case TypeCode.Single:
                                break;
                            case TypeCode.String:
                                break;
                            case TypeCode.UInt16:
                                break;
                            case TypeCode.UInt32:
                                break;
                            case TypeCode.UInt64:
                                break;
                            default:
                                break;
                        }
                        #endregion

                        Cell cell = cells.Add(line_index, key_index, xls_value);

                        key_index++;
                    }
                }
                line_index++;
            }

            xls.FileName = filePath;
            xls.Save();
        }
    }
}
*/