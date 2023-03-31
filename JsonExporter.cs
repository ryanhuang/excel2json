using System;
using System.IO;
using System.Data;
using System.Text;
using System.Collections.Generic;
using Newtonsoft.Json;
using System.Net.Security;
using System.Data.SqlTypes;
using System.Runtime.CompilerServices;
using System.ComponentModel.DataAnnotations;
using System.Collections;

namespace excel2json
{
    struct SheetColumnInfo
    {
        // 标题名
        public string fieldName;
        // 数据类型
        public Type dataType;
        // 数据初始值
        public object defaultValue;
        // 列index
        public int colIndex;
        // 附带的Column数据
        public DataColumn column;
    }

    enum ConvertJsonMode
    {
        JsonArray,
        JsonDict
    }


    // 转换选项
    struct ConvertOptions
    {        
        // 头行数
        public int headerRows;
        // 是否转换为小写
        public bool lowcase;
        // 是否导出数组
        public bool exportArray;
        // 日期格式
        public string dateFormat;
        // 是否强制使用表单名
        public bool forceSheetName;
        // 排除表单前缀
        public string excludePrefix;
        // 单元格内的字符串是否按照json格式解析输出
        public bool cellJson;
        // 是否全部转换为字符串
        public bool allString;

        public bool isExcludeName(string name)
        {
             if (string.IsNullOrEmpty(excludePrefix))
                return false;
            return name.StartsWith(excludePrefix);
        }
    }


    // 用于描述一个Excel Sheet内数据的缓存类
    class ExcelSheetJsonConvertor
    {
        private DataTable _dataset;
        private ConvertOptions _options;
        private Dictionary<string, string> _mappingTypes = new()
        {
            { "string", "System.String" },
            { "int", "System.Int32"},
            { "uint", "System.UInt32"},
            { "long", "System.Int64"},
            { "ulong", "System.UInt64"},
            { "float", "System.Single"},
            { "double", "System.Double"},
            { "bool", "System.Boolean"},
            { "date", "System.DateTime"},
            { "datetime", "System.DateTime"}
        };



        public DataTable DataSet => _dataset;        
        public List<SheetColumnInfo> columns = new ();

        // 把表单内数据按照数组格式进行转换
        private object serializeToDataArray()
        {
            List<object> values = new List<object>();
            int firstDataRow = _options.headerRows;
            var sheet = DataSet;
            // 每一行都进行转换
            for (int i = firstDataRow; i < sheet.Rows.Count; i++)
            {
                DataRow row = sheet.Rows[i];
                Dictionary<string, object> value = serializeByRow(row);
                values.Add(value);
            }
            return values;
        }

        /// <summary>
        /// 以第一列为ID，转换成ID->Object的字典对象
        /// </summary>
        private object serializeToDataDict()
        {
            Dictionary<string, object> importData = new Dictionary<string, object>();
            int firstDataRow = _options.headerRows; ;
            var sheet = DataSet;
            for (int i = firstDataRow; i < sheet.Rows.Count; i++)
            {
                DataRow row = sheet.Rows[i];
                string ID = row[sheet.Columns[0]].ToString();
                if (ID.Length <= 0)
                    ID = string.Format("row_{0}", i);

                var rowObject = serializeByRow(row);
                // 多余的字段
                // rowObject[ID] = ID;
                importData[ID] = rowObject;
            }

            return importData;
        }

        private string mappingTypeString(string typeString)
        {
            string result;
            if (_mappingTypes.TryGetValue(typeString.ToLower(), out result) == false)
            {
                result = "System.String";
            }
            return result;
        }

        /// <summary>
        /// 把一行数据转换成一个对象，每一列是一个属性
        /// </summary>
        private Dictionary<string, object> serializeByRow(DataRow row)
        {
            var rowData = new Dictionary<string, object>();                                    
            foreach (SheetColumnInfo column in columns)
            {         
                // 取表格数据
                object value = row[column.colIndex];

                // 尝试将单元格字符串转换成 Json Array 或者 Json Object
                if (_options.cellJson)
                {
                    string cellText = value.ToString().Trim();
                    if (cellText.StartsWith("[") || cellText.StartsWith("{"))
                    {
                        try
                        {
                            object cellJsonObj = JsonConvert.DeserializeObject(cellText);
                            if (cellJsonObj != null)
                                value = cellJsonObj;
                        }
                        catch (Exception exp)
                        {
                        }
                    }
                }

                if (value.GetType() == typeof(System.DBNull))
                {
                    value = column.defaultValue;
                }
                else if (value.GetType() == typeof(double))
                { // 去掉数值字段的“.0”
                    double num = (double)value;
                    if ((long)num == num)
                        value = (long)num;
                }
                //全部转换为string
                if (_options.allString && !(value is string))
                {
                    value = value.ToString();
                }             
                rowData[column.fieldName] = value;                
            }

            return rowData;
        }


        public ExcelSheetJsonConvertor(DataTable excelSheetDataset, ConvertOptions options)           
        {           
            _options = options;
            _dataset = excelSheetDataset;
            initColumnCache();
        }

        public void initColumnCache()
        {
            var sheet = DataSet;            
            for (int i = 0; i < sheet.Columns.Count - 1; i++)
            {
                DataColumn column = sheet.Columns[i];                
                // 不满足条件的columen跳过
                string columnName = column.ToString();
                if (_options.isExcludeName(columnName))
                    continue;               

                // 对列进行缓存, 并提前运算好数据
                SheetColumnInfo columnCache;
                columnCache.colIndex = i;
                columnCache.column = column;                
                columnCache.dataType = typeof(string);
                // 字段TypeRow
                DataRow rowDataType = sheet.Rows[0];
                var val = rowDataType[i];
                if (val is string typeString)
                {
                    Type type = Type.GetType(mappingTypeString(typeString));
                    if (type != null)
                    {
                        columnCache.dataType = type;
                    }                    
                };                

                // 如果是字符串就赋值空串
                if (columnCache.dataType == typeof(String))
                {
                    columnCache.defaultValue = "";
                }
                else
                {
                    columnCache.defaultValue = columnCache.dataType.IsValueType ? Activator.CreateInstance(columnCache.dataType) : null;
                }
              
                // 表头自动转换成小写
                columnCache.fieldName = column.ToString();                
                if (_options.lowcase)
                    columnCache.fieldName = columnCache.fieldName.ToLower();
                if (string.IsNullOrEmpty(columnCache.fieldName))
                    columnCache.fieldName = string.Format("col_{0}", i);

                this.columns.Add(columnCache);
            }           
        }

        public object serialize()
        {
            if (_options.exportArray)
                return serializeToDataArray();
            else
                return serializeToDataDict();
        }      
    }


    /// <summary>
    /// 将DataTable对象，转换成 JSON string 并保存到文件中
    /// </summary>
    class JsonExporter
    {
        // 转换后的结果
        private string _jsonResult = "";
        // 转换选项
        private ConvertOptions _options;

        public string context {
            get {
                return _jsonResult;
            }
        }

        /// <summary>
        /// 构造函数：完成内部数据创建
        /// </summary>
        /// <param name="excel">ExcelLoader Object</param>       
        public JsonExporter(ExcelLoader excel, bool lowcase, bool exportArray, string dateFormat, bool forceSheetName, int headerRows, string excludePrefix, bool cellJson, bool allString)
        {
            _options.lowcase = lowcase; 
            _options.exportArray = exportArray;
            _options.dateFormat = dateFormat;
            _options.forceSheetName = forceSheetName;
            _options.headerRows = headerRows - 1;
            _options.excludePrefix = excludePrefix;
            _options.cellJson = cellJson;
            _options.allString = allString;

            // 按照特定的符号 过滤表单名称
            List<ExcelSheetJsonConvertor> validSheets = new List<ExcelSheetJsonConvertor>();
            for (int i = 0; i < excel.Sheets.Count; i++)
            {
                var sheet = excel.Sheets[i];
                ExcelSheetJsonConvertor sheetInfo = new(sheet, _options);                
                // 过滤掉包含特定前缀的表单
                string sheetName = sheetInfo.DataSet.TableName;
                if (excludePrefix.Length > 0 && sheetName.StartsWith(excludePrefix))
                    continue;

                if (sheet.Columns.Count > 0 && sheet.Rows.Count > 0)
                    validSheets.Add(sheetInfo);
            }

            // json 格式化的配置
            var jsonSettings = new JsonSerializerSettings
            {
                DateFormatString = dateFormat,
                Formatting = Formatting.Indented
            };

            // 如果不是强制使用的表单名
            if (!forceSheetName && validSheets.Count == 1)
            {   
                object sheetValue = validSheets[0].serialize() ;
                _jsonResult = JsonConvert.SerializeObject(sheetValue, jsonSettings);
            }
            else
            { 
                Dictionary<string, object> data = new Dictionary<string, object>();
                foreach (var sheetCache in validSheets)
                {
                    object sheetValue = sheetCache.serialize();
                    data.Add(sheetCache.DataSet.TableName, sheetValue);
                }
                _jsonResult = JsonConvert.SerializeObject(data, jsonSettings);
            }
        }       
        

        /// <summary>
        /// 将内部数据转换成Json文本，并保存至文件
        /// </summary>
        /// <param name="jsonPath">输出文件路径</param>
        public void SaveToFile(string filePath, Encoding encoding)
        {
            //-- 保存文件
            using (FileStream file = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                using (TextWriter writer = new StreamWriter(file, encoding))
                    writer.Write(_jsonResult);
            }
        }
    }
}
