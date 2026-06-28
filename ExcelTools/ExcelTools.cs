using System;
using OfficeOpenXml;
using System.Drawing;
using OfficeOpenXml.Style;
using System;
using System.Reflection;
using Newtonsoft.Json;
using System.Diagnostics;
using System.Text;
using System.IO;

namespace ExcelTools
{
    public enum EnExcelSourceType
    {
        c = 1 << 0, // client
        s = 1 << 1, // server
        e = 1 << 2,
        All = c + s + e,  
    }
    public class ExcelTools
    {
        public List<string> GetExportCfgList(string dir)
        {
            // 缓存文件目录
            var fileDic = new Dictionary<string, string>();
            var files = Directory.GetFiles(dir, $"*.xlsx", SearchOption.AllDirectories);
            for (int i = 0; i < files.Length; i++)
            {
                var filePath = files[i];
                var fileName = Path.GetFileNameWithoutExtension(filePath);
                fileDic.Add(fileName, filePath);
            }

            if (!fileDic.TryGetValue("CfgCatalog", out var catalogPath))
            {
                Console.WriteLine($"没有找到目录文件 CfgCatalog.xlsx");
                return null;
            }
            var catalogExcel = ExcelUtil.GetExccel(catalogPath);

            var workbook = catalogExcel.Workbook;
            var workSheets = workbook.Worksheets;

            var exportList = new List<string>();
            for (int i = 0; i < workSheets.Count; i++)
            {
                var worksheet = workSheets[i];
                var rowCount = worksheet.Dimension.End.Row;
                var colCount = worksheet.Dimension.End.Column;

                for (int rowI = 1; rowI <= rowCount; rowI++)
                {
                    if (!worksheet.IsValid(rowI))
                        continue;
                    for (int colI = 1; colI <= colCount; colI++)
                    {
                        var valueStr = worksheet.GetValue(rowI, colI);
                        if (!fileDic.TryGetValue($"{valueStr}", out var path))
                            continue;
                        Console.WriteLine(path);
                        exportList.Add(path);
                    }
                }
            }
            return exportList;
        }

        public List<ExportExcelInfo> GetExportExcelInfoList(List<string> cfgList, EnExcelSourceType sourceType)
        {
            var exportExcelInfoList = new List<ExportExcelInfo>();
            for (int i = 0; i < cfgList.Count; i++)
            {
                var cfgPackage = ExcelUtil.GetExccel(cfgList[i]);
                var workbook = cfgPackage.Workbook;
                var worksheets = workbook.Worksheets;
                for (int j = 0; j < worksheets.Count(); j++)
                {
                    var sheet = worksheets[j];
                    var excelInfo = GetExcelInfo(sheet, sourceType);

                    var exportExcelInfo = GetExportExcelInfo(sheet, excelInfo);
                    exportExcelInfoList.Add(exportExcelInfo);
                }
            }
            return exportExcelInfoList;
        }
        public Dictionary<string, List<ICfg>> GetExportCfgListData(List<ExportExcelInfo> exportExcelInfoList)
        {
            var cfgDic = new Dictionary<string, List<ICfg>>();
            for (int i = 0; i < exportExcelInfoList.Count; i++)
            {
                var exportInfo = exportExcelInfoList[i];
                var className = exportInfo.worksheet.Name;
                var dynamicType = ExcelUtil.CreateDynamicType(className, exportInfo.fieldList, exportInfo.fieldTypeList);
                var fieldList = dynamicType.GetFields(BindingFlags.Public | BindingFlags.Instance);

                var cfgInsList = new List<ICfg>();
                var rowCount = exportInfo.worksheet.Dimension.End.Row;
                for (int j = exportInfo.excelInfo.dataStartRow; j <= rowCount; j++)
                {
                    if (!exportInfo.worksheet.IsValid(j))
                        continue;
                    var insType = Activator.CreateInstance(dynamicType);
                    for (int k = 0; k < fieldList.Length; k++)
                    {
                        var fieldInfo = fieldList[k];
                        if (!exportInfo.field2ColList.TryGetValue(fieldInfo.Name, out var col))
                        {
                            Console.Error.WriteLine($"读取字段列数失败，type:{insType}, name:{fieldInfo.Name}");
                            continue;
                        }
                        var valueStr = exportInfo.worksheet.GetValue<string>(j, col);
                        try
                        {
                            if (!typeof(string).Equals(fieldInfo.FieldType))
                            {
                                var jsonStr = JsonConvert.DeserializeObject(valueStr, fieldInfo.FieldType);
                                var value = Convert.ChangeType(jsonStr, fieldInfo.FieldType);
                                if (value == null)
                                {
                                    Console.Error.WriteLine($"类型转化失败，type:{insType}, value:{valueStr}， row:{j}, col:{col}, ");
                                    continue;
                                }
                                fieldInfo.SetValue(insType, value);
                            }
                            else
                            {
                                fieldInfo.SetValue(insType, valueStr);
                            }
                        }
                        catch (Exception ex)
                        {
                            //Console.Error.WriteLine(ex);
                            var value = fieldInfo.FieldType.IsValueType
                                ? Activator.CreateInstance(fieldInfo.FieldType)
                                : null;
                            fieldInfo.SetValue(insType, value);
                        }
                    }
                    cfgInsList.Add(insType as ICfg);
                }
                cfgDic.Add(className, cfgInsList);
            }
            return cfgDic;
        }

        public void ExportExcelCatalogJsonFile(List<ExportExcelInfo> exportExcelInfos, string dir)
        {
            var path = Path.Combine(dir, "CfgCatalog.json");
            var jsonStr = JsonConvert.SerializeObject(exportExcelInfos);
            File.WriteAllText(path, jsonStr);
        }
        public void ExportExcel2JsonFile(Dictionary<string, List<ICfg>> cfgDic, string targetRoot)
        {
            foreach (var item in cfgDic)
            {
                var jsonStr = JsonConvert.SerializeObject(item.Value);
                var path = Path.Combine(targetRoot, $"{item.Key}.json");
                File.WriteAllText(path, jsonStr, System.Text.Encoding.UTF8);
                Console.WriteLine(path);
            }
        }

        public void ExportExcel2CSFileEcs(List<ExportExcelInfo> excelInfoList, string targetRoot)
        {

        }
        public void ExportExcel2CSFile(List<ExportExcelInfo> excelInfoList, string targetRoot)
        {
            foreach (var excelInfo in excelInfoList)
            {
                var name = excelInfo.worksheet.Name;
                var path = Path.Combine(targetRoot, $"{name}.cs");
                var csContentStr = ExcelUtil.CreateCshapFileContent(
                    name
                    , excelInfo.strDesc
                    , excelInfo.fieldList
                    , excelInfo.fieldTypeList
                    , excelInfo.descList
                    , excelInfo.excelInfo.keysCol);

                File.WriteAllText(path, csContentStr, System.Text.Encoding.UTF8);
                Console.WriteLine(path);
            }
        }

        private ExportExcelInfo GetExportExcelInfo(ExcelWorksheet sheet, ExcelInfo excelInfo)
        {
            var exportExcelInfo = new ExportExcelInfo();
            exportExcelInfo.excelInfo = excelInfo;
            exportExcelInfo.strDesc = $"{sheet.Name}";
            exportExcelInfo.worksheet = sheet;
            var colCount = sheet.Columns.Count();
            for (int col = excelInfo.dataStartCol; col <= colCount; col++)
            {
                if (!excelInfo.validCol.Contains(col))
                    continue;
                var valueStr = sheet.GetValue<string>(excelInfo.descRow, col);
                exportExcelInfo.descList.Add(valueStr);
            }
            for (int col = excelInfo.dataStartCol; col <= colCount; col++)
            {
                if (!excelInfo.validCol.Contains(col))
                    continue;
                var valueStr = sheet.GetValue<string>(excelInfo.fieldRow, col);
                exportExcelInfo.fieldList.Add(valueStr);
                exportExcelInfo.field2ColList.Add(valueStr, col);
            }
            for (int col = excelInfo.dataStartCol; col <= colCount; col++)
            {
                if (!excelInfo.validCol.Contains(col))
                    continue;
                var valueStr = sheet.GetValue<string>(excelInfo.fieldTypeRow, col);
                exportExcelInfo.fieldTypeList.Add(valueStr);
            }
            return exportExcelInfo;
        }

        public void CreateCfgUtilFileEcs(List<ExportExcelInfo> excelInfoList, string targetRoot)
        {
        }
        public void CreateCfgUtilFile(List<ExportExcelInfo> excelInfoList, string targetRoot)
        {
            var initStrBuilder = new StringBuilder();
            initStrBuilder.AppendLine("\tpublic void Initialization()");
            initStrBuilder.AppendLine("\t{");
            var filedStrBuilder = new StringBuilder();
            var methodStrBuilder = new StringBuilder();

            for (int i = 0; i < excelInfoList.Count; i++)
            {
                var excelInfo = excelInfoList[i];
                var keysList = excelInfo.excelInfo.keysCol;
                var name = excelInfo.excelInfo.excelName;
                var arrName = $"m_{name}";
                filedStrBuilder.AppendLine($"\tprivate {name}[] {arrName} = null;");
                initStrBuilder.AppendLine($"\t\tfor (int i = 0; i < m_{name}.Length; i++)");
                initStrBuilder.AppendLine("\t\t{");
                initStrBuilder.AppendLine($"\t\t\tvar cfg = m_{name}[i];");


                // index => cfg
                {
                    methodStrBuilder.AppendLine($"\tpublic {typeof(int)} Get{name}Count()");
                    methodStrBuilder.AppendLine($"\t{{");
                    methodStrBuilder.AppendLine($"\t\treturn {arrName}.{nameof(arrName.Length)};");
                    methodStrBuilder.AppendLine($"\t}}");

                    methodStrBuilder.AppendLine($"\tpublic {name} Get{name}({typeof(int)} index)");
                    methodStrBuilder.AppendLine($"\t{{");
                    methodStrBuilder.AppendLine($"\t\treturn {arrName}[index];");
                    methodStrBuilder.AppendLine($"\t}}");
                }

                for (int j = 0; j < keysList.Count; j++)
                {
                    var keyColList = keysList[j];
                    var colCount = keyColList.Count;
                    var typeStr = "{0}";
                    var dicName = $"m_Dic{name}{j}";
                    var addDic = dicName;
                    var addKey = "";
                    methodStrBuilder.Append($"\tpublic {name} Get{name}{j}");
                    methodStrBuilder.Append('(');
                    var methodContent = $"{dicName}";
                    for (int k = 0; k < colCount; k++)
                    {
                        var cfgKeyInfo = keyColList[k];
                        var typeKey = cfgKeyInfo.fieldType;
                        var type = ExcelUtil.Key2Type(typeKey);
                        var fieldName = cfgKeyInfo.fieldName;
                        typeStr = string.Format(typeStr, ExcelUtil.GetDictionaryType(type.ToString(), $"{{0}}"));
                        var keyStr = $"cfg.{fieldName}";
                        var valueStr = $"value{j}_{k}";
                        addKey = keyStr;
                        methodStrBuilder.Append($"{type} {fieldName}");
                        methodContent += $"[{fieldName}]";
                        if (k < colCount - 1)
                        {
                            initStrBuilder.AppendLine($"\t\t\tif (!{addDic}.TryGetValue({keyStr}, out var {valueStr}))" +
                                $" {{" +
                                $" {valueStr} = new();" +
                                $" {dicName}.Add({keyStr}, {valueStr});" +
                                $" }}");
                            addDic = valueStr;
                            methodStrBuilder.Append(", ");
                        }
                    }
                    methodStrBuilder.AppendLine(")");
                    methodStrBuilder.AppendLine("\t{");
                    methodStrBuilder.AppendLine($"\t\treturn {methodContent};");
                    methodStrBuilder.AppendLine("\t}");

                    initStrBuilder.AppendLine($"\t\t\t{addDic}.Add({addKey}, cfg);");
                    typeStr = string.Format(typeStr, name);
                    filedStrBuilder.AppendLine($"\tprivate {typeStr} {dicName} = new();");
                }
                initStrBuilder.AppendLine("\t\t}");
            }
            initStrBuilder.AppendLine("\t}");



            var strBuilder = new StringBuilder();
            strBuilder.AppendLine($"public partial class GameSchedule");
            strBuilder.AppendLine("{");
            strBuilder.Append(filedStrBuilder);
            strBuilder.Append(initStrBuilder);
            strBuilder.Append(methodStrBuilder);
            strBuilder.AppendLine("}");

            var path = Path.Combine(targetRoot, "GameSchedule.cs");
            File.WriteAllText(path, strBuilder.ToString());
        }

        private ExcelInfo GetExcelInfo(ExcelWorksheet worksheet, EnExcelSourceType sourceType)
        {
            var result = new ExcelInfo();
            result.excelName = worksheet.Name;

            var checkCol = 1;
            var startCol = 2;
            var rowCount = MathF.Min(10, worksheet.Dimension.End.Row);
            var colCount = worksheet.Columns.Count();
            var startRow = -1;
            for (int i = 1; i <= rowCount; i++)
            {
                if (!worksheet.IsValid(i))
                    continue;
                var valueStr = worksheet.GetValue<string>(i, checkCol);
                switch (valueStr)
                {
                    case "${link}":
                        result.linkRow = i;
                        break;
                    case "${name}":
                        result.fieldRow = i;
                        break;
                    case "${desc}":
                        result.descRow = i;
                        break;
                    case "${type}":
                        result.fieldTypeRow = i;
                        break;
                    case "${source}":
                        {
                            for (int j = startCol; j <= colCount; j++)
                            {
                                var valueSource = worksheet.GetValue<string>(i, j) ?? "";
                                foreach (var item in valueSource)
                                {
                                    if (!Enum.TryParse<EnExcelSourceType>(item.ToString(), true, out var type))
                                        continue;
                                    if ((sourceType & type) != type)
                                        continue;
                                    if (result.validCol.Contains(j))
                                        continue;
                                    result.validCol.Add(j);
                                }
                            }
                        }
                        break;
                    case "${key}":
                        {
                            var keyInfo = new List<CfgKeyInfo>();
                            result.keysCol.Add(keyInfo);
                            for (int j = startCol; j <= colCount; j++)
                            {
                                var valueSource = worksheet.GetValue<string>(i, j);
                                if (string.IsNullOrEmpty(valueSource))
                                    continue;
                                keyInfo.Add(new() { col = j });
                            }
                        }
                        break;
                    default:
                        continue;
                }
                startRow = i;
            }

            foreach (var item in result.keysCol)
            {
                foreach (var value in item)
                {
                    value.fieldName = $"{worksheet.GetValue(result.fieldRow, value.col)}";
                    value.fieldType = $"{worksheet.GetValue(result.fieldTypeRow, value.col)}";
                }
            }

            result.dataStartCol = startCol;
            result.dataStartRow = startRow + 1;
            return result;
        }
    }

    public class ExportExcelInfo
    {
        public ExcelInfo excelInfo;
        public string strDesc;
        public List<string> fieldList = new();
        public Dictionary<string, int> field2ColList = new();
        public List<string> fieldTypeList = new();
        public List<string> descList = new();
        [NonSerialized]
        public ExcelWorksheet worksheet;
    }
    public class ExcelInfo
    {
        public string excelName;
        public int dataStartRow;
        public int dataStartCol;
        public int fieldRow;
        public int fieldTypeRow;
        public int descRow;
        public int linkRow;
        public HashSet<int> validCol = new();
        public List<List<CfgKeyInfo>> keysCol = new();
    }
    public class CfgKeyInfo
    {
        public int col;
        public string fieldName;
        public string fieldType;
    }
}

