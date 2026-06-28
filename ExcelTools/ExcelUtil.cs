using System;
using OfficeOpenXml;
using System.Reflection;
using System.Reflection.Emit;
using System.Text;
using Unity.Entities;

namespace ExcelTools
{
    public static class ExcelUtil
    {
        public static string GetExecutingDir()
        {
            var curAssembly = Assembly.GetExecutingAssembly();
            var curPath = Path.GetDirectoryName(curAssembly.Location);
            return curPath;
        }
        public static ExcelPackage GetExccel(string excelPath)
        {
            var excelPackage = new ExcelPackage(excelPath);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            return excelPackage;
        }
        public static bool IsValid(this ExcelWorksheet sheet, int row)
        {
            var value = sheet.GetValue<string>(row, 1);
            var isvalid = string.IsNullOrWhiteSpace(value);
            return !isvalid;
        }
        public static Type CreateDynamicType(string className, List<string> fieldNameList, List<string> fieldTypeList)
        {
            var assemblyName = new AssemblyName("DynamicAssembly");
            AssemblyBuilder assemblyBuilder = AssemblyBuilder.DefineDynamicAssembly(assemblyName, AssemblyBuilderAccess.Run);
            ModuleBuilder moduleBuilder = assemblyBuilder.DefineDynamicModule("MainModule");
            TypeBuilder classBuilder = moduleBuilder.DefineType(className, TypeAttributes.Public);

            // 创建一个无参数的构造函数
            var constructorBuilder = classBuilder.DefineConstructor(MethodAttributes.Public, CallingConventions.Standard, Type.EmptyTypes);
            ILGenerator constructorIL = constructorBuilder.GetILGenerator();
            //constructorIL.Emit(OpCodes.Ldstr, "Hello, World!");
            //constructorIL.Emit(OpCodes.Call, typeof(Console).GetMethod("WriteLine", new Type[] { typeof(string) }));
            constructorIL.Emit(OpCodes.Ret);
            classBuilder.AddInterfaceImplementation(typeof(ICfg));

            for (int i = 0; i < fieldNameList.Count; i++)
            {
                var fileName = fieldNameList[i];
                var filedKeyType = fieldTypeList[i];
                var fieldType = Key2Type(filedKeyType);
                if (string.IsNullOrEmpty(fileName) || fieldType == null)
                {
                    Console.WriteLine($"类型转化失败 class:{className}, name:{fileName}, type:{filedKeyType}");
                    continue;
                }
                var fieldBuilder = classBuilder.DefineField(fileName, fieldType, FieldAttributes.Public);
            }

            // 创建类型
            Type dynamicType = classBuilder.CreateType();
            return dynamicType;
        }
        public static Type Key2EcsType(string key)
        {
            return key switch
            {
                "int32" => typeof(int),
                "int64" => typeof(long),
                "int32[]" => typeof(BlobAssetReference<BlobArray<int>>),
                "int64[]" => typeof(BlobAssetReference<BlobArray<long>>),
                "uint32" => typeof(uint),
                "uint64" => typeof(ulong),
                "uint32[]" => typeof(BlobAssetReference<BlobArray<uint>>),
                "uint64[]" => typeof(BlobAssetReference<BlobArray<ulong>>),
                "float" => typeof(float),
                "float[]" => typeof(BlobAssetReference<BlobArray<float>>),
                "double" => typeof(double),
                "double[]" => typeof(BlobAssetReference<BlobArray<double>>),
                "string" => typeof(string),
                "string[]" => typeof(BlobAssetReference<BlobArray<BlobString>>),
                _ => null,
            };
        }
        public static Type Key2Type(string key)
        {
            return key switch
            {
                "int32" => typeof(int),
                "int64" => typeof(long),
                "int32[]" => typeof(int[]),
                "int64[]" => typeof(long[]),
                "uint32" => typeof(uint),
                "uint64" => typeof(ulong),
                "uint32[]" => typeof(uint[]),
                "uint64[]" => typeof(ulong[]),
                "float" => typeof(float),
                "float[]" => typeof(float[]),
                "double" => typeof(double),
                "double[]" => typeof(double[]),
                "string" => typeof(string),
                "string[]" => typeof(string[]),
                _ => null,
            };
        }
        public static string GetListType(string type)
        {
            return $"System.Collections.Generic.List<{type}>";
        }
        public static string GetDictionaryType(string type1, string type2)
        {
            return $"System.Collections.Generic.Dictionary<{type1}, {type2}>";
        }

        public static string CreateCshapFileContentEcs(
            string name
            , string csDesc
            , List<string> fieldList
            , List<string> fieldTypeList
            , List<string> descList
            , List<List<CfgKeyInfo>> keyList)
        {
            var strBuilder = new StringBuilder("");
            strBuilder.AppendLine($"// {csDesc}");
            strBuilder.AppendLine($"public class Ecs{name} : ICfg");
            strBuilder.AppendLine("{");
            strBuilder.AppendLine($"\tprivate {name}() {{}}");

            for (int i = 0; i < fieldList.Count; i++)
            {
                var fieldName = fieldList[i];
                var typeStr = fieldTypeList[i];
                var fieldType = Key2EcsType(typeStr);
                var descStr = descList[i];
                if (string.IsNullOrEmpty(fieldName) || fieldType == null)
                {
                    Console.Error.WriteLine($"生成类字段失败 name:{fieldName}, typeStr:{typeStr}, type:{fieldType}");
                    continue;
                }

                strBuilder.AppendLine($"\t// {descStr}");
                strBuilder.AppendLine($"public {fieldType} {fieldName};");
            }
            for (int i = 0; i < keyList.Count; i++)
            {
                var keyInfoList = keyList[i];
                var count = keyInfoList.Count;
                if (count != 1)
                    continue;
                var returnType = "";
                var returnStr = "";
                if (count == 1)
                {
                    var info = keyInfoList[0];
                    var type = ExcelUtil.Key2EcsType(info.fieldType);
                    returnType += $"{type}";
                    returnStr += info.fieldName;
                }

                strBuilder.AppendLine($"\tpublic {returnType} GetID()");
                strBuilder.AppendLine($"\t{{");
                strBuilder.AppendLine($"\t\treturn {returnStr};");
                strBuilder.AppendLine($"\t}}");
            }
            strBuilder.AppendLine("}");
            var result = strBuilder.ToString();
            return result;
        }
        public static string CreateCshapFileContent(
            string name
            , string csDesc
            , List<string> fieldList
            , List<string> fieldTypeList
            , List<string> descList
            , List<List<CfgKeyInfo>> keyList)
        {
            var strBuilder = new StringBuilder("");
            strBuilder.AppendLine($"// {csDesc}");
            strBuilder.AppendLine($"public class {name} : ICfg");
            strBuilder.AppendLine("{");
            strBuilder.AppendLine($"\tprivate {name}() {{}}");

            for (int i = 0; i < fieldList.Count; i++)
            {
                var fieldName = fieldList[i];
                var typeStr = fieldTypeList[i];
                var fieldType = Key2Type(typeStr);
                var descStr = descList[i];
                if (string.IsNullOrEmpty(fieldName) || fieldType == null)
                {
                    Console.Error.WriteLine($"生成类字段失败 name:{fieldName}, typeStr:{typeStr}, type:{fieldType}");
                    continue;
                }

                strBuilder.AppendLine($"\t// {descStr}");
                strBuilder.Append($"\t[Newtonsoft.Json.JsonProperty()] ");
                strBuilder.AppendLine($"public readonly {fieldType} {fieldName};");
            }
            for (int i = 0; i < keyList.Count; i++)
            {
                var keyInfoList = keyList[i];
                var count = keyInfoList.Count;

                var returnType = "";
                var returnStr = "";
                if (count == 1)
                {
                    var info = keyInfoList[0];
                    var type = ExcelUtil.Key2Type(info.fieldType);
                    returnType += $"{type}";
                    returnStr += info.fieldName;
                }
                else
                {
                    returnType += "(";
                    returnStr += "(";
                    for (int j = 0; j < count; j++)
                    {
                        var info = keyInfoList[j];
                        var type = ExcelUtil.Key2Type(info.fieldType);
                        returnType += $"{type} {info.fieldName}";
                        returnStr += $"{info.fieldName}";
                        if (j != count - 1)
                        {
                            returnType += ",";
                            returnStr += ",";
                        }
                    }
                    returnStr += ");";
                    returnType += ")";
                }
                if (i == 0)
                    strBuilder.AppendLine($"\tpublic {returnType} GetID()");
                else
                    strBuilder.AppendLine($"\tpublic {returnType} GetID_{i}()");
                strBuilder.AppendLine($"\t{{");
                strBuilder.AppendLine($"\t\treturn {returnStr};");
                strBuilder.AppendLine($"\t}}");
            }
            strBuilder.AppendLine("}");
            var result = strBuilder.ToString();
            return result;
        }
        public static string CreateICfgFileContent()
        {
            var strBuilder = new StringBuilder("");
            strBuilder.AppendLine("public interface ICfg");
            strBuilder.AppendLine("{");
            strBuilder.AppendLine("}");
            var result = strBuilder.ToString();
            return result;
        }
    }
}

