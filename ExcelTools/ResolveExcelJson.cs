using System;
using Newtonsoft.Json;

namespace ExcelTools
{
    public class ResolveExcelJson
    {
        public void Execute(string jsonDir, List<string> catalog)
        {
            //for (int i = 0; i < catalog.Count; i++)
            //{
            //    var cfgName = Path.GetFileNameWithoutExtension(catalog[i]);
            //    var path = Path.Combine(jsonDir, $"{cfgName}.json");
            //    var jsonStr = File.ReadAllText(path);
            //    var type = Type.GetType(cfgName);
            //    var listType = typeof(List<>);
            //    var listT = listType.MakeGenericType(type);
            //    var insCl = JsonConvert.DeserializeObject(jsonStr, listT);
            //    Console.WriteLine(insCl);
            //}
        }
    }
}

