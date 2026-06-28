// See https://aka.ms/new-console-template for more information

using ExcelTools;
public class Program
{
    private static void Main(string[] args)
    {
        Console.WriteLine("Hello, World!");

        var dir = ExcelUtil.GetExecutingDir();
        var defaultExcelDir = Path.Combine(dir, "..", "..", "..", "Excel");
        var defaultExportJsonDir = Path.Combine(dir, "..", "..", "..", "Assets", "Abbresources", "GameCfgJson");
        var defaultClassExportDir = Path.Combine(dir, "..", "..", "..", "Assets", "Scripts", "GameCfgCS");
        var excelDir = args?.Length > 0 ? args[0] : defaultExcelDir;
        var cfgExportDir = args?.Length > 1 ? args[1] : defaultExportJsonDir;
        var classExportDir = args?.Length > 2 ? args[2] : defaultClassExportDir;
        //excelDir = "/Users/qiuxiaohui/Documents/UnityProject/abb2d/Misc/Excel";
        //cfgExportDir = "/Users/qiuxiaohui/Documents/UnityProject/abb2d/Assets/Abbresources/GameCfgJson";

        var excelTools = new ExcelTools.ExcelTools();
        var exportList = excelTools.GetExportCfgList(excelDir);

        var sourceType =
#if DEBUG
            EnExcelSourceType.c | EnExcelSourceType.e;
#else
            EnExcelSourceType.c;
#endif
        var exportExcelInfoList = excelTools.GetExportExcelInfoList(exportList, sourceType);
        var exportCfgList = excelTools.GetExportCfgListData(exportExcelInfoList);

        Directory.CreateDirectory(excelDir);
        Directory.CreateDirectory(cfgExportDir);
        Directory.CreateDirectory(classExportDir);

        excelTools.ExportExcelCatalogJsonFile(exportExcelInfoList, cfgExportDir);
        excelTools.ExportExcel2JsonFile(exportCfgList, cfgExportDir);
        excelTools.ExportExcel2CSFile(exportExcelInfoList, classExportDir);
        excelTools.CreateCfgUtilFile(exportExcelInfoList, classExportDir);

        Console.WriteLine($"excel -> {excelDir}");
        Console.WriteLine($"cfg -> {cfgExportDir}");
        Console.WriteLine($"class -> {classExportDir}");
        //Console.ReadLine();


        var resolve = new ResolveExcelJson();
        resolve.Execute(cfgExportDir, exportList);
    }
}