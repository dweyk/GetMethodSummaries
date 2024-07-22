// See https://aka.ms/new-console-template for more information

using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using Namotion.Reflection;
using OfficeOpenXml;

var actions = new List<string>();
ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

using (ExcelPackage xlPackage = new ExcelPackage(new FileInfo(@"C:\Users\v.murtazin\Desktop\Описание методов.xlsx")))
{
    var myWorksheet = xlPackage.Workbook.Worksheets.First(); //select sheet here
    var totalRows = myWorksheet.Dimension.End.Row;
    var totalColumns = 1;

    for (var rowNum = 2; rowNum <= totalRows; rowNum++) //select starting row here
    {
        var row = myWorksheet.Cells[rowNum, 1, rowNum, totalColumns].Select(c => c.Value == null ? string.Empty : c.Value.ToString());
        actions.Add(row.FirstOrDefault() ?? string.Empty);
    }
}

var rootPath = @"C:\source\mgz";
var rootDirectoryPath = new DirectoryInfo(rootPath).FullName;
var pattern = new Regex(@"\\bin\\.*Bars.*\.dll");

var dlls = Directory.GetFiles(rootDirectoryPath, "*.dll", SearchOption.AllDirectories)
    .Where(row => row.Contains("\\bin\\") && pattern.IsMatch(row))
    .Distinct();

var assemblies = new Dictionary<string, Assembly>();
var typeList = new List<Type>();

foreach (var dll in dlls)
{
    try
    {
        if (assemblies.ContainsKey(dll))
        {
            continue;
        }

        var assembly = Assembly.LoadFrom(dll);
        assemblies.Add(dll, assembly);
    }
    catch (Exception e)
    {
        Console.WriteLine(e);
    }
}

foreach (var (key, assembly) in assemblies)
{
    Type[] types;
    try
    {
        types = assembly.GetTypes();
        foreach (var type in types)
        {
            if (typeList.Any(t => t.FullName == type.FullName))
            {
                continue;
            }

            typeList.Add(type);
        }
    }
    catch (Exception e)
    {
        Console.WriteLine(e);
    }
}

var MethodInformations = new List<MethodInformation>();

if (typeList.Count > 0)
{
    foreach (var type in typeList)
    {
        if ((bool) type.FullName?.ToLower().Contains("system"))
        {
            continue;
        }

        var methods = type.GetMethods(BindingFlags.Instance | BindingFlags.Public | BindingFlags.DeclaredOnly |
                                      BindingFlags.Static);
        foreach (var member in methods)
        {
            var summary = GetSummary(member);
            if (!string.IsNullOrWhiteSpace(summary))
            {
                //Console.WriteLine($"{type.Name}.{member.Name} {GetSummary(member)}");
                var controllerOrService = type.Name.Replace("Controller", "").Replace("Service", "");
                var method = member.Name;
                MethodInformations.Add(new MethodInformation
                {
                    Action = $"{controllerOrService}/{method}",
                    Summary = summary,
                });
            }
        }

        var methods1 = type.GetMethods(BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.DeclaredOnly |
                                       BindingFlags.Static);
        foreach (var member in methods1)
        {
            var summary = GetSummary(member);
            if (!string.IsNullOrWhiteSpace(summary))
            {
                var controllerOrService = type.Name.Replace("Controller", "").Replace("Service", "");
                var method = member.Name;
                MethodInformations.Add(new MethodInformation
                {
                    Action = $"{controllerOrService}/{method}",
                    Summary = summary,
                });
            }
        }
    }
}

static string GetSummary(MethodInfo methodInfo)
{
    var summary = methodInfo.GetXmlDocsTag("summary");

    return !string.IsNullOrWhiteSpace(summary) ? summary : string.Empty;
}

actions.ForEach(action =>
{
    MethodInformations.ForEach(methodInformation =>
    {
        if (methodInformation.Action.Contains(action) || action.Contains(methodInformation.Action))
        {
            Console.WriteLine($"{methodInformation.Action} {methodInformation.Summary}");
        }
    });
});

Console.ReadLine();

public class MethodInformation
{
    public string Action { get; set; }
    public string Summary { get; set; }
}