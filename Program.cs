using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using CommandLine;
using CsvHelper;

namespace xps_helper
{
    public class Program
    {
        public static void Main(string[] args)
        {
            Console.WriteLine("XPS Helper");
            Parser.Default.ParseArguments<Options>(args).WithParsed((o) =>
            {
                foreach (var item in Directory.EnumerateDirectories(o.ParentFolder))
                {
                    ProcessFolder(o, item);
                }
            });
        }

        public static void ProcessFolder(Options o, string folderPath)
        {
            Console.WriteLine($"Processing folder '{folderPath}'");
            var files = Directory.GetFiles(folderPath);
            var mapFile = files.FirstOrDefault(x => Regex.IsMatch(Path.GetFileName(x), o.MapFileFilter));
            if (mapFile == null) return;
            var dataFiles = files.Where(x => Regex.IsMatch(Path.GetFileName(x), o.DataFileFilter));
            var mapping = ParseMapFile(mapFile);
            foreach (var item in dataFiles)
            {
                var saveName = mapping[Path.GetFileName(item)];
                var savePath = Path.Combine(Path.GetDirectoryName(item), saveName);
                Console.WriteLine($"Converting '{item}' to '{savePath}'...");
                MakeXlsxFromTxt(item, savePath);
            }
        }

        public static Dictionary<string, string> ParseMapFile(string filePath)
        {
            var lines = File.ReadLines(filePath);
            var groupsOfLines = new List<List<string>>() { new List<string>() };
            var currentGroup = groupsOfLines[0];
            foreach (var item in lines)
            {
                var trimmed = item.Trim('\r', ' ');
                if (trimmed.Length > 0)
                {
                    currentGroup.Add(trimmed);
                }
                else
                {
                    groupsOfLines.Add(new List<string>());
                    currentGroup = groupsOfLines.Last();
                }
            }
            groupsOfLines = groupsOfLines.Where(x => x.Any()).ToList();
            Dictionary<string, string> res = new();
            foreach (var item in groupsOfLines)
            {
                var sanitizedSaveName = item[1];
                foreach (var c in Path.GetInvalidFileNameChars())
                {
                    sanitizedSaveName = sanitizedSaveName.Replace(c, '_');
                }
                res.Add(item[0].Split('\\', '/', StringSplitOptions.RemoveEmptyEntries).Last(), $"{sanitizedSaveName}.xlsx");
                Console.WriteLine($"Found mapping from '{res.Last().Key}' to '{res.Last().Value}'");
            }
            return res;
        }

        public static void MakeXlsxFromTxt(string inputPath, string outputPath)
        {
            //Read
            using var reader = new StreamReader(inputPath);
            CsvHelper.Configuration.CsvConfiguration myConfig = new(CultureInfo.InvariantCulture)
            {
                Delimiter = "\t",
                Encoding = Encoding.UTF8,
                HasHeaderRecord = false
            };
            using var tsvReader = new CsvReader(reader, myConfig);
            tsvReader.Read();
            using var wb = new XLWorkbook();
            var sheet = wb.AddWorksheet();
            int rowIndex = 1;
            while (tsvReader.Read())
            {
                for (int i = 0; tsvReader.TryGetField(i, out string? value); i++)
                {
                    sheet.Cell(rowIndex, i + 1).Value = value;
                }
                rowIndex++;
            }
            sheet.Columns().AdjustToContents();
            wb.SaveAs(outputPath);
        }
    }

    public class Options
    {
        [Option('p', "parent-folder", Required = true)]
        public string ParentFolder {get;set;}
        [Option('m', "map-file-filter", Required = false, Default = "^[0-9]*$")]
        public string MapFileFilter {get;set;}
        [Option('d', "data-file-filter", Required = false, Default = "^[0-9]*\\.txt$")]
        public string DataFileFilter {get;set;}
    }
}