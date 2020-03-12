using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelToJson
{
    enum ParsingStep
    {
        TypeParsing,
        MemberNameParsing,
        DataParsing
    }

    class TypeInfo
    {
        public string TypeStr;
    }

    class ParsingContext
    {
        public ParsingStep step = ParsingStep.TypeParsing;
        public string className;
        public TypeInfo[] types;
        public string[] members;
        public string[,] dataSet;
    }

    class Program
    {
        static Dictionary<string, List<string>> _options = new Dictionary<string, List<string>>();

        static void ParseArgs(string[] args)
        {
            var currentOption = string.Empty;

            for (var i = 0; i < args.Length; ++i)
            {
                var arg = args[i];

                if (arg.StartsWith("-"))
                {
                    if (_options.ContainsKey(arg) == false)
                    {
                        _options.Add(arg, new List<string>());
                        currentOption = arg;
                    }
                }
                else
                {
                    _options[currentOption].Add(arg);
                }
            }
        }

        static void Main(string[] args)
        {
            var excelPath = string.Empty;
            var template = string.Empty;
            var dataOutputPath = string.Empty;
            var classOutputPath = string.Empty;

            try
            {
                ParseArgs(args);

                excelPath = _options["-e"][0];
                template = File.ReadAllText(_options["-t"][0]);
                dataOutputPath = _options["-d"][0];
                classOutputPath = _options["-c"][0];
            }
            catch (Exception e)
            {
                Console.WriteLine("[Usage]");
                Console.WriteLine("ExcelToJson.exe -e .\\ -t .\\template.tml -d .\\ -c .\\");
                Console.WriteLine("-e   Path where the excel(*.xlsx) file exists");
                Console.WriteLine("-t   Template file path");
                Console.WriteLine("-d   Path to where json file is extracted");
                Console.WriteLine("-c   Path to where source file is extracted");
                Console.WriteLine(e);
                return;
            }

            Directory.CreateDirectory(dataOutputPath);

            var excelFiles = Directory.GetFiles(excelPath, "*.xlsx");

            var contexts = new List<ParsingContext>();

            for (var i = 0; i < excelFiles.Length; ++i)
            {
                var bin = File.ReadAllBytes(excelFiles[i]);

                using (var stream = new MemoryStream(bin))
                using (var excelPackage = new ExcelPackage(stream))
                {
                    foreach (var worksheet in excelPackage.Workbook.Worksheets)
                    {
                        if (worksheet.Name.StartsWith("_"))
                            continue;

                        if (worksheet.Dimension == null)
                            continue;

                        var startRow = worksheet.Dimension.Start.Row;
                        var endRow = worksheet.Dimension.End.Row;
                        var startColumn = worksheet.Dimension.Start.Column;
                        var endColumn = worksheet.Dimension.End.Column;

                        var rowCount = endRow - startRow + 1;
                        var columnCount = endColumn - startColumn + 1;

                        var ctx = new ParsingContext
                        {
                            className = worksheet.Name,
                            types = new TypeInfo[columnCount],
                            members = new string[columnCount],
                            dataSet = new string[rowCount, columnCount]
                        };

                        contexts.Add(ctx);

                        for (var r = startRow; r <= endRow; r++)
                        {
                            for (var c = startColumn; c <= endColumn; c++)
                            {
                                var val = worksheet.Cells[r, c].Value;
                                var dr = r - startRow;
                                var dc = c - startColumn;

                                if (val == null && ctx.step == ParsingStep.TypeParsing)
                                {
                                    val = ":";
                                }

                                if (val != null)
                                {
                                    switch (ctx.step)
                                    {
                                        case ParsingStep.TypeParsing:
                                            var typeInfo = new TypeInfo
                                            {
                                                TypeStr = val.ToString()
                                            };

                                            ctx.types[dc] = typeInfo;

                                            if (c == worksheet.Dimension.End.Column)
                                                ctx.step = ParsingStep.MemberNameParsing;
                                            break;

                                        case ParsingStep.MemberNameParsing:
                                            ctx.members[dc] = val.ToString();
                                            if (c == worksheet.Dimension.End.Column)
                                                ctx.step = ParsingStep.DataParsing;
                                            break;

                                        case ParsingStep.DataParsing:
                                            ctx.dataSet[dr, dc] = val.ToString();
                                            break;
                                    }
                                }
                                else
                                {
                                    Console.WriteLine($"value is null {r}:{c}");
                                }
                            }
                        }

                        ExportJson(ctx, dataOutputPath);
                        ExportClass(ctx, template, classOutputPath);
                        Console.WriteLine($"{ctx.className} Pasred.");
                    }
                }
            }

            Console.WriteLine("Complete");
        }

        static void ExportJson(ParsingContext ctx, string outputPath)
        {
            if (ctx.dataSet.Length < 2)
            {
                return;
            }

            var ret = new JArray();

            for (var i = 2; i < ctx.dataSet.GetLength(0); ++i)
            {
                var obj = new JObject();

                for (var j = 0; j < ctx.dataSet.GetLength(1); ++j)
                {
                    var typeInfo = ctx.types[j];
                    var raw = ctx.dataSet[i, j];

                    var val = default(JToken);

                    switch (typeInfo.TypeStr)
                    {
                        case "int": val = int.Parse(raw); break;
                        case "float": val = float.Parse(raw); break;
                        default: val = raw; break;
                    }

                    obj.Add(ctx.members[j], val);
                }

                ret.Add(obj);
            }

            File.WriteAllText(Path.Combine(outputPath, $"{ctx.className}.json"), ret.ToString());
        }

        static void ExportClass(ParsingContext ctx, string template, string classPath)
        {
            //
            // $EXTENSION(cs)
            // $SHEET_NAME
            // $PROPERTY_LOOP[$PROPERTY_TYPE $PROPERTY_NAME]
            // $PROPERTY_SET_LOOP[$PROPERTY_NAME $PROPERTY_NAME];
            //

            // extract extension
            var extension = string.Empty;
            (template, extension) = ExtractExtension(template);

            // set class name
            template = template.Replace("$SHEET_NAME", ctx.className);

            // property loop
            template = ExtractPropertyPattern(template, ctx);

            // property set loop
            template = ExtractPropertySetPattern(template, ctx);

            File.WriteAllText(Path.Combine(classPath, $"{ctx.className}.{extension}"), template);
        }

        static (string, string) ExtractExtension(string template)
        {
            var extensionRegex = new Regex(@"\$EXTENSION\((.+?)\)", RegexOptions.Compiled);

            var matches = extensionRegex.Matches(template);

            foreach (Match match in matches)
            {
                var groups = match.Groups;

                if (groups.Count > 1)
                {
                    return (template.Replace(groups[0].Value, ""), groups[1].Value);
                }
            }

            return (string.Empty, string.Empty);
        }

        static string ExtractPropertyPattern(string template, ParsingContext ctx)
        {
            var propLoopRegex = new Regex(@"(\t.+)\$PROPERTY_LOOP\[(.+]?)\]", RegexOptions.Compiled);

            var repl = string.Empty;
            var indent = string.Empty;
            var pattern = string.Empty;

            var matches = propLoopRegex.Matches(template);

            foreach (Match match in matches)
            {
                var groups = match.Groups;

                if (groups.Count > 2)
                {
                    repl = groups[0].Value;
                    indent = groups[1].Value;
                    pattern = groups[2].Value;
                    break;
                }
            }

            var propLoopSb = new StringBuilder();

            for (var i = 0; i < ctx.members.Length; ++i)
            {
                propLoopSb.Append(indent);

                var val = pattern.Replace("$PROPERTY_TYPE", ctx.types[i].TypeStr).Replace("$PROPERTY_NAME", ctx.members[i]);

                if (i < ctx.members.Length - 1)
                {
                    propLoopSb.AppendLine(val);
                }
                else
                {
                    propLoopSb.Append(val);
                }
            }

            return template.Replace(repl, propLoopSb.ToString());
        }

        static string ExtractPropertySetPattern(string template, ParsingContext ctx)
        {
            var defPropSetLoopRegex = new Regex(@"(\t.+)\$PROPERTY_SET_LOOP\[(.+]?)\]", RegexOptions.Compiled);
            var propSetLoopRegex = new Regex(@"(\t.+)\$PROPERTY_SET_LOOP\((.+?)\)\[(.+]?)\]", RegexOptions.Compiled);

            var groupOfType = new Dictionary<string, GroupCollection>();

            var matches = propSetLoopRegex.Matches(template);

            foreach (Match match in matches)
            {
                var groups = match.Groups;

                if (groups.Count > 3)
                {
                    var type = groups[2].Value;
                    groupOfType.Add(type, groups);
                }
            }

            matches = defPropSetLoopRegex.Matches(template);

            foreach (Match match in matches)
            {
                var groups = match.Groups;

                if (groups.Count > 2)
                {
                    groupOfType.Add("@___DEF___", groups);
                    break;
                }
            }

            var replOfType = new Dictionary<string, StringBuilder>();

            for (var i = 0; i < ctx.members.Length; ++i)
            {
                var type = ctx.types[i].TypeStr;

                var repl = string.Empty;
                var indent = string.Empty;
                var pattern = string.Empty;

                if (groupOfType.ContainsKey(type))
                {
                    var groups = groupOfType[type];
                    repl = groups[0].Value;
                    indent = groups[1].Value;
                    pattern = groups[3].Value;
                }
                else
                {
                    var groups = groupOfType["@___DEF___"];
                    repl = groups[0].Value;
                    indent = groups[1].Value;
                    pattern = groups[2].Value;
                }

                if (replOfType.ContainsKey(repl) == false)
                {
                    replOfType.Add(repl, new StringBuilder());
                }

                var sb = replOfType[repl];

                sb.Append(indent);

                var val = pattern.Replace("$PROPERTY_TYPE", type).Replace("$PROPERTY_NAME", ctx.members[i]);

                if (i < ctx.members.Length - 1)
                {
                    sb.AppendLine(val);
                }
                else
                {
                    sb.Append(val);
                }
            }

            foreach (var p in replOfType)
            {
                template = template.Replace(p.Key, p.Value.ToString());
            }

            // cleanup
            foreach (var p in groupOfType)
            {
                template = template.Replace(p.Value[0].Value, "");
            }

            return template;
        }
    }
}
