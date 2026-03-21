using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Xml;
using Word = Microsoft.Office.Interop.Word;

class Program
{
    static string _outputDir;
    static string _logPath;

    static void Main()
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8;

        string rootFolder;
        while (true)
        {
            Console.WriteLine("Введите путь к корневой папке с картами: ");
            rootFolder = Console.ReadLine()?.Trim();
            if (!string.IsNullOrWhiteSpace(rootFolder) && Directory.Exists(rootFolder))
                break;
            Console.WriteLine("  Папка не найдена. Проверьте путь и попробуйте снова.");
        }

        _outputDir = Path.Combine(rootFolder, "Output");
        Directory.CreateDirectory(_outputDir);

        _logPath = Path.Combine(_outputDir, "errors.log");
        if (File.Exists(_logPath)) File.Delete(_logPath);

        ProcessFolder(rootFolder);

        Console.WriteLine("\nГотово!");

        if (File.Exists(_logPath))
            Console.WriteLine($"Некоторые файлы обработаны с ошибками. Подробности: {_logPath}");

        Stage2.Run(_outputDir);
        Console.WriteLine("\nВсе файлы обработаны и сохранены в папку Output.");
        Console.WriteLine("Нажмите любую клавишу для выхода...");
        Console.ReadKey();
    }
    //1 этап программы - ProcessFolder (конвертация, разделение, именование)
    static void ProcessFolder(string rootFolder)
    {
        var docFiles = Directory.GetFiles(rootFolder, "*.doc", SearchOption.AllDirectories);
        var docxFiles = Directory.GetFiles(rootFolder, "*.docx", SearchOption.AllDirectories);

        var allFiles = docFiles.Concat(docxFiles)
            .Where(f => !Path.GetFileName(f).StartsWith("~$"))
            .Where(f => !f.StartsWith(_outputDir, StringComparison.OrdinalIgnoreCase))
            .OrderBy(f => f)
            .ToList();

        if (allFiles.Count == 0)
        {
            Console.WriteLine("Файлы .doc/.docx не найдены.");
            return;
        }

        Word.Application wordApp = null;

        if (allFiles.Any(f => f.EndsWith(".doc", StringComparison.OrdinalIgnoreCase)))
        {
            wordApp = new Word.Application();
            wordApp.Visible = false;
        }

        int total = allFiles.Count;
        int done = 0;
        int errors = 0;

        Console.WriteLine($"Найдено файлов: {total}\n");

        foreach (var file in allFiles)
        {
            bool isDoc = file.EndsWith(".doc", StringComparison.OrdinalIgnoreCase);

            Console.WriteLine($"[{++done}/{total}] {Path.GetFileName(file)}");

            try
            {
                string docxPath;

                if (isDoc)
                {
                    string tempDocx =
                        Path.Combine(_outputDir,
                        Path.GetFileNameWithoutExtension(file) + "_converted.docx");

                    Word.Document doc = wordApp.Documents.Open(file);

                    doc.SaveAs2(tempDocx, Word.WdSaveFormat.wdFormatXMLDocument);
                    doc.Close();
                    Word.Document normalized = wordApp.Documents.Open(tempDocx);
                    normalized.Save();
                    normalized.Close();

                    Console.WriteLine("  Сконвертирован.");

                    docxPath = tempDocx;
                }
                else
                {
                    docxPath = file;
                }

                ProcessDocx(docxPath, isDoc);
            }
            catch (Exception ex)
            {
                errors++;

                string msg = $"[ОШИБКА] {file}\n{ex.Message}\n";

                Console.WriteLine(msg);

                LogError(msg);
            }
        }

        wordApp?.Quit();

        Console.WriteLine($"\nОбработано: {done}, ошибок: {errors}");
    }

    static void ProcessDocx(string docxPath, bool isDocConverted)
    {
        int cardCount = CountCards(docxPath);

        if (cardCount <= 1)
        {
            string xml = ReadDocumentXml(docxPath);

            string cardNumber = ExtractCardNumber(xml);

            string fileName =
                cardNumber != null
                ? $"карта_{cardNumber}.docx"
                : Path.GetFileNameWithoutExtension(docxPath).Replace("_converted", "") + ".docx";

            string dest = UniqueOutputPath(fileName);

            File.Copy(docxPath, dest, true);

            Console.WriteLine($"  → 1 карта → {Path.GetFileName(dest)}");

            if (isDocConverted)
                File.Delete(docxPath);

            return;
        }

        Console.WriteLine($"  → {cardCount} карт. Разделяю...");

        SplitDocx(docxPath, isDocConverted);
    }

    static string ExtractCardNumber(string xml)
    {
        if (xml == null) return null;

        string raw = null;

        //Формат А: <w:fldSimple w:instr="... DOCVARIABLE rm_number ...">
        Match mA = Regex.Match(xml,
            @"<w:fldSimple[^>]*DOCVARIABLE\s+rm_number[^>]*>(.*?)</w:fldSimple>",
            RegexOptions.Singleline | RegexOptions.IgnoreCase);

        if (mA.Success)
        {
            var parts = Regex.Matches(mA.Groups[1].Value, @"<w:t[^>]*>([^<]*)</w:t>");
            var sb = new System.Text.StringBuilder();
            foreach (Match p in parts) sb.Append(p.Groups[1].Value);
            raw = sb.ToString();
        }
        else
        {
            //Формат Б: развёрнутое поле fldChar
            Match mB = Regex.Match(xml,
                @"<w:instrText[^>]*>[^<]*DOCVARIABLE\s+rm_number[^<]*</w:instrText>" +
                @".*?<w:fldChar\s[^>]*w:fldCharType=""separate""[^/]*/>" +
                @"(.*?)" +
                @"<w:fldChar\s[^>]*w:fldCharType=""end""",
                RegexOptions.Singleline | RegexOptions.IgnoreCase);

            if (mB.Success)
            {
                var parts = Regex.Matches(mB.Groups[1].Value, @"<w:t[^>]*>([^<]*)</w:t>");
                var sb = new System.Text.StringBuilder();
                foreach (Match p in parts) sb.Append(p.Groups[1].Value);
                raw = sb.ToString();
            }
        }

        if (string.IsNullOrWhiteSpace(raw)) return null;

        //Очищаем для имени файла: оставляем буквы, цифры, дефисы; остальное -> "_"
        string clean = Regex.Replace(raw.Trim(), @"[^\w\d\-/]", "_").Trim('_');
        // "/" -> "-" (запрещён в именах файлов Windows)
        clean = clean.Replace("/", "-");
        //Убираем повторяющиеся "_"
        clean = Regex.Replace(clean, @"_+", "_").Trim('_');

        return string.IsNullOrEmpty(clean) ? null : clean;
    }

    static int CountCards(string docxPath)
    {
        string xml = ReadDocumentXml(docxPath);

        if (xml == null) return 1;

        XmlDocument doc = new XmlDocument();
        doc.LoadXml(xml);

        XmlNamespaceManager ns = BuildNs(doc);

        var nodes = doc.SelectNodes("//w:sectPr", ns);

        return nodes?.Count ?? 1;
    }

    static void SplitDocx(string docxPath, bool isDocConverted)
    {
        string xml = ReadDocumentXml(docxPath);

        if (xml == null) return;

        XmlDocument doc = new XmlDocument();
        doc.LoadXml(xml);

        XmlNamespaceManager ns = BuildNs(doc);

        XmlNode body = doc.SelectSingleNode("//w:body", ns);

        var groups = new List<List<XmlNode>>();
        var current = new List<XmlNode>();

        foreach (XmlNode node in body.ChildNodes)
        {
            current.Add(node);

            bool boundary =
                node.LocalName == "sectPr" ||
                (node.LocalName == "p" &&
                 node.SelectSingleNode(".//w:sectPr", ns) != null);

            if (boundary)
            {
                groups.Add(current);
                current = new List<XmlNode>();
            }
        }

        if (current.Count > 0)
            groups.Add(current);

        string baseName = Path.GetFileNameWithoutExtension(docxPath)
                      .Replace("_converted", "");

        for (int i = 0; i < groups.Count; i++)
        {
            if (i > 0 && groups[i].Count > 0)
            {
                var first = groups[i][0];
                if (first.LocalName == "p")
                {
                    bool hasRun = first.SelectSingleNode(".//w:r", ns) != null;
                    bool hasSectPr = first.SelectSingleNode(".//w:sectPr", ns) != null;
                    if (!hasRun && !hasSectPr)
                        groups[i].RemoveAt(0);
                }
            }

            string groupXml = string.Concat(groups[i].Select(n => n.OuterXml));

            string cardNumber = ExtractCardNumber(groupXml);

            string name =
                cardNumber != null
                ? $"карта_{cardNumber}.docx"
                : $"{baseName}_карта_{i + 1}.docx";

            string outPath = UniqueOutputPath(name);

            WriteCardDocx(docxPath, groups[i], doc, ns, outPath);

            Console.WriteLine("    Сохранён: " + Path.GetFileName(outPath));
        }

        if (isDocConverted)
            File.Delete(docxPath);
    }

    static void WriteCardDocx(string source, List<XmlNode> nodes,
    XmlDocument originalDoc, XmlNamespaceManager ns, string output)
    {
        File.Copy(source, output, true);

        using ZipArchive zip = ZipFile.Open(output, ZipArchiveMode.Update);

        var entry = zip.GetEntry("word/document.xml");
        if (entry == null) return;

        string docXml;
        using (var sr = new StreamReader(entry.Open()))
            docXml = sr.ReadToEnd();

        string newBodyContent = BuildBodyXml(nodes, ns);

        string newDocXml = Regex.Replace(docXml,
            @"(<w:body>)(.*?)(</w:body>)",
            m => m.Groups[1].Value + newBodyContent + m.Groups[3].Value,
            RegexOptions.Singleline);

        entry.Delete();
        using var sw = new StreamWriter(zip.CreateEntry("word/document.xml").Open());
        sw.Write(newDocXml);
    }

    static string BuildBodyXml(List<XmlNode> nodes, XmlNamespaceManager ns)
    {
        var result = new System.Text.StringBuilder();
        XmlNode pendingSectPr = null;

        foreach (XmlNode node in nodes)
        {
            //Случай А: <w:sectPr> прямо в <w:body> (последняя карта)
            if (node.LocalName == "sectPr")
            {
                result.Append(node.OuterXml);
                return result.ToString();
            }

            //Случай Б: параграф, внутри которого спрятан <w:sectPr>
            if (node.LocalName == "p")
            {
                XmlNode sectPr = node.SelectSingleNode("w:pPr/w:sectPr", ns);
                if (sectPr != null)
                {
                    pendingSectPr = sectPr;

                    XmlNode pPr = node.SelectSingleNode("w:pPr", ns);
                    pPr.RemoveChild(sectPr);

                    if (!pPr.HasChildNodes)
                        node.RemoveChild(pPr);

                    //Добавляем параграф только если в нём ещё что-то осталось
                    if (node.HasChildNodes)
                        result.Append(node.OuterXml);

                    continue;
                }
            }

            result.Append(node.OuterXml);
        }

        if (pendingSectPr != null)
            result.Append(pendingSectPr.OuterXml);

        return result.ToString();
    }

    static string ReadDocumentXml(string docxPath)
    {
        using ZipArchive zip = ZipFile.OpenRead(docxPath);

        var entry = zip.GetEntry("word/document.xml");

        using StreamReader sr = new StreamReader(entry.Open());

        return sr.ReadToEnd();
    }

    static XmlNamespaceManager BuildNs(XmlDocument doc)
    {
        var ns = new XmlNamespaceManager(doc.NameTable);

        ns.AddNamespace("w",
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

        return ns;
    }

    static string UniqueOutputPath(string fileName)
    {
        string path = Path.Combine(_outputDir, fileName);

        if (!File.Exists(path))
            return path;

        string name = Path.GetFileNameWithoutExtension(fileName);
        string ext = Path.GetExtension(fileName);

        int i = 2;

        while (File.Exists(path))
        {
            path = Path.Combine(_outputDir, $"{name}_{i}{ext}");
            i++;
        }

        return path;
    }

    static void LogError(string message)
    {
        File.AppendAllText(_logPath, message);
    }
}
//2 этап программы Stage2 (вставка подписей и дат, конвертация в PDF)
class Stage2
{

    public static void Run(string output)
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8;

        string sigDir;
        while (true)
        {
            Console.WriteLine("Папка с подписями PNG: ");
            sigDir = Console.ReadLine()?.Trim();
            if (!string.IsNullOrWhiteSpace(sigDir) && Directory.Exists(sigDir))
                break;
            Console.WriteLine("  Папка не найдена. Проверьте путь и попробуйте снова.");
        }

        string commissionDate;
        while (true)
        {
            Console.WriteLine("Дата комиссии (дд.мм.гггг):");
            string input = Console.ReadLine()?.Trim();
            if (DateTime.TryParseExact(input, "dd.MM.yyyy",
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out _))
            {
                commissionDate = input;
                break;
            }
            Console.WriteLine("  Неверный формат. Нужно дд.мм.гггг (например: 21.03.2026)");
        }

        //Поиск существующей даты эксперта по всем файлам
        var files = Directory.GetFiles(output, "карта_*.docx")
            .Where(f => !f.Contains("_signed")).ToArray();

        string foundExpertDate = null;
        foreach (var f in files)
        {
            string found = FindExpertDateInDoc(f);
            if (found != null) { foundExpertDate = found; break; }
        }

        //Определяем дату эксперта
        string expertDate;
        if (foundExpertDate != null)
        {
            Console.WriteLine($"  В документе найдена дата эксперта: {foundExpertDate}");

            string answer;
            while (true)
            {
                Console.WriteLine("  Оставить её? (д/н):");
                answer = Console.ReadLine()?.Trim().ToLower();
                if (answer == "д" || answer == "да" || answer == "н" || answer == "нет")
                    break;
                Console.WriteLine("  Некорректный ввод. Введите 'д' или 'н'.");
            }

            expertDate = (answer == "д" || answer == "да")
                ? foundExpertDate
                : AskDate("Введите новую дату эксперта (дд.мм.гггг):");
        }
        else
        {
            expertDate = AskDate("Дата эксперта (дд.мм.гггг):");
        }

        Word.Application word = new Word.Application();
        word.Visible = false;

        files = Directory.GetFiles(output, "карта_*.docx").Where(f => !f.Contains("_signed")).ToArray();

        foreach (var file in files)
        {
            Console.WriteLine("\nОбработка: " + Path.GetFileName(file));

            Word.Document doc = word.Documents.Open(file);

            ProcessTables(doc, sigDir, commissionDate, expertDate);

            string newDoc =
                Path.Combine(output,
                Path.GetFileNameWithoutExtension(file) + "_signed.docx");

            doc.SaveAs2(newDoc);
            doc.Save();
            doc.Close();

            //Конвертируем inline -> anchor в XML напрямую
            ConvertInlineToAnchor(newDoc);

            //Открываем исправленный файл и экспортируем PDF
            Word.Document docFixed = word.Documents.Open(newDoc);

            string pdf = Path.ChangeExtension(newDoc, ".pdf");

            docFixed.ExportAsFixedFormat(
                pdf,
                Word.WdExportFormat.wdExportFormatPDF
            );

            docFixed.Close(false);

            Console.WriteLine("Готово");
        }

        static void ConvertInlineToAnchor(string docxPath)
        {
            using var zip = System.IO.Compression.ZipFile.Open(
                docxPath,
                System.IO.Compression.ZipArchiveMode.Update);

            var entry = zip.GetEntry("word/document.xml");
            if (entry == null) return;

            string xml;
            using (var sr = new System.IO.StreamReader(entry.Open()))
                xml = sr.ReadToEnd();

            bool hasInline = xml.Contains("wp:inline");
            bool hasPict = xml.Contains("w:pict") && xml.Contains("v:shape");

            int replacements = 0;

            if (hasInline)
            {
                //ПУТЬ 1: docx исходник — современный DrawingML
                xml = System.Text.RegularExpressions.Regex.Replace(
                    xml,
                    @"<w:tc>(.*?)</w:tc>",
                    cellMatch =>
                    {
                        string cellContent = cellMatch.Groups[1].Value;
                        if (!cellContent.Contains("<wp:inline"))
                            return cellMatch.Value;

                        long cellWidthEmu = 0;
                        var tcwMatch = System.Text.RegularExpressions.Regex.Match(
                            cellContent, @"<w:tcW[^>]*w:w=""(\d+)""");
                        if (tcwMatch.Success &&
                            long.TryParse(tcwMatch.Groups[1].Value, out long twips))
                            cellWidthEmu = twips * 635L;

                        string replaced = System.Text.RegularExpressions.Regex.Replace(
                            cellContent,
                            @"<wp:inline\b[^>]*>(.*?)</wp:inline>",
                            im =>
                            {
                                string inner = im.Groups[1].Value;
                                var extMatch = System.Text.RegularExpressions.Regex.Match(
                                    inner, @"<wp:extent cx=""(\d+)"" cy=""(\d+)""");
                                long cx = extMatch.Success && long.TryParse(extMatch.Groups[1].Value, out long cxV) ? cxV : 914400L;
                                long cy = extMatch.Success && long.TryParse(extMatch.Groups[2].Value, out long cyV) ? cyV : 457200L;
                                long posH = cellWidthEmu > 0 ? Math.Max(0L, (cellWidthEmu - cx) / 2L) : 0L;
                                long posV = (-cy / 2) + 150000L; //Примерно центр ячейки (уйти от 150000L в следующей версии)

                                string innerClean = System.Text.RegularExpressions.Regex.Replace(inner, @"<wp:extent[^/]*/>", "");
                                innerClean = System.Text.RegularExpressions.Regex.Replace(innerClean, @"<wp:effectExtent[^/]*/>", "");

                                replacements++;
                                return
                                    $"<wp:anchor distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\" " +
                                    $"simplePos=\"0\" relativeHeight=\"0\" behindDoc=\"1\" locked=\"0\" " +
                                    $"layoutInCell=\"1\" allowOverlap=\"1\">" +
                                    $"<wp:simplePos x=\"0\" y=\"0\"/>" +
                                    $"<wp:positionH relativeFrom=\"column\"><wp:posOffset>{posH}</wp:posOffset></wp:positionH>" +
                                    $"<wp:positionV relativeFrom=\"paragraph\"><wp:posOffset>{posV}</wp:posOffset></wp:positionV>" +
                                    $"<wp:extent cx=\"{cx}\" cy=\"{cy}\"/>" +
                                    $"<wp:effectExtent l=\"0\" t=\"0\" r=\"0\" b=\"0\"/>" +
                                    $"<wp:wrapNone/>" +
                                    innerClean +
                                    $"</wp:anchor>";
                            },
                            System.Text.RegularExpressions.RegexOptions.Singleline
                        );
                        return $"<w:tc>{replaced}</w:tc>";
                    },
                    System.Text.RegularExpressions.RegexOptions.Singleline
                );
            }
            else if (hasPict)
            {
                //ПУТЬ 2: doc -> docx конвертированный — VML формат (w:pict/v:shape)
                //Добавляем namespace если отсутствует
                if (!xml.Contains("xmlns:wp="))
                    xml = xml.Replace("<w:document ",
                        "<w:document xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" ");
                if (!xml.Contains("xmlns:a="))
                    xml = xml.Replace("<w:document ",
                        "<w:document xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" ");

                xml = System.Text.RegularExpressions.Regex.Replace(
                    xml,
                    @"<w:tc>(.*?)</w:tc>",
                    cellMatch =>
                    {
                        string cellContent = cellMatch.Groups[1].Value;
                        if (!cellContent.Contains("w:pict"))
                            return cellMatch.Value;

                        long cellWidthEmu = 0;
                        var tcwMatch = System.Text.RegularExpressions.Regex.Match(
                            cellContent, @"<w:tcW[^>]*w:w=""(\d+)""");
                        if (tcwMatch.Success &&
                            long.TryParse(tcwMatch.Groups[1].Value, out long twips))
                            cellWidthEmu = twips * 635L;

                        string replaced = System.Text.RegularExpressions.Regex.Replace(
                            cellContent,
                            @"<w:pict>.*?<v:shape[^>]+style=""([^""]+)""[^>]*>.*?<v:imagedata r:id=""([^""]+)""[^/]*/>" +
                            @".*?</v:shape>.*?</w:pict>",
                            vm =>
                            {
                                string style = vm.Groups[1].Value;
                                string rId = vm.Groups[2].Value;

                                // Размеры из style="width:Xpt;height:Ypt" -> EMU (1pt = 12700)
                                long cx = 914400L, cy = 457200L;
                                var wM = System.Text.RegularExpressions.Regex.Match(style, @"width:([\d.]+)pt");
                                var hM = System.Text.RegularExpressions.Regex.Match(style, @"height:([\d.]+)pt");
                                if (wM.Success && double.TryParse(wM.Groups[1].Value,
                                    System.Globalization.NumberStyles.Float,
                                    System.Globalization.CultureInfo.InvariantCulture, out double wPt))
                                    cx = (long)(wPt * 12700);
                                if (hM.Success && double.TryParse(hM.Groups[1].Value,
                                    System.Globalization.NumberStyles.Float,
                                    System.Globalization.CultureInfo.InvariantCulture, out double hPt))
                                    cy = (long)(hPt * 12700);

                                long posH = cellWidthEmu > 0 ? Math.Max(0L, (cellWidthEmu - cx) / 2L) : 0L;
                                long posV = (-cy / 2) + 150000L;

                                replacements++;

                                return
                                    $"<w:drawing>" +
                                    $"<wp:anchor distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\" " +
                                    $"simplePos=\"0\" relativeHeight=\"0\" behindDoc=\"1\" locked=\"0\" " +
                                    $"layoutInCell=\"1\" allowOverlap=\"1\">" +
                                    $"<wp:simplePos x=\"0\" y=\"0\"/>" +
                                    $"<wp:positionH relativeFrom=\"column\"><wp:posOffset>{posH}</wp:posOffset></wp:positionH>" +
                                    $"<wp:positionV relativeFrom=\"paragraph\"><wp:posOffset>{posV}</wp:posOffset></wp:positionV>" +
                                    $"<wp:extent cx=\"{cx}\" cy=\"{cy}\"/>" +
                                    $"<wp:effectExtent l=\"0\" t=\"0\" r=\"0\" b=\"0\"/>" +
                                    $"<wp:wrapNone/>" +
                                    $"<wp:docPr id=\"{replacements}\" name=\"Подпись {replacements}\"/>" +
                                    $"<wp:cNvGraphicFramePr/>" +
                                    $"<a:graphic>" +
                                    $"<a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
                                    $"<pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
                                    $"<pic:nvPicPr>" +
                                    $"<pic:cNvPr id=\"{replacements}\" name=\"Подпись {replacements}\"/>" +
                                    $"<pic:cNvPicPr/>" +
                                    $"</pic:nvPicPr>" +
                                    $"<pic:blipFill>" +
                                    $"<a:blip r:embed=\"{rId}\"/>" +
                                    $"<a:stretch><a:fillRect/></a:stretch>" +
                                    $"</pic:blipFill>" +
                                    $"<pic:spPr>" +
                                    $"<a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{cx}\" cy=\"{cy}\"/></a:xfrm>" +
                                    $"<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>" +
                                    $"</pic:spPr>" +
                                    $"</pic:pic>" +
                                    $"</a:graphicData>" +
                                    $"</a:graphic>" +
                                    $"</wp:anchor>" +
                                    $"</w:drawing>";
                            },
                            System.Text.RegularExpressions.RegexOptions.Singleline
                        );
                        return $"<w:tc>{replaced}</w:tc>";
                    },
                    System.Text.RegularExpressions.RegexOptions.Singleline
                );
            }
            //Удаляем пустой абзац непосредственно перед абзацем с sectPr
            //Он создаёт лишнюю страницу при экспорте в PDF
            xml = System.Text.RegularExpressions.Regex.Replace(
                xml,
                @"(<w:p\b[^>]*/>\s*)(<w:p\b[^>]*><w:pPr><w:sectPr\b)",
                "$2",
                System.Text.RegularExpressions.RegexOptions.Singleline
            );

            //Удаляем лишний sectPr добавленный Word при SaveAs2
            xml = System.Text.RegularExpressions.Regex.Replace(
                xml,
                @"<w:p\b[^>]*\bw:rsidR=""00000000""[^>]*/>\s*<w:sectPr\b.*?</w:sectPr>\s*</w:body>",
                "</w:body>",
                System.Text.RegularExpressions.RegexOptions.Singleline
            );

            entry.Delete();
            using var sw = new System.IO.StreamWriter(
                zip.CreateEntry("word/document.xml").Open());
            sw.Write(xml);
        }

        word.Quit();

    }
    //Запрашивает дату с валидацией формата dd.MM.yyyy
    static string AskDate(string prompt)
    {
        while (true)
        {
            Console.WriteLine(prompt);
            string input = Console.ReadLine()?.Trim();
            if (DateTime.TryParseExact(input, "dd.MM.yyyy",
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out _))
                return input;
            Console.WriteLine("  Неверный формат. Нужно дд.мм.гггг (например: 21.03.2026)");
        }
    }

    //Ищет существующую дату в ячейке эксперта через Word Interop
    //Возвращает строку даты если нашёл, иначе null
    //Ищет дату в ячейке эксперта (строка данных, 7-я ячейка таблицы эксперта).
    //Таблица эксперта — та, которой предшествует абзац с текстом "Эксперт (эксперты)".
    //Возвращает строку даты dd.MM.yyyy если нашёл, иначе null.
    static string FindExpertDateInDoc(string docxPath)
    {
        try
        {
            using var zip = System.IO.Compression.ZipFile.OpenRead(docxPath);
            var entry = zip.GetEntry("word/document.xml");
            if (entry == null) return null;

            string xml;
            using (var sr = new System.IO.StreamReader(entry.Open()))
                xml = sr.ReadToEnd();

            //Ищем таблицу по её внутреннему признаку —
            //строка подписей содержит "(№ в реестре экспертов)" в ячейке 1.
            //Текст перед таблицей ненадёжен: "Эксперт (эксперты)" разбит по нескольким <w:r>.
            var tables = System.Text.RegularExpressions.Regex.Matches(
                xml, @"<w:tbl\b.*?</w:tbl>",
                System.Text.RegularExpressions.RegexOptions.Singleline);

            foreach (System.Text.RegularExpressions.Match tblMatch in tables)
            {
                string tbl = tblMatch.Value;

                //Таблица эксперта — та, где есть "реестре экспертов"
                if (!tbl.Contains("реестре экспертов") && !tbl.Contains("реестр"))
                    continue;

                //Берём первую строку (строка данных)
                var rowMatch = System.Text.RegularExpressions.Regex.Match(
                    tbl, @"<w:tr\b.*?</w:tr>",
                    System.Text.RegularExpressions.RegexOptions.Singleline);

                if (!rowMatch.Success) continue;

                var cells = System.Text.RegularExpressions.Regex.Matches(
                    rowMatch.Value, @"<w:tc>(.*?)</w:tc>",
                    System.Text.RegularExpressions.RegexOptions.Singleline);

                if (cells.Count < 7) continue;

                //Ячейка 7 (индекс 6) — дата эксперта
                string cellXml = cells[6].Groups[1].Value;
                var texts = System.Text.RegularExpressions.Regex.Matches(
                    cellXml, @"<w:t[^>]*>([^<]*)</w:t>");

                var sb = new System.Text.StringBuilder();
                foreach (System.Text.RegularExpressions.Match t in texts)
                    sb.Append(t.Groups[1].Value);

                string candidate = sb.ToString().Trim();

                if (DateTime.TryParseExact(candidate, "dd.MM.yyyy",
                        System.Globalization.CultureInfo.InvariantCulture,
                        System.Globalization.DateTimeStyles.None, out _))
                    return candidate;
            }
        }
        catch { }

        return null;
    }
        static bool LooksLikeFio(string text)
    {
        //Три слова подряд, каждое начинается с заглавной кириллической буквы
        return Regex.IsMatch(text, @"[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+");
    }

    //Определяет роль по тексту абзаца перед таблицей (scoring)
    //Возвращает "commission", "expert", "worker" или "unknown"
    static string DetectRoleByContext(string contextText)
    {
        string t = contextText.ToLower();

        int scoreCommission = 0;
        int scoreExpert = 0;
        int scoreWorker = 0;

        // Комиссия
        if (t.Contains("председатель")) scoreCommission += 2;
        if (t.Contains("член комиссии")) scoreCommission += 2;
        if (t.Contains("члены комиссии")) scoreCommission += 2;
        if (t.Contains("комиссии")) scoreCommission += 1;
        if (t.Contains("комиссия")) scoreCommission += 1;

        // Эксперт
        if (t.Contains("эксперт")) scoreExpert += 2;
        if (t.Contains("эксперты")) scoreExpert += 1; 
        if (t.Contains("(эксперты)")) scoreExpert += 1;

        // Работник
        if (t.Contains("ознакомлен")) scoreWorker += 2;   //покрывает все формы?
        if (t.Contains("работник")) scoreWorker += 1;

        int max = Math.Max(scoreCommission, Math.Max(scoreExpert, scoreWorker));
        if (max == 0) return "unknown";
        if (scoreExpert == max) return "expert";
        if (scoreWorker == max) return "worker";
        return "commission";
    }

    static void ProcessTables(
    Word.Document doc,
    string sigDir,
    string commissionDate,
    string expertDate)
    {
        foreach (Word.Table tbl in doc.Tables)
        {
            //Определяем роль таблицы по тексту абзаца перед ней
            //Контекст берём из Range перед таблицей: до 3 абзацев назад (пересмотреть...)
            string tableContext = "";
            try
            {
                Word.Range before = tbl.Range;
                before.MoveStart(Word.WdUnits.wdParagraph, -3);
                before.MoveEnd(Word.WdUnits.wdParagraph, -3);
                tableContext = before.Text ?? "";
            }
            catch { }

            string tableRole = DetectRoleByContext(tableContext);

            //Таблицы без распознанной роли (заголовки, данные) — пропускаем
            if (tableRole == "unknown") continue;
            if (tableRole == "worker") continue;

            int rows = tbl.Rows.Count;

            for (int r = 1; r <= rows; r++)
            {
                Word.Row row;
                try { row = tbl.Rows[r]; }
                catch { continue; }

                string rowText = row.Range.Text
                    .Replace("\r", "").Replace("\a", "").Trim();

                //Ищем строку подписей: содержит "(подпись)"
                if (!rowText.ToLower().Contains("(подпись)")) continue;

                //Строка подписей найдена — строка данных стоит выше
                if (row.Index <= 1) continue;

                Word.Row dataRow;
                try { dataRow = tbl.Rows[row.Index - 1]; }
                catch { continue; }

                string dataText = dataRow.Range.Text
                    .Replace("\r", "").Replace("\a", "").Trim();

                //Проверяем что в строке данных есть ФИО
                if (!LooksLikeFio(dataText)) continue;

                //Извлекаем ФИО для поиска файла подписи
                Match fioMatch = Regex.Match(
                    dataText,
                    @"[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+\s+[А-ЯЁ][а-яё]+");
                string fio = fioMatch.Success ? fioMatch.Value.Trim() : "";

                if (string.IsNullOrWhiteSpace(fio)) continue;

                //Ищем колонку "(подпись)" в строке подписей
                int signColumn = 3; // fallback (нужен ли fallback? Может оставить только логирование...)
                for (int c = 1; c <= row.Cells.Count; c++)
                {
                    try
                    {
                        string cellText = row.Cells[c].Range.Text
                            .Replace("\r", "").Replace("\a", "").Trim().ToLower();
                        if (cellText.Contains("(подпись)"))
                        {
                            signColumn = c;
                            break;
                        }
                    }
                    catch { }
                }

                //Проверяем наличие файла подписи
                string signPath = Path.Combine(sigDir, fio + ".png");
                if (!File.Exists(signPath))
                {
                    Console.WriteLine($"  Нет подписи: {fio}");
                    continue;
                }

                //Вставляем подпись и дату
                InsertSignature(dataRow.Cells[signColumn].Range, signPath);

                //Дата: эксперту — expertDate, остальным — commissionDate
                //Колонка даты — следующая после (дата) в строке подписей
                int dateColumn = signColumn; // fallback — та же колонка (см. комментарий к signColumn)
                for (int c = 1; c <= row.Cells.Count; c++)
                {
                    try
                    {
                        string cellText = row.Cells[c].Range.Text
                            .Replace("\r", "").Replace("\a", "").Trim().ToLower();
                        if (cellText.Contains("(дата)"))
                        {
                            dateColumn = c;
                            break;
                        }
                    }
                    catch { }
                }

                string dateToInsert = tableRole == "expert" ? expertDate : commissionDate;
                try { dataRow.Cells[dateColumn].Range.Text = dateToInsert; }
                catch { }

                Console.WriteLine($"  Подписано: {fio}");
            }
        }
    }

    static void InsertSignature(
    Word.Range range,
    string img)
    {
        range.Text = "";

        //Выравнивание по левому краю (?)
        range.ParagraphFormat.Alignment =
            Word.WdParagraphAlignment.wdAlignParagraphLeft;

        //Вставляем в оригинальном размере — без ограничений
        range.InlineShapes.AddPicture(
            FileName: img,
            LinkToFile: false,
            SaveWithDocument: true,
            Range: range
        );
    }

}