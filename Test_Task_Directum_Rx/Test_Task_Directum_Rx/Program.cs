using System;
using System.IO;
using System.IO.Compression;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using static System.Net.WebRequestMethods;
using System.Globalization;
using System.Reflection;
using IronWord;
using IronWord.Models;
using System.Runtime.InteropServices.JavaScript;
using IronWord.Models.Enums;
using System.Reflection.Metadata;
using System.Diagnostics;
using System.Linq;

namespace Test_Task_Directum_RX
{
    class Program
    {
        static async Task Main()
        {
            //Логирование процесса работы программы
            TextWriterTraceListener tr1 = new TextWriterTraceListener(System.IO.File.CreateText("log.txt"));
            Trace.Listeners.Add(tr1);
            Trace.AutoFlush = true;

            //Коллекции для хранения данных для отчёта
            List<AdrObject> adrObjectsList = new List<AdrObject>();
            List<AdrObjectLevel> adrObjectsLevelsList = new List<AdrObjectLevel>();
            List<String> tempFiles = new List<String>();

            //Путь к ФИАС
            string lastDownloadFileInfoUrl = "https://fias.nalog.ru/WebServices/Public/DownloadService.asmx/GetLastDownloadFileInfo";

            // Шаг 1: Получение последней ссылки garXMLDeltaURL
            var garXMLDeltaURL = await GetLatestGarXMLDeltaURLAsync(lastDownloadFileInfoUrl);

            // Шаг 2: Скачать архив gar_delta_xml
            string zipFilePath = await DownloadPackageAsync(garXMLDeltaURL);
            tempFiles.Add(zipFilePath);

            // Шаг 3: Разархивировать архив gar_delta_xml
            string extractPath = ExtractPackage(zipFilePath);
            tempFiles.Add(extractPath);

            // Шаг 4: Получение добавленных значений
            ProcessFiles(extractPath, adrObjectsList);
            
            // Получение значения уровней
            GetAsObjectLevels(extractPath, adrObjectsLevelsList);

            // Получение даты изменений
            var date = GetDate(extractPath);

            // Шаг 5: Создание таблиц
            var tables = CreateTables(adrObjectsList, adrObjectsLevelsList);

            // Шаг 6: Формирование отчёта
            var path = BuildDocument(tables, date);

            // Шаг 7: Удаление временных файлов
            DeleteTemporaryFiles(tempFiles);

            // Шаг 8: Открытие полученного отчёта
            Trace.WriteLine("Шаг 8: Process.Start.");
            try
            {
                Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                Trace.WriteLine("$Шаг 8: Process.Start. {ex.Message}");
            }
        }

        static async Task<string> GetLatestGarXMLDeltaURLAsync(string latestUrl)
        {
            Trace.WriteLine("Шаг 1: Метод - GetLatestGarXMLDeltaURLAsync.");
            string garXmlDeltaUrl = "";

            using (var client = new HttpClient())
            {
                using HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, latestUrl);

                using HttpResponseMessage response = await client.SendAsync(request);
                try
                {
                    response.EnsureSuccessStatusCode();

                    var xml = await response.Content.ReadAsStringAsync();

                    var doc = XDocument.Parse(xml);

                    var xmlDoc = new XmlDocument();
                    xmlDoc.LoadXml(xml);

                    var node = xmlDoc.SelectSingleNode("//*[local-name()='GarXMLDeltaURL']");
                    if (node != null)
                    {
                        garXmlDeltaUrl = node.InnerText;
                    }
                    else
                    {
                        Trace.WriteLine("Шаг 1: Узел не найден.");
                    }
                }
                catch (Exception ex)
                {
                    Trace.WriteLine($"Шаг 1: {ex.Message}");
                    throw;
                }
            }
            Trace.WriteLine("Шаг 1: GetLatestGarXMLDeltaURLAsync. Завершено успешно.");
            return garXmlDeltaUrl;
        }

        static async Task<string> DownloadPackageAsync(string garUrl)
        {
            Trace.WriteLine("Шаг 2: Метод - DownloadPackageAsync.");
            string zipFilePath = "";

            using (var client = new HttpClient())
            {
                using HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, garUrl);

                using HttpResponseMessage response = await client.SendAsync(request);
                try
                {
                    response.EnsureSuccessStatusCode();

                    var zipData = response.Content.ReadAsByteArrayAsync().Result;

                    zipFilePath = Path.Combine(Directory.GetCurrentDirectory(), "gar_delta_xml.zip");
                    System.IO.File.WriteAllBytes(zipFilePath, zipData);
                }
                catch (Exception ex)
                {
                    Trace.WriteLine($"Шаг 2: {ex.Message}");
                    throw;
                }
            }
            Trace.WriteLine("Шаг 2: DownloadPackageAsync. Завершено успешно.");
            return zipFilePath;
        }

        static string ExtractPackage(string zipFilePath)
        {
            Trace.WriteLine("Шаг 3: Метод - ExtractPackage.");
            string extractPath = Path.Combine(Directory.GetCurrentDirectory(), "extracted");
            
            try
            {
                if (!Directory.Exists(extractPath))
                {
                    Directory.CreateDirectory(extractPath);
                }
                ZipFile.ExtractToDirectory(zipFilePath, extractPath);
            }
            catch (Exception ex) 
            {
                Trace.WriteLine($"Шаг 3: {ex.Message}");
                throw;
            }

            Trace.WriteLine("Шаг 3: ExtractPackage. Завершено успешно.");
            return extractPath;
        }

        static void ProcessFiles(string extractPath, List<AdrObject> adrObjectsList)
        {
            Trace.WriteLine("Шаг 4: Метод - ProcessFiles.");
            foreach (var dir in Directory.GetDirectories(extractPath))
            {
                foreach (var file in Directory.GetFiles(dir, "AS_ADDR_OBJ*.xml"))
                {
                    ProcessFile(file, adrObjectsList);
                }
            }
            Trace.WriteLine("Шаг 4: ProcessFiles. Завершено успешно.");
        }

        static void ProcessFile(string filePath, List<AdrObject> adrObjectsList)
        {
            Trace.WriteLine("Шаг 4: Метод - ProcessFile.");
            XmlDocument xmlDoc = new XmlDocument();
            try
            {
                xmlDoc.Load(filePath);
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"Шаг 4: {ex.Message}");
                throw;
            }
            adrObjectsList.AddRange(from XmlNode node in xmlDoc.SelectNodes("//OBJECT")
                                    let objectType = node.Attributes["TYPENAME"].Value
                                    let objectName = node.Attributes["NAME"].Value
                                    let level = int.Parse(node.Attributes["LEVEL"].Value)
                                    let isActive = int.Parse(node.Attributes["ISACTIVE"].Value)
                                    where isActive == 1
                                    let adrObject = new AdrObject(level, objectType, objectName)
                                    select adrObject);
            Trace.WriteLine("Шаг 4: ProcessFile. Завершено успешно.");
        }

        static void GetAsObjectLevels(string extractPath, List<AdrObjectLevel> adrObjectsLevelsList)
        {
            Trace.WriteLine("Шаг 4: Метод - GetAsObjectLevels.");

                foreach (var file in Directory.GetFiles(extractPath, "AS_OBJECT_LEVELS*.xml"))
                {
                    XmlDocument xmlDoc = new XmlDocument();
                    try
                    {
                        xmlDoc.Load(file);
                    }
                    catch (Exception ex)
                    {
                        Trace.WriteLine($"Шаг 4: {ex.Message}");
                        throw;
                    }
                adrObjectsLevelsList.AddRange(from XmlNode node in xmlDoc.SelectNodes("//OBJECTLEVEL")
                                              let objectName = node.Attributes["NAME"].Value
                                              let level = int.Parse(node.Attributes["LEVEL"].Value)
                                              let isActive = bool.Parse(node.Attributes["ISACTIVE"].Value)
                                              where isActive
                                              let adrLevel = new AdrObjectLevel(level, objectName)
                                              select adrLevel);
            }

            Trace.WriteLine("Шаг 4: GetAsObjectLevels. Завершено успешно.");
        }

        static string GetDate(string extractPath)
        {
            string date;
            Trace.WriteLine("Шаг 4: Метод - GetDate.");
            try
            {
                using (StreamReader sr = new StreamReader(Path.Combine(extractPath, "version.txt")))
                {
                    date = sr.ReadLine();

                }
                Trace.WriteLine("Шаг 4: GetDate. Завершено успешно.");
                return date;
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"Шаг 4: {ex.Message}");
                throw;
            }
        }

        static Dictionary<AdrObjectLevel, List<AdrObject>> CreateTables(List<AdrObject> adrObjectsList, List<AdrObjectLevel> adrObjectsLevelsList)
        {
            Trace.WriteLine("Шаг 5: CreateTables.");
            var tableDictionary = new Dictionary<AdrObjectLevel, List<AdrObject>>();
            foreach (var adrObjectLevel in adrObjectsLevelsList)
            {
                var level = adrObjectLevel.Level;
                var adrs = adrObjectsList
                        .Where(x => x.Level == level)
                        .OrderBy(x => x.Name, StringComparer.Create(CultureInfo.CreateSpecificCulture("ru-RU"), true))
                        .ToList();
                if (adrs.Count > 0 && level < 9)
                {
                    tableDictionary.Add(adrObjectLevel, adrs);
                }
            }
            Trace.WriteLine("Шаг 5: CreateTables. Завершено успешно.");
            return tableDictionary;
        }

        static string BuildDocument(Dictionary<AdrObjectLevel, List<AdrObject>> tableDictionary, string date)
        {
            Trace.WriteLine("Шаг 6: BuildDocument.");
            var docx_new = new WordDocument();

            Font font = new()
            {
                FontFamily = "Calibri",
                FontSize = 32
            };

            Font tableFont = new()
            {
                FontFamily = "Arial",
                FontSize = 24
            };

            TextStyle headStyle = new()
            {
                Color = Color.Black,
                TextFont = font,
                IsBold = true,
                IsItalic = false
            };

            TextStyle tableStyle = new()
            {
                Color = Color.Black,
                TextFont = tableFont,
                IsBold = true,
                IsItalic = false
            };

            Text header = new($"Отчёт по добавленным адресным объектам за {date}") { Style = headStyle  };

            Paragraph paragraph = new();
            paragraph.AddText(header);
            docx_new.AddParagraph(paragraph);


            BorderStyle borderStyle = new BorderStyle();
            borderStyle.BorderColor = Color.Black;
            borderStyle.BorderValue = BorderValues.Single;
            borderStyle.BorderSize = 2;

            TableBorders tableBorders = new TableBorders()
            {
                TopBorder = borderStyle,
                RightBorder = borderStyle,
                BottomBorder = borderStyle,
                LeftBorder = borderStyle,
            };

            foreach (var table in tableDictionary)
            {
                docx_new.AddParagraph(new Paragraph());

                Paragraph tableParagraph = new();
                Text tableHeader = new(table.Key.Name) { Style = headStyle };
                tableParagraph.AddText(tableHeader);
                docx_new.AddParagraph(tableParagraph);

                Table wordTable = new Table();
                wordTable.Zebra = new ZebraColor("EEEEEE", "CCCCCC");
                wordTable.Borders = tableBorders;

                TableRow zeroRow = new TableRow();
                zeroRow.AddCell(new TableCell(new Text("Тип") { Style = tableStyle }));
                zeroRow.AddCell(new TableCell(new Text("Наименование") { Style = tableStyle }));
                wordTable.AddRow(zeroRow);
               
                foreach (var adr in table.Value)
                {
                    TableRow row = new TableRow();
                    row.AddCell(new TableCell(new Text(adr.Type) { Style = tableStyle }));
                    row.AddCell(new TableCell(new Text(adr.Name) { Style = tableStyle }));
                    wordTable.AddRow(row);
                }
                docx_new.AddTable(wordTable);
            }
            var savingPath = $"Отчёт_по_добавленным_адресам_за_{date}.docx";

            try
            {
                docx_new.SaveAs(savingPath);
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"Шаг 6: {ex.Message}");
                throw;
            }

            Trace.WriteLine("Шаг 6: BuildDocument. Отчёт был сохранен по пути: " + savingPath);
            Trace.WriteLine("Шаг 6: BuildDocument. Завершено успешно.");
            return savingPath;
        }
        
        static void DeleteTemporaryFiles(List<string> tempFiles)
        {
            Trace.WriteLine("Шаг 7: DeleteTemporaryFiles.");
            foreach (var tempFile in tempFiles)
            {
                try
                {
                    if (System.IO.File.Exists(tempFile))
                    {
                        System.IO.File.Delete(tempFile);
                        Trace.WriteLine($"Шаг 7: Файл удален: {tempFile}");
                    }
                    else if (Directory.Exists(tempFile))
                    {
                        Directory.Delete(tempFile, true);
                        Trace.WriteLine($"Шаг 7: Директория удалена: {tempFile}");
                    }
                    else
                    {
                        Trace.WriteLine($"Шаг 7: Путь не существует: {tempFile}");
                    }
                }
                catch (Exception ex)
                {
                    Trace.WriteLine("Шаг 7: Произошла ошибка: " + ex.Message);
                }
            }
            Trace.WriteLine("Шаг 7: DeleteTemporaryFiles. Завершено успешно.");
        }
    }

    public class AdrObjectLevel
    {
        public int Level { get; set; }
        public string Name { get; set; }

        public AdrObjectLevel(int level, string name)
        {
            Level = level;
            Name = name;
        }
    }

    public class AdrObject
    {
        public int Level { get; set; }
        public string Type { get; set; }
        public string Name { get; set; }

        public AdrObject(int level, string type, string name)
        {
            Level = level;
            Type = type;
            Name = name;
        }
    }
}