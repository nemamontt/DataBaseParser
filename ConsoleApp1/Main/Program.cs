using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.Encodings.Web;
using System.Text.Json;
using Excel = Microsoft.Office.Interop.Excel;

namespace ConsoleApp1.Main
{
    class Program
    {
        static private Application? application; // объект для работы с Word и Exel
        static private int action = 0; // выбранное действие
        static private HttpClient? httpClient; // объект для http запроса
        private static readonly string API_KEY = "6cbe6e35-f52e-410f-a627-444352adf9c3";
        private static readonly string PATH_RES = Path.Combine(Environment.CurrentDirectory, "Resource"); // путь к каталогу "DBV"
        private static readonly string PATH_EXE = Path.Combine(Environment.CurrentDirectory, "VulDBReader.exe"); //путь к исполняемому файлу 
        private static readonly string PATH_FSTEC = Path.Combine(Environment.CurrentDirectory, "Resource", "FSTEC"); // путь к каталогу "resource"
        private const string HTTP_REQUEST_FSTEC = "https://bdu.fstec.ru/files/documents/vullist.xlsx"; //http-запрос для БДУ ФСТЭК
        private const string HTTP_REQUEST_NVD = "https://services.nvd.nist.gov/rest/json/cves/2.0/?resultsPerPage=1000&startIndex=0"; //http-запрос для БДУ NVD                                                                                          //
        private static async Task Main()
        {                                
            Start();

            while (true)
            {
                if (action == 0)
                {
                    Console.ForegroundColor = ConsoleColor.White;
                    try //проверка введенных символов
                    {
                        action = Convert.ToInt32(Console.ReadLine());
                        if (action != 1 && action != 2 && action != 3 && action != 4 && action != 5 && action != 6 && action != 7)
                            throw new Exception();
                    }
                    catch //обработка исключения при неверно указанном символе
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("Вы ввели недопустимый символ!");
                        Console.ForegroundColor = ConsoleColor.White;
                        action = 0;
                    }
                } // конструктор ввода
                else if (action == 1)
                {
                    Pars();
                    //WorkingWithNVD(HTTP_REQUEST_NVD);
                    //await ParsingFSTEC(HTTP_REQUEST_FSTEC); //парсинг БДУ ФСТЭК
                    action = 0;
                } // парсинг БДУ
                else if (action == 2)
                {
                    while (true)
                    {
                        Console.Write("[1]Обновлять БДУ каждый день при запуске компьютера\n[2]Обновлять БДУ в определенное время(учтите, что при этом компьютер должен быть включен!)\n[3]Удалить автоматизацию\n");
                        int actionCMD = Convert.ToInt32(Console.ReadLine());

                        if (actionCMD == 1)
                        {
                            CreationAutomation(1);
                            break;
                        }
                        else if (actionCMD == 2)
                        {
                            Console.WriteLine($"Укажите в какой час обновлять БДУ(формат: ЧЧ)");
                            try
                            {
                                int hour = Convert.ToInt32(Console.ReadLine());
                                if (hour < 0 | hour > 24)
                                    throw new Exception();
                                CreationAutomation(2, hour);
                            }
                            catch 
                            { 
                                Console.ForegroundColor = ConsoleColor.Red;
                                Console.WriteLine("Вы ввели недопустимый символ");
                            }                            
                            break;
                        }
                        else if (actionCMD == 3)
                        {
                            CreationAutomation(3);
                            break;
                        }
                    }
                    action = 0;
                } // создание автоватизаций
                else if (action == 3)
                {
                    Console.ForegroundColor = ConsoleColor.DarkMagenta;
                    action = 0;
                    try
                    {
                        GetInfoDB("FSTEC");
                        GetInfoDB("NVD");
                    }
                    catch
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("Произошла ошибка, скорее всего отсутствует или нарушен Json-объект");
                    }
                   
                } // получение информации о БДУ
                else if (action == 4)
                {                    
                    List<string> missingVulnerabilities = new();                    
                    int counter = 0;
                    bool foundID = false;
                    try
                    {
                        var dtoFSTEC = GetObjectDTO("FSTEC");
                        var dtoNVD = GetObjectDTO("NVD");

                        foreach (var itemFSTEK in dtoNVD.Vulnerabilities)
                        {
                            foreach (var itemNVD in dtoFSTEC.Vulnerabilities)
                            {
                                if (itemFSTEK.CVEidentifier == itemNVD.CVEidentifier)
                                {
                                    foundID = true;
                                }
                            }
                            if (!foundID)
                            {
                                counter++;
                                missingVulnerabilities.Add(itemFSTEK.CVEidentifier);
                            }
                        }
                        Console.WriteLine($"В БДУ ФСТЭК отсутвуют {counter} записи:");
                        for (int i = 0; i < missingVulnerabilities.Count; i++)
                        {
                            if (i % 3 == 0)
                                Console.WriteLine();
                            Console.Write($"{missingVulnerabilities[i]}\t");
                        }
                        CreatFileJSON("Resource", CreatJsonObject(missingVulnerabilities), "DBV_INFO");
                    }
                    catch
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("Произошла ошибка, скорее всего отсутствует или нарушен Json-объект");
                    }
                    action = 0;
                } //анализ БДУ и выявление различий
                else if (action == 5)
                {
                    action = 0;
                    Console.Clear();
                    Start();
                } // отрисовка начального экрана
                else if (action == 6)
                {
                    action = 0;
                } // создание рапорта
                else if(action == 7)
                {
                    Console.ForegroundColor = ConsoleColor.DarkBlue;
                    Console.WriteLine("До скорых встреч!");
                    break;
                } //завершение работы программы
            }
            Console.ForegroundColor = ConsoleColor.White;
            return;
        } //главный метод
        private struct VulnerabilitiesDTO
        {
            public struct VulnerabilityDescription
            {
                public string CVEidentifier { get; set; }
                public string CVEDescription { get; set; }
            }
            public List<VulnerabilityDescription> Vulnerabilities { get; set; }
            public string LastUpdateDate { get; set; }
        }
        private static void WorkingWithNVD(string httpRequest)
        {
            VulnerabilitiesDTO vulnerabilities = new()
            {
                Vulnerabilities = new(),
                LastUpdateDate = DateTime.Now.ToString("G")
            };
            string threatLines = GetResponseNVD(httpRequest).Result;
            if (threatLines != string.Empty)
            {
                while (true)
                {
                    try
                    {
                        var json = JObject.Parse(threatLines);
                        var trips = json["vulnerabilities"];
                        int counter = default;
                        foreach (JToken trip in trips)
                        {
                            VulnerabilitiesDTO.VulnerabilityDescription vulDescription = new();
                            counter++;
                            var counterTrips = trips.Count();
                            float progress = counter * 100 / counterTrips;
                            var cve = trip["cve"];
                            var id = cve["id"];
                            var description = cve["descriptions"].First["value"];
                            vulDescription.CVEidentifier = (string)id;
                            vulDescription.CVEDescription = (string)description;
                            vulnerabilities.Vulnerabilities.Add(vulDescription);                           
                        }
                        ClearLine();
                        //CreatFileJSON("Resource", CreatJsonObject(vulnerabilities), "NVD");
                        Axc(vulnerabilities);
                        Console.WriteLine("БДУ NVD успешно обновлена!");
                        break;
                    }
                    catch
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("Произошла ошибка, код - [001]");
                        Thread.Sleep(10000);
                        WorkingWithNVD(httpRequest);
                        break;
                    }
                }
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Произошла ошибка, код - [002]");
            }
        }
        private static async Task<string> GetResponseNVD(string address)
        {
            httpClient = new HttpClient() { Timeout = TimeSpan.FromSeconds(120) };
            using HttpRequestMessage request = new(HttpMethod.Get, address);
            request.Headers.Add("User-Agent", $"{API_KEY}");

            try
            {
                Console.ForegroundColor = ConsoleColor.Cyan;
                Console.WriteLine("Установка соединения с сервером NVD...");
                using HttpResponseMessage response = await httpClient.SendAsync(request);
                if (response.IsSuccessStatusCode)
                {
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("Соединение установлено");
                    string content = await response.Content.ReadAsStringAsync();
                    return content;
                }
                else
                {
                    return string.Empty;
                }
                
            }
            catch
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Произошла ошибка, код - [003]");
                return string.Empty;
            }
            finally
            {
                httpClient.Dispose();
            }
        }
        private static void ClearLine()
        {
            Console.SetCursorPosition(0, Console.CursorTop - 1);
            Console.Write(new string(' ', Console.WindowWidth));
            Console.SetCursorPosition(0, Console.CursorTop);
        }
        private static async Task ParsingFSTEC(string address)
        {
            var handler = new HttpClientHandler
            {
                ClientCertificateOptions = ClientCertificateOption.Manual,
                ServerCertificateCustomValidationCallback =
                (httpRequestMessage, cert, cetChain, policyErrors) =>
                { return true; }
            };

            httpClient = new HttpClient(handler) { Timeout = TimeSpan.FromSeconds(60) };
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("Установка соединения с сервером ФСТЭК...");
            using HttpRequestMessage request = new(HttpMethod.Get, address);
            var response = await httpClient.GetAsync(address);
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Соединение установлено");
            if (response.IsSuccessStatusCode)
            {
                var stream = await response.Content.ReadAsStreamAsync();

                Process[] processList;
                FileStream? fileStream = null;
                processList = Process.GetProcessesByName("EXCEL");
                foreach (Process proc in processList) { proc.Kill(); }
                try
                {
                    fileStream = new FileStream(PATH_RES + "\\FSTEC.xlsx", FileMode.Create);
                    await stream.CopyToAsync(fileStream);                                                            
                }
                catch 
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Произошла ошибка, код - [005]");
                }  
                finally 
                {
                    stream.Close();
                    fileStream?.Close();
                    httpClient.Dispose();
                }

                try { WorkingWithFSTEC(); }
                catch
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Произошла ошибка, код - [004]");
                }
            }
            return;
        }
        private static void WorkingWithFSTEC()
        {
            VulnerabilitiesDTO vulnerabilities = new()
            {
                Vulnerabilities = new(),
                LastUpdateDate = DateTime.Now.ToString("G")
            };           
            try
            {
                application = new Application();
                Workbook ObjWorkBook = application.Workbooks.Open(PATH_FSTEC,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing);
                Worksheet ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];

                Excel.Range range2 = ObjWorkSheet.UsedRange.Columns["A", Type.Missing];               
                int countRow = range2.Rows.Count;
                Excel.Range currentFind = null;
                Excel.Range firstFind = null;
                Excel.Range rangeCVE = application.get_Range("S1", "S" + countRow);

                currentFind = rangeCVE.Find("CVE", Type.Missing,
                XlFindLookIn.xlValues, XlLookAt.xlPart,
                XlSearchOrder.xlByRows, XlSearchDirection.xlNext, false,
                Type.Missing, Type.Missing);

                int iteration = 0;
                while (currentFind != null)
                {
                    if (firstFind == null)
                        firstFind = currentFind;
                    else if (currentFind.get_Address(XlReferenceStyle.xlA1) == firstFind.get_Address(XlReferenceStyle.xlA1))
                        break;

                    var cveID = currentFind.Value;
                    VulnerabilitiesDTO.VulnerabilityDescription vulDescription = new()
                    {
                        CVEidentifier = IDProcessing(cveID),
                        CVEDescription = ObjWorkSheet.Cells[currentFind.Row, "C"].Value?.ToString()
                    };             
                    vulnerabilities.Vulnerabilities.Add(vulDescription);
                    currentFind = rangeCVE.FindNext(currentFind);
                    iteration++;
                    int progress = 100 * iteration / countRow;
                    Console.ForegroundColor = ConsoleColor.Cyan;
                    ClearLine();
                    Console.WriteLine($"Прогресс: {progress}% ({countRow}/{iteration})");
                }
                ClearLine();
                CreatFileJSON("Resource", CreatJsonObject(vulnerabilities), "FSTEC");
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("БДУ FSTEC успешно обновлена!");
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(ex.ToString());
            }
            finally
            {                
                application?.Quit();
                Marshal.ReleaseComObject(application);
                Process[] processList = Process.GetProcessesByName("EXCEL");
                foreach (Process proc in processList) { proc.Kill(); }
                File.Delete(Path.Combine(PATH_RES, "FSTEC.xlsx"));
            }
        } 
        private static string IDProcessing(string sourceString)
        {
            string formattedString = string.Empty;
            for (int i = 0; i < sourceString.Length; i++)
            {
                if (sourceString[i] is 'C')
                {
                    for (int j = 0; j < 13; j++)
                    {
                        formattedString += sourceString[i];
                        i++;
                    }
                    break;
                }
            }
            return formattedString;
        }
        private static void Start()
        {
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine("Добро пожаловать в парсер!");
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("Выбери действие:\n [1] Обновить все БДУ\n [2] Создать автоматизацию\n [3] Информация о БДУ\n [4] Анализ БДУ\n [5] Очистить консоль\n [6] Создать рапорт\n [7] Завершить работу");
            Console.ForegroundColor = ConsoleColor.White;
        } //отрисовка меню
        private static void GetInfoDB(string nameDB)
        {
            string lastUpdateDate = GetObjectDTO(nameDB).LastUpdateDate;
            Console.WriteLine($"Последнее обновление БДУ {nameDB} : {lastUpdateDate}");
        } // получение информации о БДУ
        private static VulnerabilitiesDTO GetObjectDTO(string nameJsonFile)
        {
            string jsonString = File.ReadAllText(Path.Combine(PATH_RES, nameJsonFile + ".json"));
            return JsonSerializer.Deserialize<VulnerabilitiesDTO>(jsonString);
        } // десериализация объекта JSON
        private static void CreationAutomation(int actionCMD, int hour = 10)
        {           
            using Process proc = new();
            proc.StartInfo.UseShellExecute = true;
            proc.StartInfo.CreateNoWindow = true;
            proc.StartInfo.FileName = "schtasks";
            proc.StartInfo.Verb = "runas";
            
            string hourString;
            if (hour < 10)
            {
                hourString = Convert.ToString(hour);
                hourString = hourString.Insert(0, "0");
            }
            else { hourString = Convert.ToString(hour); }

            if (actionCMD == 1)
            {
                proc.StartInfo.Arguments = $"/create /tn DatabaseUpdate /tr {PATH_EXE} /sc onstart /f";
                proc.Start();
            }
            else if (actionCMD == 2)
            {
                proc.StartInfo.Arguments = $"/create /tn DatabaseUpdate /tr {PATH_EXE} /sc daily /st {hourString}:00 /f";
                proc.Start();
            }
            else if (actionCMD == 3)
            {
                proc.StartInfo.Arguments = "/delete /tn DatabaseUpdate /f";
                proc.Start();
            }

            proc.WaitForExit();
            if (proc.ExitCode == 0)
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Успешно");
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Что-то пошло не так");
            }
        } //создание автоматизации
        private static void CreatFileJSON(string nameDirectory, string json, string nameJson)
        {
            var directoryPath = Path.Combine(Environment.CurrentDirectory, nameDirectory);
            if (!Directory.Exists(directoryPath))
                Directory.CreateDirectory(directoryPath);

            var pathFile = Path.Combine(directoryPath, $"{nameJson}.json");
            File.WriteAllText(pathFile, json);
        } //создание файла json
        private static void Axc(VulnerabilitiesDTO qwe)
        {
            var directoryPath = Path.Combine(Environment.CurrentDirectory, "AAA");
            var pathFile = Path.Combine(directoryPath, $"BBB.json");
           
            if (!Directory.Exists(directoryPath))
                Directory.CreateDirectory(directoryPath);

            if(!File.Exists(pathFile))
                File.Create(pathFile).Close();  

            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
            };

            string jsonString = File.ReadAllText(pathFile);
            var disDTO = JsonSerializer.Deserialize<VulnerabilitiesDTO>(jsonString);
            disDTO.Vulnerabilities.AddRange(qwe.Vulnerabilities);
            var json = JsonSerializer.Serialize(disDTO, options);

            File.WriteAllText(pathFile, json);
        }
        private static string CreatJsonObject(object file)
        {
            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
            };

            var json = JsonSerializer.Serialize(file, options);
            return json;
        } // создание json-объекта
        private static void Pars()
        {
            string jsonString = GetResponseNVD("https://services.nvd.nist.gov/rest/json/cves/2.0/?resultsPerPage=1&startIndex=0").Result;
            if (jsonString != string.Empty)
            {
                var json = JObject.Parse(jsonString);
                var trips = json["totalResults"];
                int countIteration = Convert.ToInt32((string)trips) / 2000;
                int startIndex = 0;
                for (int i = 0; i <= countIteration; i++)
                {
                    if (i % 5 == 0)
                        Thread.Sleep(60000);

                    startIndex += 2000;
                    string request = $"https://services.nvd.nist.gov/rest/json/cves/2.0/?resultsPerPage=2000&startIndex={startIndex}";
                    Console.WriteLine(request);
                    WorkingWithNVD(request);
                    Thread.Sleep(10000);
                }
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Произошла ошибка");
            };                     
        }
    }
}