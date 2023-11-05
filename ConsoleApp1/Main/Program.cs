using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.Encodings.Web;
using System.Text.Json;
using Excel = Microsoft.Office.Interop.Excel;

namespace ConsoleApp1.Main
{
    class Program
    {
        static private Application? application;
        static private int action = default;
        static private HttpClient? httpClient;
        private static readonly string PATH_DBV = Path.Combine(AppContext.BaseDirectory + "DBV"); // путь к каталогу с БДУ
        private static void Main()
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
                        if (action != 1 && action != 2 && action != 3 && action != 4 && action != 5 && action != 6)
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
                    action = 0;
                    GenerationJSON_NVD("https://services.nvd.nist.gov/rest/json/cves/2.0"); //парсинг БДУ NVD                                     
                    Get_FSTEC("https://bdu.fstec.ru/files/documents/vullist.xlsx"); //парсинг БДУ ФСТЭК
                } // парсинг БДУ
                else if (action == 2)
                {
                    action = 0;
                    Console.WriteLine("Здесь вы можете настриоть автообновление БДУ");
                    //Console.ForegroundColor = ConsoleColor.DarkBlue;
                    //Console.WriteLine("Введите начало периода\n формат ГГГГ-ММ-ДД");
                    //Console.ForegroundColor = ConsoleColor.White;
                    //string? beginningPeriod = Console.ReadLine();
                    //Console.ForegroundColor = ConsoleColor.DarkBlue;
                    //Console.WriteLine("Введите конец периода\n формат ГГГГ-ММ-ДД");
                    //Console.ForegroundColor = ConsoleColor.White;
                    //string? endPeriod = Console.ReadLine();
                    //address = $"https://services.nvd.nist.gov/rest/json/cves/2.0/?lastModStartDate={beginningPeriod}T13:00:00.000%2B01:00&lastModEndDate={endPeriod}T13:36:00.000%2B01:00";
                } // создание автоватизаций
                else if (action == 3)
                {
                    Console.ForegroundColor = ConsoleColor.DarkMagenta;
                    action = 0;
                    GetInfoDB("FSTEC");
                    GetInfoDB("NVD");                    
                } // получение информации о БДУ
                else if(action == 4)
                {
                    action = 0;
                    List<string> unVul = new();
                    var dtoFSTEC = GetObjectDTO("FSTEC");
                    var dtoNVD = GetObjectDTO("NVD");
                    int counter = 0;
                    for (int i = 0; i < dtoNVD.CVEidentifier.Count - 1; i++)
                    {
                        bool foundID = false;
                        for (int j = 0; j < dtoFSTEC.CVEidentifier.Count - 1; j++)
                        {
                            if (dtoNVD.CVEidentifier[i] == dtoFSTEC.CVEidentifier[j])
                            {
                                foundID = true;
                            }                            
                        }            
                        if(!foundID)
                        {
                            counter++;
                            unVul.Add(dtoNVD.CVEidentifier[i]);
                        }
                    }
                    Console.WriteLine($"В БДУ ФСТЭК отсутвуют {counter} записи:");
                    for (int i = 0; i < unVul.Count; i++)
                        Console.WriteLine(unVul[i]);

                    Console.WriteLine("Желаете создать рапорт? (ДА - 1 | НЕТ - 2)");
                } //анализ БДУ и выявление различий
                else if (action == 5) 
                {
                    action = 0;
                    Console.Clear();
                    Start();
                } // отрисовка начального экрана
                else if(action == 6)
                {
                    Console.ForegroundColor = ConsoleColor.DarkBlue;
                    Console.WriteLine("До скорых встреч!");                   
                    break;
                } //завершение работы программы          
            }
            Console.ForegroundColor = ConsoleColor.White;
        }
        public static void GenerationJSON_NVD(string address)
        {
            VulnerabilitiesDTO vulnerabilities = new()
            {
                CVEidentifier = new(),
                LastUpdateDate = DateTime.Now.ToString("G")
            };
            string threatLines = Get_NVD(address).Result;
            if (threatLines != string.Empty)
            {
                var json = JObject.Parse(threatLines);
                var trips = json["vulnerabilities"];

                int counter = default;
                foreach (JToken trip in trips)
                {
                    counter++;
                    var counterTrips = trips.Count();
                    float progress = counter * 100 / counterTrips;
                    var cve = trip["cve"];
                    var id = cve["id"];
                    vulnerabilities.CVEidentifier.Add((string)id);
                    Console.ForegroundColor = ConsoleColor.Cyan;
                    Console.WriteLine($"Прогресс: {progress}%");
                    ClearLine();
                }
                CreateDBVFile(PATH_DBV, "NVD", vulnerabilities);
            }
        }
        public static async Task<string> Get_NVD(string address)
        {           
            var socketsHandler = new SocketsHttpHandler
            {
                PooledConnectionLifetime = TimeSpan.FromMinutes(2) 
            };
            httpClient = new HttpClient(socketsHandler) { Timeout = TimeSpan.FromSeconds(30) };
            using HttpRequestMessage request = new(HttpMethod.Get, address);

            try
            {
                Console.ForegroundColor = ConsoleColor.Cyan;
                Console.WriteLine("Установка соединения с сервером NVD...");
                using HttpResponseMessage response = await httpClient.SendAsync(request);
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Соединение установлено");
                string content = await response.Content.ReadAsStringAsync();
                return content;
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Что-то пошло не так, повторное подключение...");
                return Get_NVD(address).Result;               
            }
        }
        public static void CreateDBVFile(string path, string DBVNamem, VulnerabilitiesDTO vulnerabilities)
        {
            if (!Directory.Exists(PATH_DBV))
                Directory.CreateDirectory(PATH_DBV);

            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
            };

            var pathJVul = Path.Combine(path, $"{DBVNamem}.json");
            var json = JsonSerializer.Serialize(vulnerabilities, options);
            File.WriteAllText(pathJVul, json);
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"БДУ {DBVNamem} успешно обновлена!");
        }
        public static void ClearLine()
        {
            Console.SetCursorPosition(0, Console.CursorTop - 1);
            Console.Write(new string(' ', Console.WindowWidth));
            Console.SetCursorPosition(0, Console.CursorTop);
        }
        public struct VulnerabilitiesDTO
        {
            public List<string> CVEidentifier { get; set; }
            public string LastUpdateDate { get; set; }
        }
        public static async Task Get_FSTEC(string address)
        {
            var handler = new HttpClientHandler
            {
                ClientCertificateOptions = ClientCertificateOption.Manual,
                ServerCertificateCustomValidationCallback =
                (httpRequestMessage, cert, cetChain, policyErrors) =>
                { return true; }
            };

            httpClient = new HttpClient(handler) { Timeout = TimeSpan.FromSeconds(30) };
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
                processList = Process.GetProcessesByName("EXCEL");
                foreach (Process proc in processList)
                {
                    proc.Kill();
                }

                var fileStream = new FileStream(PATH_DBV + "\\FSTEC.xlsx", FileMode.Create);                
                await stream.CopyToAsync(fileStream);
                fileStream.Close();
                stream.Close();
                ParseFSTEC();
                return;
            }
        }
        public static void ParseFSTEC()
        {
            VulnerabilitiesDTO vulnerabilities = new()
            {
                CVEidentifier = new(),
                LastUpdateDate = DateTime.Now.ToString("G")
            };
            try
            {
                var path = Path.Combine(PATH_DBV + "\\FSTEC");
                application = new Excel.Application();
                Workbook ObjWorkBook = application.Workbooks.Open(path,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing);
                Worksheet ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];
                var lastCell = ObjWorkSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell);
                var countRow = lastCell.Row;

                Excel.Range currentFind = null;
                Excel.Range firstFind = null;
                Excel.Range range = application.get_Range("S1", "S" + countRow);

                currentFind = range.Find("CVE", Type.Missing,
                XlFindLookIn.xlValues, XlLookAt.xlPart,
                XlSearchOrder.xlByRows, XlSearchDirection.xlNext, false,
                Type.Missing, Type.Missing);

                int iteration = 0;
                while (currentFind != null)
                {
                    if (firstFind == null)
                    {
                        firstFind = currentFind;
                    }
                    else if (currentFind.get_Address(XlReferenceStyle.xlA1) == firstFind.get_Address(XlReferenceStyle.xlA1))
                    {
                        break;
                    }
                    var cveID = currentFind.Value;
                    vulnerabilities.CVEidentifier.Add(IDProcessing(cveID));
                    currentFind = range.FindNext(currentFind);
                    iteration++;
                    int progress = 100 * iteration / countRow;
                    Console.ForegroundColor = ConsoleColor.Cyan;
                    ClearLine();
                    Console.WriteLine($"Прогресс: {progress}%");                   
                }                
                application.Quit();               
            }
            catch (Exception ex)
            { 
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                ClearLine();
                Marshal.ReleaseComObject(application);
                File.Delete(Path.Combine(PATH_DBV, "FSTEC.xlsx"));
                CreateDBVFile(PATH_DBV, "FSTEC", vulnerabilities);
            }                     
        }
        public static string IDProcessing(string sourceString)
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
        public static void Start()
        {
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine("Добро пожаловать в парсер!");
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("Выбери действие:\n [1] Обновить все БДУ\n [2] Создать автоматизацию\n [3] Информация о БДУ\n [4] Анализ БДУ\n [5] Вернутья назад\n [6] Завершить работу");
            Console.ForegroundColor = ConsoleColor.White;
        }
        public static void GetInfoDB(string nameDB)
        {
            string lastUpdateDate = GetObjectDTO(nameDB).LastUpdateDate;
            Console.WriteLine($"Последнее обновление БДУ {nameDB}: {lastUpdateDate}");
        }
        public static VulnerabilitiesDTO GetObjectDTO(string nameJsonFile)
        {
            string jsonString = File.ReadAllText(Path.Combine(PATH_DBV, nameJsonFile + ".json"));
            return JsonSerializer.Deserialize<VulnerabilitiesDTO>(jsonString);
        }
    }
}