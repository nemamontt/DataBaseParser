using Newtonsoft.Json.Linq;
using System.Text.Encodings.Web;
using System.Text.Json;

namespace ConsoleApp1.Main
{
    class Program
    {
        static private int action = default;
        static private string address = string.Empty;
        static private HttpClient? httpClient;
        private static readonly string pathDBV = Path.Combine(AppContext.BaseDirectory + "DBV");
        private static void Main()
        {                        
            Console.WriteLine("Добро пожаловать в парсер!");
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("Выбери действие:\n [1] Обновить всю БДУ\n [2] Обновить по определнному периоду\n [3] Автообновление\n [4] Завершить работу");
            Console.ForegroundColor = ConsoleColor.White;
            action = Convert.ToInt32(Console.ReadLine());
            while(true)
            {
                if (action == 1)
                {
                    action = 0;
                    address = "https://services.nvd.nist.gov/rest/json/cves/2.0";
                    GenerationJSON();
                }
                else if (action == 2)
                {
                    action = 0;
                    Console.ForegroundColor = ConsoleColor.DarkBlue;
                    Console.WriteLine("Введите начало периода\n формат ГГГГ-ММ-ДД");
                    Console.ForegroundColor = ConsoleColor.White;
                    string? beginningPeriod = Console.ReadLine();
                    Console.ForegroundColor = ConsoleColor.DarkBlue;
                    Console.WriteLine("Введите конец периода\n формат ГГГГ-ММ-ДД");
                    Console.ForegroundColor = ConsoleColor.White;
                    string? endPeriod = Console.ReadLine();
                    address = $"https://services.nvd.nist.gov/rest/json/cves/2.0/?lastModStartDate={beginningPeriod}T13:00:00.000%2B01:00&lastModEndDate={endPeriod}T13:36:00.000%2B01:00";
                    GenerationJSON();
                }
                else if (action == 3)
                {
                    action = 0;
                    GenerationJSON();
                }
                else if (action == 4)
                {
                    action = 0;
                    Console.ForegroundColor = ConsoleColor.DarkBlue;
                    Console.WriteLine("До скорых встреч!");
                    break;
                }
                else if (action == 0)
                {
                    action = Convert.ToInt32(Console.ReadLine());
                }                    
            }
            Console.ForegroundColor = ConsoleColor.White;
        }
        public static void GenerationJSON()
        {
            VulnerabilitiesDTO vulnerabilities = new()
            {
                CVEidentifier = new(),
                LastUpdateDate = DateTime.Now.ToString("G")
            };
            string threatLines = GetResponce(address).Result;
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

                CreateDBVFile(pathDBV, "NVD", vulnerabilities);
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Обновление БДУ прошло успешно!");
                Console.ForegroundColor = ConsoleColor.White;
                action = 0;
            }
        }
        public static async Task<string> GetResponce(string address)
        {
            var socketsHandler = new SocketsHttpHandler { PooledConnectionLifetime = TimeSpan.FromMinutes(2) };
            httpClient = new HttpClient(socketsHandler) { Timeout = TimeSpan.FromSeconds(30) };
            using HttpRequestMessage request = new(HttpMethod.Get, address);

            try
            {
                Console.ForegroundColor = ConsoleColor.Cyan;
                Console.WriteLine("Установка соединения с сервером...");
                using HttpResponseMessage response = await httpClient.SendAsync(request);
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Соединение создано");
                string content = await response.Content.ReadAsStringAsync();
                return content;
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Что-то пошло не так, повторное подключение...");
                return string.Empty;
            }
        }
        public static void CreateDBVFile(string path, string DBVNamem, VulnerabilitiesDTO vulnerabilities)
        {
            if (!Directory.Exists(pathDBV))
                Directory.CreateDirectory(pathDBV);

            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
            };

            var pathJVul = Path.Combine(path, $" {DBVNamem}.json");
            var json = JsonSerializer.Serialize(vulnerabilities, options);
            File.WriteAllText(pathJVul, json);
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
    }
}