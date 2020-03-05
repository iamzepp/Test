using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Test.Classes
{
    /*Класс, представляет собой обертку на Googlee Sheets Api.
     Предназначен для создания документа, листов, записей, а также 
     обновления данных с заданным 
     интервалом времени.*/

    public class SheetBilder
    {

        #region ПОЛЯ И ПЕРЕМЕННЫЕ

       //Экземпляр интерфейса для работы с БД
        public IDataBase DataBase { get; set; }

        //Название токена, полученного через Google Api для системы авторизации OAuth2
        private string clientToken = "client_secret_sheets.json";

        //Права для работы с данными
        private string[] scopes = { SheetsService.Scope.Spreadsheets };

        //Название приложения
        private string applicationName = "TestDocument";

        //Cервис для работы с таблицами
        private SheetsService service;

        //Данные учетной записи пользователя
        private UserCredential credential;

        //Id созданного документа
        private string sheetId;

        //Коллекция названий серверов
        private List<string> serverNames;

        //Коллекция названий БД
        private List<string> dbNames;

        //Коллекция размеров БД
        private decimal[] dbSizes;

        //Cчетчик номера строки, которая должна обновиться
        private int n = 1;

        #endregion

        #region КОНСТРУКТОР
        public SheetBilder(IDataBase dataBase)
        {
            this.DataBase = dataBase;

            //Получаем название серверов
            serverNames = CreateServerNames();
            //Получаем название БД
            dbNames = GetDbNames();
            //Задаем размер массива через количество строк подключения через конфигурационный файл app.config
            dbSizes = new decimal[ConfigurationManager.ConnectionStrings.Count - 1];
        }
        #endregion

        #region ОСНОВНЫЕ МЕТОДЫ

        //Метод для создания нового документа со всеми требованиями
        public void CreateSheetDocument()
        {
            //Получаем учетные данные пользователя
            credential = GetUserCredential();

            //Получаем сервис для работы с Api
            service = GetSheetsService();

            Console.WriteLine("Успешная регистрация!");

            //Cоздаем новый документ и возвращаем его Id
            sheetId = CreateNewSpreadsheet();

            //Cоздаем новый лист на каждый сервер
            CreateNewLists();

            //Добавляем заголовки на листы
            AddHeaders();
        }

        //Метод для обновления записей на листах c заданным интервалом
        public void UpdateSheetDocument(int timeInterval)
        {
            while (true)
            {
                Console.WriteLine(" ");

                //Создаем новые записи о размерах БД на листах
                AddNewEntry();

                //Создаем новые записи о свободном пространстве на диске
                AddFooters();

                Console.WriteLine("Пауза....");

                //Ожидаем заданное время
                Thread.Sleep(timeInterval);
            }
        }
        #endregion

        #region СЛУЖЕБНЫЕ МЕТОДЫ

        //Метод для получения прав для работы с таблицами GoogleSheets пользователя 
        private UserCredential GetUserCredential()
        {
            UserCredential credential = default;

            try
            {
                if (clientToken != null)
                {
                    //Открываем в потоке файл с токеном
                    using (var stream = new FileStream(clientToken, FileMode.Open, FileAccess.Read))
                    {
                        //Cоздаем путь для сохранения туда объетка  DataStore 
                        var credPath = Path.Combine(Directory.GetCurrentDirectory(), "token.json");

                        //Получаем разрешения на работу
                        credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                            GoogleClientSecrets.Load(stream).Secrets,
                            scopes,
                            "user",
                            CancellationToken.None,
                            new FileDataStore(credPath, true)).Result;
                    }
                }
            }
            catch (Exception ex)
            {
                //Выводим сообщение об исключении
                Console.WriteLine(ex.Message);
            }

            return credential;
        }

        //Метод для получения сервиса для работы с таблицами GoogleSheets пользователя 
        private SheetsService GetSheetsService()
        {
            return new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = applicationName,
            });
        }
        
        //Метод для создания нового документа
        private string CreateNewSpreadsheet()
        {
            //Id созданной таблицы
            string SpreadsheetId = default;

            try
            {
                //Cоздаем объект документа
                Spreadsheet sheet = new Spreadsheet();
                SpreadsheetProperties prop = new SpreadsheetProperties();
                
                //Устанавливаем название документа
                prop.Title = "Информация о серверах " + DateTime.Now.ToString();
                sheet.Properties = prop;

                //Делаем запрос на создание таблицы
                var response = service.Spreadsheets.Create(sheet).Execute();

                Console.WriteLine("Документ создан!");

                //Через объект ответа получаем id таблицы
                SpreadsheetId = response.SpreadsheetId;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return SpreadsheetId;
        }

        //Метод получает название все БД указанных в app.config
        private List<string> GetDbNames()
        {
            List<string> list = new List<string>();

            //Получаем количество строк подключения
            int count_con_strings = ConfigurationManager.ConnectionStrings.Count;

            //Проходит по всем БД
            for (int i = 1; i < count_con_strings; i++)
            {
                //Заполняет коллекцию результатами SQL запросов
                list.Add(DataBase.GetDbName(ConfigurationManager.ConnectionStrings[i].ToString()));
            }
            return list;
        }

        //Метод создает имена серверов
        private List<string> CreateServerNames()
        {
            List<string> list = new List<string>();

            //Получаем количество строк подключения
            int count_con_strings = ConfigurationManager.ConnectionStrings.Count;

            //Cоздаем название для каждого
            for (int i = 1; i < count_con_strings; i++)
            {
                list.Add("Server " + i.ToString());
            }

            return list;

        }

        //Метод для создания листов по количеству строк подключения в app.config
        private void CreateNewLists()
        {
            try
            {
                foreach (var name in serverNames)
                {

                    //Создаем объект запроса на добавления листа с заданным именем
                    var addSheetRequest = new Request
                    {
                        AddSheet = new AddSheetRequest
                        {
                            Properties = new SheetProperties
                            {
                                Title = name
                            }
                        }
                    };

                    List<Request> requests = new List<Request> { addSheetRequest };

                    //Создаем запрос
                    BatchUpdateSpreadsheetRequest batchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest();
                    batchUpdateSpreadsheetRequest.Requests = requests;

                    //Выполняем запрос к указанной таблице
                    service.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, sheetId).Execute();
                }

                Console.WriteLine("Листы на каждый сервер созданы!");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        //Метод для добавления новых записей
        private void AddNewEntry()
        {
            //Проходит по всем листам
            for (int i = 0; i < serverNames.Count; i++)
            {
                //Задаем диапазон обновления через счетчик n
                var range = $"{serverNames[i]}!A{n}:E{n}";

                var valueRange = new ValueRange();

                //Получаем SQL запросом размер БД по строкам подключения в конфигурационном файле
                var size = DataBase.GetDbSize(ConfigurationManager.ConnectionStrings[i + 1].ToString(), dbNames[i]);
                dbSizes[i] = size;

                //Преобразуем ячейки в объект valueRange.Values
                var objectList = new List<object> { serverNames[i], dbNames[i], size, DateTime.Now.ToString() };
                valueRange.Values = new List<IList<object>> { objectList };

                //Cоздаем объект запроса на обновление строки в таблицу с указанным id
                var updateRequest = service.Spreadsheets.Values.Update(valueRange, sheetId, range);
                updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;

                //Выполняем запрос
                updateRequest.Execute();
            }

            //Обновляем счетчик строки
            n++;

            Console.WriteLine($"{DateTime.Now} Записи добавлены!");
        }

        //Метод для создания заголовков
        private void AddHeaders()
        {
            //Проходим по всем листам
            foreach (var name in serverNames)
            {
                //Задаем диапазон
                var range = $"{name}!A:E";

                var valueRange = new ValueRange();

                //Преобразуем ячейки в объект valueRange.Values
                var objectList = new List<object> { "Сервер", "База данных", "Размер в ГБ", "Дата обновления" };
                valueRange.Values = new List<IList<object>> { objectList };

                //Cоздаем объект запроса на добавление новой строки в таблицу с указанным id
                var appendRequest = service.Spreadsheets.Values.Append(valueRange, sheetId, range);
                appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;

                //Выполняем запрос
                appendRequest.Execute();
            }

            //Обновляем счетчик строки
            n++;
        }

        //Метод для создания новыx записей о свободном пространстве на диске
        private void AddFooters()
        {
            try
            {
                //Проходит по всем листам документа
                for (int i = 0; i < dbNames.Count; i++)
                {
                    //Задаем диапозон добавления данных
                    var range = $"{serverNames[i]}!A:D";

                    var valueRange = new ValueRange();

                    //Получает свобное пространство на диски вычитанием размера БД от объема дискового пространства
                    decimal driveFreeSize = Convert.ToDecimal(ConfigurationManager.AppSettings[ConfigurationManager.AppSettings.Keys[i]]) - dbSizes[i];

                    //Преобразуем ячейки в объект valueRange.Values
                    var objectList = new List<object> { serverNames[i], "Свободно", driveFreeSize.ToString(), DateTime.Now.ToString() };
                    valueRange.Values = new List<IList<object>> { objectList };

                    //Cоздаем объект запроса на добавление новой строки в таблицу с указанным id
                    var appendRequest = service.Spreadsheets.Values.Append(valueRange, sheetId, range);
                    appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;

                    //Выполняем запрос
                    appendRequest.Execute();
                }
            }
            catch (Exception ex)
            {
                //Вывод сообщения об ошибке
                Console.WriteLine(ex.Message);
            }
        }

        #endregion
    }
}
