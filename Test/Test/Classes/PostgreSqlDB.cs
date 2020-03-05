using Npgsql;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Test.Classes
{
    //Класс представляет собой обертку над СУБД PostgreSQL с изпольвованием библиотеки Npgsql
    public class PostgreSqlDB : IDataBase
    {
        //Метод, возвращает размер база данных в ГБ
        public decimal GetDbSize(string connectionString, string dbName)
        {
            //Cтроковое представление размера БД в байтах
            string sizeResultInBytes = default;

            //Размер БД в ГБ
            decimal ConvertResultInGb = default;

            try
            {
                //Открываем соединени с БД по строке подключения
                using (NpgsqlConnection npgSqlConnection = new NpgsqlConnection(connectionString))
                {
                    npgSqlConnection.Open();

                    //Cоздаем объект SQL запроса 
                    using (NpgsqlCommand npgSqlCommand = new NpgsqlCommand($"SELECT pg_database_size ('{dbName}')", npgSqlConnection))
                    {
                        NpgsqlDataReader npgSqlDataReader = npgSqlCommand.ExecuteReader();

                        //Проверяем наличие строк в объекте ответа от БД и считываем ответ в dbName
                        if (npgSqlDataReader.HasRows)
                            while (npgSqlDataReader.Read())
                                sizeResultInBytes = npgSqlDataReader[0].ToString();

                        npgSqlDataReader.Close();
                    }
                }

                //Преобразуем размер базы данных из байтов в ГБ и устанавливаем две значащие цифры после запятой
                ConvertResultInGb = (Math.Round((Convert.ToDecimal(sizeResultInBytes) * 9.313225746154785E-10M), 2));

             
            }
            catch (Exception ex)
            {
                //Вывод информация о исключении
                Console.WriteLine(ex.Message.ToString());
            }
           

            return ConvertResultInGb;
        }

        //Метод, возвращает имя базы данных
        public string GetDbName(string connectionString)
        {
            string dbName = default;

                try
                {
                //Открываем соединени с БД по строке подключения
                using (NpgsqlConnection npgSqlConnection = new NpgsqlConnection(connectionString))
                {
                    npgSqlConnection.Open();

                    //Cоздаем объект SQL запроса 
                    using (NpgsqlCommand npgSqlCommand = new NpgsqlCommand("SELECT current_database ()", npgSqlConnection))
                    {

                        NpgsqlDataReader npgSqlDataReader = npgSqlCommand.ExecuteReader();

                        //Проверяем наличие строк в объекте ответа от БД и считываем ответ в dbName
                        if (npgSqlDataReader.HasRows)
                            while (npgSqlDataReader.Read())
                            {
                               dbName=npgSqlDataReader[0].ToString();
                            }

                        npgSqlDataReader.Close();
                    }
                }
                }
                catch (Exception ex)
                {
                    //Выдаем текст исклбчения, если есть
                    Console.WriteLine(ex.Message);
                }
            

            return dbName;
        }
    }
}
