using System;
using Test.Classes;

namespace Test
{
    /*Пред запуском приложения нужно указать строки подключения к БД и размеры дисковых пространств в App.config (connectionStrings и appSettings).
      При запуске проекта необходиме будет войди в Google аккаунт и дать разрешение на добавление таблиц в Google Docs.
      Далее приложение создаст документ c названием "Информация о серверах {текущая дата}" в Google Sheets и 
      начнет обновлять его с интервалом 10 секунд записями 
      о размераз БД и свободном месте на диске.*/

    class Program
    {
        static void Main(string[] args)
        {
            //Запуск
            Start();

            Console.ReadKey();
        }

        //Метод запускает создание документа и дальнейшее его обновление
        static void Start()
        {
            //Создается объект для работы с документам, который в конструкторе принимает обертку над PostgreSQL
            SheetBilder sheetBilder = new SheetBilder(new PostgreSqlDB());

            //Инициируется создание документа
            sheetBilder.CreateSheetDocument();

            //Инициируется обновление документа с интервалом 10 сек.
            sheetBilder.UpdateSheetDocument(10000);

        }

    }

}
