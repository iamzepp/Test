using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Test.Classes
{
    //Интерфейс, который дает возможность работать с различными БД
    public interface IDataBase
    {
        //Метод для получения размера БД
        decimal GetDbSize(string connectionString, string dbName);

        //Метод для получения названия текучей БД
        string GetDbName(string connectionString);
    }
}
