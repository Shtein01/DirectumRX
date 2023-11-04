using Aspose.Cells;
using System.Reflection;

namespace Test
{
   class Program
   {
      static void Main(string[] args)
      {
         int product_id = 0;
         var product_price = 0.00;

         //int client_id = 0;
         string client_name = string.Empty;
         string client_address = string.Empty;

         int request_id = 0;
         string request_product = String.Empty;
         int request_client = 0;
         var request_count = 0.00;
         DateTime request_date;

         List<string> arr = new List<string>();

         var path = @"C:\Users\Rail\Desktop\124\DirectumRX\Практика.xlsx"; //string.Empty; 
         var action = 0;
         while (action != 1)
         {
            Console.WriteLine("1. Выход");
            Console.WriteLine("2. Запрос на ввод пути до файла с данными (в качестве документа с данными использовать Приложение 2).");
            Console.WriteLine("3. По наименованию товара выводить информацию о клиентах, заказавших этот товар,");
            Console.WriteLine("с указанием информации по количеству товара, цене и дате заказа.");
            Console.WriteLine("4. Запрос на изменение контактного лица клиента с указанием параметров: ");
            Console.WriteLine("Название организации, ФИО нового контактного лица.");
            Console.WriteLine("В результате информация должна быть занесена в этот же документ, ");
            Console.WriteLine("в качестве ответа пользователю необходимо выдавать информацию о результате изменений.");
            Console.WriteLine("5. Запрос на определение золотого клиента, клиента с наибольшим количеством заказов, за указанный год, месяц.");
            Console.WriteLine("Введите цифру с методом: ");

            var act = int.TryParse(Console.ReadLine(), out action);


            switch (action)
            {
               case 2:
                  Console.WriteLine("Введите полный путь до файла: ");
                  path = Console.ReadLine();
                  break;
               case 3:
                  if (path != String.Empty) OpenFile();
                  else Console.WriteLine("Выберите пункт 2 и введите полный путь до файла.");
                  Console.WriteLine();
                  break;
               case 4:
                  RewriteFile();
                  break;
               default:
                  break;
            }
         }

         void RewriteFile() //в процессе
         {
            Console.WriteLine("Введите данные для изменения контактного лица клиента:");
            Console.Write("Название организации: ");
            string? rewrite_client_name = Console.ReadLine();
            Console.Write("ФИО нового контактного лица: ");
            string? rewrite_fio = Console.ReadLine();
         }

         void OpenFile()
         {
            string product_name = string.Empty;
            while (string.IsNullOrWhiteSpace(product_name)) //.Length == 0)
            {
               Console.Write("Наименование товара: ");
               product_name = Console.ReadLine();
            }

            try
            {
               // Загрузить файл Excel
               Workbook wb = new Workbook(path);

               // Получить все рабочие листы
               WorksheetCollection collection = wb.Worksheets;
               Worksheet worksheet;

               // Перебрать все рабочие листы
               for (int worksheetIndex = 0; worksheetIndex < collection.Count; worksheetIndex++)
               {
                  // Получить рабочий лист, используя его индекс
                  worksheet = collection[worksheetIndex];

                  //Поиск товара по имени в листе "Товары"
                  if (worksheet.Name == "Товары")
                  {
                     arr = ReadArray(worksheet, worksheet.Cells.MaxDataRow, worksheet.Cells.MaxDataColumn, product_name);
                     if (arr.Count == 0)
                     {
                        Console.WriteLine("Не найден товар с таким именем!");
                        break;
                     }
                     var product_id_bool = int.TryParse(arr[0], out product_id);
                     var product_price_bool = double.TryParse(arr[3], out product_price);
                     Console.WriteLine("Товары: Код: " + product_id + ", Товар: " + arr[1] + ", Цена за ед.: " + product_price);
                  }

                  //Поиск заявки по id товара
                  if (worksheet.Name == "Заявки")
                  {
                     arr = ReadArray(worksheet, worksheet.Cells.MaxDataRow, worksheet.Cells.MaxDataColumn, int_id: product_id);
                     request_id = int.Parse(arr[0]);
                     request_product = product_name;
                     request_client = int.Parse(arr[2]);
                     request_count = int.Parse(arr[4]);
                     var bool_request_date = DateTime.TryParse(arr[5], out request_date);
                     Console.WriteLine("Заявки: Код: " + request_id.ToString() + ", Товар: " + request_product + ", Код клиента: " +
                        request_client.ToString() + ", Количество: " + request_count.ToString() + ", Дата заявки: " + $"{request_date.ToString("d")}");
                  }
               }


               // Перебрать все рабочие листы
               for (int worksheetIndex = 0; worksheetIndex < collection.Count; worksheetIndex++)
               {
                  //Если не был найден товар по имени пропускаю цикл поиска Клиента
                  if (product_id == 0) break;

                  // Получить рабочий лист, используя его индекс
                  worksheet = collection[worksheetIndex];

                  //Поиск клиента по id из листа "Заявки"
                  if (worksheet.Name == "Клиенты")
                     {
                        arr = ReadArray(worksheet, worksheet.Cells.MaxDataRow, worksheet.Cells.MaxDataColumn, int_id: request_client);
                        client_name = arr[1];
                        client_address = arr[2];
                        Console.WriteLine("Клиент: Код: " + request_client + ", Наименование организации: " + client_name + ", Адрес: " + client_address);
                     }
               }

               if (request_id == 0 && product_id != 0)
               {
                  Console.WriteLine("Заявки на данный товар не найдены.");
               }
            }
            catch (System.IO.FileNotFoundException)
            {
               Console.WriteLine("File not found");
               return;
            }
            catch (Exception ex)
            {
               Console.WriteLine("Error: " + ex.Message);
               //throw ex;
            }
         }

         List<string> ReadArray(Worksheet worksheet, int rows, int cols,
            string? str1 = default, int? int_id = default)
         {
            if (worksheet == null)
            {
               arr.Clear();
               return arr;
            }
            // Цикл по строкам
            for (int i = 1; i <= rows; i++)
            {
               //Очищаю массив перед записью
               arr.Clear();
               // Перебрать каждый столбец в выбранной строке
               for (int j = 0; j <= cols; j++)
               {
                  arr.Add(worksheet.Cells[i, j].Value.ToString());
               }
               // Записать листы в отдельные таблицы
               if (worksheet.Name == "Товары")
               {
                  if (arr[1] == str1)
                     return arr;
                  else arr.Clear();
               }
               if (worksheet.Name == "Заявки")
               {
                  if (int.Parse(arr[1]) == int_id)
                     return arr;
                  else arr.Clear();
               }
               if (worksheet.Name == "Клиенты")
               {
                  if (int.Parse(arr[0]) == int_id)
                     return arr;
                  else arr.Clear();
               }
            }
            return arr;
         }
      }
   }
}
