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

         string client_name = string.Empty;
         string client_address = string.Empty;

         int request_id = 0;
         string request_product = String.Empty;
         int request_client = 0;
         var request_count = 0.00;
         DateTime request_date = default;

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
                  if (path != String.Empty) 
                     Console.WriteLine(OpenFile());
                  else Console.WriteLine("Выберите пункт 2 и введите полный путь до файла.");
                  Console.WriteLine();
                  break;
               case 4:
                  Console.WriteLine(RewriteFile());
                  break;
               default:
                  break;
            }
         }

         string RewriteFile() 
         {
            Console.WriteLine("Введите данные для изменения контактного лица клиента:");
            string? rewrite_client_name = string.Empty;
            string? rewrite_fio = string.Empty;

            while (string.IsNullOrWhiteSpace(rewrite_client_name))
            {
               Console.Write("Название организации: ");
               rewrite_client_name = Console.ReadLine();
            }
            while (string.IsNullOrWhiteSpace(rewrite_fio))
            {
               Console.Write("ФИО нового контактного лица: ");
               rewrite_fio = Console.ReadLine();
            }

            try
            {
               // Загрузить файл Excel
               Workbook wb = new Workbook(path);

               // Получить рабочий лист, используя его индекс
               Worksheet worksheet = wb.Worksheets[1];

               //Поиск товара по имени в листе "Товары"
               if (worksheet.Name == "Клиенты")
               {
                  arr = ReadArray(worksheet, worksheet.Cells.MaxDataRow, worksheet.Cells.MaxDataColumn, rewrite_client_name);
                  if (arr.Count == 0)
                  {
                     return "Не найден клиент с таким именем!";
                  }
                  else
                  {
                     string[] names = new string[] { rewrite_fio };
                     var rewrite_str_bool = int.TryParse(arr[4], out int rewrite_str);
                     worksheet.Cells.ImportArray(names, rewrite_str, 3, true);
                     wb.Save(path);
                  }
                  //Console.WriteLine(arr[1].ToString());
               }
               else return "2 лист книги называться не 'Клиенты'! Проверьте!";
            }
            catch (System.IO.FileNotFoundException)
            {
               return "File not found";
            }
            catch (Exception ex)
            {
               return String.Concat("Error: " + ex.Message);
            }
            return "Данные изменены.";
         }

         string OpenFile()
         {
            bool product_id_bool;
            bool product_price_bool;
            bool bool_request_date;

            string product_name = string.Empty;
            while (string.IsNullOrWhiteSpace(product_name))
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

               // Получить рабочий лист "Товары", используя его индекс
               Worksheet worksheetProduct = wb.Worksheets[0];
               //Клиенты
               Worksheet worksheetClient = wb.Worksheets[1];
               //Заявки
               Worksheet worksheetRequest = wb.Worksheets[2];

               //Поиск товара по имени в листе "Товары"
               if (worksheetProduct.Name == "Товары")
               {
                  arr = ReadArray(worksheetProduct, worksheetProduct.Cells.MaxDataRow, worksheetProduct.Cells.MaxDataColumn, product_name);
                  if (arr.Count == 0)
                  {
                     return "Не найден товар с таким именем!";
                  }
                  product_id_bool = int.TryParse(arr[0], out product_id);
                  product_price_bool = double.TryParse(arr[3], out product_price);
               }

               //Поиск заявки по id товара
               if (worksheetRequest.Name == "Заявки")
               {
                  arr = ReadArray(worksheetRequest, worksheetRequest.Cells.MaxDataRow, worksheetRequest.Cells.MaxDataColumn, int_id: product_id);
                  request_id = int.Parse(arr[0]);
                  request_product = product_name;
                  request_client = int.Parse(arr[2]);
                  request_count = int.Parse(arr[4]);
                  bool_request_date = DateTime.TryParse(arr[5], out request_date);
               }

               //Поиск клиента по id из листа "Заявки"
               if (worksheetClient.Name == "Клиенты")
               {
                  arr = ReadArray(worksheetClient, worksheetClient.Cells.MaxDataRow, worksheetClient.Cells.MaxDataColumn, int_id: request_client);
                  client_name = arr[1];
                  client_address = arr[2];
               }

               if (request_id == 0 && product_id != 0)
               {
                  return "Заявки на данный товар не найдены.";
               }
            }
            catch (System.IO.FileNotFoundException)
            {
               return "File not found";
            }
            catch (Exception ex)
            {
               return String.Concat("Error: " + ex.Message);
            }
            return String.Concat("Товары: Код: " + product_id + ", Товар: " + arr[1] + ", Цена за ед.: " + product_price + "\n" +
               "Заявки: Код: " + request_id.ToString() + ", Товар: " + request_product + ", Код клиента: " +
                request_client.ToString() + ", Количество: " + request_count.ToString() + ", Дата заявки: " + $"{request_date.ToString("d")}" + "\n" +
                "Клиент: Код: " + request_client + ", Наименование организации: " + client_name + ", Адрес: " + client_address);
         }

         List<string> ReadArray(Worksheet worksheet, int rows, int cols,
            string? str1 = default, int? int_id = 0)
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
               // Записать строку в отдельный лист
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
                  if (int_id != 0 && int.Parse(arr[0]) == int_id)
                     return arr;
                  else if (!string.IsNullOrEmpty(str1) && arr[1] == str1)
                  { 
                     arr.Add(i.ToString());
                     return arr;
                  }
                  else arr.Clear();
               }
            }
            return arr;
         }
      }
   }
}
