using Aspose.Cells;

namespace Test
{
   class Program
   {
      static void Main(string[] args)
      {
         var path = @"C:\Users\Rail\Desktop\124\DirectumRX\DirRX\Практика.xlsx"; //string.Empty;
         var action = 0;
         while (action != 1)
         {
            Console.WriteLine("1. Выход");
            Console.WriteLine("2. Запрос на ввод пути до файла с данными (в качестве документа с данными использовать Приложение 2).");
            Console.WriteLine("3. По наименованию товара выводить информацию о клиентах, заказавших этот товар, " +
               "с указанием информации по количеству товара, цене и дате заказа.");
            Console.WriteLine("4. Запрос на изменение контактного лица клиента с указанием параметров: " +
               "Название организации, ФИО нового контактного лица." +
               "В результате информация должна быть занесена в этот же документ, " +
               "в качестве ответа пользователю необходимо выдавать информацию о результате изменений.");
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
                  OpenFile();
                  Console.WriteLine();
                  break;
               default:
                  break;
            }
         }
         void OpenFile()
         {
            string product_name = string.Empty;
            while (product_name.Length == 0)
            {
               Console.Write("Наименование товара: ");
               product_name = Console.ReadLine();
            }

            try
            {
               int product_id = 0;
               var product_price = 0.00;

               int client_id = 0;
               string client_name = string.Empty;
               string client_address = string.Empty;

               int request_id = 0;
               string request_product = String.Empty;
               int request_client = 0;
               var request_count = 0.00;
               DateTime request_date;

               // Загрузить файл Excel
               Workbook wb = new Workbook(path);

               // Получить все рабочие листы
               WorksheetCollection collection = wb.Worksheets;

               List<string> arr = new List<string>();

               // Перебрать все рабочие листы
               for (int worksheetIndex = 0; worksheetIndex < collection.Count; worksheetIndex++)
               {
                  // Получить рабочий лист, используя его индекс
                  Worksheet worksheet = collection[worksheetIndex];

                  // Получить количество строк и столбцов
                  int rows = worksheet.Cells.MaxDataRow;
                  int cols = worksheet.Cells.MaxDataColumn;

                  // Цикл по строкам
                  for (int i = 1; i <= rows; i++)
                  {
                     //Очищаю массив перед записью
                     arr.Clear();
                     // Перебрать каждый столбец в выбранной строке
                     for (int j = 0; j <= cols; j++)
                     {
                        arr.Add(worksheet.Cells[i, j].Value.ToString());
                        //Console.Write(worksheet.Cells[i, j].Value + " | ");
                     }
                     // Записать листы в отдельные таблицы
                     if (worksheet.Name == "Товары" && arr[1] == product_name)
                     {
                        product_id = int.Parse(arr[0]);
                        product_price = double.Parse(arr[3]);
                        Console.WriteLine("Товары: Код: " + product_id + ", Товар: " + product_name + ", Цена за ед.: " + product_price);
                     }
                     if (worksheet.Name == "Заявки" && int.Parse(arr[1]) == product_id)
                     {
                        request_id = int.Parse(arr[0]);
                        request_product = product_name;
                        request_client = int.Parse(arr[2]);
                        request_count = int.Parse(arr[4]);
                        var bool_request_date = DateTime.TryParse(arr[5], out request_date);
                        Console.WriteLine("Заявки: Код: " + request_id.ToString() + ", Товар: " + request_product + ", Код клиента: " +
                           request_client.ToString() + ", Количество: " + request_count.ToString() + ", Дата заявки: " + $"{request_date.ToString("d")}");
                     }
                     //Console.WriteLine(" ");
                  }
               }

               // Перебрать все рабочие листы
               for (int worksheetIndex = 0; worksheetIndex < collection.Count; worksheetIndex++)
               {
                  // Получить рабочий лист, используя его индекс
                  Worksheet worksheet = collection[worksheetIndex];

                  // Цикл по строкам
                  for (int i = 1; i <= worksheet.Cells.MaxDataRow; i++)
                  {
                     //Очищаю массив перед записью
                     arr.Clear();
                     // Перебрать каждый столбец в выбранной строке
                     for (int j = 0; j <= worksheet.Cells.MaxDataColumn; j++)
                     {
                        arr.Add(worksheet.Cells[i, j].Value.ToString());
                     }
                     // Записать листы в отдельные 
                     if (worksheet.Name == "Клиенты" && int.Parse(arr[0]) == request_client)
                     {
                        client_name = arr[1];
                        client_address = arr[2];
                        Console.WriteLine("Клиент: Код: " + request_client + ", Наименование организации: " + client_name + ", Адрес: " + client_address);
                     }
                  }
               }

               if (request_id == 0)
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
      }
   }
}
