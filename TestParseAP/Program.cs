using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Runtime.Serialization;
using Newtonsoft.Json;
using TestParseAP.Data;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop;
using System.Runtime.InteropServices;
using System.Threading;
using System.IO;

namespace TestParseAP
{

    class Program
    {
        private static Application xlApp1;
        private static Workbook xlWorkbook1;
        private static dynamic xlWorksheet1;
        private static dynamic xlRange1;
        private static Range excelcells;
        private static Range excelcells_key;
        private static Range excelcells_value;
        private static Range excelcells361;

        public class pairs
        {
          public  int start { get; set; }
          public  int end { get; set; }
            public pairs(int _start,int _end)
            {
                this.start = _start;
                this.end = _end;
            } }
        static void Main(string[] args)
        {
            Microsoft.Office.Interop.Excel.Application xlApp;
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook;
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet;
            Microsoft.Office.Interop.Excel.Range xlRange;
            Microsoft.Office.Interop.Excel.Range excelcells1;
            Microsoft.Office.Interop.Excel.Range excelcells2; 
            Microsoft.Office.Interop.Excel.Range excelcells3;
            Microsoft.Office.Interop.Excel.Range excelcells4;
            Microsoft.Office.Interop.Excel.Range excelcells5;
            Microsoft.Office.Interop.Excel.Range excelcells6;
            Microsoft.Office.Interop.Excel.Range excelcells7;
            Microsoft.Office.Interop.Excel.Range excelcells8;
            Microsoft.Office.Interop.Excel.Range excelcells9;
            Microsoft.Office.Interop.Excel.Range excelcells10;
            Microsoft.Office.Interop.Excel.Range excelcells11;
            Microsoft.Office.Interop.Excel.Range excelcells12;
            Microsoft.Office.Interop.Excel.Range excelcells13;
            Microsoft.Office.Interop.Excel.Range excelcells14;
            Microsoft.Office.Interop.Excel.Range excelcells15;
            Microsoft.Office.Interop.Excel.Range excelcells16;
            Microsoft.Office.Interop.Excel.Range excelcells17;
            Microsoft.Office.Interop.Excel.Range excelcells18;
            Microsoft.Office.Interop.Excel.Range excelcells19;
            Microsoft.Office.Interop.Excel.Range excelcells20;
            Microsoft.Office.Interop.Excel.Range excelcells21;
            Microsoft.Office.Interop.Excel.Range excelcells22;
            Microsoft.Office.Interop.Excel.Range excelcells23;
            Microsoft.Office.Interop.Excel.Range excelcells24;
            Microsoft.Office.Interop.Excel.Range excelcells25;
            Microsoft.Office.Interop.Excel.Range excelcells26;
            Microsoft.Office.Interop.Excel.Range excelcells27;
            Microsoft.Office.Interop.Excel.Range excelcells28;
            Microsoft.Office.Interop.Excel.Range excelcells29;
            Microsoft.Office.Interop.Excel.Range excelcells30;
            Microsoft.Office.Interop.Excel.Range excelcells31;
            Microsoft.Office.Interop.Excel.Range excelcells32;
            Microsoft.Office.Interop.Excel.Range excelcells33;
            Microsoft.Office.Interop.Excel.Range excelcells34;
            Microsoft.Office.Interop.Excel.Range excelcells35;
            Microsoft.Office.Interop.Excel.Range excelcells36;
            Microsoft.Office.Interop.Excel.Range excelcells37;
            Microsoft.Office.Interop.Excel.Range excelcells38;
            Microsoft.Office.Interop.Excel.Range excelcells39;
            Microsoft.Office.Interop.Excel.Range excelcells40;
            Microsoft.Office.Interop.Excel.Range excelcells41;
            Microsoft.Office.Interop.Excel.Range excelcells42;
            Microsoft.Office.Interop.Excel.Range excelcells43;
            Microsoft.Office.Interop.Excel.Range excelcells44;
            Microsoft.Office.Interop.Excel.Range excelcells45;
            Microsoft.Office.Interop.Excel.Range excelcells46;
            Microsoft.Office.Interop.Excel.Range excelcells47;
            Microsoft.Office.Interop.Excel.Range excelcells48;
            Microsoft.Office.Interop.Excel.Range excelcells49;
            Microsoft.Office.Interop.Excel.Range excelcells50;
            Microsoft.Office.Interop.Excel.Range excelcells51;
            Microsoft.Office.Interop.Excel.Range excelcells52;
            Microsoft.Office.Interop.Excel.Range excelcells53;
            Microsoft.Office.Interop.Excel.Range excelcells54;
            Microsoft.Office.Interop.Excel.Range excelcells55;
            Microsoft.Office.Interop.Excel.Range excelcells56;
            Microsoft.Office.Interop.Excel.Range excelcells57;
            Microsoft.Office.Interop.Excel.Range excelcells58;
            Microsoft.Office.Interop.Excel.Range excelcells59;
            Microsoft.Office.Interop.Excel.Range excelcells60;
            Microsoft.Office.Interop.Excel.Range excelcells61;
            Microsoft.Office.Interop.Excel.Range excelcells62;
            Microsoft.Office.Interop.Excel.Range excelcells63;
            Microsoft.Office.Interop.Excel.Range excelcells64;
            Microsoft.Office.Interop.Excel.Range excelcells65;
            Microsoft.Office.Interop.Excel.Range excelcells66;
            Microsoft.Office.Interop.Excel.Range excelcells67;
            Microsoft.Office.Interop.Excel.Range excelcells68;
            Microsoft.Office.Interop.Excel.Range excelcells69;
            Microsoft.Office.Interop.Excel.Range excelcells70;
            Microsoft.Office.Interop.Excel.Range excelcells71;
            Microsoft.Office.Interop.Excel.Range excelcells72;
            Microsoft.Office.Interop.Excel.Range excelcells73;
            Microsoft.Office.Interop.Excel.Range excelcells74;
            Microsoft.Office.Interop.Excel.Range excelcells75;

            //
            Dictionary<pairs, int> keyValuePairs = new Dictionary<pairs, int>();
            keyValuePairs.Add(new pairs(50,99),45);
            keyValuePairs.Add(new pairs(100, 199), 80);
            keyValuePairs.Add(new pairs(200, 299), 120);
            keyValuePairs.Add(new pairs(300, 349), 120);
            keyValuePairs.Add(new pairs(350, 399), 120);
            keyValuePairs.Add(new pairs(400, 449), 160);
            keyValuePairs.Add(new pairs(450, 499), 160);
            keyValuePairs.Add(new pairs(500, 549), 185);
            keyValuePairs.Add(new pairs(550, 599), 185);
            keyValuePairs.Add(new pairs(600, 699), 210);
            keyValuePairs.Add(new pairs(700, 799), 230);
            keyValuePairs.Add(new pairs(800, 899), 250);
            keyValuePairs.Add(new pairs(900, 999), 275);
            keyValuePairs.Add(new pairs(1000, 1199), 325);
            keyValuePairs.Add(new pairs(1200, 1299), 350);
            keyValuePairs.Add(new pairs(1300, 1399), 365);
            keyValuePairs.Add(new pairs(1400, 1499), 380);
            keyValuePairs.Add(new pairs(1500, 1599), 390);
            keyValuePairs.Add(new pairs(1600, 1699), 410);
            keyValuePairs.Add(new pairs(1700, 1799), 420);
            keyValuePairs.Add(new pairs(1800, 1999), 440);
            keyValuePairs.Add(new pairs(2000, 2399), 470);
            keyValuePairs.Add(new pairs(2400, 2499), 480);
            keyValuePairs.Add(new pairs(2500, 2999), 500);
            keyValuePairs.Add(new pairs(3000, 3499), 600);
            keyValuePairs.Add(new pairs(3500, 3999), 650);
            keyValuePairs.Add(new pairs(4000, 4499), 675);
            keyValuePairs.Add(new pairs(4500, 4999), 700);
            //
            List<string> car_id = new List<string>();
            string pathfile = @"C:\Users\kurulo\Desktop\autousa\1\cars_copart.xlsx";
            //           
            xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlWorkbook = xlApp.Workbooks.Open(pathfile);
            xlWorksheet = xlWorkbook.Sheets[1];
            xlRange = xlWorksheet.UsedRange;
            int colCount = xlRange.Columns.Count;
            int colRows = xlRange.Rows.Count;
            for (int j = 2; j <= colRows; j++)
            {
                excelcells1 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[j, 2];
                car_id.Add(Convert.ToString(excelcells1.Value2));
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            //
            List<Car_atr> car_Atrs1 = new List<Car_atr>();
            List<Car_img> car_Imgs1 = new List<Car_img>();
            string url = "";
            for (int i = 1; i < car_id.Count; i++)
            {
                url = "https://www.copart.com/public/data/lotdetails/solr/" + car_id[i] + "";
                string json_atr = new WebClient().DownloadString(url);
                var car_atr = JsonConvert.DeserializeObject<Car_atr>(json_atr);
                string json_img = new WebClient().DownloadString("https://www.copart.com/public/data/lotdetails/solr/lotImages/" + car_id[i] + "/USA");
                var car_img = JsonConvert.DeserializeObject<Car_img>(json_img);
                car_Atrs1.Add(car_atr);
                car_Imgs1.Add(car_img);
                Console.WriteLine(car_id.Count - i);
                Thread.Sleep(1000);
            }
            Console.WriteLine("end parsing");

            //serial
            using (StreamWriter file = File.CreateText(@"car_Atrs.json"))
            {
                JsonSerializer serializer = new JsonSerializer();
                serializer.Serialize(file, car_Atrs1);
            }
            using (StreamWriter file = File.CreateText(@"car_Imgs.json"))
            {
                JsonSerializer serializer = new JsonSerializer();
                serializer.Serialize(file, car_Imgs1);
            }

            //desereal
            List<Car_atr> car_Atrs = new List<Car_atr>();
            List<Car_img> car_Imgs = new List<Car_img>();
            using (StreamReader file = File.OpenText(@"car_Atrs.json"))
            {
                JsonSerializer serializer = new JsonSerializer();
                car_Atrs = (List<Car_atr>)serializer.Deserialize(file, typeof(List<Car_atr>));
            }
            Console.WriteLine("serial");
            using (StreamReader file = File.OpenText(@"car_Imgs.json"))
            {
                JsonSerializer serializer = new JsonSerializer();
                car_Imgs = (List<Car_img>)serializer.Deserialize(file, typeof(List<Car_img>));
            }
            //
            //string pathfile2 = @"C:\Users\kurulo\Desktop\autousa\1\export-products-06-05-19_14-43-35.xlsx";
            //xlApp1 = new Microsoft.Office.Interop.Excel.Application();
            //xlWorkbook1 = xlApp1.Workbooks.Open(pathfile2);
            //xlWorksheet1 = xlWorkbook1.Sheets[2];
            //xlRange1 = xlWorksheet1.UsedRange;
            //int colCount11 = xlRange1.Columns.Count;
            //int colRows11 = xlRange1.Rows.Count;
            Dictionary<string, string> group_names = new Dictionary<string, string>();
            //for (int i = 1; i < colRows11; i++)
            //{
            //    excelcells_key = (Microsoft.Office.Interop.Excel.Range)xlWorksheet1.Cells[i, 1];
            //    excelcells_value = (Microsoft.Office.Interop.Excel.Range)xlWorksheet1.Cells[i, 2];
            //    group_names.Add(key: Convert.ToString(excelcells_key.Value2), value: excelcells_value.Value2);
            //}
            //using (StreamWriter file = File.CreateText(@"car_groups.json"))
            //{
            //    JsonSerializer serializer = new JsonSerializer();
            //    serializer.Serialize(file, group_names);
            //}
            //using (StreamReader file = File.OpenText(@"car_groups.json"))
            //{
            //    JsonSerializer serializer = new JsonSerializer();
            //    group_names = (Dictionary<string, string>)serializer.Deserialize(file, typeof(Dictionary<string, string>));
            //}
            //            
            string pathfile1 = @"C:\Users\kurulo\Desktop\autousa\1\export-prom.xlsx";
            xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlWorkbook = xlApp.Workbooks.Open(pathfile1);
            xlWorksheet = xlWorkbook.Sheets[1];
            xlRange = xlWorksheet.UsedRange;
            int colCount1 = xlRange.Columns.Count;
            int colRows1 = xlRange.Rows.Count;
            {
                excelcells1 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 1];
                excelcells1.Value2 = "Код_товара";
                excelcells2 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 2];
                excelcells2.Value2 = "Название_позиции";
                excelcells3 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 3];
                excelcells3.Value2 = "Поисковые_запросы";
                excelcells4 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 4];
                excelcells4.Value2 = "Описание";
                excelcells5 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 5];
                excelcells5.Value2 = "Тип_товара";
                excelcells6 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 6];
                excelcells6.Value2 = "Цена";
                excelcells7 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1,7];
                excelcells7.Value2 = "Валюта";
                excelcells8 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 8];
                excelcells8.Value2 = "Единица_измерения";
                excelcells9 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 9];
                excelcells9.Value2 = "Минимальный_объем_заказа";
                excelcells10 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 10];
                excelcells10.Value2 = "Оптовая_цена";
                excelcells11 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 11];
                excelcells11.Value2 = "Минимальный_заказ_опт";
                excelcells12 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 12];
                excelcells12.Value2 = "Ссылка_изображения";
                excelcells13 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 13];
                excelcells13.Value2 = "Наличие";
                excelcells14 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 14];
                excelcells14.Value2 = "Количество";
                excelcells15 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 15];
                excelcells15.Value2 = "Номер_группы";
                excelcells16 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 16];
                excelcells16.Value2 = "Название_группы";
                excelcells17 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 17];
                excelcells17.Value2 = "Адрес_подраздела";
                excelcells18 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 18];
                excelcells18.Value2 = "Возможность_поставки";
                excelcells19 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 19];
                excelcells19.Value2 = "Срок_поставки";
                excelcells20 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 20];
                excelcells20.Value2 = "Способ_упаковки";
                excelcells21 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 21];
                excelcells21.Value2 = "Уникальный_идентификатор";
                excelcells22 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 22];
                excelcells22.Value2 = "Идентификатор_товара";
                excelcells23 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 23];
                excelcells23.Value2 = "Идентификатор_подраздела";
                excelcells24 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 24];
                excelcells24.Value2 = "Идентификатор_группы";
                excelcells25 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 25];
                excelcells25.Value2 = "Производитель";
                excelcells26 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 26];
                excelcells26.Value2 = "Страна_производитель";
                excelcells27 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 27];
                excelcells27.Value2 = "Скидка";
                excelcells28 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 28];
                excelcells28.Value2 = "ID_группы_разновидностей";
                excelcells29 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 29];
                excelcells29.Value2 = "Личные_заметки";
                excelcells30 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 30];
                excelcells30.Value2 = "Продукт_на_сайте";
                excelcells31 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 31];
                excelcells31.Value2 = "Cрок действия скидки от";
                excelcells32 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 32];
                excelcells32.Value2 = "Cрок действия скидки до";
                excelcells33 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1,33];
                excelcells33.Value2 = "Цена от";
                excelcells34 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 34];
                excelcells34.Value2 = "Ярлык";
                excelcells35 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 35];
                excelcells35.Value2 = "HTML_заголовок";
                excelcells36 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 36];
                excelcells36.Value2 = "HTML_описание";
                excelcells37 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 37];
                excelcells37.Value2 = "HTML_ключевые_слова";
                //excelcells38 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 38];
                //excelcells38.Value2 = "Название_Характеристики";//Тип кузова
                //excelcells39 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 39];
                //excelcells39.Value2 = "Измерение_Характеристики";
                //excelcells40 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 40];
                //excelcells40.Value2 = "Значение_Характеристики";
                excelcells41 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 41];
                excelcells41.Value2 = "Название_Характеристики"; //Коробка переключения передач
                excelcells42 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 42];
                excelcells42.Value2 = "Измерение_Характеристики";
                excelcells43 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 43];
                excelcells43.Value2 = "Значение_Характеристики";
                excelcells44 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 44];
                excelcells44.Value2 = "Название_Характеристики";//Тип привода колес
                excelcells45 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 45];
                excelcells45.Value2 = "Измерение_Характеристики";
                excelcells46 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 46];
                excelcells46.Value2 = "Значение_Характеристики";
                excelcells47 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 47];
                excelcells47.Value2 = "Название_Характеристики";//Цвет
                excelcells48 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 48];
                excelcells48.Value2 = "Измерение_Характеристики";
                excelcells49 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 49];
                excelcells49.Value2 = "Значение_Характеристики";
                excelcells50 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 50];
                excelcells50.Value2 = "Название_Характеристики";//Объем двигателя
                excelcells51 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 51];
                excelcells51.Value2 = "Измерение_Характеристики";
                excelcells52 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 52];
                excelcells52.Value2 = "Значение_Характеристики";
                excelcells53 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 53];
                excelcells53.Value2 = "Название_Характеристики";//Тип легкового автомобиля
                excelcells54 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 54];
                excelcells54.Value2 = "Измерение_Характеристики";
                excelcells55 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1,55];
                excelcells55.Value2 = "Значение_Характеристики";
                excelcells56 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 56];
                excelcells56.Value2 = "Название_Характеристики";//Тип двигателя
                excelcells57 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 57];
                excelcells57.Value2 = "Измерение_Характеристики";
                excelcells58 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 58];
                excelcells58.Value2 = "Значение_Характеристики";
                excelcells59 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 59];
                excelcells59.Value2 = "Название_Характеристики";//Марка
                excelcells60 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 60];
                excelcells60.Value2 = "Измерение_Характеристики";
                excelcells61 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 61];
                excelcells61.Value2 = "Значение_Характеристики";
                excelcells62 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 62];
                excelcells62.Value2 = "Название_Характеристики";//Год выпуска
                excelcells63 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 63];
                excelcells63.Value2 = "Измерение_Характеристики";
                excelcells64 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 64];
                excelcells64.Value2 = "Значение_Характеристики";
                excelcells65 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 65];
                excelcells65.Value2 = "Название_Характеристики";//Состояние
                excelcells66 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 66];
                excelcells66.Value2 = "Измерение_Характеристики";
                excelcells67 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 67];
                excelcells67.Value2 = "Значение_Характеристики";
                excelcells68 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 68];
                excelcells68.Value2 = "Название_Характеристики";//Пробіг
                excelcells69 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 69];
                excelcells69.Value2 = "Измерение_Характеристики";
                excelcells70 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 70];
                excelcells70.Value2 = "Значение_Характеристики";
                //
                excelcells71 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 71];
                excelcells71.Value2 = "Название_Характеристики";//Пробіг
                excelcells72 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 72];
                excelcells72.Value2 = "Измерение_Характеристики";
                excelcells73 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[1, 73];
                excelcells73.Value2 = "Значение_Характеристики";
            }
                for (int i = 2; i < car_Imgs.Count; i++)
                {
                try
                {
                    if (car_Atrs[i].data.lotDetails.la == 0 || car_Atrs[i].data.lotDetails.egn == null || car_Atrs[i].data.lotDetails == null) { goto end; }
                }
                catch { goto end; }
                    string images_path = "";
                //
                    excelcells1 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 1];
                    excelcells1.Value2 = car_Atrs[i].data.lotDetails.lotNumberStr+"-"+i.ToString();
                //
                string car_name = car_Atrs[i].data.lotDetails.mkn + " " + car_Atrs[i].data.lotDetails.lm + " "+ car_Atrs[i].data.lotDetails.lcy.ToString();
                    excelcells2 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 2];
                    excelcells2.Value2 = car_name;
                //
                //

                excelcells3 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 3];//ключы
                excelcells3.Value2 =car_Atrs[i].data.lotDetails.mkn + ", "+ car_Atrs[i].data.lotDetails.lm + ", авто, машина, автомобиль, аукцион, США, Америка, растаможка, расстаможка, таможня, рассчитать, акциз, налог, бу, б у, б/у, новая, битая, ремонт, под, и, копарт, копард, таможенная, очистка, грузовых, деклараций, консалтинг, аукциона, заказать, бит, copart, автоаукцион, пригнать, авто из Америки, из, ставка, розмитнити, розмитнення, аукціон, та, автомобіль, штатов, ретро, придбати, замовити, подержанные, поддержанные, Канады, вживані, старые, Львів, Киев, авто сша, податок, регистрация, реєстрація, iaai, Manheim, бмв, bmw, мерседес, mersedes, Cars, брокерские, брокер, услуги, сертификат, постановка, на учёт, ford, volkswagen, шевроле, субару, хонда, honda, nissan, ниссан, ауди, audi, электрокар, электро, вольт, volt, приус, Prius, мото, таможенных, услуга, авто из Америки в Украине, пром, prom, авто из сша в Украине, авто из США, пригон, пригон авто из сша, подбор авто из сша, авто из сша, Украина, купить, предложение, цена, експорт, машини";
                //               //
                excelcells5 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 5];
                    excelcells5.Value2 = "r";
                //
                //
                excelcells7 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 7];
                    excelcells7.Value2 = "USD";
                //
                excelcells8 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 8];
                    excelcells8.Value2 = "шт.";
                //
                excelcells9 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 9];
                    excelcells9.Value2 = " ";
                //
                excelcells10 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 10];
                    excelcells10.Value2 =" ";
                //
                excelcells11 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 11];
                    excelcells11.Value2 = " ";
                //img
                excelcells12 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 12];
                try
                {
                    if (car_Imgs[i].data.imagesList.HIGH_RESOLUTION_IMAGE != null)
                    {
                        for (int k = 0; k < car_Imgs[i].data.imagesList.HIGH_RESOLUTION_IMAGE.Count; k++)
                        {
                            images_path += car_Imgs[i].data.imagesList.HIGH_RESOLUTION_IMAGE[k].url + ", ";
                        }
                    }
                    else
                    {
                        for (int k = 0; k < car_Imgs[i].data.imagesList.FULL_IMAGE.Count; k++)
                        {
                            images_path += car_Imgs[i].data.imagesList.FULL_IMAGE[k].url + ", ";
                        }
                    }
                }
                catch { goto end; }
                excelcells12.Value2 = images_path;
                //колір
                excelcells13 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 13];
                    excelcells13.Value2 = "45";
                //
                excelcells14 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 14];
                    excelcells14.Value2 =" ";
                //номер групи
                excelcells15 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 15];
                foreach (var item in group_names)
                {
                    if (item.Value.ToLower() == car_Atrs[i].data.lotDetails.mkn.ToLower())
                    {
                        excelcells15.Value2 = item.Key;
                    }
                }
                //
                excelcells16 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 16];
                    excelcells16.Value2 =" ";
                //
                excelcells17 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 17];
                    excelcells17.Value2 = @"https://prom.ua/Legkovye-avtomobili";
                //     
                excelcells18 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 18];
                excelcells18.Value2 = " ";
                //
                excelcells19 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 19];
                excelcells19.Value2 = " ";
                //ключи ?
                excelcells20 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 20];
                excelcells20.Value2 = " ";
                excelcells21 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 21];
                excelcells21.Value2 = " ";
                excelcells22 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 22];
                excelcells22.Value2 = " ";
                excelcells23 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 23];
                excelcells23.Value2 = " ";
                excelcells24 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 24];
                excelcells24.Value2 = " ";
                excelcells25 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 25];
                excelcells25.Value2 = " ";
                excelcells26 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 26];
                excelcells26.Value2 = " ";
                excelcells27 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 27];
                excelcells27.Value2 = "100.00";
                excelcells28 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 28];
                excelcells28.Value2 = " ";
                excelcells29 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 29];
                excelcells29.Value2 = " ";
                excelcells30 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 30];
                excelcells30.Value2 = "";
                excelcells31 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 31];
                excelcells31.Value2 = "03.05.2019";
                excelcells32 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 32];
                excelcells32.Value2 = "17.06.2019";
                excelcells33 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 33];
                excelcells33.Value2 = "+";
                excelcells34 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 34];
                excelcells34.Value2 = "Новинка";
                excelcells35 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 35];
                excelcells35.Value2 = " ";
                excelcells36 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 36];
                excelcells36.Value2 = " ";

                excelcells361 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 37];
                excelcells361.Value2 = " ";
                /////////////////////******
                //excelcells37 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 38];
                //excelcells37.Value2 = "Тип кузова";
                //excelcells38 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 39];
                //excelcells38.Value2 = " ";
                //excelcells39 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 40];
                //excelcells39.Value2 = car_Atrs[i].data.lotDetails.bstl!=" "? car_Atrs[i].data.lotDetails.bstl:" ";
                excelcells40 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 41];
                excelcells40.Value2 = "Коробка переключения передач";
                excelcells41 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 42];
                excelcells41.Value2 = " ";
                excelcells42 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 43];
                excelcells42.Value2 = car_Atrs[i].data.lotDetails.tsmn != null ? car_Atrs[i].data.lotDetails.tsmn : "__";
                excelcells43 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 44];
                excelcells43.Value2 = "Тип привода колес";
                excelcells44 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 45];
                excelcells44.Value2 = " ";
                excelcells45 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 46];
                excelcells45.Value2 = car_Atrs[i].data.lotDetails.drv != null ? car_Atrs[i].data.lotDetails.drv : "__";
                excelcells46 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 47];
                excelcells46.Value2 = "Цвет";
                excelcells47 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 48];
                excelcells47.Value2 = " ";
                excelcells48 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 49];
                excelcells48.Value2 = car_Atrs[i].data.lotDetails.clr != null ? car_Atrs[i].data.lotDetails.clr : "__";
                excelcells49 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 50];
                excelcells49.Value2 = "Объем двигателя";
                excelcells50 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 51];
                excelcells50.Value2 = "куб. см";
                excelcells51 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 52];
                string[] words1 = car_Atrs[i].data.lotDetails.egn.Split(new char[] { 'L' });
                double eng1 = Convert.ToDouble(words1[0]);
                excelcells51.Value2 = words1 != null ? (eng1*1000).ToString() : "__";
                excelcells52 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 53];
                excelcells52.Value2 = "Тип легкового автомобиля";
                excelcells53 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 54];
                excelcells53.Value2 = " ";
                excelcells54 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 55];
                excelcells54.Value2 = "Б/у";
                excelcells55 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 56];
                excelcells55.Value2 = "Тип двигателя";
                excelcells56 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 57];
                excelcells56.Value2 = " ";
                excelcells57 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 58];
                excelcells57.Value2 = car_Atrs[i].data.lotDetails.ft != null ? car_Atrs[i].data.lotDetails.ft : "__";
                excelcells58 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 59];
                excelcells58.Value2 = "Марка";
                excelcells59 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 60];
                excelcells59.Value2 = " ";
                excelcells60 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 61];
                excelcells60.Value2 = car_Atrs[i].data.lotDetails.mkn != null ? car_Atrs[i].data.lotDetails.mkn : "__";
                excelcells61 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 62];
                excelcells61.Value2 = "Год выпуска";
                excelcells62 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 63];
                excelcells62.Value2 = " ";
                excelcells63 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 64];
                excelcells63.Value2 = car_Atrs[i].data.lotDetails.lcy.ToString() !=null ? car_Atrs[i].data.lotDetails.lcy.ToString() : "__";
                excelcells64 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 65];
                excelcells64.Value2 = "Состояние";
                excelcells65 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 66];
                excelcells65.Value2 = " ";
                excelcells66 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 67];
                excelcells66.Value2 = "Б/У";
                excelcells67 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 68];
                excelcells67.Value2 = "Модель";
                excelcells68 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 69];
                excelcells68.Value2 = " ";
                excelcells69 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 70];
                excelcells69.Value2 = car_Atrs[i].data.lotDetails.lm!= null ? car_Atrs[i].data.lotDetails.lm : "__";
                excelcells70 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 71];
                excelcells70.Value2 = "Пробіг";
                excelcells71 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 72];
                excelcells71.Value2 = "";
                excelcells72 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 73];
                excelcells72.Value2 = car_Atrs[i].data.lotDetails.orr.ToString() != null ? Convert.ToString(Math.Round(car_Atrs[i].data.lotDetails.orr*1.6))+" км." : "__";
                //опис
                //
                int price1 = 0;
                string price = "";
                int copartfee=0;
                int sum1 = 0;
                double fee1 = 0;
                double fee2 = 0;
                double fee3 = 0;
                if (car_Atrs[i].data.lotDetails.bnp != 0.0)
                {
                    price = Convert.ToString(car_Atrs[i].data.lotDetails.bnp) + " купити вже!";
                    price1 =Convert.ToInt32( car_Atrs[i].data.lotDetails.bnp);
                }
                else
                {
                    price1 = Convert.ToInt32(Math.Round(car_Atrs[i].data.lotDetails.la / 200, 0)*100);
                    if (price1 < 10000)
                    {
                        price = Convert.ToString(price1 - 500) + " - " + Convert.ToString(price1 + 500) + " орієнтована виграшна";
                    }
                    else { price = Convert.ToString(price1 - 1000) + " - " + Convert.ToString(price1 + 1000) + " орієнтована виграшна"; }
                    
                }                
                foreach (var item in keyValuePairs)
                {
                    if (price1 < 5000)
                    {
                        if (price1 >= item.Key.start && price1 <= item.Key.end)
                        {
                            copartfee = item.Value;
                        }
                    }
                    else
                    {
                        copartfee = (price1 *15)/100;//15%
                    }
                }
                sum1 = 1000 + 59 + copartfee + price1;
                //
                if (car_Atrs[i].data.lotDetails.egn == "U" || car_Atrs[i].data.lotDetails.ft == "ELECTRIC")
                {
                    fee1 = 0;
                    fee2 = 0;
                    fee3 = 0;
                    goto electro;
                }
                string[] words = car_Atrs[i].data.lotDetails.egn.Split(new char[] { 'L' });
                double eng = Convert.ToDouble(words[0]);
                //DIESEL
                //GAS
                //ELECTRIC
                //FLEXIBLE FUEL
                //HYBRID ENGINE
                
                int year = 2019 - car_Atrs[i].data.lotDetails.lcy;
                if (car_Atrs[i].data.lotDetails.lcy == 2019) { year = 1; }
                fee2 = (price1 + copartfee) / 10;//Ввізне мито
                fee3 = (int)(fee2 + fee1 + price1 + copartfee) / 5;//ПДВ 20% 
                if (car_Atrs[i].data.lotDetails.ft == "GAS" || car_Atrs[i].data.lotDetails.ft== "FLEXIBLE FUEL" || car_Atrs[i].data.lotDetails.ft == "HYBRID ENGINE")
                {
                    if (eng <= 3.0) { fee1 = 57*year*eng; } else { fee1 = 113 * year * eng; }
                }
                if (car_Atrs[i].data.lotDetails.ft == "DIESEL")
                {
                    if (eng <= 3.5) { fee1 = 85 * year * eng; } else { fee1 = 170 * year * eng; }
                }
                electro:
                if (car_Atrs[i].data.lotDetails.ft == "ELECTRIC")
                {
                    fee1 = 0;
                }          
                string description = "<div class=\"ck-image-text-right ck-image-text-right_type_advanced ck-theme-blue\"> <div class=\"ck-image-text-right__text\"> <div class=\"ck-image-text-right__title\"><span class=\"fa fa-file-text-o\"style=\"font - size:24px; color:#1A12FF\"></span>   Інформація про машину</div>"+
"<p>VIN -&nbsp;<strong>" +car_Atrs[i].data.lotDetails.fv+"</strong></p>"+
"<p>Пробіг(км) -&nbsp;<strong>"+ Convert.ToInt32(car_Atrs[i].data.lotDetails.orr*1.6) + "</strong></p>"+
"<p>Рік -&nbsp;<strong>" + car_Atrs[i].data.lotDetails.lcy + "</strong></p>" +
"<p>Колір -&nbsp;<strong>" +car_Atrs[i].data.lotDetails.clr+"</strong></p>" +
"<p>Двигун -&nbsp;<strong>"+ car_Atrs[i].data.lotDetails.egn + "</strong></p>"+
"<p>Кількість циліндрів -&nbsp;<strong>"+ car_Atrs[i].data.lotDetails.cy + "</strong></p>"+
"<p>Привід -&nbsp;<strong>"+ car_Atrs[i].data.lotDetails.drv + "</strong></p>"+
"<p>Тип палива -&nbsp;<strong>"+ car_Atrs[i].data.lotDetails.ft + "</strong></p>"+
"<p>Наявність ключів -&nbsp;<strong>"+ car_Atrs[i].data.lotDetails.hk + "</strong></p>"+
"<p>Штат/Тип сертифіката -&nbsp;<strong>"+ car_Atrs[i].data.lotDetails.syn+ "</strong></p>"+
"<p>Статус -&nbsp;<strong>"+ car_Atrs[i].data.lotDetails.lcd + "</strong></p>"+
"<p>Основні&nbsp;пошкодження -&nbsp;<strong>"+ car_Atrs[i].data.lotDetails.dd + "</strong></p>"+
"<p>Інше &nbsp;пошкодження -&nbsp;<strong>" + car_Atrs[i].data.lotDetails.sdd + "</strong></p>" +
"</div>" +
"</div>"+
"<p>&nbsp;</p>"+
"<div class=\"ck-image-text-right ck-image-text-right_type_advanced ck-theme-green\">"+
"<div class=\"ck-image-text-right__text\">"+
"<div class=\"ck-image-text-right__title\"><span class=\"fa fa-gavel\" style=\"font-size:24px; color:#1A12FF\"></span>   Розрахунок вартості</div>"+
"<p>Ставка на аукціоні ($)&nbsp;<strong>" + price + "</strong></p>"+
"<p>Налог аукціону&nbsp;Копарт&nbsp;<strong> $"+copartfee.ToString()+"</strong></p>"+
"<p>Збір за обробку (Gate Fee) <strong>$59</strong></p>" +
"<p>Наша комісія&nbsp;&nbsp;<strong>&nbsp;$1000</strong>&nbsp;</p>"+
"</div>"+
"</div>" +
"<p><span style = \"font-size:18px; \" > <u> <strong> Сума вартості "+sum1.ToString()+" $</strong></u></span></p>" +
"<div class=\"ck-image-text-right ck-image-text-right_type_advanced ck-theme-green\">" +
"<div class=\"ck-image-text-right__text\">" +
"<div class=\"ck-image-text-right__title\"><span class=\"fa fa-anchor\" style=\"font - size:24px; color:#1A12FF\"></span>   Розрахунок доставки</div>" +
"<p>Суша&nbsp; &nbsp;(<strong>$) 200 - 500</strong>&nbsp;(Доставка з " + car_Atrs[i].data.lotDetails.yn + " в найближчий порт)</p>" +
"<p>Море&nbsp; &nbsp;<strong>$1200&nbsp; </strong>(США -&nbsp;Клайпеда , Литва)</p>" +
"<p>Клайпеда-Львів<strong> 500$&nbsp; </strong>(Доставка в Городок)</p>" +
"</div>" +
"</div>" +
"<p><span style = \"font-size:18px; \" > <u> <strong> Сума доставки 1900 $ </strong></u></span></p>" +
"<div class=\"ck-image-text-right ck-image-text-right_type_advanced ck-theme-green\">" +
"<div class=\"ck-image-text-right__text\">" +
"<div class=\"ck-image-text-right__title\"><span class=\"fa fa-share-square-o\" style=\"font-size:24px; color:#1A12FF\"></span>   Розрахунок розмитнення</div>" +
"<p>Ввізне мито&nbsp; &nbsp;<strong>$" + fee2.ToString() + "</strong></p>" +
"<p>Акцизний збір&nbsp; &nbsp;<strong>$"+fee1.ToString()+"</strong></p>" +
"<p>ПДВ &nbsp; &nbsp;<strong>$"+fee3.ToString()+"</strong></p>" +
"<p>Експедиція порт Клайпеда&nbsp; &nbsp;<strong>$300</strong></p>" +
"<p>Митний брокер&nbsp; &nbsp;<strong>$450</strong></p>" +
"</div> <p> <span style = \"font-size:18px; \" > <u> <strong> Сума розмитнення "+Math.Round(fee1+fee2+fee3+300+450).ToString()+" $ </strong> </u></span> </p>" +
"</div> <div class=\"ck-alert ck-alert_theme_red\">" +
"<div style =\"text-align: center;\"> <span class=\"ck-alert__title\"></span><span style =\"font-size:20px;\"> <strong> Загалом: "+ Math.Round(fee1 + fee2 + fee3 + 300 + 450 +1900+sum1).ToString()+ "$ </strong></span></div>" +
"<div style =\"text-align: center;\"> <span style=\"font-size:20px;\"></span><span style =\"font-size:12px;\"> <strong> (без сертифікації&nbsp;і&nbsp;постановки на облік) </strong></span></div>" +
"</div>" +
"<div class=\"ck-alert ck-alert_theme_blue\" style=\"text-align:center\"><span style = \"font-size:18px\" > Як відбувається замовлення авто?</span></div>" +
"<div class=\"ck-list-horizontal ck-list-horizontal_type_lite ck-theme-grey\">" +
"<div class=\"ck-list-horizontal__table\">" +
"<div class=\"ck-list-horizontal__table-item\">" +
"<div class=\"ck-list-horizontal__image-wrapper\"><img alt = \"\" class=\"ck-list-horizontal__image\" src=\"https://images.ua.prom.st/1745875270_1745875270.jpg?PIMAGE_ID=1745875270\" style=\"width:72px;height:72px\" /></div>"+
"<div class=\"ck-list-horizontal__text\">" +
"<div class=\"ck-list-horizontal__title\">Підбір авто</div>" +
"Ми підберемо для Вас декілька варіантів авто з аукціонів США відповідно до Ваших вимог і бюджету</div>" +
"</div>" +
"<div class=\"ck-list-horizontal__table-item ck-list-horizontal__table-item_type_narrow-10\">&nbsp;</div>" +
"<div class=\"ck-list-horizontal__table-item\">" +
"<div class=\"ck-list-horizontal__image-wrapper\"><img alt = \"\" class=\"ck-list-horizontal__image\" src=\"https://images.ua.prom.st/1745876128_1745876128.jpg?PIMAGE_ID=1745876128\" style=\"width:72px;height:72px\" /></div>"+
"<div class=\"ck-list-horizontal__text\">" +
"<div class=\"ck-list-horizontal__title\">Підписання договору</div>" +
"Підписання угоди дає можливість не турбуватись замовнику за гарантоване надання послуг</div>" +
"</div>" +
"<div class=\"ck-list-horizontal__table-item ck-list-horizontal__table-item_type_narrow-10\">&nbsp;</div>" +
"<div class=\"ck-list-horizontal__table-item\">" +
"<div class=\"ck-list-horizontal__image-wrapper\"><img alt = \"\" class=\"ck-list-horizontal__image\" src=\"https://images.ua.prom.st/1745876420_1745876420.jpg?PIMAGE_ID=1745876420\" style=\"width:72px;height:72px\" /></div>"+
"<div class=\"ck-list-horizontal__text\">" +
"<div class=\"ck-list-horizontal__title\">Торги та викуп авто</div>" +
"Торги і викуп проводяться згідно укладеного договору, в якому зазначена модель машини, бюджет та інші параметри</div>" +
"</div>" +
"<div class=\"ck-list-horizontal__table-item ck-list-horizontal__table-item_type_narrow-10\">&nbsp;</div>" +
"<div class=\"ck-list-horizontal__table-item\">" +
"<div class=\"ck-list-horizontal__image-wrapper\"><img alt = \"\" class=\"ck-list-horizontal__image\" src=\"https://images.ua.prom.st/1745879220_1745879220.jpg?PIMAGE_ID=1745879220\" style=\"width:72px;height:72px\" /></div>"+
"<div class=\"ck-list-horizontal__text\">" +
"<div class=\"ck-list-horizontal__title\">Доставка&nbsp;<br />" +
"Розмитнення</div>" +
"<p>Доставка відбувається на протязі 6-8 тижнів з дня покупки на аукціоні.Розмитнюється авто на майбутнього власника, відповідно до законів України.</p>" +
"</div>" +
"</div>" +
"</div>" +
"</div>" +
"<div class=\"ck-alert ck-alert_theme_blue\" style=\"text-align:center\"><span style = \"font-size:18px\" > ЯКУ ВИГОДУ ВИ ОТРИМАЄТЕ?</span></div>" +
"<div class=\"ck-list-horizontal ck-list-horizontal_type_lite ck-theme-grey\">" +
"<div class=\"ck-list-horizontal__table\">" +
"<div class=\"ck-list-horizontal__table-item\">" +
"<div class=\"ck-list-horizontal__image-wrapper\"><img alt = \"\" class=\"ck-list-horizontal__image\" src=\"https://images.ua.prom.st/1745949931_1745949931.jpg?PIMAGE_ID=1745949931\" style=\"width:72px;height:72px\" /></div>"+
"<div class=\"ck-list-horizontal__text\">" +
"<div class=\"ck-list-horizontal__title\">Ціна на авто нижча за ринкову на 20-40%</div>" +
"Завдяки змінам в законі 8487 реально зекономити до 40% на авто</div>" +
"</div>" +
"<div class=\"ck-list-horizontal__table-item ck-list-horizontal__table-item_type_narrow-45\">&nbsp;</div>" +
"<div class=\"ck-list-horizontal__table-item\">" +
"<div class=\"ck-list-horizontal__image-wrapper\"><img alt = \"\" class=\"ck-list-horizontal__image\" src=\"https://images.ua.prom.st/1745950183_1745950183.jpg?PIMAGE_ID=1745950183\" style=\"width:72px;height:72px\" /></div>"+
"<div class=\"ck-list-horizontal__text\">" +
"<div class=\"ck-list-horizontal__title\">Свіжі роки,&nbsp;мінімальний пробіг</div>" +
"Авто з аукціонів США не старше 2010 року та мають невеликий пробіг відносно європейців</div>" +
"</div>" +
"<div class=\"ck-list-horizontal__table-item ck-list-horizontal__table-item_type_narrow-45\">&nbsp;</div>" +
"<div class=\"ck-list-horizontal__table-item\">" +
"<div class=\"ck-list-horizontal__image-wrapper\"><img alt = \"\" class=\"ck-list-horizontal__image\" src=\"https://images.ua.prom.st/1745950524_1745950524.jpg?PIMAGE_ID=1745950524\" style=\"width:72px;height:72px\" /></div>"+
"<div class=\"ck-list-horizontal__text\">" +
"<div class=\"ck-list-horizontal__title\">Прозора процедура &quot;Під ключ&quot;</div>" +
"<p>Ми беремо всі питання по доставці з Америки, розмитненю в Україні та ремонту автомобіля на себе.Всі платежі прозорі та фіксовані</p>" +
"</div>" +
"</div>" +
"</div>" +
"</div>" +
"<div class=\"ck-list-horizontal ck-list-horizontal_type_lite ck-theme-grey\">" +
"<div class=\"ck-list-horizontal__table\">" +
"<div class=\"ck-list-horizontal__table-item\">" +
"<div class=\"ck-list-horizontal__image-wrapper\"><img alt = \"\" class=\"ck-list-horizontal__image\" src=\"https://images.ua.prom.st/1745950792_1745950792.jpg?PIMAGE_ID=1745950792\" style=\"width:72px;height:72px\" /></div>"+
"<div class=\"ck-list-horizontal__text\">" +
"<div class=\"ck-list-horizontal__title\">Ексклюзивні комплектації</div>" +
"В Американських авто йде набагато краща комплектація ніж аналогів в Україні, або Європі</div>" +
"</div>" +
"<div class=\"ck-list-horizontal__table-item ck-list-horizontal__table-item_type_narrow-45\">&nbsp;</div>" +
"<div class=\"ck-list-horizontal__table-item\">" +
"<div class=\"ck-list-horizontal__image-wrapper\"><img alt = \"\" class=\"ck-list-horizontal__image\" src=\"https://images.ua.prom.st/1745956989_1745956989.jpg?PIMAGE_ID=1745956989\" style=\"width:72px;height:72px\" /></div>"+
"<div class=\"ck-list-horizontal__text\">" +
"<div class=\"ck-list-horizontal__title\">Великий вибір різних моделей</div>" +
"На аукціонах США щодня продається десятки тисяч авто. Різноманіття моделей просто вражає</div>" +
"</div>" +
"<div class=\"ck-list-horizontal__table-item ck-list-horizontal__table-item_type_narrow-45\">&nbsp;</div>" +
"<div class=\"ck-list-horizontal__table-item\">" +
"<div class=\"ck-list-horizontal__image-wrapper\"><img alt = \"\" class=\"ck-list-horizontal__image\" src=\"https://images.ua.prom.st/1745950946_1745950946.jpg?PIMAGE_ID=1745950946\" style=\"width:72px;height:72px\" /></div>"+
"<div class=\"ck-list-horizontal__text\">" +
"<div class=\"ck-list-horizontal__title\">Повна історія автомобіля</div>" +
"Всі авто в США обслуговуються в офіційних дилерських авто сервісах.Надається повна історія обслуговування по авто</div>" +
"</div>" +
"</div>" +
"</div>" +
"<div class=\"ck-alert ck-alert_theme_blue\">" +
"<div style =\"text-align:center\"> <span style=\"font-size:18px\">Чому саме зараз варто замовляти авто з Америки?</span></div>" +
"</div>" +
"<div class=\"ck-image-text-left ck-image-text-left_type_lite ck-theme-grey\">" +
"<div class=\"ck-image-text-left__image-wrapper\"><img alt = \"\" class=\"ck-image-text-left__image\" src=\"https://images.ua.prom.st/1745844194_1745844194.jpg?PIMAGE_ID=1745844194\" style=\"width:417px;height:324px\" /></div>"+
"<div class=\"ck-image-text-left__text\">08 листопада 2018 року Верховна Рада України прийняла законопроекти 8487 і 8488 &laquo;Про внесення змін до Податкового кодексу України щодо оподаткування акцизним податком легкових транспортних засобів&raquo; в цілому, 23.11.2018 закон був підписаний президентом.Дані зміни в першу чергу спрямовані на врегулювання проблеми з автомобілями на європейських номерах, і тепер дозволяють узаконити так звані &laquo; евробляхи&raquo;. <u>Однак, дані зміни стосуються і всього автомобільного ринку України, в тому числі автомобілів з США.</u>" +
"<p>&nbsp;</p>" +
"</div>" +
"</div>";
                excelcells4 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 4];
                excelcells4.Value2 = description;
                excelcells6 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 6];
                excelcells6.Value2 = Math.Round(fee1 + fee2 + fee3 + 300 + 550 + 1900 + sum1).ToString();
                
                excelcells22 = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, 22];
                excelcells22.Value2 = car_Atrs[i].data.lotDetails.lotNumberStr;
                end:
                Console.WriteLine(car_Atrs.Count-i);
                
                }
            Console.WriteLine("point");
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            Console.ReadKey();
        }

        //
        
        //
    }  
}
