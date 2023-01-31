using System;
using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using System.IO;


namespace doc_for_TO
{
    class Program
    {
        public static List<Device> devices = new List<Device>();
        static void Main(string[] args)
        {
            try
            {
                List<string> places = new List<string>();
                places.Add("ЦИТ");
                places.Add("погз Каменный Лог");
                places.Add("погп Крейванцы");
                places.Add("погп Клевица");
                foreach (string place in places)
                {
                    devices = new List<Device>();
                    read_excel_file(place);
                    create_word_file(place);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadLine();
            }
        }

        public static Dictionary<int, string> order_name = new Dictionary<int, string>() //словарь соответствия названий параметров и строк
        {
            {0, "Сервер,СХД,Диск.полка,лент библ" }, //название параметра, номер строки
            {1, "ПЭВМ" },
            {2, "ЖК-панель,навигатор,ШТК" },
            {3, "Неттоп" },
            {4, "Ноутбук" },
            {5, "Планшет" },
            {6, "Моноблок" },
            {7, "Плоттер" },
            {8, "Принтеры" },
            {9, "МФУ" },
            {10, "Сканер" },
            {11, "Считыватель документов" },
            {12, "Маршрутизатор" },
            {13, "Модем,мод.стойка" },
            {14, "Марм" },
            {15, "Парм " },
            {16, "ИБП к ПЭВМ" },
            {17, "ИБП к ВК" },
            {18, "Коммутатор,ПАК" },
            {19, "Межсетевой экран" },
            {20, "Веб-камера" },
            {21, "Wi-fi адаптер" },
            {22, "Консоль" },
            {23, "Переключатель,разветв." },
            {24, "СКС ЛВС" },
            {25, "Инструмент" },
            {26, "Имущ-во на карт-ках общ." },
            {27, "АВС,СКЖ,Масш.,Пот.код.,модуль м" },
            {28, "Беспроводная точка" },
            {29, "Копир.аппарат" },
            {30, "Сетевой накопитель" },
            {31, "Батарейный модуль" }
        };

        static void read_excel_file(string search_place)
        {
            String directory = "\\\\domane.by\\fileserver\\VCH2044\\IT\\ЦИТ\\ДОКУМЕНТЫ\\";
            string[] files = Directory.GetFiles(directory);
            foreach (string file in files)
            {
                if (file.Contains("ИНВЕТАРНЫЕ НОМЕРА"))
                {
                    directory = file;
                    break;
                }
            }
            Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application();
            try
            {
                Workbook workbook = application.Workbooks.Open(directory, Type.Missing, true);
                try
                {
                    for (int j = 0; j < order_name.Count; j++)
                    {
                        foreach (Worksheet worksheet in workbook.Worksheets)
                        {
                            if (worksheet.Name == order_name[j])
                            {
                                int i = 1; //текущая строка
                                int n = 0; //кол-во пустых строк подряд
                                int place_col = 10;
                                int name_col = 5;
                                int number_col = 3;
                                if (j == 24 || j == 32 || j == 33)
                                {
                                    place_col = 9;
                                    name_col = 4;
                                    number_col = 3;
                                }
                                if (j == 26)
                                {
                                    place_col = 4;
                                    name_col = 2;
                                    number_col = 1;
                                }
                                do
                                {
                                    Device device = new Device();
                                    string place = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[place_col][++i]).Text.ToString();
                                    if (place == "")
                                        n++;
                                    else
                                        n = 0;
                                    if (place.StartsWith(search_place))
                                    {
                                        if (place.Contains("/"))
                                            place = place.Substring(0, place.IndexOf("/"));
                                        string name = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[name_col][i]).Text.ToString();
                                        if (name.Contains("("))
                                            name = name.Substring(0, name.IndexOf("("));
                                        string serial_number = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[number_col][i]).Text.ToString();
                                        device.place = place;
                                        device.name = name;
                                        device.serial_number = serial_number;
                                        devices.Add(device);
                                    }
                                } while (n < 10);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
                finally
                {
                    workbook.Close();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                application.Quit();
            }
        }

        static void create_word_file(string search_place)
        {
            Microsoft.Office.Interop.Word.Application application = new Microsoft.Office.Interop.Word.Application();
            try
            {
                //Microsoft.Office.Interop.Word.Application application = new Microsoft.Office.Interop.Word.Application();
                //Object catalog = "C:\\Users\\ZHUCHKOVSKIJ_MYA\\Desktop\\TO\\План-график ТО.doc";
                //application.Documents.Open(ref catalog);
                //Document document = application.ActiveDocument;
                //List<string> words = new List<string>();
                //for (int i = 1; i < document.Paragraphs.Count; i++)
                //{
                //    string text = document.Paragraphs[i].Range.Text;
                //    words.Add(text);
                //}
                List<int> days = get_work_days();
                DateTime today = DateTime.Now;
                Object catalog = Directory.GetCurrentDirectory() + "\\План-график ТО.doc";
                Object new_catalog = Directory.GetCurrentDirectory() + "\\TO\\" + today.AddMonths(1).Year.ToString() + "\\" + today.AddMonths(1).ToString("MMMM");
                if (!Directory.Exists(new_catalog.ToString()))
                {
                    Directory.CreateDirectory(new_catalog.ToString());
                }
                new_catalog = new_catalog + "\\План-график ТО " + search_place + ".doc";
                Document document = application.Documents.Open(ref catalog);
                document.SaveAs(ref new_catalog);
                try
                {
                    //////////////// ПОИСК МЕСТ ДЛЯ ВСТАВОК////////////
                    Microsoft.Office.Interop.Word.Range search_area = document.Range(0, document.Characters.Count);
                    search_area.TextRetrievalMode.IncludeHiddenText = false;
                    search_area.TextRetrievalMode.IncludeFieldCodes = false;
                    object start = search_area.Text.IndexOf("ПЛАН-ГРАФИК") - 12;
                    object end = (int)start + 10;
                    document.Range(ref start, ref end).Text = today.ToString("dd.MM.yyyy");

                    search_area = document.Range(0, document.Characters.Count);
                    search_area.TextRetrievalMode.IncludeHiddenText = false;
                    search_area.TextRetrievalMode.IncludeFieldCodes = false;
                    start = search_area.Text.IndexOf("года") - 1;
                    end = search_area.Text.IndexOf("года") - 1;
                    document.Range(ref start, ref end).Text = " на " + today.AddMonths(1).ToString("MMMM") + " " + today.Year.ToString();

                    start = document.Characters.Count - 10;
                    end = document.Characters.Count;
                    document.Range(ref start, ref end).Text = today.ToString("dd.MM.yyyy");

                    search_area = document.Range(0, document.Characters.Count);
                    search_area.TextRetrievalMode.IncludeHiddenText = false;
                    search_area.TextRetrievalMode.IncludeFieldCodes = false;
                    start = search_area.Text.IndexOf("года") + 6;
                    end = (int)start;
                    Microsoft.Office.Interop.Word.Range table_location = document.Range(ref start, ref end);
                    Table table = document.Tables.Add(table_location, 3, 35);
                    document.PageSetup.Orientation = WdOrientation.wdOrientLandscape;
                    table.Range.Font.Size = 10;
                    table.Rows[1].Range.Font.Size = 8;
                    table.Rows[2].Range.ParagraphFormat.LeftIndent = table.Rows[2].Range.Application.CentimetersToPoints((float)0);
                    table.Rows[2].Range.ParagraphFormat.RightIndent = table.Rows[2].Range.Application.CentimetersToPoints((float)-0.5);
                    table.Cell(1, 1).Range.ParagraphFormat.LeftIndent = table.Rows[1].Range.Application.CentimetersToPoints((float)0);
                    table.Cell(1, 1).Range.ParagraphFormat.RightIndent = table.Rows[1].Range.Application.CentimetersToPoints((float)0);
                    table.Cell(2, 1).Range.ParagraphFormat.LeftIndent = table.Rows[2].Range.Application.CentimetersToPoints((float)0);
                    table.Cell(2, 1).Range.ParagraphFormat.RightIndent = table.Rows[2].Range.Application.CentimetersToPoints((float)0);
                    table.Cell(3, 1).Range.ParagraphFormat.LeftIndent = table.Rows[3].Range.Application.CentimetersToPoints((float)-0.5);
                    table.Cell(3, 1).Range.ParagraphFormat.RightIndent = table.Rows[3].Range.Application.CentimetersToPoints((float)-0.5);
                    for (int i = 0; i < 30 - days.Count; i++)
                    {
                        table.Columns[table.Columns.Count].Delete();
                    }
                    ///////////// РАЗМЕР ЯЧЕЕК ////////////
                    table.Rows[1].Height = (float)25;
                    table.Columns[1].Width = (float)22;
                    table.Columns[2].Width = (float)120;
                    table.Columns[3].Width = (float)90;
                    for (int i = 4; i <= table.Columns.Count - 2; i++)
                    {
                        table.Columns[i].Width = (float)19;
                    }
                    table.Columns[table.Columns.Count - 1].Width = (float)60;
                    table.Columns[table.Columns.Count].Width = (float)59;
                    ///////////// ВЫРАВНИВАНИЕ /////////////
                    table.Rows.Alignment = WdRowAlignment.wdAlignRowCenter;
                    for (int i = 1; i <= table.Columns.Count; i++)
                    {
                        table.Columns[i].Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                    }
                    for (int i = 1; i <= table.Rows.Count; i++)
                    {
                        if (i != 2)
                            table.Rows[i].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    }
                    for (int i = 2; i <= table.Rows.Count; i++)
                    {
                        table.Cell(i, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    }
                    table.Cell(1, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    ///////////// ЗАГОЛОВКИ //////////////
                    table.Cell(1, 1).Range.Text = "s/n:\nп/п";
                    table.Cell(1, 2).Range.Text = "Техника ИТ, ее заводской\n(порядковый) номер";
                    table.Cell(1, 3).Range.Text = "Исполнители работ";
                    table.Cell(1, 4).Range.Text = "Планируемое техническое обслуживание по числам";
                    table.Cell(1, 4 + days.Count).Range.Text = "Отметка о \nвыполнении";
                    table.Cell(1, 5 + days.Count).Range.Text = "Примечание";
                    for (int i = 0; i < days.Count; i++)
                    {
                        table.Cell(2, 4 + i).Range.Text = days[i].ToString();
                    }
                    table.Rows[1].Cells[4].Merge(table.Rows[1].Cells[4 + days.Count - 1]);
                    table.Cell(1, 1).Merge(table.Cell(2, 1));
                    table.Cell(1, 2).Merge(table.Cell(2, 2));
                    table.Cell(1, 3).Merge(table.Cell(2, 3));

                    table.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
                    table.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;

                    start = table.Cell(2, 4).Range.Start;
                    end = table.Cell(2, 4 + days.Count + 1).Range.End;
                    Microsoft.Office.Interop.Word.Range selected = document.Range(ref start, ref end);
                    selected.Select();
                    application.Selection.Cells.Height = 12;
                    for (int i = 1; i <= devices.Count; i++)
                    {
                        //Paragraph new_paragraph = new_document.Paragraphs.Add(Type.Missing);
                        //new_paragraph.Range.Text = devices[i].name + " №" + devices[i].serial_number + "\n";
                        //new_paragraph.Range.InsertParagraphAfter();
                        if (i < devices.Count)
                        {
                            Row row = table.Rows.Add();
                        }
                        table.Cell(i + 2, 1).Range.Text = i.ToString();
                        string text = devices[i - 1].name;
                        if (devices[i - 1].serial_number != "")
                        {
                            text = text + " №" + devices[i - 1].serial_number;
                        }
                        table.Cell(i + 2, 2).Range.Text = text;
                    }

                    //for (int i = 100; i <= table.Rows.Count; i++)
                    //{
                    //    table.Cell(i, 1).Range.ParagraphFormat.LeftIndent = table.Rows[i].Range.Application.CentimetersToPoints((float)0);
                    //    table.Cell(i, 1).Range.ParagraphFormat.RightIndent = table.Rows[i].Range.Application.CentimetersToPoints((float)0);
                    //}
                    //document.SaveAs(ref new_catalog);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
                finally
                {
                    document.Close();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                application.Quit();
            }
        }
        public static List<int> get_work_days()
        {
            List<int> work_days = new List<int>();
            DateTime date = new DateTime(DateTime.Now.AddMonths(1).Year, DateTime.Now.AddMonths(1).Month, 1);
            for (int i = 0; i < DateTime.DaysInMonth(date.Year, date.Month); i++)
            {
                if (date.AddDays(i).DayOfWeek != DayOfWeek.Saturday && date.AddDays(i).DayOfWeek != DayOfWeek.Sunday)
                    work_days.Add(i + 1);
            }
            return work_days;
        }
    }

    class Device
    {
        public string name, serial_number, place;
    }
}
