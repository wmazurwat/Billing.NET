using ClosedXML.Excel;
using Microsoft.VisualBasic.FileIO;
using System;
using System.Collections.Generic;
using System.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;

class Program
{
    static void Main(string[] args)
    {
        var file_path = "C:/Users/input.csv";
        var dictionary_file_path = "C:/Users/Słownik.csv";
        var output_file_path = "C:/Users/output.xlsx";

        try
        {
            var data = new List<(string NumerTelefonu, int Czas)>();
            var slownikImionINazwisk = new Dictionary<string, (string Imie, string Nazwisko)>();

            // Wczytanie danych ze słownika
            using (TextFieldParser parser = new TextFieldParser(dictionary_file_path))
            {
                parser.TextFieldType = FieldType.Delimited;
                parser.SetDelimiters(";");
                while (!parser.EndOfData)
                {
                    string[] fields = parser.ReadFields();
                    if (fields.Length >= 3)
                    {
                        slownikImionINazwisk[fields[0]] = (fields[1], fields[2]); // Numer telefonu, Imię, Nazwisko
                    }
                }
            }

            // Wczytanie danych z pliku CSV
            using (TextFieldParser parser = new TextFieldParser(file_path))
            {
                parser.TextFieldType = FieldType.Delimited;
                parser.SetDelimiters(";");

                bool isFirstRow = true;
                while (!parser.EndOfData)
                {
                    string[] fields = parser.ReadFields();
                    if (isFirstRow)
                    {
                        isFirstRow = false;
                        continue;
                    }

                    var numerTelefonu = fields[3];
                    var czasAsString = fields[16].Trim();
                    var typRozmowy = fields[10];
                    var typRuchu = fields[20];

                    if (!string.IsNullOrWhiteSpace(numerTelefonu) && !string.IsNullOrWhiteSpace(czasAsString) && int.TryParse(czasAsString, out int czas)
                        && (typRozmowy == "Rozmowy krajowe" || typRozmowy == "Rozmowy międzynarodowe") && typRuchu == "Ruch")
                    {
                        data.Add((NumerTelefonu: numerTelefonu, Czas: czas));
                    }
                }
            }

            // Grupowanie danych
            var groupedData = data.GroupBy(d => d.NumerTelefonu)
                                  .Select(g =>
                                  {
                                      var numerTelefonu = g.Key;
                                      var imieNazwisko = slownikImionINazwisk.ContainsKey(numerTelefonu) ? slownikImionINazwisk[numerTelefonu] : (Imie: "", Nazwisko: "");
                                      return new
                                      {
                                          NumerTelefonu = numerTelefonu,
                                          CzasWSekundach = g.Sum(x => x.Czas),
                                          CzasWGodzinach = ConvertSecondsToHMS(g.Sum(x => x.Czas)),
                                          IloscPolaczen = g.Count(), // Liczba połączeń
                                          Imie = imieNazwisko.Imie,
                                          Nazwisko = imieNazwisko.Nazwisko
                                      };
                                  }).ToList();

            // Tworzenie arkusza kalkulacyjnego
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Summary");
                worksheet.Cell("A1").Value = "Telefon komórkowy";
                worksheet.Cell("B1").Value = "Imię";
                worksheet.Cell("C1").Value = "Nazwisko";
                worksheet.Cell("D1").Value = "Ilość telefonów poza firmę"; // Nowa kolumna
                worksheet.Cell("E1").Value = "Czas trwania rozmów";
                //worksheet.Cell("B1").Value = "CzasWSekundach";




                int row = 2;
                foreach (var item in groupedData)
                {
                    worksheet.Cell(row, 1).Value = item.NumerTelefonu;
                    worksheet.Cell(row, 2).Value = item.Imie;
                    worksheet.Cell(row, 3).Value = item.Nazwisko;
                    worksheet.Cell(row, 4).Value = item.IloscPolaczen; // Dodajemy wartość do nowej kolumny
                    worksheet.Cell(row, 5).Value = item.CzasWGodzinach;

                                                                       //worksheet.Cell(row, 2).Value = item.CzasWSekundach;

                    row++;
                }

                workbook.SaveAs(output_file_path);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Wystąpił błąd: " + ex.Message);
        }
    }

    static string ConvertSecondsToHMS(int seconds)
    {
        TimeSpan time = TimeSpan.FromSeconds(seconds);
        return time.ToString(@"hh\:mm\:ss");
    }

    // Możesz dodać tutaj metodę SendEmailWithAttachment, jeśli jest potrzebna



static void SendEmailWithAttachment(string attachmentPath, string recipient, string subject, string body)
    {
        var outlookApp = new Outlook.Application();
        var mailItem = (Outlook.MailItem)outlookApp.CreateItem(Outlook.OlItemType.olMailItem);

        mailItem.Subject = subject;
        mailItem.Body = body;
        mailItem.To = recipient;

        if (!string.IsNullOrEmpty(attachmentPath))
        {
            mailItem.Attachments.Add(attachmentPath, Outlook.OlAttachmentType.olByValue, Type.Missing, Type.Missing);
        }

        // Odkomentuj poniższą linię, aby wyświetlić okno e-maila przed wysłaniem
        // mailItem.Display(true);

        mailItem.Send();
    }

  
}
