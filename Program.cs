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
        var file_path = "C:/Users/wojciech.mazor/Desktop/input.csv";
        var output_file_path = "C:/Users/wojciech.mazor/Desktop/output.xlsx";

        try
        {
            var data = new List<(string NumerTelefonu, int Czas)>();

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
                    //Console.WriteLine(numerTelefonu);
                    var czasAsString = fields[16].Trim();
                    //Console.WriteLine(czasAsString);
                    var typRozmowy = fields[10];
                    //Console.WriteLine(typRozmowy);
                    var typRuchu = fields[20];
                    //Console.WriteLine(typRuchu);

                    if (!string.IsNullOrWhiteSpace(numerTelefonu) && !string.IsNullOrWhiteSpace(czasAsString) && int.TryParse(czasAsString, out int czas)
                        && (typRozmowy == "Rozmowy krajowe" || typRozmowy == "Rozmowy międzynarodowe") && typRuchu == "Ruch")
                    {
                        data.Add((NumerTelefonu: numerTelefonu, Czas: czas));
                    }
                }
            }

            var groupedData = data.GroupBy(d => d.NumerTelefonu)
                                  .Select(g => new
                                  {
                                      NumerTelefonu = g.Key,
                                      CzasWSekundach = g.Sum(x => x.Czas),
                                      CzasWGodzinach = ConvertSecondsToHMS(g.Sum(x => x.Czas))
                                  }).ToList();

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Summary");
                worksheet.Cell("A1").Value = "Numer Telefonu";
                worksheet.Cell("B1").Value = "CzasWSekundach";
                worksheet.Cell("C1").Value = "CzasWGodzinach";

                int row = 2;
                foreach (var item in groupedData)
                {
                    worksheet.Cell(row, 1).Value = item.NumerTelefonu;
                    worksheet.Cell(row, 2).Value = item.CzasWSekundach;
                    worksheet.Cell(row, 3).Value = item.CzasWGodzinach;
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

    static string ConvertSecondsToHMS(int seconds)
    {
        TimeSpan time = TimeSpan.FromSeconds(seconds);
        return time.ToString(@"hh\:mm\:ss");
    }
}
