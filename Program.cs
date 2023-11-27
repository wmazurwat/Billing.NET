using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;

class Program
{
    static void Main(string[] args)
    {
        var file_path = "C:/Users/wojciech.mazor/Desktop/biling_07_2023.xlsx";
        var output_file_path = "C:/Users/wojciech.mazor/Desktop/output.xlsx";

        try
        {
            using (var workbook = new XLWorkbook(file_path))
            {
                var worksheet = workbook.Worksheet(1);
                var rows = worksheet.RangeUsed().RowsUsed().Skip(1); // Skip header row

                var data = rows.Select(r =>
                {
                    var numerTelefonu = r.Cell("D").GetValue<string>();
                    var czasAsString = r.Cell("Q").GetValue<string>()?.Trim(); // Użyj operatora ?. dla bezpieczeństwa null

                    if (!string.IsNullOrWhiteSpace(czasAsString) && int.TryParse(czasAsString, out int czas))
                    {
                        return new { NumerTelefonu = numerTelefonu, Czas = czas };
                    }
                    else
                    {
                        // Pomiń ten rekord lub zareaguj odpowiednio
                        //Console.WriteLine($"Nie można przekonwertować wartości '{czasAsString}' na liczbę całkowitą. Pomijanie rekordu.");
                        return null;
                    }
                }).Where(x => x != null).ToList(); // Usuń nulle z listy


                if (data.Any(d => d.NumerTelefonu == null || d.Czas == 0))
                {
                    Console.WriteLine("Kolumny 'Numer telefonu' lub 'Czas' nie istnieją w danych.");
                    return;
                }

                var groupedData = data.GroupBy(d => d.NumerTelefonu)
                                      .Select(g => new
                                      {
                                          NumerTelefonu = g.Key,
                                          CzasWSekundach = g.Sum(x => x.Czas),
                                          CzasWGodzinach = ConvertSecondsToHMS(g.Sum(x => x.Czas))
                                      }).ToList();

                using (var newWorkbook = new XLWorkbook())
                {
                    var newWorksheet = newWorkbook.Worksheets.Add("Summary");
                    newWorksheet.Cell("A1").Value = "Numer Telefonu";
                    newWorksheet.Cell("B1").Value = "CzasWSekundach";
                    newWorksheet.Cell("C1").Value = "CzasWGodzinach";

                    int row = 2;
                    foreach (var item in groupedData)
                    {
                        newWorksheet.Cell(row, 1).Value = item.NumerTelefonu;
                        newWorksheet.Cell(row, 2).Value = item.CzasWSekundach;
                        newWorksheet.Cell(row, 3).Value = item.CzasWGodzinach;
                        row++;
                    }

                    newWorkbook.SaveAs(output_file_path);
                }
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
}
