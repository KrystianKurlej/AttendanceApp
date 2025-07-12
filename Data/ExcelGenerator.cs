using ClosedXML.Excel;
using AttendanceApp.Models;

namespace AttendanceApp.Data;

public class ExcelGenerator
{
    public void GenerateAttendanceSheet(List<Participant> participants, List<DateTime> dates, string outputPath)
    {
        var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add("Obecność");

        // Nagłówki dat
        for (int i = 0; i < dates.Count; i++)
        {
            ws.Cell(1, i + 2).Value = dates[i].Day.ToString();
            ws.Column(i + 2).Width = 3; // Ustawienie szerokości kolumny
        }

        // Wiersze z uczestnikami
        for (int i = 0; i < participants.Count; i++)
        {
            ws.Cell(i + 2, 1).Value = participants[i].FullName;

            for (int j = 0; j < dates.Count; j++)
            {
                ws.Cell(i + 2, j + 2).Value = ""; // tu będzie X lub pustka
            }

            // suma wiersza
            ws.Cell(i + 2, dates.Count + 2).FormulaA1 = $"=COUNTIF(B{i + 2}:{(char)('A' + dates.Count)}{i + 2},\"X\")";
        }

        // suma kolumn
        for (int j = 0; j < dates.Count; j++)
        {
            ws.Cell(participants.Count + 2, j + 2).FormulaA1 = $"=COUNTIF({(char)('B' + j)}2:{(char)('B' + j)}{participants.Count + 1},\"X\")";
        }

        ws.Cell(1, 1).Value = "Imię i nazwisko";
        ws.Cell(1, dates.Count + 2).Value = "Suma";
        
        // Formatowanie nagłówków - pogrubiony tekst i jasne szare tło
        ws.Row(1).Style.Font.Bold = true; // Pierwszy wiersz (nagłówki)
        ws.Row(1).Style.Fill.BackgroundColor = XLColor.LightGray;
        
        ws.Row(participants.Count + 2).Style.Font.Bold = true; // Ostatni wiersz (sumy kolumn)
        ws.Row(participants.Count + 2).Style.Fill.BackgroundColor = XLColor.LightGray;
        
        ws.Column(1).Style.Font.Bold = true; // Pierwsza kolumna (imiona i nazwiska)
        ws.Column(1).Style.Fill.BackgroundColor = XLColor.LightGray;
        
        ws.Column(dates.Count + 2).Style.Font.Bold = true; // Ostatnia kolumna (suma wierszy)
        ws.Column(dates.Count + 2).Style.Fill.BackgroundColor = XLColor.LightGray;
        
        // Obramowanie dla całej tabeli
        var dataRange = ws.Range(1, 1, participants.Count + 2, dates.Count + 2);
        dataRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        dataRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
        
        // Dopasowanie szerokości pierwszej kolumny (imiona i nazwiska)
        ws.Column(1).AdjustToContents();
        
        // Dopasowanie szerokości ostatniej kolumny (suma)
        ws.Column(dates.Count + 2).AdjustToContents();
        
        // Ustawienie skalowania strony - wpasuj w 1 str. w poziomie na 1 w pionie
        ws.PageSetup.FitToPages(1, 1);
        
        workbook.SaveAs(outputPath);
    }
}