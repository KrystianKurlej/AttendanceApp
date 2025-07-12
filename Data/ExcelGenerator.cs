using ClosedXML.Excel;
using AttendanceApp.Models;

namespace AttendanceApp.Data;

public class ExcelGenerator
{
    public void GenerateAttendanceSheet(List<Participant> participants, List<DateTime> dates, string outputPath, string monthName, int year)
    {
        var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add("Obecność");

        // Nagłówek główny - nazwa ośrodka, miesiąc i rok
        ws.Cell(1, 1).Value = $"ŚDS2 {monthName} {year}";
        ws.Cell(1, 1).Style.Font.Bold = true;
        ws.Cell(1, 1).Style.Font.FontSize = 14;
        ws.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        
        // Scalenie komórek dla nagłówka głównego
        ws.Range(1, 1, 1, dates.Count + 2).Merge();
        
        // Brak obramowania z góry, lewej i prawej dla nagłówka
        ws.Range(1, 1, 1, dates.Count + 2).Style.Border.TopBorder = XLBorderStyleValues.None;
        ws.Range(1, 1, 1, dates.Count + 2).Style.Border.LeftBorder = XLBorderStyleValues.None;
        ws.Range(1, 1, 1, dates.Count + 2).Style.Border.RightBorder = XLBorderStyleValues.None;
        
        // Nagłówki dat (przesunięte o jeden wiersz w dół)
        for (int i = 0; i < dates.Count; i++)
        {
            ws.Cell(2, i + 2).Value = dates[i].Day.ToString();
            ws.Column(i + 2).Width = 3; // Ustawienie szerokości kolumny
        }

        // Wiersze z uczestnikami (przesunięte o jeden wiersz w dół)
        for (int i = 0; i < participants.Count; i++)
        {
            ws.Cell(i + 3, 1).Value = participants[i].FullName;

            for (int j = 0; j < dates.Count; j++)
            {
                ws.Cell(i + 3, j + 2).Value = ""; // tu będzie X lub pustka
            }

            // suma wiersza
            ws.Cell(i + 3, dates.Count + 2).FormulaA1 = $"=COUNTIF(B{i + 3}:{(char)('A' + dates.Count)}{i + 3},\"X\")";
        }

        // suma kolumn (przesunięte o jeden wiersz w dół)
        for (int j = 0; j < dates.Count; j++)
        {
            ws.Cell(participants.Count + 3, j + 2).FormulaA1 = $"=COUNTIF({(char)('B' + j)}3:{(char)('B' + j)}{participants.Count + 2},\"X\")";
        }

        ws.Cell(2, 1).Value = "Imię i nazwisko";
        ws.Cell(2, dates.Count + 2).Value = "Suma";
        
        // Formatowanie nagłówków - pogrubiony tekst i jasne szare tło (przesunięte o jeden wiersz)
        ws.Row(2).Style.Font.Bold = true; // Drugi wiersz (nagłówki tabeli)
        ws.Row(2).Style.Fill.BackgroundColor = XLColor.LightGray;
        
        ws.Row(participants.Count + 3).Style.Font.Bold = true; // Ostatni wiersz (sumy kolumn)
        ws.Row(participants.Count + 3).Style.Fill.BackgroundColor = XLColor.LightGray;
        
        // Pierwsza kolumna (imiona i nazwiska) - tylko od wiersza 2 w dół
        for (int i = 2; i <= participants.Count + 3; i++)
        {
            ws.Cell(i, 1).Style.Font.Bold = true;
            ws.Cell(i, 1).Style.Fill.BackgroundColor = XLColor.LightGray;
        }
        
        // Ostatnia kolumna (suma wierszy) - tylko od wiersza 2 w dół
        for (int i = 2; i <= participants.Count + 3; i++)
        {
            ws.Cell(i, dates.Count + 2).Style.Font.Bold = true;
            ws.Cell(i, dates.Count + 2).Style.Fill.BackgroundColor = XLColor.LightGray;
        }
        
        // Obramowanie dla całej tabeli (przesunięte o jeden wiersz w dół)
        var dataRange = ws.Range(2, 1, participants.Count + 3, dates.Count + 2);
        dataRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        dataRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
        
        // Dopasowanie szerokości pierwszej kolumny (imiona i nazwiska)
        ws.Column(1).AdjustToContents();
        
        // Dopasowanie szerokości ostatniej kolumny (suma)
        ws.Column(dates.Count + 2).AdjustToContents();
        
        // Ustawienie skalowania strony - wpasuj w 1 str. w poziomie na 1 w pionie
        ws.PageSetup.FitToPages(1, 1);
        
        // Usunięcie nagłówka PageSetup (nie działał prawidłowo)
        
        workbook.SaveAs(outputPath);
    }
}