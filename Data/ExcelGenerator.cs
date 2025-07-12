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
            ws.Cell(1, i + 2).Value = dates[i].ToString("dd.MM");
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
        workbook.SaveAs(outputPath);
    }
}