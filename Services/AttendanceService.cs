namespace AttendanceApp.Services;

public class AttendanceService
{
    private readonly string[] _polishMonthNames = 
    {
        "styczeń", "luty", "marzec", "kwiecień", "maj", "czerwiec",
        "lipiec", "sierpień", "wrzesień", "październik", "listopad", "grudzień"
    };

    public List<DateTime> GetWorkdays(int year, int month)
    {
        var dates = new List<DateTime>();

        int days = DateTime.DaysInMonth(year, month);
        for (int i = 1; i <= days; i++)
        {
            var date = new DateTime(year, month, i);
            if (date.DayOfWeek != DayOfWeek.Saturday && date.DayOfWeek != DayOfWeek.Sunday)
                dates.Add(date);
        }

        return dates;
    }

    public string GetPolishMonthName(int month)
    {
        if (month < 1 || month > 12)
            throw new ArgumentOutOfRangeException(nameof(month), "Miesiąc musi być w zakresie 1-12");

        return _polishMonthNames[month - 1];
    }
}