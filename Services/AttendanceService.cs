namespace AttendanceApp.Services;

public class AttendanceService
{
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
}