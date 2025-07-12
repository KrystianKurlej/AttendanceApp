using AttendanceApp.Data;
using AttendanceApp.Services;
using AttendanceApp.Models;

var repo = new ParticipantsRepository();
var participants = repo.LoadParticipants();
var attendanceService = new AttendanceService();
var excelGenerator = new ExcelGenerator();

while (true)
{
    Console.Clear();
    Console.WriteLine("1. Wyświetl uczestników");
    Console.WriteLine("2. Dodaj uczestnika");
    Console.WriteLine("3. Usuń uczestnika");
    Console.WriteLine("4. Wygeneruj arkusz obecności");
    Console.WriteLine("5. Wyjście");
    Console.Write("Wybierz opcję: ");

    string? choice = Console.ReadLine();

    switch (choice)
    {
        case "1":
            Console.WriteLine("Lista uczestników:");
            foreach (var p in participants)
                Console.WriteLine($"- {p.FullName}");
            break;

        case "2":
            Console.Write("Podaj imię i nazwisko: ");
            participants.Add(new Participant { FullName = Console.ReadLine()! });
            repo.SaveParticipants(participants);
            break;

        case "3":
            Console.Write("Podaj nazwę do usunięcia: ");
            var nameToRemove = Console.ReadLine();
            participants.RemoveAll(p => p.FullName == nameToRemove);
            repo.SaveParticipants(participants);
            break;

        case "4":
            Console.Write("Podaj rok: ");
            int year = int.Parse(Console.ReadLine()!);
            Console.Write("Podaj miesiąc (1-12): ");
            int month = int.Parse(Console.ReadLine()!);

            var dates = attendanceService.GetWorkdays(year, month);
            var filename = $"Obecnosc_{month}_{year}.xlsx";
            excelGenerator.GenerateAttendanceSheet(participants, dates, filename);

            Console.WriteLine($"Zapisano arkusz: {filename}");
            break;

        case "5":
            return;
    }

    Console.WriteLine("\nNaciśnij Enter, aby kontynuować...");
    Console.ReadLine();
}