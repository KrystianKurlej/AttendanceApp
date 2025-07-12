using System.Text.Json;
using AttendanceApp.Models;

namespace AttendanceApp.Data;

public class ParticipantsRepository
{
    private const string FilePath = "participants.json";

    public List<Participant> LoadParticipants()
    {
        if (!File.Exists(FilePath))
            return new List<Participant>();

        string json = File.ReadAllText(FilePath);
        return JsonSerializer.Deserialize<List<Participant>>(json) ?? new List<Participant>();
    }

    public void SaveParticipants(List<Participant> participants)
    {
        string json = JsonSerializer.Serialize(participants, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(FilePath, json);
    }
}