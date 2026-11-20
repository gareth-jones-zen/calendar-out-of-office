using System.Text.Json.Nodes;

namespace OutOfOffice.Tests;

/// <summary>A calendar and mailbox in memory: events, the automatic-reply setting, and every write made.</summary>
internal sealed class FakeOffice : IOffice
{
    public List<CalEvent> Items { get; } = [];
    public Setting Current { get; set; } = Setting.Disabled;
    public List<string> Writes { get; } = [];
    public bool Writable { get; set; } = true;

    /// <summary>Exchange's habit, as far as we know it: the messages come back wrapped in a document.</summary>
    public bool WrapHtml { get; set; } = true;

    /// <summary>Throw after the write (it landed, the answer was lost) or before it (it didn't).</summary>
    public Exception? ThrowAfter { get; set; }
    public Exception? ThrowBefore { get; set; }

    /// <summary>What the mailbox holds after a write, if not what was written (someone else's change in between).</summary>
    public Setting? Instead { get; set; }

    public List<(string? Text, DateTime? Until)> Statuses { get; } = [];

    public IReadOnlyList<CalEvent> Events(DateOnly first, DateOnly last) =>
        Items.Where(e => e.End > first.ToDateTime(TimeOnly.MinValue) && e.Start < last.AddDays(1).ToDateTime(TimeOnly.MinValue)).ToList();

    public Setting Read() => Current;

    public void Schedule(Reply reply, bool keepStart)
    {
        if (ThrowBefore is { } before)
            throw before;
        Writes.Add($"schedule {reply.StartUtc:yyyy-MM-dd HH:mm} {reply.EndUtc:yyyy-MM-dd HH:mm}{(keepStart ? " keep-start" : "")}");
        var start = keepStart ? Current.StartUtc : reply.StartUtc;
        Current = Instead ?? new Setting("scheduled", start, reply.EndUtc, Wrap(reply.Internal), Wrap(reply.External), reply.Audience);
        if (ThrowAfter is { } after)
            throw after;
    }

    public void Disable()
    {
        Writes.Add("disable");
        Current = Current with { Status = "disabled" };
    }

    public bool CanWrite() => Writable;

    public void StatusMessage(string? text, DateTime? untilUtc) => Statuses.Add((text, untilUtc));

    private string Wrap(string html) => WrapHtml ? $"<html>\r\n<head><style>p {{ margin: 0 }}</style></head><body>\r\n{html.Replace("</p><p>", "</p>\r\n<p>")}\r\n</body>\r\n</html>" : html;

    /// <summary>An all-day leave entry, as an HR system's sync writes one: 00:00 to 23:59, the job title in brackets.</summary>
    public void Leave(string first, string last, string subject = "Holiday (Example Role)") =>
        Items.Add(new CalEvent(subject, DateTime.Parse(first + "T00:00"), DateTime.Parse(last + "T23:59"), false, "busy", [], true, 0));

    public void BankHoliday(string date) =>
        Items.Add(new CalEvent("Bank holiday", DateTime.Parse(date + "T09:00"), DateTime.Parse(date + "T17:00"), false, "oof", ["Bank holiday"], true, 0));

    public void Add(string subject, string start, string end, bool allDay = false, string showAs = "busy", bool mine = true, int attendees = 0) =>
        Items.Add(new CalEvent(subject, DateTime.Parse(start), DateTime.Parse(end), allDay, showAs, [], mine, attendees));
}

/// <summary>A settings file, a fake office and a clock, in a folder of their own.</summary>
internal sealed class World : IDisposable
{
    public string Dir { get; } = Path.Combine(Path.GetTempPath(), "out-of-office-tests-" + Guid.NewGuid().ToString("N"));
    public FakeOffice Office { get; } = new();
    public DateTime Now { get; set; } = new(2027, 3, 1, 9, 0, 0);   // a Monday
    public List<string> Notified { get; } = [];
    public StringWriter Out { get; private set; } = new();
    public StringWriter Err { get; private set; } = new();

    public const string Base = """
        time_zone = "GMT Standard Time"
        log_file = "ooo.log"
        state_file = "state.json"
        mode = "ask"
        [leave]
        min_working_days = 3
        internal = '''
        I'm on leave, back on {return_day}.

        Please contact {cover} & the team.
        '''
        external = "I'm away until {return_day}."
        [conference]
        subjects = '\b(conference|summit|example expo)\b'
        ignore_subjects = '^\s*book\b|\?\s*$'
        internal = "At {event} until {last_day}; back {return_day}."
        external = "At {event}."
        [placeholders]
        cover = "Alex Example"
        [review]
        file = "pending.json"
        source = "out-of-office"
        """;

    public World(string? extra = null, string settings = Base)
    {
        Directory.CreateDirectory(Dir);
        File.WriteAllText(Path.Combine(Dir, "config.toml"), settings + "\n" + (extra ?? ""));
    }

    public string ConfigPath => Path.Combine(Dir, "config.toml");

    public Config Config => Config.Load(ConfigPath);

    public CliHost Host => new()
    {
        Out = Out,
        Err = Err,
        Now = () => Now,
        Office = _ => Office,
        Notify = (title, lines) => Notified.Add(title + ": " + string.Join(" | ", lines)),
        Lock = false,
        Env = name => name == "LOCALAPPDATA" ? Dir : null,
    };

    public int Cli(params string[] args)
    {
        Out = new StringWriter();
        Err = new StringWriter();
        return OutOfOffice.Cli.Run(["--config", ConfigPath, .. args], Host);
    }

    public JsonObject State => OutOfOffice.State.Load(Path.Combine(Dir, "state.json"));

    public List<JsonObject> Items =>
        File.Exists(Path.Combine(Dir, "pending.json"))
            ? (JsonNode.Parse(File.ReadAllText(Path.Combine(Dir, "pending.json")))!["items"] as JsonArray ?? []).OfType<JsonObject>().ToList()
            : [];

    public void Dispose()
    {
        try
        {
            Directory.Delete(Dir, true);
        }
        catch (IOException)
        {
        }
    }
}
