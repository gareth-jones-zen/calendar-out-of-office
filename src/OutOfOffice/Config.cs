using System.Globalization;
using System.Text.RegularExpressions;
using CalendarDays;
using PyCompat;
using ReviewQueue;
using Tomlyn;
using Tomlyn.Model;
using WinBackground;

namespace OutOfOffice;

/// <summary>Settings that can't be used, or a request that can't be met: the command line reports the message and
/// exits 2.</summary>
public class UsageException(string message) : Exception(message);

public enum Mode { Ask, Act }

public enum ConferenceAction { Reply, Status, Both }

/// <summary>One kind of period's messages: who outside sees one, the threshold, and the two templates.</summary>
public sealed record Messages(double MinWorkingDays, string Audience, string Internal, string External);

/// <summary>The settings file. See the README and <see cref="Template"/>. Read strictly: an unknown key, or a value of
/// the wrong type, is an error that names it.</summary>
public sealed class Config
{
    public const string EnvConfig = "OUT_OF_OFFICE_CONFIG";

    public static readonly string[] Audiences = ["none", "contactsOnly", "all"];

    /// <summary>The placeholders every template may use, besides those in [placeholders].</summary>
    public static readonly string[] BuiltIn = ["return_day", "first_day", "last_day", "event"];

    public const string DefaultLeave = @"^\s*(holiday\b|annual leave|on leave|summer holidays?)";
    public const string DefaultHalfDay = @"\(half day\)";
    public const string DefaultBankHoliday = @"^\s*bank holiday\b";
    public const string DefaultConference = @"\b(conference|summit|expo|exhibition|trade show|off-?site)\b";

    // Explicit "\n" (a raw literal would take the .cs file's line endings), written with CRLF.
    public const string Template =
        "# calendar-out-of-office settings. See the README.\n" +
        "# base = \"base-config.toml\"        # settings shared with other tools\n" +
        "# login_hint = \"you@example.com\"\n" +
        "# tenant = \"organizations\"          # or your organisation's tenant id\n" +
        "# time_zone = \"GMT Standard Time\"   # default: this PC's\n" +
        "\n" +
        "enabled = true\n" +
        "mode = \"ask\"                       # ask: each period is a question first; act: written from approved messages\n" +
        "\n" +
        "[window]\n" +
        "on = \"17:30\"                       # on the last working day before\n" +
        "off = \"08:00\"                      # on the day you're back\n" +
        "\n" +
        "[leave]\n" +
        "min_working_days = 3\n" +
        "internal = '''\n" +
        "I'm on leave and back on {return_day}. For anything urgent, please contact {cover}.\n" +
        "'''\n" +
        "external = '''\n" +
        "I'm on leave and back on {return_day}; I'll reply to your message when I return.\n" +
        "'''\n" +
        "\n" +
        "[conference]\n" +
        "internal = '''\n" +
        "I'm at {event} until {last_day}, so slower to reply than usual. Back on {return_day}.\n" +
        "'''\n" +
        "external = '''\n" +
        "I'm at {event} until {last_day} with limited access to email; I'll reply when I can.\n" +
        "'''\n" +
        "\n" +
        "[placeholders]\n" +
        "cover = \"my manager\"\n";

    public static readonly string[] KnownKeys =
    [
        "login_hint", "tenant", "msal_cache", "log_file", "state_file", "time_zone", "enabled", "mode", "days_ahead", "calendar_months_ahead",
        "act_days_before", "working_days", "date_format", "max_days", "every_minutes", "after_unlock_minutes", "notify",
        "notify_working_days", "window", "leave", "bank_holiday", "conference", "placeholders", "teams", "review", "schedule",
    ];

    /// <summary>The task's schedule when the settings don't give one: every 15 minutes, at logon and a minute after each unlock.</summary>
    public static readonly ScheduleSettings ScheduleDefaults = new() { EveryMinutes = 15, AtLogon = true, AfterUnlockMinutes = 1 };

    /// <summary>The schedule keys this tool took at the top level before [schedule] (0.4.0): still read, logged as deprecated.</summary>
    public static readonly string[] OldScheduleKeys = ["every_minutes", "after_unlock_minutes"];

    public required string Path { get; init; }
    public string? LoginHint { get; init; }
    public string Tenant { get; init; } = "organizations";
    public string? MsalCache { get; init; }
    public required string LogFile { get; init; }
    public required string StateFile { get; init; }
    public required TimeZoneInfo TimeZone { get; init; }
    public bool Enabled { get; init; } = true;
    public Mode Mode { get; init; } = Mode.Ask;
    public int DaysAhead { get; init; } = 42;
    public int ActDaysBefore { get; init; } = 7;
    public IReadOnlySet<DayOfWeek> WorkingDays { get; init; } = new HashSet<DayOfWeek>
        { DayOfWeek.Monday, DayOfWeek.Tuesday, DayOfWeek.Wednesday, DayOfWeek.Thursday, DayOfWeek.Friday };
    public string DateFormat { get; init; } = "dddd d MMMM";
    public int MaxDays { get; init; } = 30;
    /// <summary>When the task runs ([schedule]; the shared keys are background-tools-lib's ScheduleSettings).</summary>
    public ScheduleSettings Schedule { get; init; } = ScheduleDefaults;
    /// <summary>The old top-level schedule keys this file still uses (empty when it uses [schedule]).</summary>
    public IReadOnlyList<string> DeprecatedScheduleKeys { get; init; } = [];
    public int EveryMinutes => Schedule.EveryMinutes ?? 15;
    public int? AfterUnlockMinutes => Schedule.AfterUnlockMinutes > 0 ? Schedule.AfterUnlockMinutes : null;
    public bool Notify { get; init; } = true;
    public int NotifyWorkingDays { get; init; } = 5;

    public TimeOnly On { get; init; } = new(17, 30);
    public TimeOnly Off { get; init; } = new(8, 0);
    public TimeOnly HalfDayAt { get; init; } = new(12, 30);

    /// <summary>The [leave] settings: background-tools-lib's shared LeaveSettings (a base file's [leave] can supply the
    /// shared keys), with [bank_holiday]'s older subjects / categories folded in.</summary>
    public LeaveSettings LeaveSettings { get; init; } = new();

    /// <summary>Notes on the [leave] / [bank_holiday] settings for the log (older keys still used).</summary>
    public IReadOnlyList<string> LeaveNotes { get; init; } = [];

    public required LeaveRules Leave { get; init; }
    public required Regex HalfDay { get; init; }
    public required Messages LeaveMessages { get; init; }

    public required Regex BankHolidaySubjects { get; init; }
    public IReadOnlyList<string> BankHolidayCategories { get; init; } = ["Bank holiday"];
    public bool BankHolidayAlone { get; init; }

    public required ElsewhereRules Conference { get; init; }
    public Regex? ConferenceIgnore { get; init; }
    public ConferenceAction ConferenceAction { get; init; } = ConferenceAction.Reply;
    public required Messages ConferenceMessages { get; init; }

    public IReadOnlyDictionary<string, string> Placeholders { get; init; } = new Dictionary<string, string>();

    public bool TeamsStatusMessage { get; init; }
    public string ConferenceStatus { get; init; } = "At {event} until {last_day}";

    public ReviewSettings? Review { get; init; }

    public static string? Environment(string name) =>
        string.IsNullOrEmpty(name) || name.Contains('\0') ? null : System.Environment.GetEnvironmentVariable(name);

    public static string Local(Func<string, string?>? env = null)
    {
        env ??= Environment;
        var local = env("LOCALAPPDATA");
        return PyPath.Pure(System.IO.Path.Combine(string.IsNullOrEmpty(local) ? PyPath.Home(env) : local, "calendar-out-of-office"));
    }

    public static string DefaultPath(Func<string, string?>? env = null)
    {
        env ??= Environment;
        var appData = env("APPDATA");
        return PyPath.Pure(System.IO.Path.Combine(string.IsNullOrEmpty(appData) ? PyPath.Home(env) : appData, "calendar-out-of-office", "config.toml"));
    }

    /// <summary>Read the settings. <paramref name="today"/> (default: today in the settings' time zone) is the day
    /// <c>calendar_months_ahead</c> counts from.</summary>
    public static Config Load(string? path = null, Func<string, string?>? env = null, DateOnly? today = null)
    {
        env ??= Environment;
        var given = !string.IsNullOrEmpty(path) ? path : env(EnvConfig) is { Length: > 0 } fromEnv ? fromEnv : DefaultPath(env);
        var full = PyPath.Resolve(given);
        if (!File.Exists(full))
            throw new UsageException($"No settings file at {full}. Run `calendar-out-of-office config` to create one.");
        TomlTable raw;
        try
        {
            var text = PyStr.ReadText(full);
            if (text.StartsWith('﻿'))
                text = text[1..];
            var read = BaseSettings.Read(text, full, env, KnownKeys);
            raw = read.Raw;
            env = read.Env;
        }
        catch (TomlException error)
        {
            throw new UsageException($"{full}: {error.Message.TrimEnd()}");
        }
        catch (FormatException error)
        {
            throw new UsageException(error.Message.TrimEnd());
        }
        catch (Exception error) when (error is not (IOException or UnauthorizedAccessException or UsageException))
        {
            throw new UsageException($"{full}: {error.Message.TrimEnd()}");
        }
        foreach (var key in raw.Keys)
            if (!KnownKeys.Contains(key))
                throw new UsageException($"{full}: unknown setting '{key}' (known: {string.Join(", ", KnownKeys)}).");

        var folder = System.IO.Path.GetDirectoryName(full)!;
        var expandWith = env;
        string PathSetting(string value)
        {
            var expanded = PyPath.ExpandVars(PyPath.ExpandUser(value, expandWith), expandWith);
            return PyPath.IsAbsolute(expanded) ? PyPath.Pure(expanded) : PyPath.ResolveAgainst(folder, expanded);
        }
        UsageException Bad(string message) => new($"{full}: {message}");
        void Strict(TomlTable table, string where, params string[] keys)
        {
            foreach (var key in table.Keys)
                if (!keys.Contains(key))
                    throw Bad($"[{where}] has an unknown setting '{key}' (known: {string.Join(", ", keys)}).");
        }
        string? Text(TomlTable table, string key, string where) =>
            table.TryGetValue(key, out var v) ? v as string is { Length: > 0 } s ? s : throw Bad($"{where} should be text, and not empty.") : null;
        bool Flag(TomlTable table, string key, string where, bool fallback) =>
            table.TryGetValue(key, out var v) ? v as bool? ?? throw Bad($"{where} should be true or false.") : fallback;
        int Whole(string key, int fallback, int min, int max) =>
            raw.TryGetValue(key, out var v) ? v is long n && n >= min && n <= max ? (int)n : throw Bad($"{key} should be a whole number from {min} to {max}.") : fallback;
        double Number(TomlTable table, string key, string where, double fallback, double min, double max) =>
            table.TryGetValue(key, out var v)
                ? v switch { long l when l >= min && l <= max => l, double d when d >= min && d <= max => d, _ => throw Bad($"{where} should be a number from {min} to {max}.") }
                : fallback;
        IReadOnlyList<string>? Strings(TomlTable table, string key, string where) =>
            table.TryGetValue(key, out var v)
                ? v is TomlArray a && a.All(x => x is string) ? a.Select(x => (string)x!).ToList() : throw Bad($"{where} should be a list of text.")
                : null;
        Regex Pattern(string text, string where)
        {
            try
            {
                return new Regex(text, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1));
            }
            catch (ArgumentException error)
            {
                throw Bad($"{where} isn't a valid pattern: {error.Message}");
            }
        }
        TomlTable Section(string name)
        {
            if (!raw.TryGetValue(name, out var v))
                return [];
            return v as TomlTable ?? throw Bad($"{name} should be a [{name}] section.");
        }
        TimeOnly Clock(TomlTable table, string key, string where, TimeOnly fallback)
        {
            var text = Text(table, key, where);
            if (text is null)
                return fallback;
            return TimeOnly.TryParseExact(text, "HH:mm", CultureInfo.InvariantCulture, DateTimeStyles.None, out var t)
                ? t : throw Bad($"{where} should be a time such as \"17:30\".");
        }
        string Audience(TomlTable table, string where, string fallback)
        {
            var a = Text(table, "external_audience", where) ?? fallback;
            return Audiences.Contains(a) ? a : throw Bad($"{where} should be one of {string.Join(", ", Audiences)}.");
        }

        TimeZoneInfo zone;
        var zoneName = Text(raw, "time_zone", "time_zone");
        try
        {
            zone = zoneName is null ? TimeZoneInfo.Local : TimeZoneInfo.FindSystemTimeZoneById(zoneName);
        }
        catch (Exception error) when (error is TimeZoneNotFoundException or InvalidTimeZoneException)
        {
            throw Bad($"time_zone '{zoneName}' isn't a time zone this PC knows (a Windows name, such as \"GMT Standard Time\").");
        }

        var mode = Text(raw, "mode", "mode") ?? "ask";
        if (mode is not ("ask" or "act"))
            throw Bad("mode should be \"ask\" or \"act\".");

        var days = new HashSet<DayOfWeek>();
        foreach (var d in Strings(raw, "working_days", "working_days") ?? ["mon", "tue", "wed", "thu", "fri"])
            days.Add(d.ToLowerInvariant() switch
            {
                "mon" => DayOfWeek.Monday, "tue" => DayOfWeek.Tuesday, "wed" => DayOfWeek.Wednesday, "thu" => DayOfWeek.Thursday,
                "fri" => DayOfWeek.Friday, "sat" => DayOfWeek.Saturday, "sun" => DayOfWeek.Sunday,
                _ => throw Bad($"working_days: '{d}' isn't a day (mon, tue, wed, thu, fri, sat, sun)."),
            });
        if (days.Count == 0)
            throw Bad("working_days can't be empty.");

        var window = Section("window");
        Strict(window, "window", "on", "off", "half_day_at");

        var placeholders = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var (key, value) in Section("placeholders"))
        {
            if (BuiltIn.Contains(key))
                throw Bad($"placeholders.{key} is filled in by calendar-out-of-office itself; choose another name.");
            placeholders[key] = value as string ?? throw Bad($"placeholders.{key} should be text.");
        }
        var known = BuiltIn.Concat(placeholders.Keys).ToHashSet();
        string Message(TomlTable table, string key, string where, string fallback)
        {
            var text = table.TryGetValue(key, out var v) ? v as string ?? throw Bad($"{where} should be text.") : fallback;
            foreach (Match m in Regex.Matches(text, @"\{([A-Za-z_][A-Za-z0-9_]*)\}"))
                if (!known.Contains(m.Groups[1].Value))
                    throw Bad($"{where} uses {{{m.Groups[1].Value}}}, which isn't a placeholder (known: {string.Join(", ", known.Order())}).");
            return text.Trim();
        }

        var leave = Section("leave");
        var bank = Section("bank_holiday");
        Strict(bank, "bank_holiday", "subjects", "categories", "alone");
        LeaveSettings leaveSettings;
        var leaveNotes = new List<string>();
        try
        {
            leaveSettings = LeaveSettings.Parse(leave, full, ["min_working_days", "external_audience", "internal", "external"]);
        }
        catch (FormatException error)
        {
            throw new UsageException(error.Message);
        }
        // [bank_holiday] subjects / categories (before 0.5.0): the same as [leave]'s bank_holiday_subjects /
        // bank_holiday_categories, which a base file can share with the other tools. Still read; both at once is an error.
        foreach (var (old, shared) in new[] { ("subjects", "bank_holiday_subjects"), ("categories", "bank_holiday_categories") })
        {
            if (!bank.ContainsKey(old))
                continue;
            if (leave.ContainsKey(shared))
                throw Bad($"bank_holiday.{old} and leave.{shared} are both given: keep leave.{shared} only.");
            leaveNotes.Add($"settings: bank_holiday.{old} is deprecated: move it to [leave] as {shared}");
        }
        if (bank.ContainsKey("subjects"))
            leaveSettings = leaveSettings with { BankHolidaySubjects = Pattern(Text(bank, "subjects", "bank_holiday.subjects")!, "bank_holiday.subjects") };
        if (bank.ContainsKey("categories"))
            leaveSettings = leaveSettings with { BankHolidayCategories = Strings(bank, "categories", "bank_holiday.categories") ?? [] };
        var conference = Section("conference");
        Strict(conference, "conference", "min_working_days", "subjects", "show_as_signals", "min_hours", "ignore_subjects", "action",
            "external_audience", "internal", "external");
        var teams = Section("teams");
        Strict(teams, "teams", "status_message", "conference_status");

        var action = Text(conference, "action", "conference.action") ?? "reply";
        if (action is not ("reply" or "status" or "both"))
            throw Bad("conference.action should be \"reply\", \"status\" or \"both\".");
        var conferenceSubjects = conference.TryGetValue("subjects", out var cs)   // "" turns it off
            ? cs as string ?? throw Bad("conference.subjects should be text (a pattern), or \"\" for none.")
            : DefaultConference;

        ReviewSettings? review = null;
        if (raw.TryGetValue("review", out var reviewSection))
            review = ReviewSettings.Read(reviewSection, full, PathSetting);

        // How far ahead: the tool's own days_ahead wins; otherwise calendar_months_ahead (usually from a base file shared
        // with other calendar tools), as today to the same date that many calendar months on; otherwise 42 days.
        int Ahead()
        {
            if (raw.ContainsKey("days_ahead"))
                return Whole("days_ahead", 42, 1, 365);
            if (!raw.ContainsKey("calendar_months_ahead"))
                return 42;
            var months = Whole("calendar_months_ahead", 3, 1, 12);
            var from = today ?? DateOnly.FromDateTime(TimeZoneInfo.ConvertTime(DateTime.UtcNow, zone));
            return from.AddMonths(months).DayNumber - from.DayNumber;
        }

        var statusTemplate = Text(teams, "conference_status", "teams.conference_status") ?? "At {event} until {last_day}";
        Message(new TomlTable { ["x"] = statusTemplate }, "x", "teams.conference_status", "");

        var oldScheduleKeys = OldScheduleKeys.Where(raw.ContainsKey).ToList();
        ScheduleSettings schedule;
        try
        {
            if (raw.TryGetValue("schedule", out var scheduleTable))
            {
                if (oldScheduleKeys.Count > 0)
                    throw Bad($"{string.Join(" and ", oldScheduleKeys)} and [schedule] are both given: keep [schedule] only.");
                schedule = ScheduleSettings.Parse(scheduleTable as TomlTable ?? throw Bad("schedule should be a [schedule] section."),
                    ScheduleDefaults, TaskKind.Job, $"{full}: [schedule]");
            }
            else
            {
                var old = ScheduleDefaults;
                if (raw.ContainsKey("every_minutes"))
                    old = old with { EveryMinutes = Whole("every_minutes", 15, 1, 1440) };
                if (raw.ContainsKey("after_unlock_minutes"))
                    old = old with { AfterUnlockMinutes = Whole("after_unlock_minutes", 1, 0, 60) };
                schedule = ScheduleSettings.Parse(null, old, TaskKind.Job, $"{full}: [schedule]");
            }
        }
        catch (FormatException error)
        {
            throw new UsageException(error.Message);
        }

        return new Config
        {
            Path = full,
            LoginHint = Text(raw, "login_hint", "login_hint"),
            Tenant = Text(raw, "tenant", "tenant") ?? "organizations",
            MsalCache = Text(raw, "msal_cache", "msal_cache") is { } cache ? PathSetting(cache) : null,
            LogFile = Text(raw, "log_file", "log_file") is { } log ? PathSetting(log) : PyPath.Pure(System.IO.Path.Combine(Local(env), "calendar-out-of-office.log")),
            StateFile = Text(raw, "state_file", "state_file") is { } state ? PathSetting(state) : PyPath.Pure(System.IO.Path.Combine(Local(env), "state.json")),
            TimeZone = zone,
            Enabled = Flag(raw, "enabled", "enabled", true),
            Mode = mode == "act" ? Mode.Act : Mode.Ask,
            DaysAhead = Ahead(),
            ActDaysBefore = Whole("act_days_before", 7, 0, 60),
            WorkingDays = days,
            DateFormat = Text(raw, "date_format", "date_format") ?? "dddd d MMMM",
            MaxDays = Whole("max_days", 30, 1, 366),
            Schedule = schedule,
            DeprecatedScheduleKeys = oldScheduleKeys,
            Notify = Flag(raw, "notify", "notify", true),
            NotifyWorkingDays = Whole("notify_working_days", 5, 0, 60),
            On = Clock(window, "on", "window.on", new TimeOnly(17, 30)),
            Off = Clock(window, "off", "window.off", new TimeOnly(8, 0)),
            HalfDayAt = Clock(window, "half_day_at", "window.half_day_at", new TimeOnly(12, 30)),
            LeaveSettings = leaveSettings,
            LeaveNotes = leaveNotes,
            Leave = leaveSettings.Rules,
            HalfDay = leaveSettings.HalfDay,
            LeaveMessages = new Messages(Number(leave, "min_working_days", "leave.min_working_days", 3, 0, 366), Audience(leave, "leave.external_audience", "all"),
                Message(leave, "internal", "leave.internal", "I'm on leave and back on {return_day}."),
                Message(leave, "external", "leave.external", "I'm on leave and back on {return_day}; I'll reply to your message when I return.")),
            BankHolidaySubjects = leaveSettings.BankHolidaySubjects,
            BankHolidayCategories = leaveSettings.BankHolidayCategories,
            BankHolidayAlone = Flag(bank, "alone", "bank_holiday.alone", false),
            Conference = new ElsewhereRules(conferenceSubjects.Length == 0 ? null : Pattern(conferenceSubjects, "conference.subjects"),
                Strings(conference, "show_as_signals", "conference.show_as_signals") ?? ["workingElsewhere"],
                Number(conference, "min_hours", "conference.min_hours", 6, 0, 24), SubjectsAllDayOnly: false),
            ConferenceIgnore = Text(conference, "ignore_subjects", "conference.ignore_subjects") is { } ig ? Pattern(ig, "conference.ignore_subjects") : null,
            ConferenceAction = action switch { "status" => ConferenceAction.Status, "both" => ConferenceAction.Both, _ => ConferenceAction.Reply },
            ConferenceMessages = new Messages(Number(conference, "min_working_days", "conference.min_working_days", 1, 0, 366),
                Audience(conference, "conference.external_audience", "contactsOnly"),
                Message(conference, "internal", "conference.internal", "I'm at {event} until {last_day}. Back on {return_day}."),
                Message(conference, "external", "conference.external", "I'm at {event} until {last_day}; I'll reply when I can.")),
            Placeholders = placeholders,
            TeamsStatusMessage = Flag(teams, "status_message", "teams.status_message", false),
            ConferenceStatus = statusTemplate,
            Review = review,
        };
    }

    public static string CreateDefault(string? path, Func<string, string?>? env = null)
    {
        var full = PyPath.Resolve(!string.IsNullOrEmpty(path) ? path : DefaultPath(env));
        if (!File.Exists(full))
        {
            Directory.CreateDirectory(System.IO.Path.GetDirectoryName(full)!);
            File.WriteAllText(full, Template.Replace("\n", "\r\n"), new System.Text.UTF8Encoding(false));
        }
        return full;
    }
}
