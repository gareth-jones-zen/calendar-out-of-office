using System.Globalization;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using CalendarDays;

namespace OutOfOffice;

/// <summary>A calendar event as read, local time; <see cref="Declined"/> or <see cref="Cancelled"/> ones are ignored.</summary>
public sealed record CalEvent(string Subject, DateTime Start, DateTime End, bool AllDay, string ShowAs,
    IReadOnlyList<string> Categories, bool Mine, int Attendees, bool Cancelled = false, bool Declined = false)
{
    public DayEvent ForRules() => new(Subject, Start, End, AllDay, ShowAs, Categories, Mine, Attendees);
}

public enum Kind { Working, Weekend, BankHoliday, Leave, HalfDay, Conference }

/// <summary>A day and what it is, with why (and the conference's name for a conference day).</summary>
public sealed record Day(DateOnly Date, Kind Kind, string Why, string? Event = null);

/// <summary>A run of days away: the days (weekends and bank holidays joined), whether it's leave or a conference, the
/// weekdays away (half days 0.5), and the window the reply is on for, in local time.</summary>
public sealed record Period(string Id, bool IsLeave, IReadOnlyList<Day> Days, double WeekdaysAway, DateTime On, DateTime Off,
    DateOnly ReturnDay, string? Event, bool OpenEnded)
{
    public DateOnly First => Days.First(d => d.Kind is not (Kind.Weekend or Kind.BankHoliday)).Date;
    public DateOnly Last => Days.Last(d => d.Kind is not (Kind.Weekend or Kind.BankHoliday)).Date;
    public string KindName => IsLeave ? "leave" : "conference";
}

/// <summary>What would be (or was) written: the window as UTC, the two messages as HTML, and who outside sees one.</summary>
public sealed record Reply(DateTime StartUtc, DateTime EndUtc, string Internal, string External, string Audience)
{
    /// <summary>What's compared to tell whether the mailbox still holds this: the times to the minute, the messages as
    /// text (Exchange may rewrite the HTML), the audience.</summary>
    public string Fingerprint() => Planner.Hash(
        $"{StartUtc:yyyy-MM-ddTHH:mm}|{EndUtc:yyyy-MM-ddTHH:mm}|{Planner.Plain(Internal)}|{Planner.Plain(External)}|{Audience}");
}

public static class Planner
{
    private static readonly CultureInfo En = CultureInfo.InvariantCulture;   // English names; invariant globalisation has no en-GB

    /// <summary>Each day from first to last: what it is. Pure.</summary>
    public static List<Day> Classify(Config config, IReadOnlyList<CalEvent> events, DateOnly first, DateOnly last)
    {
        var live = events.Where(e => !e.Cancelled && !e.Declined).ToList();
        // conference.ignore_subjects leaves events out of the conference rules: a reminder to book one ("Book Digital
        // City Festival?") isn't being there, and another tool's "Working from ..." entries say nothing.
        var conferenceEvents = live.Where(e => config.ConferenceIgnore?.IsMatch(e.Subject) != true).Select(e => e.ForRules()).ToList();
        var rules = live.Select(e => e.ForRules()).ToList();
        var days = new List<Day>();
        for (var d = first; d <= last; d = d.AddDays(1))
        {
            var bank = rules.FirstOrDefault(e => e.Mine && e.Attendees == 0 && DayRules.Covers(e, d)
                && (config.BankHolidaySubjects.IsMatch(e.Subject) || e.Categories.Any(c => config.BankHolidayCategories.Contains(c, StringComparer.OrdinalIgnoreCase))));
            if (!config.WorkingDays.Contains(d.DayOfWeek))
            {
                days.Add(new Day(d, Kind.Weekend, "not a working day"));
                continue;
            }
            if (bank is not null)
            {
                days.Add(new Day(d, Kind.BankHoliday, $"bank holiday: {bank.Subject}"));
                continue;
            }
            var leave = DayRules.Leave(config.Leave, rules, d);
            if (leave.Entries.Count > 0)
            {
                var half = leave.Entries.All(e => config.HalfDay.IsMatch(e.Subject)) || !leave.DayOff;
                days.Add(new Day(d, half ? Kind.HalfDay : Kind.Leave, $"{(half ? "half day" : "leave")}: {leave.Entries[0].Subject}"));
                continue;
            }
            var away = DayRules.Elsewhere(config.Conference, conferenceEvents, d);
            if (away is not null)
            {
                days.Add(new Day(d, Kind.Conference, away.Why, EventName(away.Event.Subject)));
                continue;
            }
            days.Add(new Day(d, Kind.Working, "working"));
        }
        return days;
    }

    /// <summary>"At XConf Europe, limited availability" is "XConf Europe, limited availability" in a message.</summary>
    public static string EventName(string subject)
    {
        var s = subject.Trim();
        return s.StartsWith("At ", StringComparison.OrdinalIgnoreCase) ? s[3..].TrimStart() : s;
    }

    /// <summary>The periods away in the days given: runs of days that aren't working days, holding leave, a half day or
    /// a conference (a bank holiday alone only if bank_holiday.alone), over the kind's threshold. A run still going on
    /// the last day given is open-ended.</summary>
    public static List<Period> Periods(Config config, IReadOnlyList<Day> days)
    {
        var periods = new List<Period>();
        var i = 0;
        while (i < days.Count)
        {
            if (days[i].Kind == Kind.Working)
            {
                i++;
                continue;
            }
            var j = i;
            while (j < days.Count && days[j].Kind != Kind.Working)
                j++;
            var run = days.Skip(i).Take(j - i).ToList();
            var openStart = i == 0;
            var openEnd = j == days.Count;
            i = j;

            var core = run.Where(d => d.Kind is Kind.Leave or Kind.HalfDay or Kind.Conference).ToList();
            if (core.Count == 0 && !(config.BankHolidayAlone && run.Any(d => d.Kind == Kind.BankHoliday)))
                continue;
            if (openStart && core.Count == 0)
                continue;
            // Trim the weekends at either end, so the first and last days are days away.
            var firstAway = run.FindIndex(d => d.Kind != Kind.Weekend);
            var lastAway = run.FindLastIndex(d => d.Kind != Kind.Weekend);
            if (firstAway < 0)
                continue;
            var isLeave = run.Any(d => d.Kind is Kind.Leave or Kind.HalfDay) || core.Count == 0;
            var weekdays = run.Sum(d => d.Kind switch
            {
                Kind.Leave or Kind.Conference or Kind.BankHoliday => 1.0,
                Kind.HalfDay => 0.5,
                _ => 0.0,
            });
            var messages = isLeave ? config.LeaveMessages : config.ConferenceMessages;
            if (weekdays < messages.MinWorkingDays - 1e-9)
                continue;

            var first = run[firstAway];
            var last = run[lastAway];
            // On: the last working day before, at window.on; a half day at the start, that day at half_day_at.
            var on = first.Kind == Kind.HalfDay
                ? first.Date.ToDateTime(config.HalfDayAt)
                : run[0].Date.AddDays(-1).ToDateTime(config.On);
            // Off: the day you're back, at window.off; a half day at the end, that day at half_day_at.
            DateOnly back;
            DateTime off;
            if (last.Kind == Kind.HalfDay && lastAway > firstAway)
            {
                back = last.Date;
                off = last.Date.ToDateTime(config.HalfDayAt);
            }
            else
            {
                back = run[^1].Date.AddDays(1);
                off = back.ToDateTime(config.Off);
            }
            var eventName = run.FirstOrDefault(d => d.Event is not null)?.Event;
            periods.Add(new Period($"{(isLeave ? "leave" : "conference")}-{first.Date:yyyy-MM-dd}", isLeave, run, weekdays, on, off, back,
                eventName, openEnd));
        }
        return periods;
    }

    /// <summary>The placeholders' values for a period.</summary>
    public static Dictionary<string, string> Values(Config config, Period p)
    {
        var values = new Dictionary<string, string>(config.Placeholders, StringComparer.Ordinal)
        {
            ["return_day"] = Date(config, p.ReturnDay),
            ["first_day"] = Date(config, p.First),
            ["last_day"] = Date(config, p.IsLeave ? p.Last : p.Days.Last(d => d.Kind == Kind.Conference || d.Kind is Kind.Leave or Kind.HalfDay).Date),
            ["event"] = p.Event ?? "a conference",
        };
        return values;
    }

    public static string Date(Config config, DateOnly day) => day.ToString(config.DateFormat, En);

    /// <summary>A template filled in: {name} replaced, everything else left as written.</summary>
    public static string Fill(string template, IReadOnlyDictionary<string, string> values) =>
        Regex.Replace(template, @"\{([A-Za-z_][A-Za-z0-9_]*)\}", m => values.TryGetValue(m.Groups[1].Value, out var v) ? v : m.Value);

    /// <summary>Plain text as HTML: escaped, each paragraph (a blank line between) a &lt;p&gt;, its lines joined.</summary>
    public static string Html(string text)
    {
        var paragraphs = Regex.Split(text.Replace("\r\n", "\n").Trim(), @"\n\s*\n")
            .Select(p => string.Join(" ", p.Split('\n').Select(l => l.Trim()).Where(l => l.Length > 0)))
            .Where(p => p.Length > 0);
        return string.Concat(paragraphs.Select(p => $"<p>{WebUtility.HtmlEncode(p)}</p>"));
    }

    /// <summary>HTML as comparable text: tags out, entities decoded, whitespace collapsed.</summary>
    public static string Plain(string html)
    {
        var noTags = Regex.Replace(Regex.Replace(html ?? "", @"<(style|script)\b[^>]*>.*?</\1>", " ", RegexOptions.Singleline | RegexOptions.IgnoreCase), "<[^>]+>", " ");
        return Regex.Replace(WebUtility.HtmlDecode(noTags), @"\s+", " ").Trim();
    }

    /// <summary>The reply for a period: the window in UTC, the two messages, and who outside gets one (the audience
    /// chosen for the period, or the settings' for its kind).</summary>
    public static Reply Render(Config config, Period p, string? audience = null)
    {
        var messages = p.IsLeave ? config.LeaveMessages : config.ConferenceMessages;
        var values = Values(config, p);
        return new Reply(Utc(config, p.On), Utc(config, p.Off), Html(Fill(messages.Internal, values)), Html(Fill(messages.External, values)),
            audience ?? messages.Audience);
    }

    /// <summary>The settings' audience for a period's kind.</summary>
    public static string SettingsAudience(Config config, Period p) => (p.IsLeave ? config.LeaveMessages : config.ConferenceMessages).Audience;

    /// <summary>A period's first and last days away (weekends at either end left out), never throwing.</summary>
    public static (DateOnly First, DateOnly Last) Span(Period p)
    {
        var away = p.Days.Where(d => d.Kind != Kind.Weekend).ToList();
        return away.Count > 0 ? (away[0].Date, away[^1].Date) : (p.Days[0].Date, p.Days[^1].Date);
    }

    public static DateTime Utc(Config config, DateTime local) =>
        TimeZoneInfo.ConvertTimeToUtc(DateTime.SpecifyKind(local, DateTimeKind.Unspecified), config.TimeZone);

    /// <summary>The templates a period's reply comes from, as a hash: act mode writes only from ones approved.</summary>
    public static string TemplateHash(Config config, bool leave)
    {
        var m = leave ? config.LeaveMessages : config.ConferenceMessages;
        var placeholders = string.Join("|", config.Placeholders.OrderBy(kv => kv.Key, StringComparer.Ordinal).Select(kv => $"{kv.Key}={kv.Value}"));
        return Hash($"{(leave ? "leave" : "conference")}|{m.Internal}|{m.External}|{m.Audience}|{placeholders}");
    }

    public static string Hash(string text) => Convert.ToHexStringLower(SHA256.HashData(Encoding.UTF8.GetBytes(text)))[..16];
}
