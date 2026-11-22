using System.Globalization;
using System.Text.Json.Nodes;
using M365Graph;
using WinBackground;

namespace OutOfOffice;

/// <summary>The automatic-reply setting as the mailbox holds it: the status (disabled, alwaysEnabled, scheduled), the
/// window in UTC (null when not set), the two messages (HTML) and who outside gets one.</summary>
public sealed record Setting(string Status, DateTime? StartUtc, DateTime? EndUtc, string Internal, string External, string Audience)
{
    public static readonly Setting Disabled = new("disabled", null, null, "", "", "none");

    /// <summary>The same comparison as <see cref="Reply.Fingerprint"/>, for a scheduled setting.</summary>
    public string Fingerprint() => Planner.Hash(
        $"{StartUtc:yyyy-MM-ddTHH:mm}|{EndUtc:yyyy-MM-ddTHH:mm}|{Planner.Plain(Internal)}|{Planner.Plain(External)}|{Audience}");

    /// <summary>Whether it's sending nothing now or later: disabled, or a schedule already over.</summary>
    public bool Idle(DateTime nowUtc) => Status == "disabled" || (Status == "scheduled" && EndUtc is { } end && end <= nowUtc);
}

/// <summary>What a run reads and writes. GraphOffice is the real one; tests give a fake.</summary>
public interface IOffice
{
    /// <summary>Every event from first to last (inclusive), local time.</summary>
    IReadOnlyList<CalEvent> Events(DateOnly first, DateOnly last);

    /// <summary>The automatic-reply setting now (MailboxSettings.Read).</summary>
    Setting Read();

    /// <summary>Schedule a reply (MailboxSettings.ReadWrite). Only the start is left out when it has already passed.</summary>
    void Schedule(Reply reply, bool keepStart);

    /// <summary>Turn the reply off (MailboxSettings.ReadWrite).</summary>
    void Disable();

    /// <summary>Whether a write can be made without a sign-in: the write permission is in the cached sign-in.</summary>
    bool CanWrite();

    /// <summary>A Teams status message until the given time (Presence.ReadWrite), or null to clear it.</summary>
    void StatusMessage(string? text, DateTime? untilUtc);
}

/// <summary>Your calendar and mailbox settings through Microsoft Graph.
///
/// <para><b>Reads</b> need Calendars.ReadWrite (shared with the other calendar tools' sign-in) and MailboxSettings.Read;
/// <b>writes</b> MailboxSettings.ReadWrite, and Presence.ReadWrite for the optional Teams status message.</para></summary>
public sealed class GraphOffice(GraphClient graph, TimeZoneInfo zone) : IOffice
{
    public static readonly string[] CalendarScopes = ["Calendars.ReadWrite"];
    public static readonly string[] ReadScopes = ["MailboxSettings.Read"];
    public static readonly string[] WriteScopes = ["MailboxSettings.ReadWrite"];
    public static readonly string[] PresenceScopes = ["Presence.ReadWrite"];

    private const string Select = "subject,start,end,isAllDay,showAs,categories,isOrganizer,attendees,isCancelled,responseStatus";

    private Dictionary<string, string> Headers => new() { ["Prefer"] = $"outlook.timezone=\"{zone.Id}\"" };

    private string Utc(DateOnly day) =>
        TimeZoneInfo.ConvertTimeToUtc(day.ToDateTime(TimeOnly.MinValue), zone).ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture);

    public IReadOnlyList<CalEvent> Events(DateOnly first, DateOnly last)
    {
        var url = $"/me/calendarView?startDateTime={Utc(first)}&endDateTime={Utc(last.AddDays(1))}&$select={Select}&$top=250";
        return graph.PagesAsync(url, CalendarScopes, Headers).ToBlockingEnumerable().OfType<JsonObject>().Select(ReadEvent).ToList();
    }

    public Setting Read()
    {
        var o = graph.GetAsync("/me/mailboxSettings/automaticRepliesSetting", ReadScopes).GetAwaiter().GetResult();
        return Parse(o);
    }

    public void Schedule(Reply reply, bool keepStart)
    {
        var setting = new JsonObject
        {
            ["status"] = "scheduled",
            ["scheduledEndDateTime"] = Stamp(reply.EndUtc),
            ["internalReplyMessage"] = reply.Internal,
            ["externalReplyMessage"] = reply.External,
            ["externalAudience"] = reply.Audience,
        };
        if (!keepStart)
            setting["scheduledStartDateTime"] = Stamp(reply.StartUtc);
        graph.RequestAsync("/me/mailboxSettings", HttpMethod.Patch, new JsonObject { ["automaticRepliesSetting"] = setting }, WriteScopes)
            .GetAwaiter().GetResult();
    }

    public void Disable() =>
        graph.RequestAsync("/me/mailboxSettings", HttpMethod.Patch,
            new JsonObject { ["automaticRepliesSetting"] = new JsonObject { ["status"] = "disabled" } }, WriteScopes).GetAwaiter().GetResult();

    public bool CanWrite()
    {
        try
        {
            graph.TokenAsync(WriteScopes, interactive: false).GetAwaiter().GetResult();
            return true;
        }
        catch (SignInNeededException)
        {
            return false;
        }
    }

    public void StatusMessage(string? text, DateTime? untilUtc)
    {
        var me = graph.GetAsync("/me?$select=id", ["User.Read"]).GetAwaiter().GetResult();
        var id = me["id"]?.GetValue<string>() ?? throw new InvalidOperationException("Graph didn't say who you are.");
        var message = new JsonObject { ["message"] = new JsonObject { ["content"] = text ?? "", ["contentType"] = "text" } };
        if (text is not null && untilUtc is { } until)
            message["expiryDateTime"] = Stamp(until);
        graph.RequestAsync($"/users/{Uri.EscapeDataString(id)}/presence/setStatusMessage", HttpMethod.Post,
            new JsonObject { ["statusMessage"] = message }, PresenceScopes).GetAwaiter().GetResult();
    }

    private static JsonObject Stamp(DateTime utc) => new()
    {
        ["dateTime"] = utc.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture),
        ["timeZone"] = "UTC",
    };

    /// <summary>The setting from Graph's automaticRepliesSetting.</summary>
    internal static Setting Parse(JsonObject o) => new(
        Text(o, "status") ?? "disabled",
        When(o["scheduledStartDateTime"]), When(o["scheduledEndDateTime"]),
        Text(o, "internalReplyMessage") ?? "", Text(o, "externalReplyMessage") ?? "", Text(o, "externalAudience") ?? "none");

    /// <summary>A dateTimeTimeZone as UTC: "UTC" or a Windows (or IANA) zone name.</summary>
    internal static DateTime? When(JsonNode? node)
    {
        if (node is not JsonObject o || Text(o, "dateTime") is not { } text)
            return null;
        var local = DateTime.Parse(text[..Math.Min(19, text.Length)], CultureInfo.InvariantCulture, DateTimeStyles.None);
        var name = Text(o, "timeZone") ?? "UTC";
        if (name.Equals("UTC", StringComparison.OrdinalIgnoreCase) || name.Equals("Etc/UTC", StringComparison.OrdinalIgnoreCase))
            return DateTime.SpecifyKind(local, DateTimeKind.Utc);
        try
        {
            return TimeZoneInfo.ConvertTimeToUtc(DateTime.SpecifyKind(local, DateTimeKind.Unspecified), TimeZoneInfo.FindSystemTimeZoneById(name));
        }
        catch (Exception error) when (error is TimeZoneNotFoundException or InvalidTimeZoneException)
        {
            return DateTime.SpecifyKind(local, DateTimeKind.Utc);
        }
    }

    internal static CalEvent ReadEvent(JsonObject o) => new(
        Text(o, "subject") ?? "", Time(o["start"]), Time(o["end"]), Flag(o, "isAllDay"), Text(o, "showAs") ?? "unknown",
        (o["categories"] as JsonArray)?.Select(c => c?.GetValue<string>() ?? "").Where(c => c.Length > 0).ToList() ?? [],
        o["isOrganizer"] is not JsonValue v || !v.TryGetValue<bool>(out var organiser) || organiser,
        (o["attendees"] as JsonArray)?.OfType<JsonObject>().Count(a => Text(a, "type") != "resource") ?? 0,
        Flag(o, "isCancelled"),
        (o["responseStatus"] as JsonObject) is { } r && Text(r, "response") == "declined");

    internal static DateTime Time(JsonNode? node)
    {
        var text = (node as JsonObject) is { } o ? Text(o, "dateTime") : null;
        return text is null ? default : DateTime.Parse(text[..Math.Min(19, text.Length)], CultureInfo.InvariantCulture, DateTimeStyles.None);
    }

    internal static string? Text(JsonObject o, string key) =>
        o[key] is JsonValue v && v.TryGetValue<string>(out var s) ? s : null;

    private static bool Flag(JsonObject o, string key) =>
        o[key] is JsonValue v && v.TryGetValue<bool>(out var b) && b;

    public static GraphClient Client(Config config) => new(new GraphClientOptions
    {
        CachePath = config.MsalCache ?? M365Graph.Graph.DefaultCachePath,
        Tenant = config.Tenant,
        LoginHint = config.LoginHint,
        Scopes = CalendarScopes,
        SignInHint = "At a console, run: calendar-out-of-office sign-in",
        Log = Logger.Create(config.LogFile),
    });
}
