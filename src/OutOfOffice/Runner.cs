using System.Globalization;
using System.Text.Json.Nodes;
using M365Graph;
using ReviewQueue;

namespace OutOfOffice;

/// <summary>What the mailbox's reply is, as far as calendar-out-of-office is concerned.</summary>
public enum Owner { Idle, Ours, Theirs }

/// <summary>Everything a run, plan, apply or clear works from: the periods ahead, the setting now, and whose it is.</summary>
public sealed record Picture(List<Period> Periods, Setting Setting, Owner Owner, JsonObject? Written, DateTime Now, DateTime NowUtc,
    IReadOnlyDictionary<string, Audience> Audiences);

/// <summary>Who outside gets a period's reply, and where that came from: chosen for the period (`apply
/// --external-audience`, or `audience` ahead of time), perhaps carried from the id it had before its first day moved;
/// or the settings' for its kind.</summary>
public sealed record Audience(string Value, string Settings, bool Chosen = false, string? CarriedFrom = null)
{
    public static string Words(string value) => value switch
    {
        "all" => "everyone",
        "contactsOnly" => "your contacts only",
        "none" => "no one",
        _ => value,
    };

    /// <summary>"your contacts only (chosen for this period; the settings say everyone)".</summary>
    public string Describe()
    {
        if (!Chosen)
            return $"{Words(Value)} (the settings)";
        var source = CarriedFrom is { } from ? $"chosen for {from}, carried over when the first day moved" : "chosen for this period";
        return $"{Words(Value)} ({source}{(Value == Settings ? "" : $"; the settings say {Words(Settings)}")})";
    }

    /// <summary>A value as Graph spells it (none, contactsOnly, all), any case; null if it isn't one.</summary>
    public static string? Parse(string text) => Config.Audiences.FirstOrDefault(a => a.Equals(text, StringComparison.OrdinalIgnoreCase));
}

// The state file is the only record of which reply is ours (no marker is written to the mailbox): an upgrade or rename must
// carry it over, or a reply written before is treated as set by hand and left alone.
/// <summary>The state file: what calendar-out-of-office wrote (as fingerprints and times, never the messages), what it has raised,
/// and which messages have been approved. Small, so a backup of it holds nothing personal.</summary>
public static class State
{
    public static JsonObject Load(string path)
    {
        try
        {
            return File.Exists(path) ? JsonNode.Parse(File.ReadAllText(path)) as JsonObject ?? [] : [];
        }
        catch (Exception error) when (error is IOException or System.Text.Json.JsonException)
        {
            return [];
        }
    }

    public static void Save(string path, JsonObject state)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        var temp = path + ".tmp";
        File.WriteAllText(temp, state.ToJsonString(new System.Text.Json.JsonSerializerOptions { WriteIndented = true }));
        File.Move(temp, path, true);
    }

    public static JsonObject Table(JsonObject state, string key) =>
        state[key] as JsonObject ?? (JsonObject)(state[key] = new JsonObject());

    public static string Stamp(DateTime t) => t.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture);

    public static DateTime? Time(JsonNode? node) =>
        node is JsonValue v && v.TryGetValue<string>(out var s)
            && DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal, out var t)
            ? DateTime.SpecifyKind(t, DateTimeKind.Utc) : null;

    /// <summary>A period's chosen audience as kept in the state's "audience" table: the value and the period's first and
    /// last days, so a choice can follow a period whose first day (and so its id) moves.</summary>
    public static JsonObject Choice(string audience, Period p)
    {
        var (first, last) = Planner.Span(p);
        return new JsonObject
        {
            ["audience"] = audience,
            ["first"] = first.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
            ["last"] = last.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
        };
    }

    public static string? ChoiceValue(JsonNode? node) =>
        node is JsonObject o && o["audience"] is JsonValue v && v.TryGetValue<string>(out var s) ? Audience.Parse(s) : null;

    public static DateOnly? Day(JsonNode? node) =>
        node is JsonValue v && v.TryGetValue<string>(out var s)
            && DateOnly.TryParseExact(s, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var d) ? d : null;

    public static bool Approved(JsonObject state, string hash) =>
        state["approved"] is JsonArray a && a.Any(x => x?.GetValue<string>() == hash);

    public static void Approve(JsonObject state, string hash)
    {
        var a = state["approved"] as JsonArray ?? (JsonArray)(state["approved"] = new JsonArray());
        if (!a.Any(x => x?.GetValue<string>() == hash))
            a.Add((JsonNode?)JsonValue.Create(hash));
    }
}

/// <summary>The work of each command, given the office (Graph, or a fake) and the host.</summary>
public sealed class Runner(Config config, IOffice office, CliHost host)
{
    private static readonly CultureInfo En = CultureInfo.InvariantCulture;
    private readonly Action<string> log = WinBackground.Logger.Create(config.LogFile, null, clock: host.Now);
    public const string Item = "proposed";

    /// <summary>The periods from five weeks back (one in progress) to days_ahead on, plus max_days to see where the
    /// last one ends; those already over left out.</summary>
    public Picture Look(JsonObject state)
    {
        var now = host.Now();
        var nowUtc = Planner.Utc(config, now);
        var today = DateOnly.FromDateTime(now);
        var first = today.AddDays(-35);
        var last = today.AddDays(config.DaysAhead + config.MaxDays + 14);
        var days = Planner.Classify(config, office.Events(first, last), first, last);
        var periods = Planner.Periods(config, days).Where(p => p.Off > now && p.On <= now.AddDays(config.DaysAhead)).ToList();
        var setting = office.Read();
        var written = state["written"] as JsonObject;
        Owner owner;
        // Ours while it's the schedule we wrote and not yet over; a finished one is idle, whoever wrote it.
        if (written is not null && setting.Status == "scheduled" && !setting.Idle(nowUtc)
            && setting.Fingerprint() == written["fingerprint"]?.GetValue<string>())
            owner = Owner.Ours;
        else
            owner = setting.Idle(nowUtc) ? Owner.Idle : Owner.Theirs;
        return new Picture(periods, setting, owner, written, now, nowUtc, Audiences(state, periods, written));
    }

    /// <summary>Each period's audience: chosen for its id; else a choice kept for an id no longer in the calendar whose
    /// days overlap it (its first day moved); else the settings'.</summary>
    private Dictionary<string, Audience> Audiences(JsonObject state, List<Period> periods, JsonObject? written)
    {
        var table = state["audience"] as JsonObject;
        var ids = periods.Select(p => p.Id).ToHashSet(StringComparer.Ordinal);
        var result = new Dictionary<string, Audience>(StringComparer.Ordinal);
        foreach (var p in periods)
        {
            var settings = Planner.SettingsAudience(config, p);
            if (State.ChoiceValue(table?[p.Id]) is { } mine)
            {
                result[p.Id] = new Audience(mine, settings, true);
                continue;
            }
            var (first, last) = Planner.Span(p);
            var carried = table?.FirstOrDefault(kv => !ids.Contains(kv.Key) && State.ChoiceValue(kv.Value) is not null
                && State.Day(kv.Value?["first"]) is { } f && State.Day(kv.Value?["last"]) is { } l && f <= last && l >= first);
            if (carried is { Key: { } from } c)
                result[p.Id] = new Audience(State.ChoiceValue(c.Value)!, settings, true, from);
            // The reply written for this period records its choice too (kept in the table before the write, so this
            // is only a fallback): never compared against the settings' instead, which would rewrite it.
            else if (written?["period"]?.GetValue<string>() == p.Id && Audience.Parse(written["audience"]?.GetValue<string>() ?? "") is { } w)
                result[p.Id] = new Audience(w, settings, true);
            else
                result[p.Id] = new Audience(settings, settings);
        }
        return result;
    }

    /// <summary>A period's audience (the settings' for one the picture doesn't hold).</summary>
    public Audience AudienceOf(Period p, Picture pic) =>
        pic.Audiences.TryGetValue(p.Id, out var a) ? a : new Audience(Planner.SettingsAudience(config, p), Planner.SettingsAudience(config, p));

    private string? Chosen(Period p, Picture pic) => AudienceOf(p, pic) is { Chosen: true } a ? a.Value : null;

    /// <summary>A period's reply, its start moved a few minutes ahead if the window has already begun (a reply can
    /// only be scheduled from the future), or kept as written when it's ours and already running.</summary>
    public Reply Want(Period p, Picture pic, DateTime? keepStartUtc = null, string? audience = null)
    {
        var reply = Planner.Render(config, p, audience ?? AudienceOf(p, pic).Value);
        if (keepStartUtc is { } kept)
            return reply with { StartUtc = kept };
        var soonest = pic.NowUtc.AddMinutes(5);
        return reply.StartUtc < soonest ? reply with { StartUtc = new DateTime(soonest.Ticks - soonest.Ticks % TimeSpan.TicksPerMinute, DateTimeKind.Utc) } : reply;
    }

    /// <summary>Why a period can't be written, or null.</summary>
    public string? Refused(Period p, JsonObject state)
    {
        if (p.OpenEnded)
            return $"it runs past the {config.DaysAhead + config.MaxDays + 14} days read, so its end isn't known";
        if (p.WeekdaysAway > config.MaxDays)
            return $"{p.WeekdaysAway:0.#} weekdays away is more than max_days ({config.MaxDays}): check the calendar";
        if (State.Table(state, "hands_off")[p.Id] is not null)
            return "you changed or turned off its reply yourself, so it's yours";
        if (State.Table(state, "status_set")[p.Id] is not null)
            return "its Teams status message is set";
        return null;
    }

    /// <summary>The period that holds the reply we wrote: the same id, or one whose window overlaps it.</summary>
    public Period? Holding(Picture pic)
    {
        if (pic.Written is not { } w)
            return null;
        var id = w["period"]?.GetValue<string>();
        var start = State.Time(w["start"]);
        var end = State.Time(w["end"]);
        return pic.Periods.FirstOrDefault(p => p.Id == id)
               ?? pic.Periods.FirstOrDefault(p => start is { } s && end is { } e && Planner.Utc(config, p.On) < e && Planner.Utc(config, p.Off) > s);
    }

    /// <summary>One run: settle a lost write, follow what happened to ours, then the next period: proposed (ask) or
    /// written (act). Returns the lines a dry run prints.</summary>
    public List<string> Run(JsonObject state, bool dryRun)
    {
        var said = new List<string>();
        if (!config.Enabled)
        {
            said.Add("Switched off (enabled = false in the settings): nothing done.");
            return said;
        }
        if (!dryRun && config.DeprecatedScheduleKeys.Count > 0)
            log($"settings: {string.Join(" and ", config.DeprecatedScheduleKeys)} at the top level are deprecated: move them to [schedule]");
        if (!dryRun)
            foreach (var note in config.LeaveNotes)
                log(note);
        if (!dryRun && !config.Schedule.IsActive(host.Now(), out var why))
        {
            log($"outside the [schedule] ({why}): nothing done");
            said.Add($"Outside the [schedule] ({why}): nothing done.");
            return said;
        }
        var pic = Look(state);
        pic = Settle(state, pic, dryRun, said);

        // What happened to the reply we wrote.
        if (pic.Written is { } w && pic.Owner != Owner.Ours)
        {
            var end = State.Time(w["end"]);
            var id = w["period"]?.GetValue<string>() ?? "";
            if (pic.Setting.Idle(pic.NowUtc) && end is { } e && e <= pic.NowUtc.AddMinutes(1))
                said.Add($"The reply for {id} has finished.");
            else
            {
                said.Add(pic.Setting.Idle(pic.NowUtc)
                    ? $"The reply for {id} was turned off outside calendar-out-of-office: left off, and that period is yours now."
                    : $"The reply for {id} was changed outside calendar-out-of-office: left as it is, and that period is yours now.");
                if (!dryRun)
                {
                    State.Table(state, "hands_off")[id] = State.Stamp(pic.NowUtc);
                    Raise("changed", $"Out of office: your reply for {id} was changed by hand", id, null, "Nothing to do: calendar-out-of-office won't touch that period again.");
                }
            }
            if (!dryRun)
                state.Remove("written");
            pic = pic with { Written = null };
        }

        KeepChoices(state, pic, dryRun, said);
        if (pic.Owner == Owner.Ours)
            Follow(state, pic, dryRun, said);
        else if (pic.Owner == Owner.Theirs)
            Theirs(state, pic, dryRun, said);
        else
            Next(state, pic, dryRun, said);
        if (!dryRun)
            State.Save(config.StateFile, state);
        return said;
    }

    /// <summary>The audiences chosen per period: carried to a period's new id when its first day moved (said, and kept),
    /// each one's days kept up to date, and pruned once its period is over.</summary>
    private void KeepChoices(JsonObject state, Picture pic, bool dryRun, List<string> said)
    {
        var carried = new HashSet<string>(StringComparer.Ordinal);
        foreach (var p in pic.Periods)
        {
            if (AudienceOf(p, pic) is not { Chosen: true } a)
                continue;
            if (a.CarriedFrom is { } from)
            {
                said.Add($"{p.Id}: who outside gets a reply ({Audience.Words(a.Value)}) is carried over from {from}, whose first day moved.");
                carried.Add(from);
            }
            if (!dryRun)
                State.Table(state, "audience")[p.Id] = State.Choice(a.Value, p);
        }
        if (dryRun || state["audience"] is not JsonObject table)
            return;
        foreach (var from in carried)
            table.Remove(from);
        var today = DateOnly.FromDateTime(pic.Now);
        var writtenId = pic.Written?["period"]?.GetValue<string>();
        foreach (var (id, node) in table.ToList())
            if (id != writtenId && !pic.Periods.Any(p => p.Id == id) && (State.Day(node?["last"]) is not { } last || last < today))
                table.Remove(id);
        if (table.Count == 0)
            state.Remove("audience");
    }

    /// <summary>A write whose answer was lost: settled by reading back, never by writing blind.</summary>
    private Picture Settle(JsonObject state, Picture pic, bool dryRun, List<string> said)
    {
        if (state["writing"] is not JsonObject writing)
            return pic;
        var expected = writing["fingerprint"]?.GetValue<string>();
        var previous = writing["previous"]?.GetValue<string>();
        var current = pic.Setting.Idle(pic.NowUtc) ? "idle" : pic.Setting.Fingerprint();
        var id = writing["period"]?.GetValue<string>() ?? "";
        if (current == expected || (expected == "idle" && current == "idle"))
        {
            said.Add($"An earlier write for {id} did land: recorded.");
            if (!dryRun)
            {
                if (expected == "idle")
                    state.Remove("written");
                else
                    state["written"] = writing.DeepClone();
                state.Remove("writing");
            }
            var written = expected == "idle" ? null : (JsonObject)writing.DeepClone();
            return pic with { Written = written, Owner = written is null ? Owner.Idle : Owner.Ours };
        }
        if (current == previous && (writing["attempts"]?.GetValue<int>() ?? 1) < 2)
        {
            said.Add($"An earlier write for {id} didn't land: it may be made once more.");
            return pic;
        }
        said.Add($"An earlier write for {id} can't be confirmed: the mailbox holds something else. Nothing more is written for it.");
        if (!dryRun)
        {
            state.Remove("writing");
            Raise("failed", $"Out of office: a write for {id} can't be confirmed", id, null,
                "Check your automatic replies in Outlook (`calendar-out-of-office show`); nothing more is written for this period.");
        }
        return pic;
    }

    /// <summary>Our reply is on: keep it matching its period (decision: a cancelled, shortened or extended period is
    /// adjusted, and you're told in a line).</summary>
    private void Follow(JsonObject state, Picture pic, bool dryRun, List<string> said)
    {
        var w = pic.Written!;
        var id = w["period"]?.GetValue<string>() ?? "";
        var start = State.Time(w["start"])!.Value;
        var period = Holding(pic);
        if (period is null)
        {
            said.Add($"{id}: no longer in the calendar, so the reply is turned off.");
            if (!dryRun)
            {
                Write(state, pic, id, null, null);
                Raise("changed", $"Out of office: turned off the reply for {id}", id, null, "Its leave or conference is no longer in your calendar, so the reply was turned off.");
            }
            return;
        }
        var running = start <= pic.NowUtc;
        var want = Want(period, pic, running ? start : null);
        if (want.Fingerprint() == w["fingerprint"]?.GetValue<string>())
        {
            said.Add($"{period.Id}: the reply is on as written ({Window(want)}).");
            return;
        }
        var oldEnd = State.Time(w["end"]);
        var windowMoved = State.Time(w["start"]) != want.StartUtc || oldEnd != want.EndUtc;
        if (!windowMoved && w["templates"]?.GetValue<string>() != Planner.TemplateHash(config, period.IsLeave)
            && !(config.Mode == Mode.Act && State.Approved(state, Planner.TemplateHash(config, period.IsLeave))))
        {
            said.Add($"{period.Id}: the messages in the settings have changed; asking before the reply is rewritten.");
            if (!dryRun)
                Propose(state, period, want, pic);
            return;
        }
        var how = oldEnd is { } e && want.EndUtc > e ? "longer" : oldEnd is { } e2 && want.EndUtc < e2 ? "shorter" : "changed";
        var audience = AudienceOf(period, pic);
        var moved = period.Id != id && audience.Chosen
            ? $" It starts on a different day now, so it's {period.Id}; who outside gets a reply ({Audience.Words(audience.Value)}, chosen for {id}) is carried over."
            : "";
        said.Add($"{period.Id}: the period is {how}, so the reply is rewritten: {Window(want)}.{moved}");
        if (!dryRun)
        {
            Write(state, pic, period.Id, want, period, keepStart: running, chosen: Chosen(period, pic));
            Raise("changed", $"Out of office: your {period.KindName} is {how}, so the reply now runs {Window(want)}", period.Id, want,
                "Nothing to do; just so you know." + moved, audience);
        }
    }

    /// <summary>A reply calendar-out-of-office didn't write: never touched; raised once.</summary>
    private void Theirs(JsonObject state, Picture pic, bool dryRun, List<string> said)
    {
        var fp = pic.Setting.Fingerprint();
        said.Add($"Your automatic replies are {pic.Setting.Status} ({SettingWindow(pic.Setting)}), set outside calendar-out-of-office: left alone.");
        var next = pic.Periods.FirstOrDefault(p => Refused(p, state) is null);
        if (next is not null)
            said.Add($"{next.Id}: waits until that's off ({Window(Want(next, pic))}).");
        if (dryRun || state["hand_set"]?.GetValue<string>() == fp)
            return;
        state["hand_set"] = fp;
        if (next is not null)
        {
            Raise("hand-set", $"Out of office: a reply you set is on, so {next.Id} wasn't scheduled", next.Id, null,
                "Your own automatic reply is left alone. When it's off, `calendar-out-of-office plan` shows the next one; or `calendar-out-of-office clear --force --yes` turns yours off.");
            NotifyOnce(state, "hand-set:" + fp, "Out of office: your own automatic reply is on",
                [$"So {next.Id} wasn't scheduled. Nothing was changed."]);
        }
    }

    /// <summary>No reply is on: the next period, proposed or (act mode, approved messages) written.</summary>
    private void Next(JsonObject state, Picture pic, bool dryRun, List<string> said)
    {
        Period? next = null;
        foreach (var p in pic.Periods)
        {
            var why = Refused(p, state);
            if (why is null)
            {
                next = p;
                break;
            }
            said.Add($"{p.Id}: not scheduled: {why}.");
        }
        if (next is null)
        {
            if (pic.Periods.Count == 0)
                said.Add("No leave or conference in the next " + config.DaysAhead + " days that needs a reply.");
            return;
        }
        var want = Want(next, pic);
        if (next is { IsLeave: false } && config.ConferenceAction == ConferenceAction.Status)
        {
            said.Add($"{next.Id}: a Teams status message only (conference.action = \"status\").");
            if (!config.TeamsStatusMessage)
            {
                said.Add("  ...but teams.status_message is false, so nothing is set.");
                return;
            }
        }
        var hash = Planner.TemplateHash(config, next.IsLeave);
        var due = DateOnly.FromDateTime(next.On).AddDays(-config.ActDaysBefore) <= DateOnly.FromDateTime(pic.Now);
        if (config.Mode == Mode.Act && State.Approved(state, hash))
        {
            if (!due)
            {
                said.Add($"{next.Id}: written on {DateOnly.FromDateTime(next.On).AddDays(-config.ActDaysBefore):ddd d MMM} (act mode): {Window(want)}.");
                return;
            }
            said.Add($"{next.Id}: writing (act mode, messages approved): {Window(want)}.");
            if (!dryRun)
            {
                Write(state, pic, next.Id, want, next, chosen: Chosen(next, pic));
                Raise("changed", $"Out of office: scheduled your reply for {Describe(next)}, {Window(want)}", next.Id, want, "Nothing to do; just so you know.",
                    AudienceOf(next, pic));
            }
            return;
        }
        said.Add($"{next.Id}: to ask about{(config.Mode == Mode.Act ? " (act mode, but these messages haven't been approved yet)" : "")}: {Window(want)}. "
                 + $"`calendar-out-of-office apply {next.Id} --yes` schedules it.");
        if (!dryRun)
            Propose(state, next, want, pic);
    }

    /// <summary>A period raised as a question (once for each version of its reply), and a notification when it starts
    /// within notify_working_days.</summary>
    private void Propose(JsonObject state, Period p, Reply want, Picture pic)
    {
        var raised = State.Table(state, "raised");
        // Keyed on the reply as the calendar gives it, not the start moved to now for one already running, so a
        // period isn't raised again on every run; with its audience, so a choice made ahead updates the question.
        var audience = AudienceOf(p, pic);
        var fp = Planner.Render(config, p, audience.Value).Fingerprint();
        if (raised[p.Id]?.GetValue<string>() != fp)
        {
            raised[p.Id] = fp;
            Raise(Item, $"Out of office: schedule your reply for {Describe(p)}?", p.Id, want,
                $"Check the dates, who outside gets one and both messages, then `calendar-out-of-office apply {p.Id} --yes` (on your OK only); "
                + $"`calendar-out-of-office apply {p.Id} --external-audience none|contactsOnly|all --yes` chooses who outside gets one for this period.",
                audience);
        }
        if (WorkingDaysUntil(DateOnly.FromDateTime(pic.Now), DateOnly.FromDateTime(p.On)) <= config.NotifyWorkingDays)
            NotifyOnce(state, "proposed:" + p.Id, "Out of office: a reply to approve",
                [$"{Describe(p)}: {Window(want)}.", $"calendar-out-of-office apply {p.Id} --yes"]);
    }

    private int WorkingDaysUntil(DateOnly from, DateOnly to)
    {
        var n = 0;
        for (var d = from; d < to; d = d.AddDays(1))
            if (config.WorkingDays.Contains(d.DayOfWeek))
                n++;
        return n;
    }

    /// <summary>Apply: schedule one period's reply on the user's say-so. Without --yes, shows what it would write.</summary>
    public int Apply(JsonObject state, string id, bool yes, string? audience = null)
    {
        var pic = Look(state);
        var p = Find(pic, id);
        if (Refused(p, state) is { } why && !why.StartsWith("you changed", StringComparison.Ordinal))
            throw new UsageException($"{p.Id} can't be scheduled: {why}.");
        if (pic.Owner == Owner.Theirs)
            throw new UsageException($"Your automatic replies are {pic.Setting.Status} ({SettingWindow(pic.Setting)}), set outside calendar-out-of-office, so nothing was written. Turn yours off first (in Outlook, or `calendar-out-of-office clear --force --yes`).");
        var keep = pic.Owner == Owner.Ours && Holding(pic)?.Id == p.Id && State.Time(pic.Written!["start"]) is { } s && s <= pic.NowUtc ? s : (DateTime?)null;
        if (pic.Owner == Owner.Ours && Holding(pic)?.Id != p.Id)
            throw new UsageException($"The reply for {pic.Written!["period"]} is on; only one can be scheduled at a time. `calendar-out-of-office clear --yes` turns it off first.");
        var current = AudienceOf(p, pic);
        var a = audience is null ? current : new Audience(audience, current.Settings, true);
        var want = Want(p, pic, keep, a.Value);
        host.Out.WriteLine(Preview(p, want, a));
        if (!yes)
        {
            host.Out.WriteLine(audience is null
                ? $"\nNothing written. `calendar-out-of-office apply {p.Id} --yes` schedules it; `--external-audience none|contactsOnly|all` chooses who outside gets one for this period."
                : $"\nNothing written. `calendar-out-of-office apply {p.Id} --external-audience {audience} --yes` schedules it.");
            return 0;
        }
        if (!office.CanWrite())
            throw new SignInNeededException("Scheduling a reply needs the MailboxSettings.ReadWrite permission, which this sign-in doesn't have yet. "
                                            + "Run `calendar-out-of-office sign-in` at a console (your organisation may need to approve it first).");
        // The choice is kept before the write, so a write whose answer is lost (settled by reading back later) is
        // compared with the audience it was written with, never the settings'.
        if (a.Chosen)
        {
            var table = State.Table(state, "audience");
            table[p.Id] = State.Choice(a.Value, p);
            if (current.CarriedFrom is { } from)
                table.Remove(from);
        }
        Write(state, pic, p.Id, want, p, keepStart: keep is not null, chosen: a.Chosen ? a.Value : null);
        // Approves the messages (the templates' hash) for act mode; who outside gets them is never part of approving.
        State.Approve(state, Planner.TemplateHash(config, p.IsLeave));
        State.Table(state, "hands_off").Remove(p.Id);
        DoneItems(p.Id);
        State.Save(config.StateFile, state);
        host.Out.WriteLine($"\nScheduled: {Window(want)}.");
        return 0;
    }

    private Period Find(Picture pic, string id) =>
        pic.Periods.FirstOrDefault(x => string.Equals(x.Id, id, StringComparison.OrdinalIgnoreCase))
        ?? throw new UsageException($"No period '{id}' in the next {config.DaysAhead} days. `calendar-out-of-office plan` lists them.");

    /// <summary>Audience: choose who outside gets a period's reply ahead of time (none, contactsOnly, all), or go back to
    /// the settings (default). Local only: nothing is sent. A reply already written is changed only by apply.</summary>
    public int SetAudience(JsonObject state, string id, string value)
    {
        var pic = Look(state);
        var p = Find(pic, id);
        var current = AudienceOf(p, pic);
        var wanted = value == "default" ? null : value;
        var target = wanted ?? current.Settings;
        if (pic.Owner == Owner.Ours && Holding(pic)?.Id == p.Id)
        {
            if (target == current.Value)
            {
                host.Out.WriteLine($"{p.Id}: its reply is written already, to {Audience.Words(current.Value)} outside: nothing to change.");
                return 0;
            }
            throw new UsageException($"{p.Id}: its reply is written already, to {Audience.Words(current.Value)} outside, so nothing was changed. "
                                     + $"`calendar-out-of-office apply {p.Id} --external-audience {target} --yes` rewrites it (without --yes, it shows it first).");
        }
        var table = State.Table(state, "audience");
        if (current.CarriedFrom is { } from)
            table.Remove(from);
        if (wanted is null)
        {
            table.Remove(p.Id);
            host.Out.WriteLine($"{p.Id}: who outside gets a reply is back to the settings: {Audience.Words(current.Settings)}.");
        }
        else
        {
            table[p.Id] = State.Choice(wanted, p);
            host.Out.WriteLine($"{p.Id}: who outside gets a reply: {new Audience(wanted, current.Settings, true).Describe()}.");
        }
        if (table.Count == 0)
            state.Remove("audience");
        State.Save(config.StateFile, state);
        host.Out.WriteLine($"Nothing is sent now: it's used when the reply is written (`calendar-out-of-office apply {p.Id} --yes`"
                           + (config.Mode == Mode.Act ? ", or act mode when it's due" : "") + ").");
        return 0;
    }

    /// <summary>Clear: turn off the reply calendar-out-of-office wrote; another one only with force.</summary>
    public int Clear(JsonObject state, bool yes, bool force)
    {
        var pic = Look(state);
        if (pic.Owner == Owner.Idle)
        {
            host.Out.WriteLine("Automatic replies are already off (or finished): nothing to do.");
            return 0;
        }
        if (pic.Owner == Owner.Theirs && !force)
            throw new UsageException($"The reply that's on ({pic.Setting.Status}, {SettingWindow(pic.Setting)}) wasn't written by calendar-out-of-office, so it's left alone. --force turns it off anyway.");
        host.Out.WriteLine($"Would turn off: {pic.Setting.Status}, {SettingWindow(pic.Setting)}.");
        if (!yes)
        {
            host.Out.WriteLine("Nothing changed. Add --yes to turn it off.");
            return 0;
        }
        var id = pic.Written?["period"]?.GetValue<string>() ?? "by-hand";
        Write(state, pic, id, null, null);
        if (pic.Owner == Owner.Ours)
            State.Table(state, "hands_off")[id] = State.Stamp(pic.NowUtc);
        State.Save(config.StateFile, state);
        host.Out.WriteLine("Turned off.");
        return 0;
    }

    /// <summary>Schedule (or, with no reply, turn off), then read back. "writing" is saved first, so an answer that never
    /// comes is settled by the next run's read, never by writing again blind.</summary>
    private void Write(JsonObject state, Picture pic, string id, Reply? reply, Period? period, bool keepStart = false, string? chosen = null)
    {
        var expected = reply?.Fingerprint() ?? "idle";
        var previous = pic.Setting.Idle(pic.NowUtc) ? "idle" : pic.Setting.Fingerprint();
        var attempts = state["writing"] is JsonObject w && w["period"]?.GetValue<string>() == id ? (w["attempts"]?.GetValue<int>() ?? 1) + 1 : 1;
        var record = new JsonObject
        {
            ["period"] = id,
            ["fingerprint"] = expected,
            ["previous"] = previous,
            ["attempts"] = attempts,
            ["at"] = State.Stamp(pic.NowUtc),
        };
        if (reply is not null)
        {
            record["start"] = State.Stamp(reply.StartUtc);
            record["end"] = State.Stamp(reply.EndUtc);
            record["kind"] = period?.KindName;
            record["templates"] = period is null ? null : Planner.TemplateHash(config, period.IsLeave);
            if (chosen is not null)
                record["audience"] = chosen;   // a value, not a message
        }
        state["writing"] = record;
        State.Save(config.StateFile, state);

        Exception? failure = null;
        try
        {
            if (reply is null)
                office.Disable();
            else if (!(period is { IsLeave: false } && config.ConferenceAction == ConferenceAction.Status))
                office.Schedule(reply, keepStart);
        }
        catch (SignInNeededException)
        {
            state.Remove("writing");
            State.Save(config.StateFile, state);
            throw;
        }
        catch (Exception error)
        {
            failure = error;
        }

        // A write sent once whose answer was a 5xx or never came (GraphDeliveryUnknownException) may or may not have
        // landed: this read says which, and it's never resent blind.
        var unknown = failure is GraphDeliveryUnknownException;
        Setting back;
        try
        {
            back = office.Read();
        }
        catch (Exception error) when (failure is not null && error is not SignInNeededException)
        {
            // "writing" is saved, so the next run reads it back before writing anything.
            throw new InvalidOperationException($"The write for {id} {(unknown ? "may or may not have landed" : "failed")} ({failure.Message}), "
                                                + $"and reading it back failed too ({error.Message}); the next run reads it back before writing anything.", error);
        }
        var now = back.Idle(pic.NowUtc) ? "idle" : back.Fingerprint();
        var statusOnly = period is { IsLeave: false } && config.ConferenceAction == ConferenceAction.Status;
        if (now == expected || statusOnly)
        {
            state.Remove("writing");
            if (statusOnly)
                State.Table(state, "status_set")[id] = State.Stamp(pic.NowUtc);
            if (reply is null || statusOnly)
                state.Remove("written");
            else
            {
                record.Remove("previous");
                record.Remove("attempts");
                state["written"] = record;
            }
            log(reply is null ? $"turned off the reply ({id})" : $"scheduled {id}: {Window(reply)}");
        }
        else if (failure is not null && now == previous)
        {
            State.Save(config.StateFile, state);
            // Not there when read back. Making it once more (on a later run, after reading back again) is safe: it
            // sets the same reply, and a write that lands late is found by that read and recorded, not repeated.
            throw new InvalidOperationException(unknown
                ? $"Microsoft didn't confirm the write for {id} ({failure.Message}) and the mailbox doesn't show it; "
                  + "the next run reads it back again, and it may be made once more only if it still isn't there."
                : $"The write for {id} didn't land ({failure.Message}); it may be made once more.", failure);
        }
        else
        {
            state.Remove("writing");
            State.Save(config.StateFile, state);
            Raise("failed", $"Out of office: a write for {id} can't be confirmed", id, null,
                "Check your automatic replies in Outlook (`calendar-out-of-office show`); nothing more is written for this period.");
            throw new InvalidOperationException($"After writing {id}, the mailbox holds something else; nothing more is written for it."
                                                + (failure is null ? "" : $" ({failure.Message})"));
        }

        if (period is { IsLeave: false } && reply is not null && config.TeamsStatusMessage && config.ConferenceAction != ConferenceAction.Reply)
        {
            var text = Planner.Fill(config.ConferenceStatus, Planner.Values(config, period));
            office.StatusMessage(text, reply.EndUtc);
            log($"Teams status message for {id}: \"{text}\"");
        }
    }

    /// <summary>A review item for the session (the previous one for the same period and type replaced).</summary>
    private void Raise(string type, string title, string period, Reply? reply, string todo, Audience? audience = null)
    {
        if (config.Review is not { } review)
            return;
        var queue = review.Open(host.Now);
        var old = queue.Items(review.Source).Where(i => i["period"]?.ToString() == period && i["type"]?.ToString() == type)
            .Select(i => i["id"]!.ToString()).ToList();
        if (old.Count > 0)
            queue.Done(old);
        var fields = new JsonObject { ["title"] = title, ["period"] = period, ["todo"] = todo };
        if (reply is not null)
        {
            fields["window"] = Window(reply);
            fields["internal"] = Planner.Plain(reply.Internal);
            fields["external"] = Planner.Plain(reply.External);
            fields["audience"] = reply.Audience;
            if (audience is not null)
                fields["outside"] = audience.Describe();
            fields["command"] = $"calendar-out-of-office apply {period} --yes";
            if (type == Item)
                fields["command_audience"] = $"calendar-out-of-office apply {period} --external-audience none|contactsOnly|all --yes";
        }
        queue.Add(review.Source, type, fields);
    }

    private void DoneItems(string period)
    {
        if (config.Review is not { } review)
            return;
        var queue = review.Open(host.Now);
        var ids = queue.Items(review.Source).Where(i => i["period"]?.ToString() == period && i["type"]?.ToString() is Item or "hand-set")
            .Select(i => i["id"]!.ToString()).ToList();
        if (ids.Count > 0)
            queue.Done(ids);
    }

    private void NotifyOnce(JsonObject state, string key, string title, IEnumerable<string> lines)
    {
        var notified = State.Table(state, "notified");
        if (notified[key] is not null)
            return;
        notified[key] = State.Stamp(host.Now());
        if (config.Notify)
            host.Notify(title, lines);
    }

    public string Window(Reply r) =>
        $"{TimeZoneInfo.ConvertTimeFromUtc(r.StartUtc, config.TimeZone).ToString("ddd d MMM HH:mm", En)} to "
        + $"{TimeZoneInfo.ConvertTimeFromUtc(r.EndUtc, config.TimeZone).ToString("ddd d MMM HH:mm", En)}";

    public string SettingWindow(Setting s) => s.Status switch
    {
        "scheduled" when s.StartUtc is { } a && s.EndUtc is { } b =>
            $"{TimeZoneInfo.ConvertTimeFromUtc(a, config.TimeZone).ToString("ddd d MMM HH:mm", En)} to {TimeZoneInfo.ConvertTimeFromUtc(b, config.TimeZone).ToString("ddd d MMM HH:mm", En)}",
        "alwaysEnabled" => "on until turned off",
        _ => "off",
    };

    public string Describe(Period p) =>
        p.IsLeave
            ? $"leave from {Planner.Date(config, p.First)}, back {Planner.Date(config, p.ReturnDay)}"
            : $"{p.Event ?? "a conference"}, {Planner.Date(config, p.First)}" + (p.Last != p.First ? $" to {Planner.Date(config, p.Last)}" : "");

    public string Preview(Period p, Reply want, Audience audience) =>
        $"{p.Id}: {Describe(p)} ({p.WeekdaysAway:0.#} weekday(s) away)\n"
        + $"  window    {Window(want)}\n"
        + $"  outside   {audience.Describe()}\n"
        + $"  internal  {Planner.Plain(want.Internal)}\n"
        + $"  external  {(want.Audience == "none" ? "not sent (no one outside gets a reply)" : Planner.Plain(want.External))}";
}
