using System.Text.Json.Nodes;

namespace OutOfOffice.Tests;

/// <summary>Runs, apply and clear against a fake mailbox: what's written, raised and left alone.</summary>
public sealed class RunnerTests : IDisposable
{
    private readonly World w = new();

    public RunnerTests() => w.Office.Leave("2027-03-10", "2027-03-12");   // Wed to Fri; Monday 1 March now

    public void Dispose() => w.Dispose();

    private static string Field(JsonObject item, string key) => item[key]?.ToString() ?? "";

    [Fact]
    public void Ask_mode_raises_a_period_once_and_writes_nothing()
    {
        Assert.Equal(0, w.Cli("run"));
        Assert.Equal(0, w.Cli("run"));
        Assert.Empty(w.Office.Writes);
        var item = Assert.Single(w.Items);
        Assert.Equal("proposed", Field(item, "type"));
        Assert.Equal("leave-2027-03-10", Field(item, "period"));
        Assert.Equal("calendar-out-of-office apply leave-2027-03-10 --yes", Field(item, "command"));
        Assert.Equal("Tue 9 Mar 17:30 to Mon 15 Mar 08:00", Field(item, "window"));
        Assert.Contains("Alex Example", Field(item, "internal"));
        Assert.Empty(w.Notified);   // more than five working days ahead
    }

    [Fact]
    public void A_period_to_approve_soon_is_notified_once()
    {
        w.Now = new DateTime(2027, 3, 5, 9, 0, 0);
        w.Cli("run");
        w.Cli("run");
        Assert.Single(w.Notified);
    }

    [Fact]
    public void The_state_holds_no_message_text()
    {
        w.Cli("run");
        w.Cli("apply", "leave-2027-03-10", "--yes");
        var text = File.ReadAllText(Path.Combine(w.Dir, "state.json"));
        Assert.DoesNotContain("Alex Example", text);
        Assert.DoesNotContain("on leave", text);
    }

    [Fact]
    public void Apply_without_yes_changes_nothing()
    {
        Assert.Equal(0, w.Cli("apply", "leave-2027-03-10"));
        Assert.Empty(w.Office.Writes);
        Assert.Contains("Nothing written", w.Out.ToString());
        Assert.Contains("Tue 9 Mar 17:30 to Mon 15 Mar 08:00", w.Out.ToString());
    }

    [Fact]
    public void Apply_needs_the_write_permission()
    {
        w.Office.Writable = false;
        Assert.Equal(2, w.Cli("apply", "leave-2027-03-10", "--yes"));
        Assert.Contains("MailboxSettings.ReadWrite", w.Err.ToString());
        Assert.Empty(w.Office.Writes);
        Assert.Null(w.State["writing"]);
    }

    [Fact]
    public void Apply_schedules_reads_back_and_clears_the_question()
    {
        w.Cli("run");
        Assert.Equal(0, w.Cli("apply", "leave-2027-03-10", "--yes"));
        Assert.Equal(["schedule 2027-03-09 17:30 2027-03-15 08:00"], w.Office.Writes);
        Assert.Equal("scheduled", w.Office.Current.Status);
        Assert.Equal("leave-2027-03-10", w.State["written"]!["period"]!.ToString());
        Assert.Null(w.State["writing"]);
        Assert.Empty(w.Items);
        // Then runs leave it alone: Exchange's wrapping doesn't make it look changed.
        w.Cli("run");
        Assert.Single(w.Office.Writes);
    }

    [Fact]
    public void A_reply_set_by_hand_is_never_touched_and_is_raised_once()
    {
        w.Office.Current = new Setting("alwaysEnabled", null, null, "<p>Back soon</p>", "<p>Back soon</p>", "all");
        w.Cli("run");
        w.Cli("run");
        Assert.Empty(w.Office.Writes);
        var item = Assert.Single(w.Items);
        Assert.Equal("hand-set", Field(item, "type"));
        Assert.Single(w.Notified);
        Assert.Equal(2, w.Cli("apply", "leave-2027-03-10", "--yes"));
        Assert.Equal(2, w.Cli("clear", "--yes"));
        Assert.Empty(w.Office.Writes);
        Assert.Equal(0, w.Cli("clear", "--yes", "--force"));
        Assert.Equal(["disable"], w.Office.Writes);
    }

    [Fact]
    public void A_schedule_that_has_ended_is_free()
    {
        w.Office.Current = new Setting("scheduled", new DateTime(2027, 2, 1, 17, 30, 0, DateTimeKind.Utc),
            new DateTime(2027, 2, 8, 8, 0, 0, DateTimeKind.Utc), "<p>x</p>", "<p>x</p>", "all");
        w.Cli("run");
        Assert.Equal("proposed", Field(Assert.Single(w.Items), "type"));
    }

    [Fact]
    public void Leave_extended_after_writing_is_rewritten_and_told()
    {
        w.Cli("apply", "leave-2027-03-10", "--yes");
        w.Office.Leave("2027-03-15", "2027-03-16");
        w.Cli("run");
        Assert.Equal("schedule 2027-03-09 17:30 2027-03-17 08:00", w.Office.Writes[^1]);
        var item = Assert.Single(w.Items);
        Assert.Equal("changed", Field(item, "type"));
        Assert.Contains("longer", Field(item, "title"));
        Assert.Contains("Wednesday 17 March", Field(item, "internal"));
    }

    [Fact]
    public void Leave_shortened_is_rewritten_and_cancelled_is_turned_off()
    {
        w.Office.Leave("2027-03-15", "2027-03-16");   // Wed 10 to Tue 16
        w.Cli("apply", "leave-2027-03-10", "--yes");
        w.Office.Items.Clear();
        w.Office.Leave("2027-03-10", "2027-03-12");
        w.Office.Leave("2027-03-15", "2027-03-15", "Holiday (half day) (Example Role)");
        w.Cli("run");
        Assert.Equal("schedule 2027-03-09 17:30 2027-03-15 12:30", w.Office.Writes[^1]);
        Assert.Contains("shorter", Field(w.Items.Single(), "title"));
        w.Office.Items.Clear();
        w.Cli("run");
        Assert.Equal("disable", w.Office.Writes[^1]);
        Assert.Null(w.State["written"]);
        Assert.Contains("turned off", Field(w.Items.Last(), "title"));
    }

    [Fact]
    public void Extended_while_already_away_keeps_the_start()
    {
        w.Cli("apply", "leave-2027-03-10", "--yes");
        w.Now = new DateTime(2027, 3, 11, 10, 0, 0);
        w.Office.Leave("2027-03-15", "2027-03-15");
        w.Cli("run");
        Assert.Equal("schedule 2027-03-09 17:30 2027-03-16 08:00 keep-start", w.Office.Writes[^1]);
        Assert.Equal("scheduled", w.Office.Current.Status);
        Assert.Equal(new DateTime(2027, 3, 9, 17, 30, 0, DateTimeKind.Utc), w.Office.Current.StartUtc);
    }

    [Fact]
    public void Our_reply_edited_by_hand_becomes_theirs_for_that_period()
    {
        w.Cli("apply", "leave-2027-03-10", "--yes");
        w.Office.Current = w.Office.Current with { Internal = "<p>Edited by me</p>" };
        w.Cli("run");
        w.Office.Leave("2027-03-15", "2027-03-16");
        w.Cli("run");
        Assert.Single(w.Office.Writes);
        Assert.Equal("changed", Field(w.Items.Single(i => Field(i, "type") == "changed"), "type"));
        Assert.NotNull(w.State["hands_off"]!["leave-2027-03-10"]);
    }

    [Fact]
    public void Our_reply_turned_off_by_hand_stays_off()
    {
        w.Cli("apply", "leave-2027-03-10", "--yes");
        w.Office.Current = w.Office.Current with { Status = "disabled" };
        w.Cli("run");
        w.Cli("run");
        Assert.Single(w.Office.Writes);
        Assert.DoesNotContain(w.Items, i => Field(i, "type") == "proposed");
    }

    [Fact]
    public void Act_mode_writes_approved_messages_when_due_and_asks_about_new_ones()
    {
        using var act = new World(null, World.Base.Replace("mode = \"ask\"", "mode = \"act\""));
        act.Office.Leave("2027-03-10", "2027-03-12");
        act.Office.Leave("2027-04-07", "2027-04-09");
        act.Cli("run");
        Assert.Empty(act.Office.Writes);   // not approved yet: asked
        Assert.Equal("proposed", Field(Assert.Single(act.Items), "type"));
        act.Cli("apply", "leave-2027-03-10", "--yes");   // approves the leave messages
        act.Now = new DateTime(2027, 3, 15, 9, 0, 0);   // that one's over
        act.Office.Current = act.Office.Current with { Status = "scheduled", EndUtc = new DateTime(2027, 3, 15, 8, 0, 0, DateTimeKind.Utc) };
        act.Cli("run");
        Assert.Single(act.Office.Writes);   // the March one finished by itself; the April one isn't due (7 days before)
        act.Now = new DateTime(2027, 3, 31, 9, 0, 0);
        act.Cli("run");
        Assert.Equal("schedule 2027-04-06 16:30 2027-04-12 07:00", act.Office.Writes[^1]);
        Assert.Contains(act.Items, i => Field(i, "type") == "changed" && Field(i, "title").Contains("scheduled"));
    }

    [Fact]
    public void A_lost_answer_that_landed_is_recorded_without_writing_again()
    {
        w.Office.ThrowAfter = new HttpRequestException("connection reset");
        Assert.Equal(0, w.Cli("apply", "leave-2027-03-10", "--yes"));
        Assert.Single(w.Office.Writes);
        Assert.NotNull(w.State["written"]);
        Assert.Null(w.State["writing"]);
    }

    [Fact]
    public void A_write_that_didnt_land_may_be_made_once_more_and_no_more()
    {
        w.Office.ThrowBefore = new HttpRequestException("connection reset");
        Assert.Equal(1, w.Cli("apply", "leave-2027-03-10", "--yes"));
        Assert.Equal(1, w.State["writing"]!["attempts"]!.GetValue<int>());
        w.Cli("run", "--dry-run");
        Assert.Contains("didn't land: it may be made once more", w.Out.ToString());
        Assert.Equal(1, w.Cli("apply", "leave-2027-03-10", "--yes"));   // the second and last try fails too
        w.Cli("run");
        Assert.Null(w.State["writing"]);
        Assert.Contains(w.Items, i => Field(i, "type") == "failed");
        Assert.Empty(w.Office.Writes);
        w.Office.ThrowBefore = null;
        Assert.Equal(0, w.Cli("apply", "leave-2027-03-10", "--yes"));   // on the user's say-so, again
        Assert.Equal("scheduled", w.Office.Current.Status);
    }

    [Fact]
    public void A_mailbox_holding_something_else_after_a_write_is_a_failure()
    {
        w.Office.Instead = new Setting("alwaysEnabled", null, null, "<p>someone else</p>", "", "none");
        Assert.Equal(1, w.Cli("apply", "leave-2027-03-10", "--yes"));
        Assert.Contains("holds something else", w.Err.ToString());
        Assert.Null(w.State["written"]);
    }

    [Fact]
    public void Clear_turns_off_ours_and_that_period_is_left_alone_after()
    {
        w.Cli("apply", "leave-2027-03-10", "--yes");
        Assert.Equal(0, w.Cli("clear"));
        Assert.Single(w.Office.Writes);
        Assert.Equal(0, w.Cli("clear", "--yes"));
        Assert.Equal("disable", w.Office.Writes[^1]);
        w.Cli("run");
        Assert.Equal(2, w.Office.Writes.Count);
    }

    [Fact]
    public void Only_one_reply_at_a_time()
    {
        w.Office.Leave("2027-03-24", "2027-03-26");
        w.Cli("apply", "leave-2027-03-10", "--yes");
        Assert.Equal(2, w.Cli("apply", "leave-2027-03-24", "--yes"));
        Assert.Contains("only one can be scheduled at a time", w.Err.ToString());
    }

    [Fact]
    public void An_implausibly_long_period_is_refused()
    {
        using var short_ = new World(null, World.Base.Replace("mode = \"ask\"", "mode = \"ask\"\nmax_days = 5"));
        short_.Office.Leave("2027-03-10", "2027-03-19");
        short_.Cli("run");
        Assert.Empty(short_.Items);
        Assert.Equal(2, short_.Cli("apply", "leave-2027-03-10", "--yes"));
        Assert.Contains("max_days", short_.Err.ToString());
    }

    [Fact]
    public void Teams_status_is_off_unless_asked_for()
    {
        w.Office.Add("Example Summit", "2027-03-03T00:00", "2027-03-04T00:00", allDay: true, showAs: "free");
        w.Cli("apply", "conference-2027-03-03", "--yes");
        Assert.Empty(w.Office.Statuses);

        using var teams = new World(null, World.Base.Replace("[placeholders]", "[teams]\nstatus_message = true\n[placeholders]")
            .Replace("external = \"At {event}.\"", "external = \"At {event}.\"\naction = \"both\""));
        teams.Office.Add("Example Summit", "2027-03-03T00:00", "2027-03-04T00:00", allDay: true, showAs: "free");
        teams.Cli("apply", "conference-2027-03-03", "--yes");
        var (text, until) = Assert.Single(teams.Office.Statuses);
        Assert.Equal("At Example Summit until Wednesday 3 March", text);
        Assert.Equal(new DateTime(2027, 3, 4, 8, 0, 0, DateTimeKind.Utc), until);
    }

    [Fact]
    public void Plan_writes_nothing_and_says_whether_a_write_is_possible()
    {
        w.Office.Writable = false;
        Assert.Equal(0, w.Cli("plan"));
        Assert.Empty(w.Office.Writes);
        Assert.Empty(w.Items);
        Assert.False(File.Exists(Path.Combine(w.Dir, "state.json")));
        var output = w.Out.ToString();
        Assert.Contains("leave-2027-03-10: leave from Wednesday 10 March, back Monday 15 March (3 weekday(s) away)", output);
        Assert.Contains("would need MailboxSettings.ReadWrite", output);
    }

    [Fact]
    public void Disabled_does_nothing()
    {
        using var off = new World(null, World.Base.Replace("mode = \"ask\"", "mode = \"ask\"\nenabled = false"));
        off.Office.Leave("2027-03-10", "2027-03-12");
        off.Cli("run");
        Assert.Empty(off.Items);
    }
}
