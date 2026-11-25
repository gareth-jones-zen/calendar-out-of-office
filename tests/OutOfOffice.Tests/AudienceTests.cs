using System.Text.Json.Nodes;

namespace OutOfOffice.Tests;

/// <summary>Who outside gets a reply, chosen per period (`apply --external-audience`, `audience`): kept in the state, part
/// of the fingerprint (so a reply written with a chosen audience is never mistaken for one changed by hand), kept
/// through rewrites and carried when a period's first day moves, and never an approval of the messages.</summary>
public sealed class AudienceTests : IDisposable
{
    private readonly World w = new();   // [leave] external_audience is the default, "all"

    public AudienceTests() => w.Office.Leave("2027-03-10", "2027-03-12");   // Wed to Fri; Monday 1 March now

    public void Dispose() => w.Dispose();

    private const string Id = "leave-2027-03-10";

    private static string Field(JsonObject item, string key) => item[key]?.ToString() ?? "";

    private JsonObject? Chosen(string id) => w.State["audience"]?[id] as JsonObject;

    [Fact]
    public void Apply_with_a_chosen_audience_writes_it_and_keeps_it_with_the_period()
    {
        Assert.Equal(0, w.Cli("apply", Id, "--external-audience", "contactsOnly", "--yes"));
        Assert.Equal("contactsOnly", w.Office.Current.Audience);
        Assert.Equal("contactsOnly", Chosen(Id)!["audience"]!.ToString());
        Assert.Equal("2027-03-10", Chosen(Id)!["first"]!.ToString());
        Assert.Equal("2027-03-12", Chosen(Id)!["last"]!.ToString());
        Assert.Equal("contactsOnly", w.State["written"]!["audience"]!.ToString());
        Assert.Contains("outside   your contacts only (chosen for this period; the settings say everyone)", w.Out.ToString());
    }

    [Fact]
    public void A_reply_written_with_a_chosen_audience_is_still_ours_on_the_next_runs()
    {
        w.Cli("apply", Id, "--external-audience", "contactsOnly", "--yes");
        w.Cli("run");
        w.Cli("run");
        Assert.Single(w.Office.Writes);   // not rewritten to the settings' audience
        Assert.Equal("contactsOnly", w.Office.Current.Audience);
        Assert.Null(w.State["hands_off"]?[Id]);   // not mistaken for a reply changed by hand
        Assert.DoesNotContain(w.Items, i => Field(i, "type") == "changed");
        Assert.NotNull(w.State["written"]);
        w.Cli("plan");
        Assert.Contains("written by calendar-out-of-office", w.Out.ToString());
        Assert.Contains($"{Id}: the reply is on as written", w.Out.ToString());
    }

    [Fact]
    public void A_reply_written_with_the_settings_audience_and_then_one_chosen_differ_in_fingerprint()
    {
        var config = w.Config;
        var p = Assert.Single(Planner.Periods(config, Planner.Classify(config, w.Office.Items, new DateOnly(2027, 3, 1), new DateOnly(2027, 4, 30))));
        Assert.NotEqual(Planner.Render(config, p).Fingerprint(), Planner.Render(config, p, "contactsOnly").Fingerprint());
        Assert.Equal(Planner.Render(config, p).Fingerprint(), Planner.Render(config, p, "all").Fingerprint());
    }

    [Fact]
    public void A_longer_period_is_rewritten_with_its_chosen_audience_and_then_left_alone()
    {
        w.Cli("apply", Id, "--external-audience", "contactsOnly", "--yes");
        w.Office.Leave("2027-03-15", "2027-03-16");
        w.Cli("run");
        Assert.Equal("schedule 2027-03-09 17:30 2027-03-17 08:00", w.Office.Writes[^1]);
        Assert.Equal("contactsOnly", w.Office.Current.Audience);
        var item = Assert.Single(w.Items);
        Assert.Equal("changed", Field(item, "type"));
        Assert.Equal("contactsOnly", Field(item, "audience"));
        Assert.Equal("2027-03-16", Chosen(Id)!["last"]!.ToString());   // the days kept up to date
        w.Cli("run");
        Assert.Equal(2, w.Office.Writes.Count);   // the steady state after the rewrite
        Assert.Null(w.State["hands_off"]?[Id]);
    }

    [Fact]
    public void A_moved_first_day_carries_the_choice_to_the_new_period_and_says_so()
    {
        w.Cli("apply", Id, "--external-audience", "contactsOnly", "--yes");
        w.Office.Leave("2027-03-09", "2027-03-09");   // now Tue 9 to Fri 12: a new id
        w.Cli("plan");
        Assert.Contains("leave-2027-03-09: who outside gets a reply (your contacts only) is carried over from leave-2027-03-10, whose first day moved.", w.Out.ToString());
        Assert.Contains("outside   your contacts only (chosen for leave-2027-03-10, carried over when the first day moved; the settings say everyone)", w.Out.ToString());
        Assert.Single(w.Office.Writes);   // plan writes nothing
        Assert.NotNull(Chosen(Id));

        w.Cli("run");
        Assert.Equal("schedule 2027-03-08 17:30 2027-03-15 08:00", w.Office.Writes[^1]);
        Assert.Equal("contactsOnly", w.Office.Current.Audience);
        var item = Assert.Single(w.Items);
        Assert.Equal("changed", Field(item, "type"));
        Assert.Contains("it's leave-2027-03-09; who outside gets a reply (your contacts only, chosen for leave-2027-03-10) is carried over", Field(item, "todo"));
        Assert.Null(Chosen(Id));
        Assert.Equal("contactsOnly", Chosen("leave-2027-03-09")!["audience"]!.ToString());
        Assert.Equal("leave-2027-03-09", w.State["written"]!["period"]!.ToString());
        Assert.Equal("contactsOnly", w.State["written"]!["audience"]!.ToString());

        w.Cli("run");
        Assert.Equal(2, w.Office.Writes.Count);
        Assert.Null(w.State["hands_off"]?["leave-2027-03-09"]);
    }

    [Fact]
    public void A_choice_made_ahead_follows_a_period_not_written_yet_whose_first_day_moves()
    {
        Assert.Equal(0, w.Cli("audience", Id, "none"));
        w.Office.Items.Clear();
        w.Office.Leave("2027-03-11", "2027-03-15");   // Thu 11 to Mon 15: overlaps
        w.Cli("run");
        Assert.Null(Chosen(Id));
        Assert.Equal("none", Chosen("leave-2027-03-11")!["audience"]!.ToString());
        var item = Assert.Single(w.Items);
        Assert.Equal("leave-2027-03-11", Field(item, "period"));
        Assert.Equal("none", Field(item, "audience"));
    }

    [Fact]
    public void A_choice_is_not_carried_to_a_period_whose_days_dont_overlap()
    {
        w.Cli("audience", Id, "contactsOnly");
        w.Office.Items.Clear();
        w.Office.Leave("2027-03-17", "2027-03-19");
        w.Cli("plan");
        Assert.Contains("outside   everyone (the settings)", w.Out.ToString());
        Assert.DoesNotContain("carried", w.Out.ToString());
    }

    [Fact]
    public void Audience_ahead_is_local_updates_the_question_and_is_used_by_apply()
    {
        w.Cli("run");
        Assert.Equal("all", Field(Assert.Single(w.Items), "audience"));
        Assert.Equal(0, w.Cli("audience", Id, "contactsOnly"));
        Assert.Contains("Nothing is sent now", w.Out.ToString());
        Assert.Empty(w.Office.Writes);
        Assert.Equal("disabled", w.Office.Current.Status);
        w.Cli("run");
        var item = Assert.Single(w.Items);
        Assert.Equal("contactsOnly", Field(item, "audience"));
        Assert.Equal("your contacts only (chosen for this period; the settings say everyone)", Field(item, "outside"));
        Assert.Equal($"calendar-out-of-office apply {Id} --yes", Field(item, "command"));
        Assert.Equal($"calendar-out-of-office apply {Id} --external-audience none|contactsOnly|all --yes", Field(item, "command_audience"));
        Assert.Contains("--external-audience", Field(item, "todo"));

        w.Cli("apply", Id);
        Assert.Contains("your contacts only (chosen for this period", w.Out.ToString());
        Assert.Empty(w.Office.Writes);
        w.Cli("apply", Id, "--yes");
        Assert.Equal("contactsOnly", w.Office.Current.Audience);
        w.Cli("run");
        Assert.Single(w.Office.Writes);
    }

    [Fact]
    public void Audience_default_goes_back_to_the_settings()
    {
        w.Cli("audience", Id, "contactsOnly");
        Assert.Equal(0, w.Cli("audience", Id, "default"));
        Assert.Contains("back to the settings: everyone", w.Out.ToString());
        Assert.Null(w.State["audience"]);
        w.Cli("apply", Id, "--yes");
        Assert.Equal("all", w.Office.Current.Audience);
        Assert.Null(w.State["written"]!["audience"]);
    }

    [Fact]
    public void Audience_for_a_written_reply_changes_nothing_and_gives_the_apply_command()
    {
        w.Cli("apply", Id, "--yes");
        Assert.Equal(2, w.Cli("audience", Id, "contactsOnly"));
        Assert.Contains("written already, to everyone outside, so nothing was changed", w.Err.ToString());
        Assert.Contains($"`calendar-out-of-office apply {Id} --external-audience contactsOnly --yes` rewrites it", w.Err.ToString());
        Assert.Null(w.State["audience"]);
        Assert.Equal(0, w.Cli("audience", Id, "all"));
        Assert.Contains("nothing to change", w.Out.ToString());
        w.Cli("run");
        Assert.Single(w.Office.Writes);
    }

    [Fact]
    public void Apply_again_with_another_audience_rewrites_a_written_reply_once()
    {
        w.Cli("apply", Id, "--yes");
        Assert.Equal(0, w.Cli("apply", Id, "--external-audience", "none", "--yes"));
        Assert.Equal(2, w.Office.Writes.Count);
        Assert.Equal("none", w.Office.Current.Audience);
        w.Cli("run");
        w.Cli("run");
        Assert.Equal(2, w.Office.Writes.Count);
        Assert.Null(w.State["hands_off"]?[Id]);
    }

    [Fact]
    public void None_shows_the_external_message_as_not_sent()
    {
        w.Cli("audience", Id, "none");
        w.Cli("plan");
        Assert.Contains("outside   no one (chosen for this period; the settings say everyone)", w.Out.ToString());
        Assert.Contains("external  not sent (no one outside gets a reply)", w.Out.ToString());
        Assert.Contains("internal  I'm on leave", w.Out.ToString());
    }

    [Fact]
    public void Plan_shows_the_settings_audience_when_nothing_is_chosen()
    {
        w.Cli("plan");
        Assert.Contains("outside   everyone (the settings)", w.Out.ToString());
        Assert.Contains("external  I'm away until Monday 15 March.", w.Out.ToString());
    }

    [Fact]
    public void A_chosen_audience_is_not_an_approval_in_act_mode()
    {
        using var act = new World(null, World.Base.Replace("mode = \"ask\"", "mode = \"act\""));
        act.Office.Leave("2027-03-03", "2027-03-05");
        act.Cli("audience", "leave-2027-03-03", "contactsOnly");
        act.Cli("run");
        Assert.Empty(act.Office.Writes);   // due, but the messages aren't approved: asked
        var item = Assert.Single(act.Items);
        Assert.Equal("proposed", Field(item, "type"));
        Assert.Equal("contactsOnly", Field(item, "audience"));
    }

    [Fact]
    public void Act_mode_writes_with_the_audience_chosen_ahead()
    {
        using var act = new World(null, World.Base.Replace("mode = \"ask\"", "mode = \"act\""));
        act.Office.Leave("2027-03-03", "2027-03-05");
        act.Office.Leave("2027-04-07", "2027-04-09");
        act.Cli("apply", "leave-2027-03-03", "--yes");   // approves the leave messages, to everyone
        Assert.Equal("all", act.Office.Current.Audience);
        act.Cli("audience", "leave-2027-04-07", "contactsOnly");
        act.Now = new DateTime(2027, 3, 31, 9, 0, 0);   // March's is over; April's is due
        act.Cli("run");
        Assert.Equal("schedule 2027-04-06 16:30 2027-04-12 07:00", act.Office.Writes[^1]);
        Assert.Equal("contactsOnly", act.Office.Current.Audience);
        Assert.Equal("contactsOnly", act.State["written"]!["audience"]!.ToString());
        act.Cli("run");
        Assert.Equal(2, act.Office.Writes.Count);
    }

    [Fact]
    public void The_choice_is_kept_before_the_write_so_a_lost_answer_settles_against_it()
    {
        w.Office.ThrowBefore = new HttpRequestException("connection reset");
        Assert.Equal(1, w.Cli("apply", Id, "--external-audience", "contactsOnly", "--yes"));
        Assert.Equal("contactsOnly", Chosen(Id)!["audience"]!.ToString());
        Assert.Equal("contactsOnly", w.State["writing"]!["audience"]!.ToString());
        // It lands after all (late), with the chosen audience: the next run records it rather than writing again.
        w.Office.ThrowBefore = null;
        w.Office.Current = new Setting("scheduled", new DateTime(2027, 3, 9, 17, 30, 0, DateTimeKind.Utc), new DateTime(2027, 3, 15, 8, 0, 0, DateTimeKind.Utc),
            "<p>I'm on leave, back on Monday 15 March.</p><p>Please contact Alex Example &amp; the team.</p>", "<p>I'm away until Monday 15 March.</p>", "contactsOnly");
        w.Cli("run");
        Assert.Empty(w.Office.Writes);
        Assert.Equal("contactsOnly", w.State["written"]!["audience"]!.ToString());
        Assert.Null(w.State["writing"]);
        w.Cli("run");
        Assert.Empty(w.Office.Writes);
        Assert.Null(w.State["hands_off"]?[Id]);
    }

    [Fact]
    public void A_choice_is_pruned_once_its_period_is_over()
    {
        w.Cli("audience", Id, "contactsOnly");
        w.Now = new DateTime(2027, 3, 16, 9, 0, 0);
        w.Cli("run");
        Assert.Null(w.State["audience"]);
    }

    [Fact]
    public void The_state_holds_the_choice_but_no_message_text()
    {
        w.Cli("apply", Id, "--external-audience", "contactsOnly", "--yes");
        var text = File.ReadAllText(Path.Combine(w.Dir, "state.json"));
        Assert.Contains("contactsOnly", text);
        Assert.DoesNotContain("Alex Example", text);
        Assert.DoesNotContain("away until", text);
    }

    [Fact]
    public void Show_says_what_was_chosen()
    {
        w.Office.Leave("2027-03-24", "2027-03-26");
        w.Cli("apply", Id, "--external-audience", "contactsOnly", "--yes");
        w.Cli("audience", "leave-2027-03-24", "none");
        Assert.Equal(0, w.Cli("show"));
        var output = w.Out.ToString();
        Assert.Contains("external to   contactsOnly (your contacts only)", output);
        Assert.Contains($"Last written by calendar-out-of-office: {Id}", output);
        Assert.Contains("to your contacts only outside (chosen for that period)", output);
        Assert.Contains("Chosen for leave-2027-03-24: no one outside (none).", output);
    }

    [Theory]
    [InlineData("apply", Id, "--external-audience", "everyone")]
    [InlineData("apply", Id, "--external-audience")]
    [InlineData("audience", Id, "some")]
    [InlineData("audience", Id)]
    [InlineData("audience")]
    [InlineData("clear", "--external-audience", "all")]
    public void Bad_audience_arguments_are_refused(params string[] args)
    {
        Assert.Equal(2, w.Cli(args));
        Assert.Empty(w.Office.Writes);
        Assert.Null(w.State["audience"]);
    }

    [Fact]
    public void Audience_values_take_any_case_and_the_equals_form()
    {
        Assert.Equal(0, w.Cli("audience", Id, "CONTACTSONLY"));
        Assert.Equal("contactsOnly", Chosen(Id)!["audience"]!.ToString());
        Assert.Equal(0, w.Cli("apply", Id, "--external-audience=None"));
        Assert.Contains($"apply {Id} --external-audience none --yes` schedules it", w.Out.ToString());
        Assert.Equal("contactsOnly", Chosen(Id)!["audience"]!.ToString());   // a preview changes nothing
    }

    [Fact]
    public void Audience_for_an_unknown_period_is_refused()
    {
        Assert.Equal(2, w.Cli("audience", "leave-2027-05-03", "none"));
        Assert.Contains("No period 'leave-2027-05-03'", w.Err.ToString());
    }
}
