namespace OutOfOffice.Tests;

/// <summary>Days and periods from made-up calendars.</summary>
public sealed class PlannerTests : IDisposable
{
    private readonly World w = new();

    public void Dispose() => w.Dispose();

    private List<Period> Periods(string first = "2027-03-01", string last = "2027-05-31")
    {
        var config = w.Config;
        var days = Planner.Classify(config, w.Office.Events(DateOnly.Parse(first), DateOnly.Parse(last)), DateOnly.Parse(first), DateOnly.Parse(last));
        return Planner.Periods(config, days);
    }

    [Fact]
    public void Christmas_leave_joins_bank_holidays_and_weekends_and_ends_the_first_working_day_back()
    {
        // Wed 15 to Fri 31 December 2027, one entry; bank holidays Mon 27 and Tue 28 Dec, Mon 3 Jan.
        w.Office.Leave("2027-12-15", "2027-12-31");
        w.Office.BankHoliday("2027-12-27");
        w.Office.BankHoliday("2027-12-28");
        w.Office.BankHoliday("2028-01-03");
        var p = Assert.Single(Periods("2027-12-01", "2028-01-31"));
        Assert.True(p.IsLeave);
        Assert.Equal("leave-2027-12-15", p.Id);
        Assert.Equal(new DateTime(2027, 12, 14, 17, 30, 0), p.On);
        Assert.Equal(new DateOnly(2028, 1, 3), p.Days[^1].Date);   // through the bank holiday Monday
        Assert.Equal(new DateOnly(2028, 1, 4), p.ReturnDay);
        Assert.Equal(new DateTime(2028, 1, 4, 8, 0, 0), p.Off);
        // 11 leave weekdays (15-17, 20-24, 29-31) and 3 bank holidays.
        Assert.Equal(14, p.WeekdaysAway);
        Assert.Equal("Tuesday 4 January", Planner.Values(w.Config, p)["return_day"]);
    }

    [Fact]
    public void Two_days_of_leave_is_under_the_threshold_and_three_isnt()
    {
        w.Office.Leave("2027-03-03", "2027-03-04");
        Assert.Empty(Periods());
        w.Office.Leave("2027-03-10", "2027-03-12");
        var p = Assert.Single(Periods());
        Assert.Equal(3, p.WeekdaysAway);
        Assert.Equal(new DateTime(2027, 3, 9, 17, 30, 0), p.On);
        Assert.Equal(new DateTime(2027, 3, 15, 8, 0, 0), p.Off);   // Friday's leave: back Monday
    }

    [Fact]
    public void A_bank_holiday_alone_is_no_period_but_joined_to_leave_it_counts()
    {
        w.Office.BankHoliday("2027-03-29");
        Assert.Empty(Periods());
        w.Office.Leave("2027-03-30", "2027-03-31");
        var p = Assert.Single(Periods());
        Assert.Equal(3, p.WeekdaysAway);
        Assert.Equal(new DateTime(2027, 3, 26, 17, 30, 0), p.On);   // the Friday before the bank holiday Monday
    }

    [Fact]
    public void A_bank_holiday_alone_is_a_period_when_asked_for()
    {
        using var alone = new World("", World.Base.Replace("[placeholders]", "[bank_holiday]\nalone = true\n[placeholders]")
            .Replace("min_working_days = 3", "min_working_days = 1"));
        alone.Office.BankHoliday("2027-03-29");
        var config = alone.Config;
        var days = Planner.Classify(config, alone.Office.Events(new(2027, 3, 1), new(2027, 4, 30)), new(2027, 3, 1), new(2027, 4, 30));
        Assert.Single(Planner.Periods(config, days));
    }

    [Fact]
    public void Half_days_count_half_and_move_the_window_to_midday()
    {
        w.Office.Leave("2027-03-03", "2027-03-03", "Holiday (half day) (Example Role)");
        w.Office.Leave("2027-03-04", "2027-03-05");
        w.Office.Leave("2027-03-08", "2027-03-08", "Holiday (half day) (Example Role)");
        var p = Assert.Single(Periods());
        Assert.Equal(3, p.WeekdaysAway);
        Assert.Equal(new DateTime(2027, 3, 3, 12, 30, 0), p.On);    // afternoon off assumed
        Assert.Equal(new DateTime(2027, 3, 8, 12, 30, 0), p.Off);   // morning off assumed
        Assert.Equal(new DateOnly(2027, 3, 8), p.ReturnDay);
    }

    [Fact]
    public void A_few_hours_of_leave_is_a_half_day()
    {
        w.Office.Add("Annual leave", "2027-03-03T13:00", "2027-03-03T17:30");
        w.Office.Leave("2027-03-04", "2027-03-08");
        var p = Assert.Single(Periods());
        Assert.Equal(3.5, p.WeekdaysAway);
        Assert.Equal(Kind.HalfDay, p.Days[0].Kind);
    }

    [Fact]
    public void A_conference_by_name_from_one_day_with_its_name_in_the_messages()
    {
        w.Office.Add("Example Expo 2027", "2027-03-10T08:30", "2027-03-10T17:30");
        var p = Assert.Single(Periods());
        Assert.False(p.IsLeave);
        Assert.Equal("Example Expo 2027", p.Event);
        var reply = Planner.Render(w.Config, p);
        Assert.Equal("At Example Expo 2027 until Wednesday 10 March; back Thursday 11 March.", Planner.Plain(reply.Internal));
        Assert.Equal("contactsOnly", reply.Audience);
    }

    [Fact]
    public void A_reminder_to_book_a_conference_isnt_being_there()
    {
        w.Office.Add("Book Example Expo?", "2027-03-10T00:00", "2027-03-11T00:00", allDay: true, showAs: "free");
        Assert.Empty(Periods());
    }

    [Fact]
    public void Working_elsewhere_counts_only_for_most_of_a_day()
    {
        w.Office.Add("Site visit", "2027-03-10T10:00", "2027-03-10T12:00", showAs: "workingElsewhere");
        Assert.Empty(Periods());
        w.Office.Add("Partner day", "2027-03-17T08:00", "2027-03-17T17:00", showAs: "workingElsewhere");
        Assert.Equal("Partner day", Assert.Single(Periods()).Event);
    }

    [Fact]
    public void A_conference_running_into_leave_is_one_leave_period_with_the_event_named()
    {
        w.Office.Add("Example Summit", "2027-03-08T00:00", "2027-03-10T00:00", allDay: true, showAs: "free");
        w.Office.Leave("2027-03-10", "2027-03-12");
        var p = Assert.Single(Periods());
        Assert.True(p.IsLeave);
        Assert.Equal(5, p.WeekdaysAway);
        Assert.Equal("Example Summit", p.Event);
    }

    [Fact]
    public void Someone_elses_meeting_called_holiday_and_a_declined_or_cancelled_one_are_ignored()
    {
        w.Office.Add("Holiday planning", "2027-03-03T00:00", "2027-03-06T00:00", mine: false);
        w.Office.Add("Holiday rota", "2027-03-10T00:00", "2027-03-13T00:00", attendees: 4);
        w.Office.Items.Add(new CalEvent("Example Summit", DateTime.Parse("2027-03-17T00:00"), DateTime.Parse("2027-03-18T00:00"), true, "free", [], false, 3, Declined: true));
        Assert.Empty(Periods());
    }

    [Fact]
    public void Messages_are_filled_in_escaped_and_split_into_paragraphs()
    {
        w.Office.Leave("2027-03-10", "2027-03-12");
        var reply = Planner.Render(w.Config, Assert.Single(Periods()));
        Assert.Equal("<p>I&#39;m on leave, back on Monday 15 March.</p><p>Please contact Alex Example &amp; the team.</p>", reply.Internal);
        Assert.Equal("all", reply.Audience);
        Assert.Equal(new DateTime(2027, 3, 9, 17, 30, 0, DateTimeKind.Utc), reply.StartUtc);   // GMT in March
    }

    [Fact]
    public void Times_in_summer_are_converted_to_utc()
    {
        w.Office.Leave("2027-07-07", "2027-07-09");
        var reply = Planner.Render(w.Config, Assert.Single(Periods("2027-07-01", "2027-07-31")));
        Assert.Equal(new DateTime(2027, 7, 6, 16, 30, 0, DateTimeKind.Utc), reply.StartUtc);
        Assert.Equal(new DateTime(2027, 7, 12, 7, 0, 0, DateTimeKind.Utc), reply.EndUtc);
    }

    [Fact]
    public void Plain_text_ignores_the_wrapping_exchange_adds()
    {
        Assert.Equal("Hi & bye", Planner.Plain("<html><head><style>p {margin:0}</style></head><body>\r\n<p>Hi &amp;</p>\r\n<p>bye</p></body></html>"));
    }
}
