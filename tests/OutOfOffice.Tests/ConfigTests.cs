namespace OutOfOffice.Tests;

public sealed class ConfigTests
{
    private static string Error(string settings)
    {
        using var w = new World(null, settings);
        return Assert.Throws<UsageException>(() => w.Config).Message;
    }

    [Fact]
    public void An_unknown_placeholder_is_an_error()
    {
        Assert.Contains("{covr}", Error(World.Base.Replace("Please contact {cover}", "Please contact {covr}")));
    }

    [Fact]
    public void A_placeholder_cant_reuse_a_built_in_name()
    {
        Assert.Contains("filled in by calendar-out-of-office", Error(World.Base.Replace("cover = \"Alex Example\"", "return_day = \"never\"")));
    }

    [Fact]
    public void Mode_audience_action_and_times_are_checked()
    {
        Assert.Contains("mode", Error(World.Base.Replace("mode = \"ask\"", "mode = \"maybe\"")));
        Assert.Contains("external_audience", Error(World.Base.Replace("min_working_days = 3", "min_working_days = 3\nexternal_audience = \"everyone\"")));
        Assert.Contains("conference.action", Error(World.Base.Replace("external = \"At {event}.\"", "external = \"At {event}.\"\naction = \"shout\"")));
        Assert.Contains("window.on", Error(World.Base + "\n[window]\non = \"5.30pm\"\n"));
    }

    private static int DaysAhead(string toolSettings, string? baseSettings)
    {
        using var w = new World(null, (baseSettings is null ? "" : "base = \"base.toml\"\n") + toolSettings + World.Base);
        if (baseSettings is not null)
            File.WriteAllText(Path.Combine(w.Dir, "base.toml"), baseSettings);
        return Config.Load(w.ConfigPath, null, new DateOnly(2026, 11, 22)).DaysAhead;
    }

    [Fact]
    public void How_far_ahead_comes_from_the_base_files_calendar_months_when_not_set()
    {
        // 22 Nov 2026 to 22 Feb 2027: calendar months, not 3 x 30 days.
        Assert.Equal(92, DaysAhead("", "calendar_months_ahead = 3\n"));
        Assert.Equal(92, DaysAhead("calendar_months_ahead = 3\n", null));
    }

    [Fact]
    public void The_tools_own_days_ahead_wins_over_calendar_months()
    {
        Assert.Equal(42, DaysAhead("days_ahead = 42\n", "calendar_months_ahead = 3\n"));
        Assert.Equal(10, DaysAhead("days_ahead = 10\ncalendar_months_ahead = 3\n", null));
    }

    [Fact]
    public void How_far_ahead_defaults_to_42_days_with_neither()
    {
        Assert.Equal(42, DaysAhead("", null));
        Assert.Equal(42, DaysAhead("", "login_hint = \"you@example.com\"\n"));
    }

    [Fact]
    public void Calendar_months_ahead_is_checked()
    {
        Assert.Contains("calendar_months_ahead", Error("calendar_months_ahead = 0\n" + World.Base));
        Assert.Contains("calendar_months_ahead", Error("calendar_months_ahead = 13\n" + World.Base));
    }

    [Fact]
    public void An_unknown_setting_is_an_error()
    {
        Assert.Contains("unknown setting 'dayz_ahead'", Error("dayz_ahead = 3\n" + World.Base));
    }

    [Fact]
    public void Config_creates_a_template_that_loads()
    {
        var dir = Path.Combine(Path.GetTempPath(), "ooo-config-" + Guid.NewGuid().ToString("N"));
        try
        {
            var path = Config.CreateDefault(Path.Combine(dir, "config.toml"));
            var config = Config.Load(path);
            Assert.Equal(Mode.Ask, config.Mode);
            Assert.Equal(3, config.LeaveMessages.MinWorkingDays);
        }
        finally
        {
            Directory.Delete(dir, true);
        }
    }
}
