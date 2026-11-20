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
        Assert.Contains("filled in by out-of-office", Error(World.Base.Replace("cover = \"Alex Example\"", "return_day = \"never\"")));
    }

    [Fact]
    public void Mode_audience_action_and_times_are_checked()
    {
        Assert.Contains("mode", Error(World.Base.Replace("mode = \"ask\"", "mode = \"maybe\"")));
        Assert.Contains("external_audience", Error(World.Base.Replace("min_working_days = 3", "min_working_days = 3\nexternal_audience = \"everyone\"")));
        Assert.Contains("conference.action", Error(World.Base.Replace("external = \"At {event}.\"", "external = \"At {event}.\"\naction = \"shout\"")));
        Assert.Contains("window.on", Error(World.Base + "\n[window]\non = \"5.30pm\"\n"));
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
