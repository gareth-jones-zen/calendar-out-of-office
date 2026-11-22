using OutOfOffice;

namespace OutOfOffice.Tests;

/// <summary>[leave] read through background-tools-lib's shared LeaveSettings (0.5.0): a base file's [leave] supplying
/// the shared keys, [leave]'s bank_holiday_subjects / bank_holiday_categories, and [bank_holiday]'s older subjects /
/// categories still read the same.</summary>
public sealed class LeaveSharedTests
{
    private static List<string> Days(World world)
    {
        world.Office.Leave("2027-03-24", "2027-03-26");
        world.Office.BankHoliday("2027-03-29");
        world.Office.Add("Holiday (half day) (Example Role)", "2027-03-30T00:00", "2027-03-30T23:59");
        var config = world.Config;
        var days = Planner.Classify(config, world.Office.Events(new(2027, 3, 1), new(2027, 4, 30)), new(2027, 3, 1), new(2027, 4, 30));
        return days.Select(d => $"{d.Date:yyyy-MM-dd} {d.Kind}").ToList();
    }

    private static List<string> Expected()
    {
        using var world = new World();
        return Days(world);
    }

    [Fact]
    public void The_shared_bank_holiday_keys_in_leave_classify_as_before()
    {
        using var world = new World("", World.Base.Replace("min_working_days = 3",
            "min_working_days = 3\nbank_holiday_subjects = '^\\s*bank holiday\\b'\nbank_holiday_categories = [\"Bank holiday\"]"));
        Assert.Equal(Expected(), Days(world));
        Assert.Empty(world.Config.LeaveNotes);
    }

    [Fact]
    public void Older_bank_holiday_subjects_and_categories_still_work_with_a_note()
    {
        using var world = new World("", World.Base.Replace("[placeholders]", "[bank_holiday]\ncategories = [\"Bank holiday\"]\n[placeholders]"));
        Assert.Equal(Expected(), Days(world));
        Assert.Contains(world.Config.LeaveNotes, n => n.Contains("bank_holiday.categories is deprecated"));
        Assert.Contains("Bank holiday", world.Config.BankHolidayCategories);
    }

    [Fact]
    public void Both_forms_of_a_bank_holiday_key_are_an_error()
    {
        using var world = new World("", World.Base
            .Replace("min_working_days = 3", "min_working_days = 3\nbank_holiday_categories = [\"Bank holiday\"]")
            .Replace("[placeholders]", "[bank_holiday]\ncategories = [\"Bank holiday\"]\n[placeholders]"));
        var error = Assert.Throws<UsageException>(() => world.Config);
        Assert.Contains("both given", error.Message);
    }

    [Fact]
    public void A_base_files_leave_supplies_the_shared_keys()
    {
        using var world = new World("", "base = \"base.toml\"\n" + World.Base.Replace("min_working_days = 3", "min_working_days = 3"));
        File.WriteAllText(Path.Combine(world.Dir, "base.toml"), """
            [leave]
            subjects = '^\s*(holiday\b|annual leave|on leave|summer holidays?)'
            categories = []
            min_hours = 6
            half_day = '\(half day\)'
            bank_holiday_subjects = '^\s*bank holiday\b'
            bank_holiday_categories = ["Bank holiday"]
            """);
        Assert.Equal(Expected(), Days(world));
        Assert.Equal(6, world.Config.Leave.MinHours);
    }

    [Fact]
    public void A_base_file_changes_what_counts_as_leave()
    {
        using var world = new World("", "base = \"base.toml\"\n" + World.Base);
        File.WriteAllText(Path.Combine(world.Dir, "base.toml"), "[leave]\nsubjects = '^\\s*vacation\\b'\n");
        world.Office.Leave("2027-03-24", "2027-03-26", "Vacation");
        var config = world.Config;
        var days = Planner.Classify(config, world.Office.Events(new(2027, 3, 1), new(2027, 4, 30)), new(2027, 3, 1), new(2027, 4, 30));
        Assert.Contains(days, d => d.Date == new DateOnly(2027, 3, 24) && d.Kind == Kind.Leave);
    }

    [Fact]
    public void An_unknown_leave_setting_is_an_error()
    {
        using var world = new World("", World.Base.Replace("min_working_days = 3", "min_working_days = 3\nhalf_days = 'x'"));
        var error = Assert.Throws<UsageException>(() => world.Config);
        Assert.Contains("half_days", error.Message);
    }
}
