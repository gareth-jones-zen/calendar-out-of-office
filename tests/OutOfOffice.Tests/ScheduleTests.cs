using WinBackground;

namespace OutOfOffice.Tests;

/// <summary>The task's schedule: [schedule] (background-tools-lib's shared keys), the old top-level keys, and drift.</summary>
public sealed class ScheduleTests
{
    private sealed class RecordingTask(string? description = null) : ITask
    {
        public TaskSchedule? Schedule { get; private set; }
        public string Name => "calendar-out-of-office run";
        public string Enable(string exe, IReadOnlyList<string> args, string? workingDir) => throw new NotSupportedException();
        public string Enable(string exe, IReadOnlyList<string> args, string? workingDir, TaskSchedule schedule)
        {
            Schedule = schedule;
            return exe;
        }
        public bool Disable() => true;
        public bool Stop() => true;
        public IReadOnlyDictionary<string, string>? Status() => description is null ? null : new Dictionary<string, string>
        {
            ["state"] = "Ready", ["command"] = "x", ["description"] = description,
        };
    }

    private static (int Code, string Out, string Err) Cmd(World w, ITask task, params string[] args)
    {
        var output = new StringWriter { NewLine = "\n" };
        var errors = new StringWriter { NewLine = "\n" };
        var code = Cli.Run(["--config", w.ConfigPath, .. args], new CliHost
        {
            Out = output, Err = errors, Now = () => w.Now, Office = _ => w.Office, Notify = (_, _) => { }, Lock = false,
            Env = name => name == "LOCALAPPDATA" ? w.Dir : null, Task = () => task, ProgramDir = w.Dir,
        });
        return (code, output.ToString(), errors.ToString());
    }

    [Fact]
    public void Without_a_schedule_section_the_task_is_as_before_and_its_fingerprint_recorded()
    {
        using var w = new World();
        var task = new RecordingTask();
        var (code, output, error) = Cmd(w, task, "autostart", "on");
        Assert.True(code == 0, error);
        Assert.True(task.Schedule!.AtLogon);
        Assert.Equal(TimeSpan.FromMinutes(15), task.Schedule.Every);
        Assert.Equal(TimeSpan.FromMinutes(1), task.Schedule.AfterUnlock);
        Assert.Equal("v1 job every=15 daily=- logon=yes unlock=1", task.Schedule.Fingerprint);
        Assert.Contains("runs every 15 minutes, at logon, 1 minute(s) after each unlock, and has run now", output);
    }

    [Fact]
    public void The_schedule_section_sets_the_task()
    {
        using var w = new World("[schedule]\nevery_minutes = 30\nafter_unlock_minutes = 0\n");
        var task = new RecordingTask();
        var (code, output, error) = Cmd(w, task, "autostart", "on");
        Assert.True(code == 0, error);
        Assert.Equal(TimeSpan.FromMinutes(30), task.Schedule!.Every);
        Assert.Null(task.Schedule.AfterUnlock);
        Assert.Contains("runs every 30 minutes, at logon, and has run now", output);
    }

    [Fact]
    public void The_old_top_level_keys_still_work_and_are_logged_as_deprecated()
    {
        using var w = new World(null, "every_minutes = 20\nafter_unlock_minutes = 0\n" + World.Base);
        var task = new RecordingTask();
        var (code, _, error) = Cmd(w, task, "autostart", "on");
        Assert.True(code == 0, error);
        Assert.Equal(TimeSpan.FromMinutes(20), task.Schedule!.Every);
        Assert.Null(task.Schedule.AfterUnlock);
        Assert.Equal(["every_minutes", "after_unlock_minutes"], w.Config.DeprecatedScheduleKeys);
        Assert.Equal(0, w.Cli("run"));
        Assert.Contains("every_minutes and after_unlock_minutes at the top level are deprecated: move them to [schedule]",
            File.ReadAllText(Path.Combine(w.Dir, "ooo.log")));
    }

    [Theory]
    [InlineData("every_minutes = 20\n", "[schedule]\nevery_minutes = 15\n", "every_minutes and [schedule] are both given")]
    [InlineData("", "[schedule]\ncolour = 1\n", "unknown setting colour")]
    [InlineData("", "[schedule]\nevery_minutes = 15\ndaily_at = \"07:00\"\n", "set every_minutes or daily_at, not both")]
    public void A_bad_schedule_is_refused(string top, string extra, string message)
    {
        using var w = new World(extra, top + World.Base);
        Assert.Contains(message, Assert.Throws<UsageException>(() => w.Config).Message);
    }

    [Fact]
    public void Status_says_when_the_task_no_longer_matches_its_schedule()
    {
        using var same = new World();
        var registered = new RecordingTask("calendar-out-of-office run (schedule: v1 job every=15 daily=- logon=yes unlock=1)");
        Assert.DoesNotContain("settings changed", Cmd(same, registered, "autostart", "status").Out);
        using var changed = new World("[schedule]\nafter_unlock_minutes = 0\n");
        Assert.Contains("settings changed since this task was registered: run `calendar-out-of-office autostart on`",
            Cmd(changed, registered, "autostart", "status").Out);
        Assert.Contains("registered before its schedule was recorded",
            Cmd(same, new RecordingTask("calendar-out-of-office run"), "autostart", "status").Out);
    }

    [Fact]
    public void A_run_outside_the_schedule_does_nothing()
    {
        using var w = new World("[schedule]\nend_date = \"2027-02-01\"\n");
        Assert.Equal(0, w.Cli("run"));
        Assert.Contains("outside the [schedule] (", File.ReadAllText(Path.Combine(w.Dir, "ooo.log")));
        Assert.Empty(w.Items);
    }
}
