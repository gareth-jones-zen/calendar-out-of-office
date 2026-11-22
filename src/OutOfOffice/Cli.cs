using System.Globalization;
using System.Reflection;
using System.Text.Json.Nodes;
using M365Graph;
using WinBackground;

namespace OutOfOffice;

/// <summary>Everything the command line touches outside itself, so tests can give it fakes.</summary>
public sealed class CliHost
{
    public TextWriter Out { get; init; } = Console.Out;
    public TextWriter Err { get; init; } = Console.Error;
    public Func<string, string?> Env { get; init; } = Config.Environment;
    public Func<DateTime> Now { get; init; } = () => DateTime.Now;
    public Func<Config, IOffice> Office { get; init; } = config => new GraphOffice(GraphOffice.Client(config), config.TimeZone);
    public Action<Config> SignIn { get; init; } = config =>
    {
        var graph = GraphOffice.Client(config);
        graph.TokenAsync(GraphOffice.CalendarScopes, true).GetAwaiter().GetResult();
        graph.TokenAsync(GraphOffice.WriteScopes, true).GetAwaiter().GetResult();
        if (config.TeamsStatusMessage)
            graph.TokenAsync(GraphOffice.PresenceScopes, true).GetAwaiter().GetResult();
    };
    public Action<string, IEnumerable<string>> Notify { get; init; } = Toast.Notify;
    public Func<ITask> Task { get; init; } = () => new ScheduledTask(Cli.TaskName,
        "Schedules your Outlook automatic replies from your calendar's leave and conferences (calendar-out-of-office run).");
    public string ProgramDir { get; init; } = AppContext.BaseDirectory;
    public bool Lock { get; init; } = true;
}

/// <summary>calendar-out-of-office: your Outlook automatic replies scheduled from your calendar. Exit codes: 0 done; 1 failed;
/// 2 settings, arguments or sign-in.</summary>
public static class Cli
{
    public const string Prog = "calendar-out-of-office";
    public const string TaskName = "calendar-out-of-office run";
    public const string BackgroundExe = "calendar-out-of-office-background.exe";

    private const string Usage = "usage: calendar-out-of-office [-h] [--version] [--config CONFIG] {plan,show,apply,clear,run,autostart,sign-in,config} ...";

    // Explicit "\n" throughout (a raw literal would take the .cs file's line endings).
    public const string Help =
        Usage + "\n\n" +
        "calendar-out-of-office: your Outlook automatic replies, scheduled from your calendar: leave of a few days or more, and\n" +
        "conferences, on the evening before and off the morning you're back. Teams shows them too. It never touches a\n" +
        "reply you set yourself, and in ask mode (the default) writes nothing without `apply ... --yes`.\n\n" +
        "    calendar-out-of-office plan                 each period away, its window and messages, and what a run would do\n" +
        "    calendar-out-of-office show                 your automatic replies now, and what calendar-out-of-office last wrote\n" +
        "    calendar-out-of-office apply ID [--yes]     schedule one period's reply (without --yes: show it, change nothing)\n" +
        "    calendar-out-of-office clear [--yes]        turn off the reply it wrote (--force: whoever wrote it)\n" +
        "    calendar-out-of-office run [--dry-run]      what the scheduled task does: propose, adjust, or (act mode) write\n" +
        "    calendar-out-of-office autostart on         run every few minutes, at logon and after unlocking (off, restart, status)\n" +
        "    calendar-out-of-office sign-in              sign in to Microsoft, asking for the permissions it needs\n" +
        "    calendar-out-of-office config               create the settings file (if missing) and say where it is\n\n" +
        "options:\n" +
        "  -h, --help       show this help message and exit\n" +
        "  --version        show program's version number and exit\n" +
        "  --config CONFIG  settings file (default %APPDATA%\\calendar-out-of-office\\config.toml)\n";

    private static readonly string[] Commands = ["plan", "show", "apply", "clear", "run", "autostart", "sign-in", "config"];

    public static string Version =>
        typeof(Cli).Assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion ?? "0.0.0";

    internal sealed class Options
    {
        public string? Config;
        public string Command = "";
        public List<string> Args = [];
        public bool DryRun;
        public bool Yes;
        public bool Force;
    }

    private sealed class ArgumentsException(string message) : Exception(message);

    public static int Run(IReadOnlyList<string> argv, CliHost? host = null)
    {
        host ??= new CliHost();
        Options options;
        try
        {
            var parsed = Parse(argv, host);
            if (parsed is null)
                return 0;
            options = parsed;
        }
        catch (ArgumentsException error)
        {
            host.Err.WriteLine(Usage);
            host.Err.WriteLine($"{Prog}: error: {error.Message}");
            return 2;
        }

        Config? config = null;
        try
        {
            if (options.Command == "config")
            {
                host.Out.WriteLine(OutOfOffice.Config.CreateDefault(options.Config ?? host.Env(OutOfOffice.Config.EnvConfig), host.Env));
                return 0;
            }
            config = OutOfOffice.Config.Load(options.Config, host.Env);
            switch (options.Command)
            {
                case "sign-in":
                    host.SignIn(config);
                    host.Out.WriteLine("Signed in.");
                    return 0;
                case "autostart":
                    return Autostart(config, options.Args[0], host);
            }
            using var held = host.Lock ? Locks.Exclusive("calendar-out-of-office-" + Locks.Key(config.Path), 300) : null;
            var office = host.Office(config);
            var runner = new Runner(config, office, host);
            var state = State.Load(config.StateFile);
            var code = options.Command switch
            {
                "plan" => Plan(config, runner, office, state, host),
                "show" => Show(config, runner, office, state, host),
                "apply" => runner.Apply(state, options.Args[0], options.Yes),
                "clear" => runner.Clear(state, options.Yes, options.Force),
                _ => RunOnce(config, runner, state, host, options.DryRun),
            };
            if (options.Command == "run" && !options.DryRun)
                ForgetFailure(config);
            return code;
        }
        catch (Exception error) when (error is UsageException or SignInNeededException)
        {
            host.Err.WriteLine(error.Message);
            if (options.Command == "run" && config is not null && !options.DryRun)
                Failed(config, host, error);
            return 2;
        }
        catch (Exception error)
        {
            host.Err.WriteLine($"{error.GetType().Name}: {error.Message}");
            if (options.Command == "run" && config is not null && !options.DryRun)
                Failed(config, host, error);
            return 1;
        }
    }

    private static int RunOnce(Config config, Runner runner, JsonObject state, CliHost host, bool dryRun)
    {
        var said = runner.Run(state, dryRun);
        if (dryRun)
        {
            foreach (var line in said)
                host.Out.WriteLine(line);
            host.Out.WriteLine("(nothing changed)");
        }
        return 0;
    }

    private static int Plan(Config config, Runner runner, IOffice office, JsonObject state, CliHost host)
    {
        var pic = runner.Look(state);
        host.Out.WriteLine($"Mode: {config.Mode.ToString().ToLowerInvariant()}. Automatic replies now: {pic.Setting.Status} ({runner.SettingWindow(pic.Setting)})"
                           + pic.Owner switch { Owner.Ours => ", written by calendar-out-of-office.", Owner.Theirs => ", set outside calendar-out-of-office (left alone).", _ => "." });
        if (pic.Periods.Count == 0)
            host.Out.WriteLine($"No leave or conference in the next {config.DaysAhead} days that needs a reply.");
        foreach (var p in pic.Periods)
        {
            host.Out.WriteLine("");
            host.Out.WriteLine(runner.Preview(p, runner.Want(p, pic)));
            host.Out.WriteLine($"  days      {string.Join(", ", p.Days.Select(d => $"{d.Date:ddd d}: {Short(d.Kind)}" + (d.Kind == Kind.Conference ? $" ({d.Why})" : "")))}");
            if (runner.Refused(p, state) is { } why)
                host.Out.WriteLine($"  not written: {why}");
        }
        host.Out.WriteLine("");
        foreach (var line in runner.Run(state, dryRun: true))
            host.Out.WriteLine(line);
        host.Out.WriteLine(office.CanWrite()
            ? "Writing is possible (MailboxSettings.ReadWrite is in the sign-in)."
            : "A write would need MailboxSettings.ReadWrite, which the sign-in doesn't have yet (`calendar-out-of-office sign-in`; your organisation may need to approve it).");
        return 0;
    }

    private static string Short(Kind kind) => kind switch
    {
        Kind.Weekend => "weekend",
        Kind.BankHoliday => "bank holiday",
        Kind.HalfDay => "half day",
        Kind.Leave => "leave",
        Kind.Conference => "conference",
        _ => "working",
    };

    private static int Show(Config config, Runner runner, IOffice office, JsonObject state, CliHost host)
    {
        var pic = runner.Look(state);
        var s = pic.Setting;
        host.Out.WriteLine($"Automatic replies: {s.Status} ({runner.SettingWindow(s)})"
                           + pic.Owner switch { Owner.Ours => ", written by calendar-out-of-office", Owner.Theirs => ", set outside calendar-out-of-office", _ => "" });
        if (s.Status != "disabled" || s.Internal.Length > 0)
        {
            host.Out.WriteLine($"  external to   {s.Audience}");
            host.Out.WriteLine($"  internal      {Planner.Plain(s.Internal)}");
            host.Out.WriteLine($"  external      {Planner.Plain(s.External)}");
        }
        if (state["written"] is JsonObject w)
            host.Out.WriteLine($"Last written by calendar-out-of-office: {w["period"]} at {w["at"]} (UTC).");
        if (state["writing"] is JsonObject wr)
            host.Out.WriteLine($"A write for {wr["period"]} hasn't been confirmed yet; the next run reads it back.");
        return 0;
    }

    /// <summary>A run that failed: logged, and notified on the third in a row (at once for a sign-in), once until a run
    /// succeeds again.</summary>
    private static void Failed(Config config, CliHost host, Exception error)
    {
        try
        {
            Logger.Create(config.LogFile, null, clock: host.Now)($"run failed: {error.GetType().Name}: {error.Message}");
            var state = State.Load(config.StateFile);
            var failures = (state["failures"]?.GetValue<int>() ?? 0) + 1;
            state["failures"] = failures;
            if (state["failure_notified"] is null && (failures >= 3 || error is SignInNeededException))
            {
                if (config.Notify)
                    host.Notify("calendar-out-of-office couldn't run", [error.Message, $"Log: {config.LogFile}"]);
                state["failure_notified"] = State.Stamp(host.Now());
            }
            State.Save(config.StateFile, state);
        }
        catch (Exception)
        {
            // nothing more can be done: the exit code says it failed
        }
    }

    private static void ForgetFailure(Config config)
    {
        var state = State.Load(config.StateFile);
        if (state.Remove("failures") | state.Remove("failure_notified"))
            State.Save(config.StateFile, state);
    }

    private static int Autostart(Config config, string action, CliHost host)
    {
        var task = host.Task();
        switch (action)
        {
            case "on":
                var background = Path.Combine(host.ProgramDir, BackgroundExe);
                var exe = File.Exists(background) ? background : Path.Combine(host.ProgramDir, "calendar-out-of-office.exe");
                var schedule = new TaskSchedule
                {
                    AtLogon = true,
                    Every = TimeSpan.FromMinutes(config.EveryMinutes),
                    AfterUnlock = config.AfterUnlockMinutes is { } u ? TimeSpan.FromMinutes(u) : null,
                    KeepRunning = false,
                    TimeLimit = TimeSpan.FromMinutes(5),
                };
                var line = task.Enable(exe, ["--config", config.Path, "run"], Path.GetDirectoryName(config.Path), schedule);
                host.Out.WriteLine($"Autostart is on: '{TaskName}' runs every {config.EveryMinutes} minutes, at logon" +
                                   (config.AfterUnlockMinutes is { } m ? $" and {m} minute(s) after each unlock" : "") +
                                   $", and has run now.\n  {line}");
                return 0;
            case "off":
                host.Out.WriteLine(task.Disable() ? "Autostart is off." : "Autostart was already off.");
                return 0;
            case "restart":
                try
                {
                    host.Out.WriteLine(task.Restart() ? $"Started '{TaskName}' now." : "Autostart is off, so there's nothing to restart.");
                }
                catch (InvalidOperationException error)
                {
                    host.Err.WriteLine(error.Message);
                    return 1;
                }
                return 0;
            default:
                var info = task.Status();
                if (info is null)
                    host.Out.WriteLine("Autostart is off. `calendar-out-of-office autostart on` turns it on.");
                else
                {
                    host.Out.WriteLine($"Autostart is on ('{TaskName}').");
                    foreach (var name in new[] { "state", "last_run", "last_result", "next_run", "command" })
                        host.Out.WriteLine($"  {name.Replace('_', ' '),-12} {info.GetValueOrDefault(name, "")}");
                }
                return 0;
        }
    }

    private static Options? Parse(IReadOnlyList<string> argv, CliHost host)
    {
        var options = new Options();
        var i = 0;
        for (; i < argv.Count && argv[i].StartsWith('-'); i++)
        {
            var arg = argv[i];
            if (arg is "-h" or "--help")
            {
                host.Out.Write(Help);
                return null;
            }
            if (arg == "--version")
            {
                host.Out.WriteLine(Version);
                return null;
            }
            if (arg == "--config")
            {
                if (++i >= argv.Count)
                    throw new ArgumentsException("argument --config: expected one argument");
                options.Config = argv[i];
            }
            else if (arg.StartsWith("--config=", StringComparison.Ordinal))
                options.Config = arg["--config=".Length..];
            else
                throw new ArgumentsException($"unrecognized arguments: {arg}");
        }
        if (i >= argv.Count)
            throw new ArgumentsException("the following arguments are required: command");
        options.Command = argv[i++];
        if (!Commands.Contains(options.Command))
            throw new ArgumentsException($"argument command: invalid choice: '{options.Command}' (choose from {string.Join(", ", Commands.Select(c => $"'{c}'"))})");
        for (; i < argv.Count; i++)
        {
            var arg = argv[i];
            switch (options.Command, arg)
            {
                case (_, "-h" or "--help"):
                    host.Out.Write(Help);
                    return null;
                case ("run", "--dry-run"):
                    options.DryRun = true;
                    break;
                case ("apply" or "clear", "--yes"):
                    options.Yes = true;
                    break;
                case ("clear", "--force"):
                    options.Force = true;
                    break;
                default:
                    if (arg.StartsWith('-') && arg != "-")
                        throw new ArgumentsException($"unrecognized arguments: {arg}");
                    options.Args.Add(arg);
                    break;
            }
        }
        var max = options.Command is "autostart" or "apply" ? 1 : 0;
        if (options.Command == "autostart" && options.Args.Count == 0)
            throw new ArgumentsException("the following arguments are required: action");
        if (options.Command == "apply" && options.Args.Count == 0)
            throw new ArgumentsException("the following arguments are required: id (from `calendar-out-of-office plan`)");
        if (options.Args.Count > max)
            throw new ArgumentsException($"unrecognized arguments: {string.Join(" ", options.Args.Skip(max))}");
        if (options.Command == "autostart" && options.Args[0] is not ("on" or "off" or "restart" or "status"))
            throw new ArgumentsException($"argument action: invalid choice: '{options.Args[0]}' (choose from 'on', 'off', 'restart', 'status')");
        return options;
    }
}
