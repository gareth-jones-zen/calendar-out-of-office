// out-of-office-background.exe: no window, ever. What the scheduled task runs (`run`); its log is the settings' log_file.
return OutOfOffice.Cli.Run(args, new OutOfOffice.CliHost { Out = TextWriter.Null, Err = TextWriter.Null });
