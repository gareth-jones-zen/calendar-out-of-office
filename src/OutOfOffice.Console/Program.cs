// out-of-office.exe. Output is Unicode to a console and UTF-8 when redirected (see ConsoleText).
return OutOfOffice.Cli.Run(args, new OutOfOffice.CliHost { Out = WinBackground.ConsoleText.Out, Err = WinBackground.ConsoleText.Error });
