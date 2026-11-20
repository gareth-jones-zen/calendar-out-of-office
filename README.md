# out-of-office

Schedules your Outlook automatic replies from your calendar: when you're on leave for a few days or more, or at a
conference, the reply goes on the evening before and off the morning you're back, with the dates filled in for you.
Teams shows you as Out of office for the same window, with the same message, without anything else to set. It never
touches a reply you set yourself, and by default writes nothing until you've seen it and said yes.

```
out-of-office plan
```

```
Mode: ask. Automatic replies now: disabled (off).

leave-2027-03-10: leave from Wednesday 10 March, back Monday 15 March (3 weekday(s) away)
  window    Tue 9 Mar 17:30 to Mon 15 Mar 08:00
  external  to all
  internal  I'm on leave, back on Monday 15 March. For anything urgent, please contact Alex Example.
  external  I'm on leave and back on Monday 15 March; I'll reply to your message when I return.
  days      Wed 10: leave, Thu 11: leave, Fri 12: leave, Sat 13: weekend, Sun 14: weekend

leave-2027-03-10: to ask about: Tue 9 Mar 17:30 to Mon 15 Mar 08:00. `out-of-office apply leave-2027-03-10 --yes` schedules it.
Writing is possible (MailboxSettings.ReadWrite is in the sign-in).
```

## Install

Windows (x64) only; nothing else is needed.

1. Download `out-of-office-<version>-win-x64.zip` from [Releases](https://github.com/gareth-jones-zen/out-of-office/releases).
2. Unzip it to a folder that will stay put, such as `%LOCALAPPDATA%\Programs\out-of-office`.
3. `out-of-office config` creates the settings file; write your messages. `out-of-office sign-in` signs you in (once).
4. `out-of-office plan` shows each period and its messages; `out-of-office apply <id> --yes` schedules one;
   `out-of-office autostart on` keeps watching (every 15 minutes, at logon and a minute after each unlock), asking
   about each new period and following changes to ones it has written.

**Permissions** (Microsoft's "Graph Command Line Tools" app; `sign-in` asks for them): Calendars.ReadWrite and
MailboxSettings.Read to read; **MailboxSettings.ReadWrite to schedule a reply**, and Presence.ReadWrite only for the
optional Teams status message. Your organisation may need to approve MailboxSettings.ReadWrite; until then `plan`,
`show` and `run` in ask mode work, and `plan` says a write isn't possible yet.

## How a period is found

Each day from today is one of: a working day, a weekend (any day not in `working_days`), a bank holiday, leave, a half
day, or a conference.

- **Leave:** your own entries (no one else invited) whose subject matches `leave.subjects` (by default starting
  "Holiday", "Annual leave", "On leave" or "Summer holidays"), covering at least `leave.min_hours` (6) of the day. One
  matching `leave.half_day` (by default "(half day)"), or covering fewer hours, is a **half day**.
- **Bank holiday:** your own entries in a `bank_holiday.categories` category (default "Bank holiday") or whose subject
  starts "Bank holiday".
- **Conference:** an event whose subject matches `conference.subjects` (conference, summit, expo, exhibition, trade
  show, off-site, and any names you add), or one shown as *Working elsewhere*, covering at least `conference.min_hours`
  (6) or all day. `conference.ignore_subjects` leaves out what isn't being there: a reminder to book one ("Book
  XConf?"), another tool's own entries.
- The leave and conference rules are [background-tools-lib](https://github.com/gareth-jones-zen/background-tools-lib)'s
  `CalendarDays`, shared with [work-location](https://github.com/gareth-jones-zen/work-location), so the two agree.

**A period** is a run of days that aren't working days, holding leave, a half day or a conference: weekends and bank
holidays inside or at either end are part of it, so you're "back" on the first working day after. A bank holiday on
its own isn't a period (unless `bank_holiday.alone = true`). **A period needs `min_working_days` weekdays away**
(leave and conferences counted 1, joined bank holidays 1, half days 0.5): by default 3 for leave, 1 for a conference.
Leave and a conference together are one leave period, with the event's name available to the message.

**The window:** on at `window.on` (17:30) on the last working day before, off at `window.off` (08:00) on the day
you're back. A half day at the start puts it on at `window.half_day_at` (12:30) that day (afternoon off assumed); a
half day at the end takes it off at 12:30 that day (morning off assumed). The calendar can't say which half; the
preview shows the times.

**Scheduled, not switched on:** the reply is written once, as a schedule (`status: scheduled`, a start and an end), and
Exchange turns it on and off by itself, so the PC can be off for the whole of your leave. A mailbox has one automatic
reply, so **one period at a time**: the next is scheduled when the current one has finished.

## The messages

Two per kind of period, internal and external (`[leave]` and `[conference]`), plain text: paragraphs separated by a
blank line, lines within a paragraph joined. They're escaped and sent as HTML. **Placeholders:** `{return_day}`,
`{first_day}`, `{last_day}`, `{event}` (the conference's name, without a leading "At "), and any you define in
`[placeholders]` (who to contact, say). An unknown placeholder is an error when the settings are read. Dates follow
`date_format` ("dddd d MMMM": "Monday 15 March"). `external_audience` is `none`, `contactsOnly` or `all`.

## What it does on its own, and what it asks

| Situation | What happens |
|---|---|
| A new period, ask mode (the default) | **A question**: a review item with the window, both messages and the command; a notification if it starts within `notify_working_days` (5). Nothing is written until `apply <id> --yes` |
| A new period, act mode | **Written** `act_days_before` (7) days ahead, **if its messages have been approved** (applied once by hand). New or edited messages are asked about first |
| A period it wrote is longer, shorter or moved | **Rewritten, and you're told** (a review item). If you're already away, the start is kept and only the end moves |
| A period it wrote is no longer in the calendar | **Turned off, and you're told** |
| A reply you set yourself (in Outlook or Teams) is on | **Left alone**; raised once, and the next period waits. `clear --force --yes` turns it off |
| You change or turn off a reply it wrote | It's yours from then on: no more writes for that period |
| A write whose answer was lost | **Settled by reading back**, never by writing blind: recorded if it landed; one more try if the mailbox still holds what was there; otherwise raised |
| Implausible (more than `max_days` weekdays) or its end not known yet | Not written; `plan` says why |

**Notifications** only for something that needs you (a period to approve soon, your own reply in the way) or a failure
(the third in a row, or at once for a sign-in). Success is silent. Review items go to the queue named in `[review]`
(see [background-tools-lib](https://github.com/gareth-jones-zen/background-tools-lib)'s ReviewQueue).

**The state file** holds fingerprints and times of what was written, and which messages were approved: never the
messages.

## Commands

| Command | What it does |
|---|---|
| `plan` | Each period: why, the window, both messages, the days; then what a run would do, and whether a write is possible. Reads only |
| `show` | Your automatic replies now, and what out-of-office last wrote |
| `apply <id> [--yes]` | Schedule one period's reply. Without `--yes`, shows it and changes nothing |
| `clear [--yes] [--force]` | Turn off the reply it wrote; `--force` turns off one it didn't |
| `run [--dry-run]` | What the scheduled task does |
| `autostart on\|off\|restart\|status` | The scheduled task: every `every_minutes` (15), at logon, and after unlocking |
| `sign-in` | Sign in, asking for the permissions it needs |
| `config` | Create the settings file if it's missing, and say where it is |

Exit codes: 0 done; 1 failed; 2 settings, arguments or sign-in (including a write without the permission).

## Settings

`out-of-office config` writes a starting file (`%APPDATA%\out-of-office\config.toml`, or `--config`). Every key:

```toml
# base = "base-config.toml"     # settings shared with other tools (login_hint, tenant, msal_cache, time_zone, [paths])
time_zone = "GMT Standard Time"
log_file = "%LOCALAPPDATA%/out-of-office/out-of-office.log"
state_file = "%LOCALAPPDATA%/out-of-office/state.json"
enabled = true
mode = "ask"                    # ask | act
days_ahead = 42                 # how far ahead it looks
act_days_before = 7             # act mode: how early a period is written
working_days = ["mon", "tue", "wed", "thu", "fri"]
date_format = "dddd d MMMM"
max_days = 30                   # weekdays away; longer is refused as implausible
notify = true
notify_working_days = 5         # ask mode: a notification when a period to approve starts this soon
every_minutes = 15
after_unlock_minutes = 1

[window]
on = "17:30"
off = "08:00"
half_day_at = "12:30"

[leave]
min_working_days = 3
subjects = '^\s*(holiday\b|annual leave|on leave|summer holidays?)'
categories = []
half_day = '\(half day\)'
min_hours = 6
external_audience = "all"
internal = '''
I'm on leave, back on {return_day}. For anything urgent, please contact {cover}.
'''
external = '''
I'm on leave and back on {return_day}; I'll reply to your message when I return.
'''

[bank_holiday]
subjects = '^\s*bank holiday\b'
categories = ["Bank holiday"]
alone = false

[conference]
min_working_days = 1
subjects = '\b(conference|summit|expo|exhibition|trade show|off-?site)\b'
show_as_signals = ["workingElsewhere"]
min_hours = 6
ignore_subjects = '^\s*(book|register)\b|\?\s*$'
action = "reply"                # reply | status (a Teams status message only) | both
external_audience = "contactsOnly"
internal = "I'm at {event} until {last_day}; back on {return_day}."
external = "I'm at {event} until {last_day}; I'll reply when I can."

[placeholders]
cover = "Alex Example"

[teams]
status_message = false          # with conference.action "status" or "both": a Teams status message until the reply ends
conference_status = "At {event} until {last_day}"

# [review]
# file = "%LOCALAPPDATA%/review/pending.json"
# source = "out-of-office"
```

**Teams:** while a scheduled reply is on, Teams shows you as Out of office with the internal message; nothing more is
needed for leave. Graph can't set Out of office presence itself, and out-of-office never sets a preferred presence.
The optional status message is for a conference you'd rather show as a note than as Out of office.

## Not verified against a real mailbox yet

The write path is tested against a fake mailbox. Not yet confirmed with Exchange: how it rewrites the messages' HTML
(compared as plain text, so it shouldn't matter), whether a schedule already running accepts a new end with the start
left out, and whether `setStatusMessage` accepts `/me` (it uses `/users/{id}`).

## Building

.NET 11 (see `global.json`), with the shared libraries from
[background-tools-lib](https://github.com/gareth-jones-zen/background-tools-lib) checked out beside this repo:

```
dotnet test -c Release
dotnet publish src/OutOfOffice.Console -c Release -r win-x64 -o out    # Native AOT
```

A tag `vX.Y.Z` builds a release (`.github/workflows/release.yml`), against the library commit pinned there.
