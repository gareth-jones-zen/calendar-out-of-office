using System.Net;
using System.Text;
using System.Text.Json.Nodes;
using M365Graph;

namespace OutOfOffice.Tests;

/// <summary>A write that may have reached Microsoft is never resent (background-tools-lib 62f653f): the automatic-reply
/// PATCH is sent once, and a 5xx or a timeout after sending is settled by reading the setting back, never by sending
/// it again. Reads, and the Teams status message (repeatable, notifies nobody), are still retried. Through the real
/// GraphOffice over a fake mailbox.</summary>
public sealed class SendOnceTests : IDisposable
{
    private readonly World w = new();
    private readonly Mailbox mailbox = new();
    private StringWriter output = new(), errors = new();

    public SendOnceTests()
    {
        mailbox.Events.Add((JsonNode)new JsonObject
        {
            ["subject"] = "Holiday (Example Role)", ["isAllDay"] = false, ["showAs"] = "busy", ["categories"] = new JsonArray(),
            ["isOrganizer"] = true, ["attendees"] = new JsonArray(),
            ["start"] = new JsonObject { ["dateTime"] = "2027-03-10T00:00:00.0000000" },
            ["end"] = new JsonObject { ["dateTime"] = "2027-03-12T23:59:00.0000000" },
        });
    }

    public void Dispose() => w.Dispose();

    /// <summary>The mailbox's automatic-reply setting and calendar as Graph serves them, and every request made.</summary>
    private sealed class Mailbox : HttpMessageHandler
    {
        public JsonObject Setting { get; set; } = new()
        {
            ["status"] = "disabled", ["externalAudience"] = "none", ["internalReplyMessage"] = "", ["externalReplyMessage"] = "",
        };
        public JsonArray Events { get; } = [];
        public List<string> Requests { get; } = [];

        /// <summary>Whether a PATCH changes the setting (before its answer goes wrong, if it does).</summary>
        public bool Applies { get; set; } = true;
        public HttpStatusCode? WriteStatus { get; set; }
        public Exception? WriteThrows { get; set; }

        /// <summary>What the mailbox holds after a PATCH instead (someone else's change in between).</summary>
        public JsonObject? Instead { get; set; }

        public JsonObject? LastPatch { get; private set; }
        public int ReadFailures { get; set; }
        public bool ReadsFailAfterWrite { get; set; }
        public int StatusFailures { get; set; }

        public void Land()
        {
            foreach (var (key, value) in LastPatch!)
                Setting[key] = value?.DeepClone();
        }

        public int Count(string request) => Requests.Count(r => r == request);

        protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancel)
        {
            var path = request.RequestUri!.AbsolutePath;
            Requests.Add($"{request.Method} {path}");
            if (request.Method == HttpMethod.Get && path.EndsWith("/me/calendarView", StringComparison.Ordinal))
                return Json(new JsonObject { ["value"] = Events.DeepClone() });
            if (request.Method == HttpMethod.Get && path.EndsWith("/automaticRepliesSetting", StringComparison.Ordinal))
            {
                if (ReadFailures > 0 || (ReadsFailAfterWrite && LastPatch is not null))
                {
                    ReadFailures--;
                    return new HttpResponseMessage(HttpStatusCode.ServiceUnavailable) { Content = new StringContent("{}") };
                }
                return Json((JsonObject)Setting.DeepClone());
            }
            if (request.Method == HttpMethod.Get && path.EndsWith("/me", StringComparison.Ordinal))
                return Json(new JsonObject { ["id"] = "me-id" });
            if (request.Method == HttpMethod.Patch && path.EndsWith("/me/mailboxSettings", StringComparison.Ordinal))
            {
                var body = JsonNode.Parse(await request.Content!.ReadAsStringAsync(cancel))!["automaticRepliesSetting"]!.AsObject();
                LastPatch = (JsonObject)body.DeepClone();
                if (Instead is not null)
                    Setting = (JsonObject)Instead.DeepClone();
                else if (Applies)
                    Land();
                if (WriteThrows is { } error)
                    throw error;
                return WriteStatus is { } status ? new HttpResponseMessage(status) { Content = new StringContent("{\"error\":{\"code\":\"busy\"}}") } : Json([]);
            }
            if (request.Method == HttpMethod.Post && path.EndsWith("/presence/setStatusMessage", StringComparison.Ordinal))
                return StatusFailures-- > 0 ? new HttpResponseMessage(HttpStatusCode.BadGateway) { Content = new StringContent("{}") }
                    : new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent("") };
            return new HttpResponseMessage(HttpStatusCode.NotFound) { Content = new StringContent("{}") };
        }

        private static HttpResponseMessage Json(JsonObject body) =>
            new(HttpStatusCode.OK) { Content = new StringContent(body.ToJsonString(), Encoding.UTF8, "application/json") };
    }

    private GraphClient Client() =>
        new(new GraphClientOptions { Log = _ => { } }, mailbox, token: (_, _) => Task.FromResult("token"), delay: (_, _) => Task.CompletedTask);

    private int Cli(params string[] args)
    {
        output = new StringWriter();
        errors = new StringWriter();
        return OutOfOffice.Cli.Run(["--config", w.ConfigPath, .. args], new CliHost
        {
            Out = output,
            Err = errors,
            Now = () => w.Now,
            Office = config => new GraphOffice(Client(), config.TimeZone),
            Notify = (title, lines) => w.Notified.Add(title + ": " + string.Join(" | ", lines)),
            Lock = false,
            Env = name => name == "LOCALAPPDATA" ? w.Dir : null,
        });
    }

    private const string Patch = "PATCH /v1.0/me/mailboxSettings";

    [Theory]
    [InlineData(HttpStatusCode.InternalServerError)]
    [InlineData(HttpStatusCode.ServiceUnavailable)]
    [InlineData(HttpStatusCode.GatewayTimeout)]
    public void A_5xx_to_the_reply_is_sent_once_and_unknown(HttpStatusCode status)
    {
        mailbox.WriteStatus = status;
        var office = new GraphOffice(Client(), w.Config.TimeZone);
        var reply = new Reply(new DateTime(2027, 3, 9, 17, 30, 0, DateTimeKind.Utc), new DateTime(2027, 3, 15, 8, 0, 0, DateTimeKind.Utc), "<p>a</p>", "<p>b</p>", "all");
        var error = Assert.Throws<GraphDeliveryUnknownException>(() => office.Schedule(reply, false));
        Assert.Equal(status, error.Status);
        Assert.Equal(1, mailbox.Count(Patch));
        Assert.Throws<GraphDeliveryUnknownException>(office.Disable);
        Assert.Equal(2, mailbox.Count(Patch));
    }

    [Fact]
    public void A_5xx_to_a_reply_that_landed_is_recorded_and_not_resent()
    {
        mailbox.WriteStatus = HttpStatusCode.ServiceUnavailable;
        Assert.Equal(0, Cli("apply", "leave-2027-03-10", "--yes"));
        Assert.Equal(1, mailbox.Count(Patch));
        Assert.Equal("scheduled", mailbox.Setting["status"]!.ToString());
        Assert.Equal("leave-2027-03-10", w.State["written"]!["period"]!.ToString());
        Assert.Null(w.State["writing"]);
    }

    [Fact]
    public void A_5xx_to_a_reply_not_there_yet_is_unknown_and_the_next_run_reads_it_back_before_anything_is_resent()
    {
        mailbox.WriteStatus = HttpStatusCode.BadGateway;
        mailbox.Applies = false;
        Assert.Equal(1, Cli("apply", "leave-2027-03-10", "--yes"));
        Assert.Equal(1, mailbox.Count(Patch));
        Assert.Contains("Microsoft didn't confirm the write for leave-2027-03-10", errors.ToString());
        Assert.Contains("may be made once more only if it still isn't there", errors.ToString());
        Assert.Equal(1, w.State["writing"]!["attempts"]!.GetValue<int>());

        mailbox.Land();   // it landed after all, late
        Assert.Equal(0, Cli("run"));
        Assert.Equal(1, mailbox.Count(Patch));   // recorded from the read, never sent again
        Assert.Null(w.State["writing"]);
        Assert.Equal("leave-2027-03-10", w.State["written"]!["period"]!.ToString());
    }

    [Fact]
    public void A_reply_timing_out_after_sending_is_sent_once_and_one_that_cant_be_confirmed_is_raised()
    {
        mailbox.WriteThrows = new TaskCanceledException("timed out", new TimeoutException());
        mailbox.Instead = new JsonObject
        {
            ["status"] = "alwaysEnabled", ["externalAudience"] = "none", ["internalReplyMessage"] = "<p>someone else</p>", ["externalReplyMessage"] = "",
        };
        Assert.Equal(1, Cli("apply", "leave-2027-03-10", "--yes"));
        Assert.Equal(1, mailbox.Count(Patch));
        Assert.Contains("holds something else", errors.ToString());
        Assert.Contains("TaskCanceledException: timed out", errors.ToString());
        Assert.Null(w.State["writing"]);
        Assert.Null(w.State["written"]);
        var failed = Assert.Single(w.Items, i => i["type"]?.ToString() == "failed");
        Assert.Equal("Out of office: a write for leave-2027-03-10 can't be confirmed", failed["title"]!.ToString());
    }

    [Fact]
    public void A_reply_timing_out_that_landed_is_recorded_and_not_resent()
    {
        mailbox.WriteThrows = new TaskCanceledException("timed out", new TimeoutException());
        Assert.Equal(0, Cli("apply", "leave-2027-03-10", "--yes"));
        Assert.Equal(1, mailbox.Count(Patch));
        Assert.NotNull(w.State["written"]);
    }

    [Fact]
    public void An_unknown_write_whose_read_back_fails_too_is_left_for_the_next_run_to_read()
    {
        mailbox.WriteStatus = HttpStatusCode.ServiceUnavailable;
        mailbox.ReadsFailAfterWrite = true;
        Assert.Equal(1, Cli("apply", "leave-2027-03-10", "--yes"));
        Assert.Equal(1, mailbox.Count(Patch));
        Assert.Contains("may or may not have landed", errors.ToString());
        Assert.Contains("the next run reads it back before writing anything", errors.ToString());
        Assert.NotNull(w.State["writing"]);
    }

    [Fact]
    public void A_read_still_retries_a_5xx()
    {
        mailbox.ReadFailures = 1;
        Assert.Equal(0, Cli("show"));
        Assert.Equal(2, mailbox.Count("GET /v1.0/me/mailboxSettings/automaticRepliesSetting"));
        Assert.Contains("Automatic replies: disabled", output.ToString());
    }

    [Fact]
    public void The_teams_status_message_is_repeatable_so_still_retried()
    {
        mailbox.StatusFailures = 1;
        new GraphOffice(Client(), w.Config.TimeZone).StatusMessage("At Example Expo", new DateTime(2027, 3, 12, 17, 0, 0, DateTimeKind.Utc));
        Assert.Equal(2, mailbox.Count("POST /v1.0/users/me-id/presence/setStatusMessage"));
    }
}
