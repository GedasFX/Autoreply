using Autoreply.TextAnalysis;
using HtmlAgilityPack;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Net.Smtp;
using MimeKit;
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace Autoreply.Imap
{
    public class ImapResponder
    {
        private readonly TextAnalysisClient _textAnalysisClient;

        private readonly string[] _greetings = new[] {
            "Hello,",
            "Hi,",
            "Greetings,",
            "Good day,"
        };
        private readonly string[] _questions = new[] {
            "I was wondering if you could elaborate on {0}?",
            "I am not quite sure what you mean by {0}. Could you try to elaborate on that?",
            "Why should I help you considering {0} is involved?",
        };
        private readonly string[] _signatures = new[] {
            "Regards,",
            "Best wishes,",
            "Respectfully,"
        };

        private CancellationTokenSource _idle;

        public ImapResponder(TextAnalysisClient textAnalysis)
        {
            _textAnalysisClient = textAnalysis;
        }

        public async Task RunImapClientAsync([NotNull] ImapOptions options)
        {
            using var imap = new ImapClient();
            using var smtp = new SmtpClient();
            using var cancel = new CancellationTokenSource();

            await imap.ConnectAsync(options.Host, options.ImapPort, true, cancel.Token).ConfigureAwait(false);
            await imap.AuthenticateAsync(options.Email, options.Password, cancel.Token).ConfigureAwait(false);

            await smtp.ConnectAsync(options.Host, options.SmtpPort, true, cancel.Token).ConfigureAwait(false);
            await smtp.AuthenticateAsync(options.Email, options.Password).ConfigureAwait(false);

            await imap.Inbox.OpenAsync(FolderAccess.ReadOnly).ConfigureAwait(false);

            var initialCount = imap.Inbox.Count;

            var emailArrived = false;
            imap.Inbox.CountChanged += (sender, e) =>
            {
                Console.WriteLine($"[{DateTime.Now:g}] New email detected.");
                emailArrived = true;
            };

            Console.WriteLine($"[{DateTime.Now:g}] Client ready. Waiting for emails.");

            while (true)
            {
                _idle = new CancellationTokenSource();
                var idl = imap.IdleAsync(_idle.Token);

                var random = new Random();
                while (!emailArrived)
                {
                    Console.WriteLine($"[{DateTime.Now:g}] No new emails detected. Sleeping for 5 minutes.");
                    await Task.Delay(300000).ConfigureAwait(false);
                }

                // Email arrived, handle it.

                _idle.Cancel();
                idl.Wait();

                if (imap.Inbox.Count > initialCount)
                {
                    Console.WriteLine($"[{DateTime.Now:g}] Generating responses for the received emails.");
                    await HandleEmailsAsync(imap, smtp, initialCount, options).ConfigureAwait(false);
                }

                initialCount = imap.Inbox.Count;
                emailArrived = false;
            }
        }

        private async Task HandleEmailsAsync(ImapClient imap, SmtpClient smtp, int prevEnd, ImapOptions options)
        {
            using var cancel = new CancellationTokenSource();

            var newEmailSummaries = await imap.Inbox.FetchAsync(prevEnd, -1, MessageSummaryItems.Full | MessageSummaryItems.UniqueId).ConfigureAwait(false);
            Console.WriteLine($"[{DateTime.Now:g}] Detected {newEmailSummaries.Count} new e-mails. Generating responses.");

            foreach (var summary in newEmailSummaries)
            {
                var message = await imap.Inbox.GetMessageAsync(summary.Index, cancel.Token).ConfigureAwait(false);
                var text = PrepareText(message.TextBody) ?? PrepareText(PrepareHtml(message.HtmlBody)) ?? "";

                if (!string.IsNullOrWhiteSpace(text))
                {
                    var keyPhrases = await _textAnalysisClient.ExtractKeyPhrasesAsync(text, cancel.Token).ConfigureAwait(false);
                    var reply = PrepareReply(message, options.Email, options.Name, GenerateReplyMessage(keyPhrases));

                    var delay = new Random().Next(300000, 3600000);
                    Console.WriteLine($"[{DateTime.Now:g}] Response generated for email '{message.Subject}'. Response is scheduled to be sent at {DateTime.Now.AddMilliseconds(delay):G}.");

                    // Send the reply.
                    _ = Task.Run(async () =>
                    {
                        await Task.Delay(delay).ConfigureAwait(false);

                        Console.WriteLine($"[{DateTime.Now:g}] Sending reply for '{message.Subject}' to '{reply.To[0].Name}'.");
                        try
                        {
                            await smtp.SendAsync(reply).ConfigureAwait(false);
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e.ToString());
                            Console.WriteLine(e.StackTrace);
                            throw e;
                        }
                    });
                }
            }
        }

        private static string PrepareText(string text)
        {
            if (text == null)
            {
                return null;
            }

            if (text.Length > 1000)
            {
                // This preparation is required for the Azure text analysis API.
                text = text.Substring(0, 1000);
            }

            return text;
        }

        private static string PrepareHtml(string html)
        {
            if (html == null)
            {
                return null;
            }

            var doc = new HtmlDocument();
            doc.LoadHtml(html);

            var dirtyText = doc.DocumentNode
                .SelectSingleNode("//body")
                .InnerText;

            return Regex.Replace(dirtyText, "&.{1,10};", " ");
        }

        private static MimeMessage PrepareReply(MimeMessage message, string senderEmail, string senderName, string replyText)
        {
            var reply = new MimeMessage();

            // https://stackoverflow.com/a/35106633
            if (message.ReplyTo.Count > 0)
            {
                reply.To.AddRange(message.ReplyTo);
            }
            else if (message.From.Count > 0)
            {
                reply.To.AddRange(message.From);
            }
            else if (message.Sender != null)
            {
                reply.To.Add(message.Sender);
            }

            // Add Re:, only when it it absolutely necessarry.
            reply.Subject = message.Subject.StartsWith("Re:", StringComparison.OrdinalIgnoreCase) ? message.Subject : $"Re: {message.Subject}";

            // construct the In-Reply-To and References headers
            if (!string.IsNullOrEmpty(message.MessageId))
            {
                reply.InReplyTo = message.MessageId;

                reply.References.AddRange(message.References);
                reply.References.Add(message.MessageId);
            }

            reply.Sender = new MailboxAddress(senderName, senderEmail);
            reply.Body = new TextPart("plain")
            {
                Text = replyText
            };

            return reply;
        }

        private string GenerateReplyMessage(IList<string> keywords)
        {
            var random = new Random();

            var message = new StringBuilder();
            message.AppendLine(_greetings[random.Next(_greetings.Length)]);
            message.AppendLine();
            message.AppendLine(string.Format(CultureInfo.InvariantCulture, _questions[random.Next(_questions.Length)], keywords[0]));
            message.AppendLine();
            message.AppendLine(_signatures[random.Next(_signatures.Length)]);
            message.AppendLine("Dainius");

            return message.ToString();
        }
    }
}
