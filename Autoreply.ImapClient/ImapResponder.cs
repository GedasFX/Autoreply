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

        private readonly string[] _greetings = new[] { "Hello,", "Hi,", "Greetings,", "Good day," };
        private readonly string[] _questions = new[] { "I was wondering if you could elaborate on {0}?", "I am not quite sure what you mean by {0}. Could you try to elaborate on that?", "Why should I help you considering {0} is involved." };
        private readonly string[] _signatures = new[] { "Regards,", "Best wishes,", "Respectfully," };

        private CancellationTokenSource _idle;

        public ImapResponder(TextAnalysisClient textAnalysis)
        {
            _textAnalysisClient = textAnalysis;
        }

        public async Task RunImapClientAsync([NotNull] ImapOptions options)
        {
            using var client = new ImapClient();
            using var cancel = new CancellationTokenSource();

            await client.ConnectAsync(options.Host, options.ImapPort, true, cancel.Token).ConfigureAwait(false);
            await client.AuthenticateAsync(options.Email, options.Password, cancel.Token).ConfigureAwait(false);

            await client.Inbox.OpenAsync(FolderAccess.ReadOnly).ConfigureAwait(false);

            var initialCount = client.Inbox.Count;

            var emailArrived = false;
            client.Inbox.CountChanged += (sender, e) =>
            {
                emailArrived = true;
            };

            Console.WriteLine("Client ready. Waiting for emails");

            while (true)
            {
                _idle = new CancellationTokenSource();
                var idl = client.IdleAsync(_idle.Token);

                var random = new Random();
                while (!emailArrived)
                {

                    // Replies should not be instantanious. This check will be performed between 5 minutes and 2 hours.
                    var delay = random.Next(300000, 7200000);

                    Console.WriteLine("No new emails detected. Sleeping.");
                    await Task.Delay(delay).ConfigureAwait(false);
                }

                // Email arrived, handle it.
                Console.WriteLine("Email detected. Generating response");

                _idle.Cancel();
                idl.Wait();

                if (client.Inbox.Count > initialCount)
                {
                    await HandleEmailsAsync(client, initialCount, options).ConfigureAwait(false);
                }

                initialCount = client.Inbox.Count;
                emailArrived = false;
            }
        }

        private async Task HandleEmailsAsync(ImapClient client, int prevEnd, ImapOptions options)
        {
            using var cancel = new CancellationTokenSource();

            foreach (var summary in await client.Inbox.FetchAsync(prevEnd, -1, MessageSummaryItems.Full | MessageSummaryItems.UniqueId).ConfigureAwait(false))
            {
                var message = await client.Inbox.GetMessageAsync(summary.Index, cancel.Token).ConfigureAwait(false);
                var text = PrepareText(message.TextBody) ?? PrepareText(PrepareHtml(message.HtmlBody)) ?? "";

                var preparedText = text;
                if (!string.IsNullOrWhiteSpace(preparedText))
                {
                    var keyPhrases = await _textAnalysisClient.ExtractKeyPhrasesAsync(text, cancel.Token).ConfigureAwait(false);
                    var reply = PrepareReply(message, options.Email, options.Name, GenerateReplyMessage(keyPhrases));

                    // Send the reply.
                    using var smtp = new SmtpClient();
                    await smtp.ConnectAsync(options.Host, options.SmtpPort, true, cancel.Token).ConfigureAwait(false);
                    await smtp.AuthenticateAsync(options.Email, options.Password).ConfigureAwait(false);
                    await smtp.SendAsync(reply).ConfigureAwait(false);
                    await smtp.DisconnectAsync(true).ConfigureAwait(false);
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

            var replyMsgWithQuotes = new StringBuilder();

            replyMsgWithQuotes.AppendLine(replyText);
            //replyMsgWithQuotes.AppendLine();

            //var sender = message.Sender ?? message.From.Mailboxes.FirstOrDefault();

            //replyMsgWithQuotes.AppendLine($"On {message.Date:f}, {(!string.IsNullOrEmpty(sender.Name) ? sender.Name : sender.Address)} wrote:");
            //using (var reader = new StringReader(message.TextBody))
            //{
            //    string line;
            //    while ((line = reader.ReadLine()) != null)
            //    {
            //        replyMsgWithQuotes.AppendLine($"> {line}");
            //    }
            //}


            reply.Body = new TextPart("plain")
            {
                Text = replyMsgWithQuotes.ToString()
            };

            return reply;
        }

        private string GenerateReplyMessage(IList<string> keywords)
        {
            var random = new Random();

            var message = new StringBuilder();
            message.AppendLine(_greetings[random.Next(_greetings.Length)]);
            message.AppendLine();
            message.AppendLine(string.Format(CultureInfo.InvariantCulture, _questions[random.Next(_questions.Length)], keywords[random.Next(keywords.Count)]));
            message.AppendLine();
            message.AppendLine(_signatures[random.Next(_signatures.Length)]);
            message.AppendLine("Dainius");

            return message.ToString();
        }
    }
}
