using System;
using System.Globalization;
using Autoreply.Imap;
using Autoreply.TextAnalysis;
using Microsoft.Extensions.Configuration;

namespace Autoreply
{
    class Program
    {
        static void Main()
        {
            var config = new ConfigurationBuilder().AddJsonFile("appsettings.json", optional: false, reloadOnChange: false).Build();
            var emailCfg = config.GetSection("EmailOptions");

            Console.WriteLine($"[{DateTime.Now:g}] Configuration Loaded successfully.");

            new ImapResponder(new TextAnalysisClient(config["CognitiveServices:APIKey"], new Uri(config["CognitiveServices:Endpoint"])), new ImapOptions
            {
                Email = emailCfg["Email"],
                Name = emailCfg["Name"],
                Password = emailCfg["Password"],
                Host = emailCfg["Host"],
                ImapPort = int.Parse(emailCfg["ImapPort"], CultureInfo.InvariantCulture),
                SmtpPort = int.Parse(emailCfg["SmtpPort"], CultureInfo.InvariantCulture)
            }).RunImapClientAsync().Wait();
        }
    }
}
