# Azure Cognitive Services based Email autoreply bot

Homework assignment for ScamLab exercise 5 - automatic reply to scam (ITC8210 Human Aspects of Cyber Security, [TalTech](https://www.taltech.ee/)).

## Premise

The entire program consists of 2 main parts - [Text Analysis](Autoreply.Keywords), and [Sending and receiving e-mails](Autoreply.ImapClient).

#### Text Analysis

We are given an arbitrary string and need to extract the keywords. For that [Microsoft Azure Cognitive Services Text Analytics API](https://azure.microsoft.com/en-us/services/cognitive-services/text-analytics/) is used. They offer 5,000 free analytics per month, which is more than enough for the purposes of this assignment. This service takes English text of 1,000 symbols or less and provides it's keywords. One of the keywords will be used to form a question in the reply email (example: "I was wondering if you could elaborate on ABC", where "ABC" was taken from the initial email).

#### Sending and receiving e-mails

This part consists of 2 clients - IMAP, and SMTP. IMAP client is used to receive emails and SMTP is used to send the replies. The clients need to be set up from `appsettings.json` file (example below).

Basic algorithm flowchart:

1. Sleep for 5 minutes.
2. Check if new emails arrived within those last 5 minutes.
3. If new e-mail(s) arrived:
   1. Run keyword text analysis on the received e-mail(s).
   2. Pick a random selection of predetermined template messages (greeting, question, signature).
   3. Schedule a response e-mail to be sent in the next 5 minutes to 2 hours (instantaneous replies appear suspicious).
4. Repeat from (1).

Rest of the code is used to sanitize and cleanup dirty email code, such as `&nbsp;` characters.

## Setup

#### Prerequisites

[Required] [.Net Core 3.1 Runtime](https://dotnet.microsoft.com/download/dotnet-core/3.1)

#### Installation

Download latest release from releases tab.

Edit `appsettings.json`:

```js
{
  "CognitiveServices": {
    "APIKey": "",				// From Azure Cognitive services resource
    "Endpoint": ""				// From Azure Cognitive services resource
  },
  "EmailOptions": {
    "Email": "",				// Email address to use for the IMTP/SMTP clients
    "Name": "",					// Name to send emails as (usually your full name)
    "Password": "",				// Password to use for the IMTP/CMTP clients
    "Host": "imap.gmail.com",	// Host. Change if not using gmail.
    "ImapPort": "993",			// IMAP port. Default should work in most cases.
    "SmtpPort": "465"			// SMTP port. Default should work in most cases.
  }
}
```

#### Launching

On Windows: Run `Autoreply.exe`.

On Limux/MacOS: Run CLI `dotnet Autoreply.dll`.
