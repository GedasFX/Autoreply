using System;
using System.Collections.Generic;
using System.Text;

namespace Autoreply.Imap
{
    public class ImapOptions
    {
        public string Host { get; set; }
        public int ImapPort { get; set; }
        public int SmtpPort { get; set; }
        public string Name { get; set; }
        public string Email { get; set; }
        public string Password { get; set; }
    }
}
