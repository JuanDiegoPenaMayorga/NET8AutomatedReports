using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Net;
using System.IO;
using MimeKit;
using MailKit.Net.Smtp;
using MailKit.Security;
using System;
using System.IO;

namespace NET8AutomatedReports
{
    public class SmtpMailer
    {
        private readonly Config _config;

        public SmtpMailer(Config config)
        {
            _config = config;
        }

        public void SendReportEmail(string subject, string bodyHtmlPath, string attachmentPath)
        {
            if (!File.Exists(bodyHtmlPath))
                throw new FileNotFoundException("No se encontró el archivo de cuerpo HTML del correo.", bodyHtmlPath);

            if (!File.Exists(attachmentPath))
                throw new FileNotFoundException("No se encontró el archivo adjunto.", attachmentPath);

            string body = File.ReadAllText(bodyHtmlPath);
            var message = new MimeMessage();
            message.From.Add(MailboxAddress.Parse(_config.Email.From));

            foreach (var to in _config.Email.To.Split(';', StringSplitOptions.RemoveEmptyEntries))
                message.To.Add(MailboxAddress.Parse(to.Trim()));

            if (!string.IsNullOrWhiteSpace(_config.Email.Cc))
            {
                foreach (var cc in _config.Email.Cc.Split(';', StringSplitOptions.RemoveEmptyEntries))
                    message.Cc.Add(MailboxAddress.Parse(cc.Trim()));
            }

            if (!string.IsNullOrWhiteSpace(_config.Email.Bcc))
            {
                foreach (var bcc in _config.Email.Bcc.Split(';', StringSplitOptions.RemoveEmptyEntries))
                    message.Bcc.Add(MailboxAddress.Parse(bcc.Trim()));
            }

            message.Subject = subject;

            var builder = new BodyBuilder
            {
                HtmlBody = File.ReadAllText(bodyHtmlPath, Encoding.UTF8)
            };

            string imagePath = Path.Combine(AppContext.BaseDirectory, "Sign", "PaysettFirma.png");

            if (!File.Exists(imagePath))
                throw new FileNotFoundException("No se encontró la imagen de firma.", imagePath);

            var image = builder.LinkedResources.Add(imagePath) as MimePart;

            if (image != null)
            {
                image.ContentId = "icono";
                image.ContentDisposition = new ContentDisposition(ContentDisposition.Inline);
                image.ContentTransferEncoding = ContentEncoding.Base64;
            }

            builder.Attachments.Add(attachmentPath);
            message.Body = builder.ToMessageBody();

            using var client = new SmtpClient();
            var socketOption = ParseEncryptionMode(_config.SMTPEncryptionMode);

            client.Connect(_config.SMTPServer, _config.SMTPPort, socketOption);
            client.Authenticate(_config.SMTPUsername, _config.SMTPPassword);
            client.Send(message);
            client.Disconnect(true);
        }


        private SecureSocketOptions ParseEncryptionMode(string mode)
        {
            return mode?.ToLower() switch
            {
                "ssl" => SecureSocketOptions.SslOnConnect,
                "starttls" => SecureSocketOptions.StartTls,
                "auto" => SecureSocketOptions.Auto,
                "none" => SecureSocketOptions.None,
                _ => SecureSocketOptions.None
            };
        }
    }
}
