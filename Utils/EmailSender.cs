
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Microsoft.Office.Interop.Excel;
using System.Threading.Tasks;
using System.IO;
using Anupama_project.Models;
using SendGrid;
using SendGrid.Helpers.Mail;
using System;
using System.Threading.Tasks;

namespace Anupama_project.Utils
{
    

    public class EmailSender
    {
        public static async Task<bool> SendEmail(string to, string subject, string body, byte[] attachmentBytes, string attachmentFileName)
        {
            try
            {
                var apiKey = ;
        var client = new SendGridClient(apiKey);

                var from = new EmailAddress("anupamashipurkar@gmail.com", "Anupama");
                var toEmail = new EmailAddress(to);
                var plainTextContent = body;
                var htmlContent = body;

                var msg = MailHelper.CreateSingleEmail(
                    from, toEmail, subject, plainTextContent, htmlContent);

                if (attachmentBytes != null)
                {
                    var attachment = new Attachment
                    {
                        Content = Convert.ToBase64String(attachmentBytes),
                        Filename = attachmentFileName,
                    };
                    msg.AddAttachment(attachment);
                }

                var response = await client.SendEmailAsync(msg);
                return response.StatusCode == System.Net.HttpStatusCode.Accepted;
            }
            catch (Exception ex)
            {
                
                return false;
            }
        }
    }

}
