using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Activities;
using System.ComponentModel;
using System.Net.Mail;

namespace Enhanced_SMTP
{
   
    public class Enhanced_Send_SMTP_Mail:CodeActivity
    {
        
        [Category("Receiver")]
        [RequiredArgument]
        [Description("Email Addresses seprated by ; e.g.: \"xyz@pqr.com;abc@uwy.com\"")]
        public InArgument<string> ToEmailAddress { get; set; }

        [Category("Receiver")]
        [Description("Email Addresses seprated by ; e.g.: \"xyz@pqr.com;abc@uwy.com\"")]
        public InArgument<string> CCEmailAddress { get; set; }

        [Category("Receiver")]
        [Description("Email Addresses seprated by ; e.g.: \"xyz@pqr.com;abc@uwy.com\"")]
        public InArgument<string> BccEmailAddress { get; set; }

        [Category("Email")]
        [RequiredArgument]
        public InArgument<string> Subject { get; set; }

        [Category("Email")]
        [RequiredArgument]
        public InArgument<string> Body { get; set; }

        [Category("Email")]
        [RequiredArgument]
        public InArgument<bool> IsBodyHtml { get; set; } = false;

        [Category("Configuration")]
        [RequiredArgument]
        public InArgument<string> SMTP_Server { get; set; }

        [Category("Configuration")]
        [RequiredArgument]
        public InArgument<int> SMTP_Port { get; set; }

        [Category("Configuration")]
        [RequiredArgument]
        public InArgument<string> SMTP_Email { get; set; }

        [Category("Configuration")]
        [RequiredArgument]
        public InArgument<string> SMTP_password { get; set; }

        [Category("Configuration")]
        [RequiredArgument]
        public InArgument<bool> EnableSsl { get; set; } = false;

        [Category("Attachments")]
        public InArgument<List<string>> MailAttachments { get; set; } = new List<string>();

        [Category("Sender")]
        public InArgument<string> FromEmailAddress { get; set; }


        [Category("Output")]
        public OutArgument<bool> IsSuccess { get; set; }

        [Category("Output")]
        public OutArgument<string> Result { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            try
            {
                MailMessage message = new MailMessage();
                SmtpClient smtp = new SmtpClient();
                if (string.IsNullOrEmpty(FromEmailAddress.Get(context)))
                {
                    message.From = new MailAddress(SMTP_Email.Get(context));
                }
                else
                {
                    message.From = new MailAddress(FromEmailAddress.Get(context));
                }
                

                foreach(string emailid in ToEmailAddress.Get(context).Trim().Split(';'))
                {
                    message.To.Add(new MailAddress(emailid));
                }

                if (! string.IsNullOrEmpty(CCEmailAddress.Get(context)))
                {
                    foreach (string emailid in CCEmailAddress.Get(context).Trim().Split(';'))
                    {
                        message.CC.Add(new MailAddress(emailid));
                    }
                }

                if (! string.IsNullOrEmpty(BccEmailAddress.Get(context)))
                {
                    foreach (string emailid in BccEmailAddress.Get(context).Trim().Split(';'))
                    {
                        message.Bcc.Add(new MailAddress(emailid));
                    }
                }


                message.Subject = Subject.Get(context).Trim();
                message.IsBodyHtml = IsBodyHtml.Get(context); //to make message body as html  
                message.Body = Body.Get(context).Trim();

                Console.WriteLine("----------------------");


                if (MailAttachments.Get(context) != null)
                {
                    Console.WriteLine("Number of Attachments: " + MailAttachments.Get(context).Count.ToString());
                    foreach (string file in MailAttachments.Get(context))
                    {
                        message.Attachments.Add(new Attachment(file));
                    }
                }
                

                smtp.Port = Convert.ToInt32(SMTP_Port.Get(context));
                smtp.Host = SMTP_Server.Get(context).Trim();
                smtp.EnableSsl = EnableSsl.Get(context);
                smtp.UseDefaultCredentials = false;
                smtp.Credentials = new System.Net.NetworkCredential(SMTP_Email.Get(context), SMTP_password.Get(context));
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtp.Send(message);
                IsSuccess.Set(context, true);
                Result.Set(context, string.Empty);
            }
            catch(Exception e)
            {
                Console.WriteLine("Error: "+e.Message+"--"+e.Source);
                IsSuccess.Set(context, false);
                Result.Set(context, "Error: " + e.Message);
            }
        }
    }

    public class DynamicAttachmentsMailMessage : CodeActivity
    {
        [Category("Receiver")]
        [RequiredArgument]
        [Description("Email Addresses seprated by ; e.g.: \"xyz@pqr.com;abc@uwy.com\"")]
        public InArgument<string> ToEmailAddress { get; set; }
        

        [Category("Receiver")]
        [Description("Email Addresses seprated by ; e.g.: \"xyz@pqr.com;abc@uwy.com\"")]
        public InArgument<string> CCEmailAddress { get; set; }

        [Category("Receiver")]
        [Description("Email Addresses seprated by ; e.g.: \"xyz@pqr.com;abc@uwy.com\"")]
        public InArgument<string> BccEmailAddress { get; set; }

        [Category("Email")]
        [RequiredArgument]
        public InArgument<string> Subject { get; set; }

        [Category("Email")]
        [RequiredArgument]
        public InArgument<string> Body { get; set; }

        [Category("Email")]
        [RequiredArgument]
        public InArgument<bool> IsBodyHtml { get; set; } = false;

        [Category("Attachments")]
        public InArgument<List<string>> MailAttachments { get; set; }

        [Category("Output")]
        [RequiredArgument]
        public OutArgument<MailMessage> out_MailMessage { get; set; }

        
        protected override void Execute(CodeActivityContext context)
        {
            try
            {
                MailMessage message = new MailMessage();

                foreach (string emailid in ToEmailAddress.Get(context).Trim().Split(';'))
                {
                    message.To.Add(new MailAddress(emailid));
                }

                if (!string.IsNullOrEmpty(CCEmailAddress.Get(context)))
                {
                    foreach (string emailid in CCEmailAddress.Get(context).Trim().Split(';'))
                    {
                        message.CC.Add(new MailAddress(emailid));
                    }
                }

                if (!string.IsNullOrEmpty(BccEmailAddress.Get(context)))
                {
                    foreach (string emailid in BccEmailAddress.Get(context).Trim().Split(';'))
                    {
                        message.Bcc.Add(new MailAddress(emailid));
                    }
                }


                message.Subject = Subject.Get(context).Trim();
                message.IsBodyHtml = IsBodyHtml.Get(context); //to make message body as html  
                message.Body = Body.Get(context).Trim();

                Console.WriteLine("----------------------");

                

                if (MailAttachments.Get(context) != null)
                {
                    Console.WriteLine("Number of Attachments: " + MailAttachments.Get(context).Count.ToString());
                    foreach (string file in MailAttachments.Get(context))
                    {
                        message.Attachments.Add(new Attachment(file));
                    }
                }
                out_MailMessage.Set(context, message);
            }
            catch (Exception e)
            {
                Console.WriteLine("Error: " + e.Message + "--" + e.Source);
                out_MailMessage.Set(context, null);
            }
        }
    }
}
