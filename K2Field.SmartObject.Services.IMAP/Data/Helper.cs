using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SourceCode.SmartObjects.Services.ServiceSDK;
using AE.Net.Mail;

namespace K2Field.SmartObject.Services.IMAP.Data
{
    public class Helper
    {
        private ServiceAssemblyBase serviceBroker = null;
        public Helper(ServiceAssemblyBase serviceBroker)
        {
            // Set local serviceBroker variable.
            this.serviceBroker = serviceBroker;
        }

        public AE.Net.Mail.ImapClient GetImapClient()
        {
            return new AE.Net.Mail.ImapClient(
                serviceBroker.Service.ServiceConfiguration["MailServer"].ToString(), 
                serviceBroker.Service.ServiceConfiguration.ServiceAuthentication.UserName, 
                serviceBroker.Service.ServiceConfiguration.ServiceAuthentication.Password,
                AE.Net.Mail.AuthMethods.Login, int.Parse(serviceBroker.Service.ServiceConfiguration["MailPort"].ToString()),
                bool.Parse(serviceBroker.Service.ServiceConfiguration["MailSSL"].ToString()));
        }
        
    }
}
