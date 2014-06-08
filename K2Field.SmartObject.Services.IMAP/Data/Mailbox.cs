using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SourceCode.SmartObjects.Services.ServiceSDK.Objects;
using SourceCode.SmartObjects.Services.ServiceSDK;
using SourceCode.SmartObjects.Services.ServiceSDK.Types;
using AE.Net.Mail;

namespace K2Field.SmartObject.Services.IMAP.Data
{
    class Mailbox
    {
        private ServiceAssemblyBase serviceBroker = null;

        public Mailbox(ServiceAssemblyBase serviceBroker)
        {
            // Set local serviceBroker variable.
            this.serviceBroker = serviceBroker;
        }
        List<Property> AllProps = new List<Property>();

        public void Create()
        {
            ServiceObject MailboxServiceObject = new ServiceObject();
            MailboxServiceObject.Name = "mailbox";
            MailboxServiceObject.MetaData.DisplayName = "Mailbox";
            MailboxServiceObject.Active = true;

            AllProps = GetProperties();

            foreach (Property prop in GetProperties())
            {
                MailboxServiceObject.Properties.Add(prop);
            }

            // add methods
            //MailboxServiceObject.Methods.Add(GetMailboxes());
            MailboxServiceObject.Methods.Add(GetMailbox());
            MailboxServiceObject.Methods.Add(GetInbox());
            serviceBroker.Service.ServiceObjects.Add(MailboxServiceObject);
        }


        public List<Property> GetProperties()
        {
            List<Property> props = new List<Property>();

            Property NameProperty = new Property();
            NameProperty.Name = "name";
            NameProperty.MetaData.DisplayName = "Name";
            NameProperty.Type = "string";
            NameProperty.SoType = SoType.Text;
            props.Add(NameProperty);

            Property NumMsgProperty = new Property();
            NumMsgProperty.Name = "numberofmessages";
            NumMsgProperty.MetaData.DisplayName = "Number of Messages";
            NumMsgProperty.Type = "int";
            NumMsgProperty.SoType = SoType.Number;
            props.Add(NumMsgProperty);

            Property NumNewMsgProperty = new Property();
            NumNewMsgProperty.Name = "numberofnewmessages";
            NumNewMsgProperty.MetaData.DisplayName = "Number of New Messages";
            NumNewMsgProperty.Type = "int";
            NumNewMsgProperty.SoType = SoType.Number;
            props.Add(NumNewMsgProperty);

            Property NumUnSeenProperty = new Property();
            NumUnSeenProperty.Name = "numberofunseenmessages";
            NumUnSeenProperty.MetaData.DisplayName = "Number of Unseen Messages";
            NumUnSeenProperty.Type = "int";
            NumUnSeenProperty.SoType = SoType.Number;
            props.Add(NumUnSeenProperty);

            Property IsWritableProperty = new Property();
            IsWritableProperty.Name = "iswritable";
            IsWritableProperty.MetaData.DisplayName = "Is Writable";
            IsWritableProperty.Type = "int";
            IsWritableProperty.SoType = SoType.Number;
            props.Add(IsWritableProperty);

            return props;
        }

        private Method GetMailboxes()
        {
            Method GetMailboxes = new Method();
            GetMailboxes.Name = "getmailboxes";
            GetMailboxes.MetaData.DisplayName = "Get Mail Boxes";
            GetMailboxes.Type = MethodType.List;

            foreach (Property prop in AllProps)
            {
                GetMailboxes.ReturnProperties.Add(prop);
            }
            return GetMailboxes;
        }

        private Method GetMailbox()
        {
            Method GetMailbox = new Method();
            GetMailbox.Name = "getmailbox";
            GetMailbox.MetaData.DisplayName = "Get Mailbox";
            GetMailbox.Type = MethodType.Read;

            GetMailbox.InputProperties.Add(AllProps.Where(p => p.Name.Equals("name")).FirstOrDefault());
            GetMailbox.Validation.RequiredProperties.Add(GetMailbox.InputProperties[0]);

            foreach (Property prop in AllProps)
            {
                GetMailbox.ReturnProperties.Add(prop);
            }
            return GetMailbox;            
        }

        private Method GetInbox()
        {
            Method GetMailbox = new Method();
            GetMailbox.Name = "getinbox";
            GetMailbox.MetaData.DisplayName = "Get Inbox";
            GetMailbox.Type = MethodType.Read;

            foreach (Property prop in AllProps)
            {
                GetMailbox.ReturnProperties.Add(prop);
            }
            return GetMailbox;

        }

        #region Execution

        public void GetMailbox(Property[] inputs, RequiredProperties required, Property[] returns, MethodType methodType, ServiceObject serviceObject)
        {
            serviceObject.Properties.InitResultTable();
            Helper h = new Helper(serviceBroker);

            try
            {
                using (var ic = h.GetImapClient())
                {
                    string mailbox = string.Empty;
                    if (serviceObject.Methods[0].Name.Equals("getmailbox"))
                    {
                        mailbox = inputs.Where(p => p.Name.Equals("name")).FirstOrDefault().Value.ToString();
                    }
                    else
                    {
                        mailbox = "INBOX";
                    }


                    AE.Net.Mail.Imap.Mailbox mb = ic.SelectMailbox(mailbox);

                    returns.Where(p => p.Name.Equals("name")).FirstOrDefault().Value = mb.Name;
                    returns.Where(p => p.Name.Equals("numberofmessages")).FirstOrDefault().Value = mb.NumMsg;
                    returns.Where(p => p.Name.Equals("numberofnewmessages")).FirstOrDefault().Value = mb.NumNewMsg;
                    returns.Where(p => p.Name.Equals("numberofunseenmessages")).FirstOrDefault().Value = mb.NumUnSeen;
                    returns.Where(p => p.Name.Equals("iswritable")).FirstOrDefault().Value = mb.IsWritable;

                    ic.Disconnect();
                }
            }
            catch (Exception ex)
            {
             
                serviceObject.Properties.BindPropertiesToResultTable();
            }
            serviceObject.Properties.BindPropertiesToResultTable();
        }

        #endregion Execution
    }
}
