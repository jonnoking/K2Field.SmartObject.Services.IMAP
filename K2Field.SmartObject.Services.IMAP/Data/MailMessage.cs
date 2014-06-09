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
    public class MailMessage
    {
        private ServiceAssemblyBase serviceBroker = null;

        public MailMessage(ServiceAssemblyBase serviceBroker)
        {
            // Set local serviceBroker variable.
            this.serviceBroker = serviceBroker;
        }
        List<Property> AllProps = new List<Property>();
        List<Property> MessageProps = new List<Property>();

        public void Create()
        {
            ServiceObject MailMessageServiceObject = new ServiceObject();
            MailMessageServiceObject.Name = "mailmessage";
            MailMessageServiceObject.MetaData.DisplayName = "Mail Message";
            MailMessageServiceObject.Active = true;

            AllProps = GetAllProperties();
            MessageProps = GetProperties();

            foreach (Property prop in AllProps)
            {
                MailMessageServiceObject.Properties.Add(prop);
            }

            // add methods
            MailMessageServiceObject.Methods.Add(GetAllMessages());
            MailMessageServiceObject.Methods.Add(GetMessages());
            MailMessageServiceObject.Methods.Add(SearchMessageBySubject());
            MailMessageServiceObject.Methods.Add(SearchMessageBySubjectUID());
            MailMessageServiceObject.Methods.Add(SearchMessageByBody());
            MailMessageServiceObject.Methods.Add(SearchMessageByBodyUID());
            MailMessageServiceObject.Methods.Add(SearchMessageByFrom());
            MailMessageServiceObject.Methods.Add(SearchMessageByFromUID());
            MailMessageServiceObject.Methods.Add(GetMessageByUID());
            MailMessageServiceObject.Methods.Add(GetMessageBySubject());
            serviceBroker.Service.ServiceObjects.Add(MailMessageServiceObject);
        }

        public List<Property> GetAllProperties()
        {
            List<Property> props = new List<Property>();
            props.AddRange(GetInputProperties());
            props.AddRange(GetProperties());

            return props;
        }

        public List<Property> GetProperties()
        {
            List<Property> props = new List<Property>();

            Property BCCProperty = new Property();
            BCCProperty.Name = "bcc";
            BCCProperty.MetaData.DisplayName = "BCC";
            BCCProperty.Type = "string";
            BCCProperty.SoType = SoType.Memo;
            props.Add(BCCProperty);

            Property BodyProperty = new Property();
            BodyProperty.Name = "body";
            BodyProperty.MetaData.DisplayName = "Body";
            BodyProperty.Type = "string";
            BodyProperty.SoType = SoType.Memo;
            props.Add(BodyProperty);

            Property CCProperty = new Property();
            CCProperty.Name = "cc";
            CCProperty.MetaData.DisplayName = "CC";
            CCProperty.Type = "string";
            CCProperty.SoType = SoType.Memo;
            props.Add(CCProperty);

            Property DateProperty = new Property();
            DateProperty.Name = "date";
            DateProperty.MetaData.DisplayName = "Date";
            //DateProperty.Type = "datetime";
            DateProperty.SoType = SoType.DateTime;
            props.Add(DateProperty);

            Property FromProperty = new Property();
            FromProperty.Name = "from";
            FromProperty.MetaData.DisplayName = "From";
            FromProperty.Type = "string";
            FromProperty.SoType = SoType.Text;
            props.Add(FromProperty);


            Property ImportanceProperty = new Property();
            ImportanceProperty.Name = "importance";
            ImportanceProperty.MetaData.DisplayName = "Importance";
            ImportanceProperty.Type = "string";
            ImportanceProperty.SoType = SoType.Text;
            props.Add(ImportanceProperty);

            Property MessageIDProperty = new Property();
            MessageIDProperty.Name = "messageid";
            MessageIDProperty.MetaData.DisplayName = "MessageID";
            MessageIDProperty.Type = "string";
            MessageIDProperty.SoType = SoType.Text;
            props.Add(MessageIDProperty);

            Property ReplyToProperty = new Property();
            ReplyToProperty.Name = "replyto";
            ReplyToProperty.MetaData.DisplayName = "ReplyTo";
            ReplyToProperty.Type = "string";
            ReplyToProperty.SoType = SoType.Memo;
            props.Add(ReplyToProperty);

            Property SenderProperty = new Property();
            SenderProperty.Name = "sender";
            SenderProperty.MetaData.DisplayName = "Sender";
            SenderProperty.Type = "string";
            SenderProperty.SoType = SoType.Text;
            props.Add(SenderProperty);

            Property SizeProperty = new Property();
            SizeProperty.Name = "size";
            SizeProperty.MetaData.DisplayName = "Size";
            SizeProperty.Type = "int";
            SizeProperty.SoType = SoType.Number;
            props.Add(SizeProperty);

            Property SubjectProperty = new Property();
            SubjectProperty.Name = "subject";
            SubjectProperty.MetaData.DisplayName = "Subject";
            SubjectProperty.Type = "string";
            SubjectProperty.SoType = SoType.Text;
            props.Add(SubjectProperty);

            Property ToProperty = new Property();
            ToProperty.Name = "to";
            ToProperty.MetaData.DisplayName = "To";
            ToProperty.Type = "string";
            ToProperty.SoType = SoType.Memo;
            props.Add(ToProperty);

            Property UIDProperty = new Property();
            UIDProperty.Name = "uid";
            UIDProperty.MetaData.DisplayName = "UID";
            UIDProperty.Type = "string";
            UIDProperty.SoType = SoType.Text;
            props.Add(UIDProperty);

            Property AttachmentsCountProperty = new Property();
            AttachmentsCountProperty.Name = "attachmentscount";
            AttachmentsCountProperty.MetaData.DisplayName = "Attachments Count";
            AttachmentsCountProperty.Type = "int";
            AttachmentsCountProperty.SoType = SoType.Number;
            props.Add(AttachmentsCountProperty);

            Property RawMessage = new Property();
            RawMessage.Name = "rawmessage";
            RawMessage.MetaData.DisplayName = "Raw Message";
            RawMessage.Type = "string";
            RawMessage.SoType = SoType.Memo;
            props.Add(RawMessage);

            Property Base64EmlMessage = new Property();
            Base64EmlMessage.Name = "base64emlmessage";
            Base64EmlMessage.MetaData.DisplayName = "Base64 Eml Message";
            Base64EmlMessage.Type = "string";
            Base64EmlMessage.SoType = SoType.Memo;
            props.Add(Base64EmlMessage);

            return props;
        }

        public List<Property> GetInputProperties()
        {
            List<Property> props = new List<Property>();

            Property MailboxProperty = new Property();
            MailboxProperty.Name = "mailbox";
            MailboxProperty.MetaData.DisplayName = "Mailbox";
            MailboxProperty.Type = "string";
            MailboxProperty.SoType = SoType.Text;
            props.Add(MailboxProperty);

            Property HeadersOnlyProperty = new Property();
            HeadersOnlyProperty.Name = "headersonly";
            HeadersOnlyProperty.MetaData.DisplayName = "Headers Only";
            HeadersOnlyProperty.Type = "bool";
            HeadersOnlyProperty.SoType = SoType.YesNo;
            props.Add(HeadersOnlyProperty);

            Property IndexStartProperty = new Property();
            IndexStartProperty.Name = "startindex";
            IndexStartProperty.MetaData.DisplayName = "Start";
            IndexStartProperty.Type = "int";
            IndexStartProperty.SoType = SoType.Number;
            props.Add(IndexStartProperty);

            Property NumProperty = new Property();
            NumProperty.Name = "numberofmessages";
            NumProperty.MetaData.DisplayName = "Number of Messages";
            NumProperty.Type = "int";
            NumProperty.SoType = SoType.Number;
            props.Add(NumProperty);

            Property FromDateProperty = new Property();
            FromDateProperty.Name = "fromdate";
            FromDateProperty.MetaData.DisplayName = "From Date";
            FromDateProperty.Type = "datetime";
            FromDateProperty.SoType = SoType.DateTime;
            props.Add(FromDateProperty);

            Property SubjectFilterProperty = new Property();
            SubjectFilterProperty.Name = "subjectfilter";
            SubjectFilterProperty.MetaData.DisplayName = "Subject Filter";
            SubjectFilterProperty.Type = "string";
            SubjectFilterProperty.SoType = SoType.Text;
            props.Add(SubjectFilterProperty);

            Property MessageFilterProperty = new Property();
            MessageFilterProperty.Name = "bodyfilter";
            MessageFilterProperty.MetaData.DisplayName = "Body Filter";
            MessageFilterProperty.Type = "string";
            MessageFilterProperty.SoType = SoType.Text;
            props.Add(MessageFilterProperty);

            Property FromFilterProperty = new Property();
            FromFilterProperty.Name = "fromfilter";
            FromFilterProperty.MetaData.DisplayName = "From Filter";
            FromFilterProperty.Type = "string";
            FromFilterProperty.SoType = SoType.Text;
            props.Add(FromFilterProperty);

            Property SetSeenProperty = new Property();
            SetSeenProperty.Name = "setseen";
            SetSeenProperty.MetaData.DisplayName = "Set Seen";
            SetSeenProperty.Type = "bool";
            SetSeenProperty.SoType = SoType.YesNo;
            props.Add(SetSeenProperty);


            return props;
        }

        private Method GetAllMessages()
        {
            Method method = new Method();
            method.Name = "getallmessages";
            method.MetaData.DisplayName = "Get All Messages";
            method.Type = MethodType.List;

            method.InputProperties.Add(AllProps.Where(p => p.Name.Equals("mailbox")).FirstOrDefault());
            method.Validation.RequiredProperties.Add(method.InputProperties[0]);
            method.InputProperties.Add(AllProps.Where(p => p.Name.Equals("headersonly")).FirstOrDefault());
            method.Validation.RequiredProperties.Add(method.InputProperties[1]);
            method.InputProperties.Add(AllProps.Where(p => p.Name.Equals("setseen")).FirstOrDefault());
            method.Validation.RequiredProperties.Add(method.InputProperties[2]);

            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("mailbox")).FirstOrDefault());
            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("headersonly")).FirstOrDefault());
            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("setseen")).FirstOrDefault());
            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("startindex")).FirstOrDefault());
            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("numberofmessages")).FirstOrDefault());


            foreach (Property prop in MessageProps)
            {
                if (!method.ReturnProperties.Contains(prop))
                {
                    method.ReturnProperties.Add(prop);
                }
            }

            return method;
        }

        private Method GetMessages()
        {
            Method method = new Method();
            method.Name = "getmessages";
            method.MetaData.DisplayName = "Get Messages";
            method.Type = MethodType.List;

            method.InputProperties.Add(AllProps.Where(p => p.Name.Equals("mailbox")).FirstOrDefault());
            method.Validation.RequiredProperties.Add(method.InputProperties[0]);

            method.InputProperties.Add(AllProps.Where(p => p.Name.Equals("headersonly")).FirstOrDefault());
            method.Validation.RequiredProperties.Add(method.InputProperties[1]);

            method.InputProperties.Add(AllProps.Where(p => p.Name.Equals("setseen")).FirstOrDefault());
            method.Validation.RequiredProperties.Add(method.InputProperties[2]);

            method.InputProperties.Add(AllProps.Where(p => p.Name.Equals("startindex")).FirstOrDefault());
            method.Validation.RequiredProperties.Add(method.InputProperties[3]);

            method.InputProperties.Add(AllProps.Where(p => p.Name.Equals("numberofmessages")).FirstOrDefault());
            method.Validation.RequiredProperties.Add(method.InputProperties[4]);

            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("mailbox")).FirstOrDefault());
            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("headersonly")).FirstOrDefault());
            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("setseen")).FirstOrDefault());
            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("startindex")).FirstOrDefault());
            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("numberofmessages")).FirstOrDefault());
            foreach (Property prop in MessageProps)
            {
                if (!method.ReturnProperties.Contains(prop))
                {
                    method.ReturnProperties.Add(prop);
                }
            }
            return method;
        }

        private Method SearchMessageBySubject()
        {
            Method method = new Method();
            method.Name = "searchmessagesbysubject";
            method.MetaData.DisplayName = "Search Messages By Subject";
            method.Type = MethodType.List;

            method.InputProperties.Add(AllProps.Where(p => p.Name.Equals("mailbox")).FirstOrDefault());
            method.Validation.RequiredProperties.Add(method.InputProperties[0]);

            method.InputProperties.Add(AllProps.Where(p => p.Name.Equals("headersonly")).FirstOrDefault());
            method.Validation.RequiredProperties.Add(method.InputProperties[1]);

            method.InputProperties.Add(AllProps.Where(p => p.Name.Equals("setseen")).FirstOrDefault());
            method.Validation.RequiredProperties.Add(method.InputProperties[2]);

            method.InputProperties.Add(AllProps.Where(p => p.Name.Equals("subjectfilter")).FirstOrDefault());
            method.Validation.RequiredProperties.Add(method.InputProperties[3]);

            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("mailbox")).FirstOrDefault());
            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("headersonly")).FirstOrDefault());
            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("setseen")).FirstOrDefault());
            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("subjectfilter")).FirstOrDefault());
            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("startindex")).FirstOrDefault());
            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("numberofmessages")).FirstOrDefault());

            foreach (Property prop in MessageProps)
            {
                if (!method.ReturnProperties.Contains(prop))
                {
                    method.ReturnProperties.Add(prop);
                }
            }
            return method;
        }

        private Method SearchMessageBySubjectUID()
        {
            Method method = new Method();
            method.Name = "searchmessagesbysubjectUID";
            method.MetaData.DisplayName = "Search Messages By Subject UID";
            method.Type = MethodType.List;

            method.InputProperties.Add(AllProps.Where(p => p.Name.Equals("mailbox")).FirstOrDefault());
            method.Validation.RequiredProperties.Add(method.InputProperties[0]);

            method.InputProperties.Add(AllProps.Where(p => p.Name.Equals("headersonly")).FirstOrDefault());
            method.Validation.RequiredProperties.Add(method.InputProperties[1]);

            method.InputProperties.Add(AllProps.Where(p => p.Name.Equals("setseen")).FirstOrDefault());
            method.Validation.RequiredProperties.Add(method.InputProperties[2]);

            method.InputProperties.Add(AllProps.Where(p => p.Name.Equals("subjectfilter")).FirstOrDefault());
            method.Validation.RequiredProperties.Add(method.InputProperties[3]);

            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("mailbox")).FirstOrDefault());
            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("headersonly")).FirstOrDefault());
            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("setseen")).FirstOrDefault());
            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("subjectfilter")).FirstOrDefault());
            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("startindex")).FirstOrDefault());
            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("numberofmessages")).FirstOrDefault());

            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("uid")).FirstOrDefault());
           
            return method;
        }

        private Method SearchMessageByBody()
        {
            Method method = new Method();
            method.Name = "searchmessagesbybody";
            method.MetaData.DisplayName = "Search Messages By Body";
            method.Type = MethodType.List;

            method.InputProperties.Add(AllProps.Where(p => p.Name.Equals("mailbox")).FirstOrDefault());
            method.Validation.RequiredProperties.Add(method.InputProperties[0]);

            method.InputProperties.Add(AllProps.Where(p => p.Name.Equals("headersonly")).FirstOrDefault());
            method.Validation.RequiredProperties.Add(method.InputProperties[1]);

            method.InputProperties.Add(AllProps.Where(p => p.Name.Equals("setseen")).FirstOrDefault());
            method.Validation.RequiredProperties.Add(method.InputProperties[2]);

            method.InputProperties.Add(AllProps.Where(p => p.Name.Equals("bodyfilter")).FirstOrDefault());
            method.Validation.RequiredProperties.Add(method.InputProperties[3]);

            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("mailbox")).FirstOrDefault());
            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("headersonly")).FirstOrDefault());
            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("setseen")).FirstOrDefault());
            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("bodyfilter")).FirstOrDefault());
            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("startindex")).FirstOrDefault());
            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("numberofmessages")).FirstOrDefault());

            foreach (Property prop in MessageProps)
            {
                if (!method.ReturnProperties.Contains(prop))
                {
                    method.ReturnProperties.Add(prop);
                }
            }
            return method;
        }

        private Method SearchMessageByBodyUID() 
        {
            Method method = new Method();
            method.Name = "searchmessagesbybodyuid";
            method.MetaData.DisplayName = "Search Messages By Body UID";
            method.Type = MethodType.List;

            method.InputProperties.Add(AllProps.Where(p => p.Name.Equals("mailbox")).FirstOrDefault());
            method.Validation.RequiredProperties.Add(method.InputProperties[0]);

            method.InputProperties.Add(AllProps.Where(p => p.Name.Equals("headersonly")).FirstOrDefault());
            method.Validation.RequiredProperties.Add(method.InputProperties[1]);

            method.InputProperties.Add(AllProps.Where(p => p.Name.Equals("setseen")).FirstOrDefault());
            method.Validation.RequiredProperties.Add(method.InputProperties[2]);

            method.InputProperties.Add(AllProps.Where(p => p.Name.Equals("bodyfilter")).FirstOrDefault());
            method.Validation.RequiredProperties.Add(method.InputProperties[3]);

            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("mailbox")).FirstOrDefault());
            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("headersonly")).FirstOrDefault());
            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("setseen")).FirstOrDefault());
            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("bodyfilter")).FirstOrDefault());
            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("startindex")).FirstOrDefault());
            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("numberofmessages")).FirstOrDefault());

            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("uid")).FirstOrDefault());

            return method;
        }

        private Method SearchMessageByFrom()
        {
            Method method = new Method();
            method.Name = "searchmessagesbyfrom";
            method.MetaData.DisplayName = "Search Messages By From";
            method.Type = MethodType.List;

            method.InputProperties.Add(AllProps.Where(p => p.Name.Equals("mailbox")).FirstOrDefault());
            method.Validation.RequiredProperties.Add(method.InputProperties[0]);

            method.InputProperties.Add(AllProps.Where(p => p.Name.Equals("headersonly")).FirstOrDefault());
            method.Validation.RequiredProperties.Add(method.InputProperties[1]);

            method.InputProperties.Add(AllProps.Where(p => p.Name.Equals("setseen")).FirstOrDefault());
            method.Validation.RequiredProperties.Add(method.InputProperties[2]);

            method.InputProperties.Add(AllProps.Where(p => p.Name.Equals("fromfilter")).FirstOrDefault());
            method.Validation.RequiredProperties.Add(method.InputProperties[3]);

            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("mailbox")).FirstOrDefault());
            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("headersonly")).FirstOrDefault());
            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("setseen")).FirstOrDefault());
            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("fromfilter")).FirstOrDefault());
            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("startindex")).FirstOrDefault());
            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("numberofmessages")).FirstOrDefault());

            foreach (Property prop in MessageProps)
            {
                if (!method.ReturnProperties.Contains(prop))
                {
                    method.ReturnProperties.Add(prop);
                }
            }
            return method;
        }

        private Method SearchMessageByFromUID()
        {
            Method method = new Method();
            method.Name = "searchmessagesbyfromUID";
            method.MetaData.DisplayName = "Search Messages By From UID";
            method.Type = MethodType.List;

            method.InputProperties.Add(AllProps.Where(p => p.Name.Equals("mailbox")).FirstOrDefault());
            method.Validation.RequiredProperties.Add(method.InputProperties[0]);

            method.InputProperties.Add(AllProps.Where(p => p.Name.Equals("headersonly")).FirstOrDefault());
            method.Validation.RequiredProperties.Add(method.InputProperties[1]);

            method.InputProperties.Add(AllProps.Where(p => p.Name.Equals("setseen")).FirstOrDefault());
            method.Validation.RequiredProperties.Add(method.InputProperties[2]);

            method.InputProperties.Add(AllProps.Where(p => p.Name.Equals("fromfilter")).FirstOrDefault());
            method.Validation.RequiredProperties.Add(method.InputProperties[3]);

            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("mailbox")).FirstOrDefault());
            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("headersonly")).FirstOrDefault());
            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("setseen")).FirstOrDefault());
            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("fromfilter")).FirstOrDefault());
            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("startindex")).FirstOrDefault());
            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("numberofmessages")).FirstOrDefault());

            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("uid")).FirstOrDefault()); 
            
            return method;
        }

        private Method GetMessageByUID()
        {
            Method method = new Method();
            method.Name = "getmessagebyuid";
            method.MetaData.DisplayName = "Get Message By UID";
            method.Type = MethodType.Read;

            method.InputProperties.Add(AllProps.Where(p => p.Name.Equals("mailbox")).FirstOrDefault());
            method.Validation.RequiredProperties.Add(method.InputProperties[0]);

            method.InputProperties.Add(AllProps.Where(p => p.Name.Equals("uid")).FirstOrDefault());
            method.Validation.RequiredProperties.Add(method.InputProperties[1]);

            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("mailbox")).FirstOrDefault());
            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("uid")).FirstOrDefault());

            foreach (Property prop in MessageProps)
            {
                if (!method.ReturnProperties.Contains(prop))
                {
                    method.ReturnProperties.Add(prop);
                }
            }
            return method;
        }

        private Method GetMessageBySubject()
        {
            Method method = new Method();
            method.Name = "getmessagebysubject";
            method.MetaData.DisplayName = "Get Message By Subject";
            method.Type = MethodType.Read;

            method.InputProperties.Add(AllProps.Where(p => p.Name.Equals("mailbox")).FirstOrDefault());
            method.Validation.RequiredProperties.Add(method.InputProperties[0]);

            method.InputProperties.Add(AllProps.Where(p => p.Name.Equals("subject")).FirstOrDefault());
            method.Validation.RequiredProperties.Add(method.InputProperties[1]);

            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("mailbox")).FirstOrDefault());
            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("subject")).FirstOrDefault());


            foreach (Property prop in MessageProps)
            {
                if (!method.ReturnProperties.Contains(prop))
                {
                    method.ReturnProperties.Add(prop);
                }
            }
            return method;
        }


        #region Execute

        public string GetAddresses(ICollection<System.Net.Mail.MailAddress> addresses)
        {
            List<System.Net.Mail.MailAddress> add = addresses.ToList<System.Net.Mail.MailAddress>();

            string address = string.Empty;
            
            for(int i=0;i<add.Count;i++)
            {
                address += add[i].Address;
                if (i < add.Count - 1)
                {
                    address += "; ";
                }                
            }

            return address;
        }

        public void GetMessages(Property[] inputs, RequiredProperties required, Property[] returns, MethodType methodType, ServiceObject serviceObject)
        {
            serviceObject.Properties.InitResultTable();

            System.Data.DataRow dr;
            Helper h = new Helper(serviceBroker);

            string subjectfilter = string.Empty;
            string bodyfilter = string.Empty;
            string fromfilter = string.Empty;

            string startindex = string.Empty;
            string numberofmessages = string.Empty;

            string mailbox = inputs.Where(p => p.Name.Equals("mailbox")).FirstOrDefault().Value.ToString();
            bool headersonly = bool.Parse(inputs.Where(p => p.Name.Equals("headersonly")).FirstOrDefault().Value.ToString());
            bool setseen = bool.Parse(inputs.Where(p => p.Name.Equals("setseen")).FirstOrDefault().Value.ToString()); ;
            AE.Net.Mail.MailMessage mtemp = new AE.Net.Mail.MailMessage();
            try
            {
                using (var ic = h.GetImapClient())
                {

                    AE.Net.Mail.MailMessage[] m = null;
                    Lazy<AE.Net.Mail.MailMessage>[] mm = null;

                    bool isLazy = false;

                    switch (serviceObject.Methods[0].Name.ToLower())
                    {
                        case "getallmessages":
                            m = ic.GetMessages(0, ic.GetMessageCount(), headersonly, setseen);
                            isLazy = false;
                            break;
                        case "getmessages":
                            startindex = inputs.Where(p => p.Name.Equals("startindex")).FirstOrDefault().Value.ToString();
                            numberofmessages = inputs.Where(p => p.Name.Equals("numberofmessages")).FirstOrDefault().Value.ToString();
                            m = ic.GetMessages(int.Parse(startindex), int.Parse(numberofmessages), headersonly, setseen);                          
                            isLazy = false;
                            break;
                        case "searchmessagesbysubject":
                            subjectfilter = inputs.Where(p => p.Name.Equals("subjectfilter")).FirstOrDefault().Value.ToString();
                            mm = ic.SearchMessages(SearchCondition.Undeleted().And(SearchCondition.Subject(subjectfilter)));
                            isLazy = true;
                            break;
                        case "searchmessagesbybody":
                            bodyfilter = inputs.Where(p => p.Name.Equals("bodyfilter")).FirstOrDefault().Value.ToString();
                            mm = ic.SearchMessages(SearchCondition.Undeleted().And(SearchCondition.Body(bodyfilter))).ToArray();
                            isLazy = true;
                            break;
                        case "searchmessagesbyfrom":
                            fromfilter = inputs.Where(p => p.Name.Equals("fromfilter")).FirstOrDefault().Value.ToString();
                            mm = ic.SearchMessages(SearchCondition.Undeleted().And(SearchCondition.From(fromfilter))).ToArray();
                            isLazy = true;
                            break;                            
                    }

                    //AE.Net.Mail.MailMessage[] mm = ic.GetMessages(0, ic.GetMessageCount(), headersonly, setseen);

                    if (isLazy)
                    {
                        foreach (System.Lazy<AE.Net.Mail.MailMessage> msg in mm)
                        {
                            AE.Net.Mail.MailMessage mmsg = msg.Value;
                            dr = serviceBroker.ServicePackage.ResultTable.NewRow();

                            MapMailMessage(dr, mmsg);
                            
                            dr["mailbox"] = mailbox;
                            dr["headersonly"] = headersonly;
                            dr["setseen"] = setseen;

                            switch (serviceObject.Methods[0].Name.ToLower())
                            {
                                case "searchmessagesbysubject":
                                    dr["subjectfilter"] = subjectfilter;
                                    break;
                                case "searchmessagesbybody":
                                    dr["bodyfilter"] = bodyfilter;
                                    break;
                                case "searchmessagesbyfrom":
                                    dr["fromfilter"] = fromfilter;
                                    break;
                            }


                            dr["startindex"] = startindex;
                            dr["numberofmessages"] = numberofmessages;

                            serviceBroker.ServicePackage.ResultTable.Rows.Add(dr);
                        }
                    }
                    else
                    {
                        foreach (AE.Net.Mail.MailMessage msg in m.OrderByDescending(p => p.Date))
                        {
                            mtemp = msg;
                            dr = serviceBroker.ServicePackage.ResultTable.NewRow();

                            MapMailMessage(dr, msg);

                            dr["mailbox"] = mailbox;
                            dr["headersonly"] = headersonly;
                            dr["setseen"] = setseen;
                            switch (serviceObject.Methods[0].Name.ToLower())
                            {
                                case "searchmessagesbysubject":
                                    dr["subjectfilter"] = subjectfilter;
                                    break;
                                case "searchmessagesbybody":
                                    dr["bodyfilter"] = bodyfilter;
                                    break;
                                case "searchmessagesbyfrom":
                                    dr["fromfilter"] = fromfilter;
                                    break;
                            }
                            dr["startindex"] = startindex;
                            dr["numberofmessages"] = numberofmessages;

                            serviceBroker.ServicePackage.ResultTable.Rows.Add(dr);
                        }
                    }

                    ic.Disconnect();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(mtemp.Subject);
                //serviceObject.Properties.BindPropertiesToResultTable();
            }
            //serviceObject.Properties.BindPropertiesToResultTable();
        }

        public void MapMailMessage(System.Data.DataRow dr, AE.Net.Mail.MailMessage mmsg)
        {
            dr["bcc"] = GetAddresses(mmsg.Bcc);
            dr["body"] = mmsg.Body;
            dr["cc"] = GetAddresses(mmsg.Cc);
            dr["date"] = mmsg.Date;
            dr["from"] = mmsg.From != null ? mmsg.From.Address : "";
            dr["importance"] = mmsg.Importance.ToString();
            dr["replyto"] = GetAddresses(mmsg.ReplyTo);
            dr["sender"] = mmsg.Sender != null ? mmsg.Sender.Address : "";
            dr["size"] = mmsg.Size;
            dr["subject"] = mmsg.Subject;
            dr["to"] = GetAddresses(mmsg.To);
            dr["uid"] = mmsg.Uid;
            //dr["rawmessage"] = m.RawBody;
            //dr["base64emlmessage"] = m.Base64Body;
            dr["attachmentscount"] = mmsg.Attachments != null ? mmsg.Attachments.Count : 0;


        }

        public void GetMessageBy(Property[] inputs, RequiredProperties required, Property[] returns, MethodType methodType, ServiceObject serviceObject)
        {
            serviceObject.Properties.InitResultTable();

            Helper h = new Helper(serviceBroker);

            string mailbox = inputs.Where(p => p.Name.Equals("mailbox")).FirstOrDefault().Value.ToString();
            string uid = string.Empty;
            string subjectfilter = string.Empty;                

            AE.Net.Mail.MailMessage m = null;
            try
            {
                using (var ic = h.GetImapClient())
                {

                    switch (serviceObject.Methods[0].Name.ToLower())
                    {
                        case "getmessagebyuid":
                            uid = inputs.Where(p => p.Name.Equals("uid")).FirstOrDefault().Value.ToString();
                            m = ic.GetMessage(uid, false);
                            break;
                        case "getmessagebysubject":
                            subjectfilter = inputs.Where(p => p.Name.Equals("subjectfilter")).FirstOrDefault().Value.ToString();
                            Lazy<AE.Net.Mail.MailMessage>[] mm = ic.SearchMessages(SearchCondition.Undeleted().And(SearchCondition.Subject(subjectfilter)));
                            Lazy<AE.Net.Mail.MailMessage> lm = mm.OrderByDescending(p => p.Value.Date).Where(p => p.Value.Subject.ToLower().Equals(subjectfilter)).FirstOrDefault();

                            if (lm != null && lm.Value != null)
                            {
                                m = lm.Value;
                            }
                            
                            break;
                    }

                    // = ic.GetMessage(uid, false);

                    if (m == null)
                    {
                        return;
                    }                    

                    returns.Where(p => p.Name.ToLower().Equals("bcc")).FirstOrDefault().Value = GetAddresses(m.Bcc);
                    returns.Where(p => p.Name.ToLower().Equals("body")).FirstOrDefault().Value = m.Body;
                    returns.Where(p => p.Name.ToLower().Equals("cc")).FirstOrDefault().Value = GetAddresses(m.Cc);
                    returns.Where(p => p.Name.ToLower().Equals("date")).FirstOrDefault().Value = m.Date;
                    returns.Where(p => p.Name.ToLower().Equals("from")).FirstOrDefault().Value = m.From != null ? m.From.Address : "";
                    returns.Where(p => p.Name.ToLower().Equals("importance")).FirstOrDefault().Value = m.Importance.ToString();
                    returns.Where(p => p.Name.ToLower().Equals("replyto")).FirstOrDefault().Value = GetAddresses(m.ReplyTo);
                    returns.Where(p => p.Name.ToLower().Equals("sender")).FirstOrDefault().Value = m.Sender != null ? m.Sender.Address : "";
                    returns.Where(p => p.Name.ToLower().Equals("size")).FirstOrDefault().Value = m.Size;
                    returns.Where(p => p.Name.ToLower().Equals("subject")).FirstOrDefault().Value = m.Subject;
                    returns.Where(p => p.Name.ToLower().Equals("to")).FirstOrDefault().Value = GetAddresses(m.To);
                    returns.Where(p => p.Name.ToLower().Equals("uid")).FirstOrDefault().Value = m.Uid;
                    //returns.Where(p => p.Name.ToLower().Equals("rawmessage")).FirstOrDefault().Value = m.RawBody;
                    //returns.Where(p => p.Name.ToLower().Equals("base64emlmessage")).FirstOrDefault().Value = m.Base64Body;
                    returns.Where(p => p.Name.ToLower().Equals("attachmentscount")).FirstOrDefault().Value = m.Attachments != null ? m.Attachments.Count : 0;

                    returns.Where(p => p.Name.ToLower().Equals("mailbox")).FirstOrDefault().Value = mailbox;
                    returns.Where(p => p.Name.ToLower().Equals("uid")).FirstOrDefault().Value = uid;
                    returns.Where(p => p.Name.ToLower().Equals("subjectfilter")).FirstOrDefault().Value = subjectfilter;

                    ic.Disconnect();
                }
            }
            catch (Exception ex)
            {
                //serviceObject.Properties.BindPropertiesToResultTable();
            }
            serviceObject.Properties.BindPropertiesToResultTable();
        }

        #endregion Execute

    }
}
