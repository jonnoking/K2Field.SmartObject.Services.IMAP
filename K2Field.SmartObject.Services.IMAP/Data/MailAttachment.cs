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
    class MailAttachment
    {
        private ServiceAssemblyBase serviceBroker = null;

        public MailAttachment(ServiceAssemblyBase serviceBroker)
        {
            // Set local serviceBroker variable.
            this.serviceBroker = serviceBroker;
        }
        List<Property> AllProps = new List<Property>();
        List<Property> MessageProps = new List<Property>();

        public void Create()
        {
            ServiceObject MailAttachmentServiceObject = new ServiceObject();
            MailAttachmentServiceObject.Name = "mailattachment";
            MailAttachmentServiceObject.MetaData.DisplayName = "Mail Attachment";
            MailAttachmentServiceObject.Active = true;

            AllProps = GetAllProperties();
            MessageProps = GetProperties();

            foreach (Property prop in AllProps)
            {
                MailAttachmentServiceObject.Properties.Add(prop);
            }

            // add methods
            MailAttachmentServiceObject.Methods.Add(GetAllAttachments());
            MailAttachmentServiceObject.Methods.Add(GetAttachment());
            serviceBroker.Service.ServiceObjects.Add(MailAttachmentServiceObject);
        }

        public List<Property> GetAllProperties()
        {
            List<Property> props = new List<Property>();
            props.AddRange(GetInputProperties());
            props.AddRange(GetProperties());

            return props;
        }
        AE.Net.Mail.Attachment a = new Attachment();
        

        public List<Property> GetProperties()
        {                     
            List<Property> props = new List<Property>();

            Property Body = new Property();
            Body.Name = "body";
            Body.MetaData.DisplayName = "Body";
            Body.Type = "string";
            Body.SoType = SoType.Memo;
            props.Add(Body);

            Property CharSet = new Property();
            CharSet.Name = "charset";
            CharSet.MetaData.DisplayName = "Charset";
            CharSet.Type = "string";
            CharSet.SoType = SoType.Text;
            props.Add(CharSet);

            Property ContentTransferEncoding = new Property();
            ContentTransferEncoding.Name = "contenttransferencoding";
            ContentTransferEncoding.MetaData.DisplayName = "Content Transfer Encoding";
            ContentTransferEncoding.Type = "string";
            ContentTransferEncoding.SoType = SoType.Text;
            props.Add(ContentTransferEncoding);

            Property ContentType = new Property();
            ContentType.Name = "contenttype";
            ContentType.MetaData.DisplayName = "Content Type";
            ContentType.Type = "string";
            ContentType.SoType = SoType.Text;
            props.Add(ContentType);

            Property Filename = new Property();
            Filename.Name = "filename";
            Filename.MetaData.DisplayName = "Filename";
            Filename.Type = "string";
            Filename.SoType = SoType.Text;
            props.Add(Filename);

            Property RawHeaders = new Property();
            RawHeaders.Name = "rawheaders";
            RawHeaders.MetaData.DisplayName = "Raw Headers";
            RawHeaders.Type = "string";
            RawHeaders.SoType = SoType.Memo;
            props.Add(RawHeaders);

            Property OnServer = new Property();
            OnServer.Name = "onserver";
            OnServer.MetaData.DisplayName = "On Server";
            OnServer.Type = "bool";
            OnServer.SoType = SoType.YesNo;
            props.Add(OnServer);

            return props;

        }

        public List<Property> GetInputProperties()
        {
            List<Property> props = new List<Property>();

            Property Mailbox = new Property();
            Mailbox.Name = "mailbox";
            Mailbox.MetaData.DisplayName = "Mailbox";
            Mailbox.Type = "string";
            Mailbox.SoType = SoType.Text;
            props.Add(Mailbox);

            Property UID = new Property();
            UID.Name = "uid";
            UID.MetaData.DisplayName = "UID";
            UID.Type = "string";
            UID.SoType = SoType.Text;
            props.Add(UID);

            Property AttachmentIndex = new Property();
            AttachmentIndex.Name = "uid";
            AttachmentIndex.MetaData.DisplayName = "UID";
            AttachmentIndex.Type = "int";
            AttachmentIndex.SoType = SoType.Number;
            props.Add(AttachmentIndex);

            return props;
        }

        private Method GetAllAttachments()
        {
            Method method = new Method();
            method.Name = "getallattachments";
            method.MetaData.DisplayName = "Get All Attachments";
            method.Type = MethodType.List;

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

        private Method GetAttachment()
        {
            Method method = new Method();
            method.Name = "getattachment";
            method.MetaData.DisplayName = "Get Attachment";
            method.Type = MethodType.Read;

            method.InputProperties.Add(AllProps.Where(p => p.Name.Equals("mailbox")).FirstOrDefault());
            method.Validation.RequiredProperties.Add(method.InputProperties[0]);
            method.InputProperties.Add(AllProps.Where(p => p.Name.Equals("uid")).FirstOrDefault());
            method.Validation.RequiredProperties.Add(method.InputProperties[1]);
            method.InputProperties.Add(AllProps.Where(p => p.Name.Equals("attachmentindex")).FirstOrDefault());
            method.Validation.RequiredProperties.Add(method.InputProperties[2]);


            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("mailbox")).FirstOrDefault());
            method.ReturnProperties.Add(AllProps.Where(p => p.Name.Equals("uid")).FirstOrDefault());
            method.InputProperties.Add(AllProps.Where(p => p.Name.Equals("attachmentindex")).FirstOrDefault());

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

        public void GetAllAttachments(Property[] inputs, RequiredProperties required, Property[] returns, MethodType methodType, ServiceObject serviceObject)
        {
            serviceObject.Properties.InitResultTable();

            System.Data.DataRow dr;
            Helper h = new Helper(serviceBroker);

            string mailbox = inputs.Where(p => p.Name.Equals("mailbox")).FirstOrDefault().Value.ToString();
            string uid = string.Empty;
            string subjectfilter = string.Empty;

            uid = inputs.Where(p => p.Name.Equals("uid")).FirstOrDefault().Value.ToString();

            AE.Net.Mail.MailMessage m = null;
            List<AE.Net.Mail.Attachment> attachments = new List<AE.Net.Mail.Attachment>();

            try
            {
                using (var ic = h.GetImapClient())
                {
                    m = ic.GetMessage(uid, false);

                    if (m == null)
                    {
                        return;
                    }

                    attachments = m.Attachments as List<AE.Net.Mail.Attachment>;

                    foreach (Attachment a in attachments)
                    {
                        dr = serviceBroker.ServicePackage.ResultTable.NewRow();
                        
                        dr["body"] = a.Body;
                        dr["charset"] = a.Charset;
                        dr["contenttransferencoding"] = a.ContentTransferEncoding;
                        dr["contenttype"] = a.ContentType;
                        dr["filename"] = a.Filename;
                        dr["rawheaders"] = a.RawHeaders;
                        dr["onserver"] = a.OnServer;

                        dr["mailbox"] = inputs.Where(p => p.Name.Equals("mailbox")).FirstOrDefault().Value.ToString();
                        dr["uid"] = inputs.Where(p => p.Name.Equals("uid")).FirstOrDefault().Value.ToString();

                        serviceBroker.ServicePackage.ResultTable.Rows.Add(dr);
                    }

                    ic.Disconnect();
                }
            }
            catch (Exception ex)
            {
                //serviceObject.Properties.BindPropertiesToResultTable();
            }
            
        }

        public void GetAttachment(Property[] inputs, RequiredProperties required, Property[] returns, MethodType methodType, ServiceObject serviceObject)
        {
            serviceObject.Properties.InitResultTable();
            
            Helper h = new Helper(serviceBroker);

            string mailbox = inputs.Where(p => p.Name.Equals("mailbox")).FirstOrDefault().Value.ToString();
            string uid = string.Empty;
            string subjectfilter = string.Empty;
            int index = -1;
            uid = inputs.Where(p => p.Name.Equals("uid")).FirstOrDefault().Value.ToString();
            index = int.Parse(inputs.Where(p => p.Name.Equals("attachmentindex")).FirstOrDefault().Value.ToString());

            AE.Net.Mail.MailMessage m = null;
            List<AE.Net.Mail.Attachment> attachments = new List<AE.Net.Mail.Attachment>();

            try
            {
                using (var ic = h.GetImapClient())
                {
                    m = ic.GetMessage(uid, false);

                    if (m == null)
                    {
                        return;
                    }

                    attachments = m.Attachments as List<AE.Net.Mail.Attachment>;

                    AE.Net.Mail.Attachment a = null;

                    if (index == -1 || index > attachments.Count)
                    {
                        return;
                    }
                    a = attachments[index];

                    if (a == null)
                    {
                        return;
                    }

                    returns.Where(p => p.Name.ToLower().Equals("body")).FirstOrDefault().Value = a.Body;
                    returns.Where(p => p.Name.ToLower().Equals("charset")).FirstOrDefault().Value = a.Charset;
                    returns.Where(p => p.Name.ToLower().Equals("contenttransferencoding")).FirstOrDefault().Value = a.ContentTransferEncoding;
                    returns.Where(p => p.Name.ToLower().Equals("contenttype")).FirstOrDefault().Value = a.ContentType;
                    returns.Where(p => p.Name.ToLower().Equals("filename")).FirstOrDefault().Value = a.Filename;
                    returns.Where(p => p.Name.ToLower().Equals("rawheaders")).FirstOrDefault().Value = a.RawHeaders;
                    returns.Where(p => p.Name.ToLower().Equals("onserver")).FirstOrDefault().Value = a.OnServer;
                    returns.Where(p => p.Name.ToLower().Equals("mailbox")).FirstOrDefault().Value = inputs.Where(p => p.Name.Equals("mailbox")).FirstOrDefault().Value.ToString();
                    returns.Where(p => p.Name.ToLower().Equals("uid")).FirstOrDefault().Value = inputs.Where(p => p.Name.Equals("uid")).FirstOrDefault().Value.ToString();;
                    returns.Where(p => p.Name.ToLower().Equals("attachmentindex")).FirstOrDefault().Value = inputs.Where(p => p.Name.Equals("attachmentindex")).FirstOrDefault().Value.ToString();;

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
