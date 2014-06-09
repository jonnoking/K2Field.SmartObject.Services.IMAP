using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Linq;
using System.Xml.Linq;

using SourceCode.SmartObjects.Services.ServiceSDK;
using SourceCode.SmartObjects.Services.ServiceSDK.Objects;
using SourceCode.SmartObjects.Services.ServiceSDK.Types;

using K2Field.SmartObject.Services.IMAP.Interfaces;
using System.Net;
using System.IO;

namespace K2Field.SmartObject.Services.IMAP.Data
{
    /// <summary>
    /// A concrete implementation of IDataConnector responsible for interacting with an underlying system or technology. The purpose of this class it to expose and represent the underlying data and services as Service Objects for consumptions by K2 SmartObjects.
    /// </summary>
    class DataConnector : IDataConnector
    {
        #region Class Level Fields

        #region Constants
        /// <summary>
        /// Constant for the Type Mappings configuration lookup in the service instance.
        /// </summary>
        private static string __TypeMappings = "Type Mappings";

        #endregion

        #region Private Fields
        /// <summary>
        /// Local serviceBroker variable.
        /// </summary>
        private ServiceAssemblyBase serviceBroker = null;

        string MAILSERVER = string.Empty;
        int MAILPORT = 0;
        string MAILUSERNAME = string.Empty;
        string MAILPASSWORD = string.Empty;
        bool MAILSSL = true;

        #endregion

        #endregion

        #region Constructor
        /// <summary>
        /// Instantiates a new DataConnector.
        /// </summary>
        /// <param name="serviceBroker">The ServiceBroker.</param>
        public DataConnector(ServiceAssemblyBase serviceBroker)
        {
            // Set local serviceBroker variable.
            this.serviceBroker = serviceBroker;
        }
        #endregion

        #region Methods

        #region void Dispose()
        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            // Add any additional IDisposable implementation code here. Make sure to dispose of any data connections.
            // Clear references to serviceBroker.
            serviceBroker = null;
        }
        #endregion

        #region void GetConfiguration()
        /// <summary>
        /// Gets the configuration from the service instance and stores the retrieved configuration in local variables for later use.
        /// </summary>
        public void GetConfiguration()
        {
            MAILSERVER = serviceBroker.Service.ServiceConfiguration["MailServer"].ToString();
            MAILPORT = int.Parse(serviceBroker.Service.ServiceConfiguration["MailPort"].ToString());
            MAILUSERNAME = serviceBroker.Service.ServiceConfiguration.ServiceAuthentication.UserName;
            MAILPASSWORD = serviceBroker.Service.ServiceConfiguration.ServiceAuthentication.Password;
            MAILSSL = bool.Parse(serviceBroker.Service.ServiceConfiguration["MailSSL"].ToString());
        }
        #endregion

        #region void SetupConfiguration()
        /// <summary>
        /// Sets up the required configuration parameters in the service instance. When a new service instance is registered for this ServiceBroker, the configuration parameters are surfaced to the appropriate tooling. The configuration parameters are provided by the person registering the service instance.
        /// </summary>
        public void SetupConfiguration()
        {
            serviceBroker.Service.ServiceConfiguration.Add("MailServer", true, "");
            serviceBroker.Service.ServiceConfiguration.Add("MailPort", true, "");
            serviceBroker.Service.ServiceConfiguration.Add("MailSSL", true, "true");
        }
        #endregion

        #region void SetupService()
        /// <summary>
        /// Sets up the service instance's default name, display name, and description.
        /// </summary>
        public void SetupService()
        {
            // Add service instance default setup code here.
            serviceBroker.Service.Name = "IMAPServiceBroker";
            serviceBroker.Service.MetaData.DisplayName = "IMAP - " + serviceBroker.Service.ServiceConfiguration["MailServer"].ToString();
            serviceBroker.Service.MetaData.Description = "IMAP Service Broker";
        }
        #endregion

        #region void DescribeSchema()
        /// <summary>
        /// Describes the schema of the underlying data and services to the K2 platform.
        /// </summary>
        public void DescribeSchema()
        {

            //MailBox
            Mailbox mailboxServiceObject = new Mailbox(serviceBroker);
            mailboxServiceObject.Create();
            //Message
            MailMessage mailmessageServiceObject = new MailMessage(serviceBroker);
            mailmessageServiceObject.Create();

            //Attachment
            //MailAttachment mailAttachmentServiceObject = new MailAttachment(serviceBroker);
            //mailAttachmentServiceObject.Create();


                //if (!serviceBroker.Service.ServiceObjects.Contains(obj))
                //{
                //    serviceBroker.Service.ServiceObjects.Add(obj);
                //}
            //}

        }
        #endregion

       
        #region TypeMappings GetTypeMappings()
        /// <summary>
        /// Gets the type mappings used to map the underlying data's types to the appropriate K2 SmartObject types.
        /// </summary>
        /// <returns>A TypeMappings object containing the ServiceBroker's type mappings which were previously stored in the service instance configuration.</returns>
        public TypeMappings GetTypeMappings()
        {
            // Lookup and return the type mappings stored in the service instance.
            return (TypeMappings)serviceBroker.Service.ServiceConfiguration[__TypeMappings];
        }
        #endregion

        #region void SetTypeMappings()
        /// <summary>
        /// Sets the type mappings used to map the underlying data's types to the appropriate K2 SmartObject types.
        /// </summary>
        public void SetTypeMappings()
        {
            // Variable declaration.
            TypeMappings map = new TypeMappings();

            map.Add("Edm.Binary", SoType.Memo);
            map.Add("Edm.Boolean", SoType.YesNo);
            map.Add("Edm.Byte", SoType.Memo);
            map.Add("Edm.DateTime", SoType.DateTime);
            map.Add("Edm.DateTimeOffset", SoType.DateTime);
            map.Add("Edm.Time", SoType.DateTime);
            map.Add("Edm.Double", SoType.Decimal);
            map.Add("Edm.Decimal", SoType.Decimal);
            map.Add("Edm.Single", SoType.Decimal);
            map.Add("Edm.Guid", SoType.Guid);
            map.Add("Edm.Int16", SoType.Number);
            map.Add("Edm.Int32", SoType.Number);
            map.Add("Edm.Int64", SoType.Number);
            map.Add("Edm.String", SoType.Text);

            // Add the type mappings to the service instance.
            serviceBroker.Service.ServiceConfiguration.Add(__TypeMappings, map);
        }
        #endregion

        #region void Execute(Property[] inputs, RequiredProperties required, Property[] returns, MethodType methodType, ServiceObject serviceObject)
        /// <summary>
        /// SmartObject execution. HTTP calls and mapping return xml to SmartObject properties
        /// </summary>
        public void Execute(Property[] inputs, RequiredProperties required, Property[] returns, MethodType methodType, ServiceObject serviceObject)
        {
            serviceObject.Properties.InitResultTable();

            if (serviceObject.Methods[0].Name.Equals("getmailbox") || serviceObject.Methods[0].Name.Equals("getinbox"))
            {
                Mailbox mailbox = new Mailbox(serviceBroker);
                mailbox.GetMailbox(inputs, required, returns, methodType, serviceObject);
            }



            if (serviceObject.Methods[0].Name.Equals("getallmessages") || serviceObject.Methods[0].Name.Equals("getmessages") || serviceObject.Methods[0].Name.Equals("searchmessagesbysubject") || serviceObject.Methods[0].Name.Equals("searchmessagesbybody") || serviceObject.Methods[0].Name.Equals("searchmessagesbyfrom"))
            {
                MailMessage msg = new MailMessage(serviceBroker);
                msg.GetMessages(inputs, required, returns, methodType, serviceObject);
            }

            if (serviceObject.Methods[0].Name.Equals("getmessagebyuid") || serviceObject.Methods[0].Name.Equals("getmessagebysubject"))
            {
                MailMessage msg = new MailMessage(serviceBroker);
                msg.GetMessageBy(inputs, required, returns, methodType, serviceObject);
            }



            if (serviceObject.Methods[0].Name.Equals("getallattachments"))
            {
                MailAttachment attach = new MailAttachment(serviceBroker);
                attach.GetAllAttachments(inputs, required, returns, methodType, serviceObject);
            }

            if (serviceObject.Methods[0].Name.Equals("getattachment"))
            {
                MailAttachment attach = new MailAttachment(serviceBroker);
                attach.GetAttachment(inputs, required, returns, methodType, serviceObject);
            }

        }
        #endregion

       #region XmlDocument DiscoverSchema()
        /// <summary>
        /// Discover OData Schema
        /// </summary>
        public XmlDocument DiscoverSchema()
        {
            return null;
        }
        #endregion

           #region void ToSystemType()
        /// <summary>
        /// Sets the type mappings used to map the underlying data's types to the appropriate K2 SmartObject types.
        /// </summary>
        public string EdmToSystemType(string edmtype)
        {
            // Variable declaration.
            TypeMappings map = new TypeMappings();
            string nettype = "";
            switch (edmtype)
            {
                case "Edm.Binary":
                case "Edm.Byte":
                case "Edm.String":
                    nettype = typeof(System.String).ToString();
                    break;
                case "Edm.Boolean":
                    nettype = typeof(System.Boolean).ToString();
                    break;
                case "Edm.DateTime":
                case "Edm.DateTimeOffset":
                case "Edm.Time":
                    nettype = typeof(System.DateTime).ToString();
                    break;
                case "Edm.Double":
                case "Edm.Decimal":
                case "Edm.Single":
                    nettype = typeof(System.Decimal).ToString();
                    break;
                case "Edm.Guid":
                    nettype = typeof(System.Guid).ToString();
                    break;
                case "Edm.Int16":
                    nettype = typeof(System.Int16).ToString();
                    break;
                case "Edm.Int32":
                    nettype = typeof(System.Int32).ToString();
                    break;
                case "Edm.Int64":
                    nettype = typeof(System.Int64).ToString();
                    break;
                default:
                    nettype = typeof(System.String).ToString();
                    break;
            }

            // Add the type mappings to the service instance.
            return nettype;
        }
        #endregion

        #endregion

    }
}