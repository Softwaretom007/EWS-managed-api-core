/*
 * Exchange Web Services Managed API
 *
 * Copyright (c) Microsoft Corporation
 * All rights reserved.
 *
 * MIT License
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this
 * software and associated documentation files (the "Software"), to deal in the Software
 * without restriction, including without limitation the rights to use, copy, modify, merge,
 * publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
 * to whom the Software is furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all copies or
 * substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
 * INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
 * PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
 * FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
 * OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
 * DEALINGS IN THE SOFTWARE.
 */

namespace Microsoft.Exchange.WebServices.Data
{
    using System.Collections.Generic;
    using System.Threading;
    using System.Threading.Tasks;

    /// <summary>
    /// Represents an e-mail message. Properties available on e-mail messages are defined in the EmailMessageSchema class.
    /// </summary>
    [Attachable]
    [ServiceObjectDefinition(XmlElementNames.Message)]
    public class EmailMessage : Item
    {
        /// <summary>
        /// Initializes an unsaved local instance of <see cref="EmailMessage"/>. To bind to an existing e-mail message, use EmailMessage.Bind() instead.
        /// </summary>
        /// <param name="service">The ExchangeService object to which the e-mail message will be bound.</param>
        public EmailMessage(ExchangeService service)
            : base(service)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="EmailMessage"/> class.
        /// </summary>
        /// <param name="parentAttachment">The parent attachment.</param>
        internal EmailMessage(ItemAttachment parentAttachment)
            : base(parentAttachment)
        {
        }

        /// <summary>
        /// Binds to an existing e-mail message and loads the specified set of properties.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="service">The service to use to bind to the e-mail message.</param>
        /// <param name="id">The Id of the e-mail message to bind to.</param>
        /// <param name="propertySet">The set of properties to load.</param>
        /// <returns>An EmailMessage instance representing the e-mail message corresponding to the specified Id.</returns>
        public static new Task<EmailMessage> Bind(
            ExchangeService service,
            ItemId id,
            PropertySet propertySet,
            CancellationToken token = default(CancellationToken))
        {
            return service.BindToItem<EmailMessage>(id, propertySet, token);
        }

        /// <summary>
        /// Binds to an existing e-mail message and loads its first class properties.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="service">The service to use to bind to the e-mail message.</param>
        /// <param name="id">The Id of the e-mail message to bind to.</param>
        /// <returns>An EmailMessage instance representing the e-mail message corresponding to the specified Id.</returns>
        public static new Task<EmailMessage> Bind(ExchangeService service, ItemId id)
        {
            return EmailMessage.Bind(
                service,
                id,
                PropertySet.FirstClassProperties);
        }

        /// <summary>
        /// Internal method to return the schema associated with this type of object.
        /// </summary>
        /// <returns>The schema associated with this type of object.</returns>
        internal override ServiceObjectSchema GetSchema()
        {
            return EmailMessageSchema.Instance;
        }

        /// <summary>
        /// Gets the minimum required server version.
        /// </summary>
        /// <returns>Earliest Exchange version in which this service object type is supported.</returns>
        internal override ExchangeVersion GetMinimumRequiredServerVersion()
        {
            return ExchangeVersion.Exchange2007_SP1;
        }

        /// <summary>
        /// Send message.
        /// </summary>
        /// <param name="parentFolderId">The parent folder id.</param>
        /// <param name="messageDisposition">The message disposition.</param>
        private async System.Threading.Tasks.Task InternalSend(FolderId parentFolderId, MessageDisposition messageDisposition, CancellationToken token)
        {
            this.ThrowIfThisIsAttachment();

            if (this.IsNew)
            {
                if ((this.Attachments.Count == 0) || (messageDisposition == MessageDisposition.SaveOnly))
                {
                    await this.InternalCreate(
                        parentFolderId,
                        messageDisposition,
                        null,
                        token);
                }
                else
                {
                    // If the message has attachments, save as a draft (and add attachments) before sending.
                    await this.InternalCreate(
                        null,                           // null means use the Drafts folder in the mailbox of the authenticated user.
                        MessageDisposition.SaveOnly,
                        null,
                        token);

                    await this.Service.SendItem(this, parentFolderId, token);
                }
            }
            else
            {
                // Regardless of whether item is dirty or not, if it has unprocessed
                // attachment changes, process them now.

                // Validate and save attachments before sending.
                if (this.HasUnprocessedAttachmentChanges())
                {
                    this.Attachments.Validate();
                    await this.Attachments.Save();
                }

                if (this.PropertyBag.GetIsUpdateCallNecessary())
                {
                    await this.InternalUpdate(
                        parentFolderId,
                        ConflictResolutionMode.AutoResolve,
                        messageDisposition,
                        null,
                        token);
                }
                else
                {
                    await this.Service.SendItem(this, parentFolderId, token);
                }
            }
        }

        /// <summary>
        /// Creates a reply response to the message.
        /// </summary>
        /// <param name="replyAll">Indicates whether the reply should go to all of the original recipients of the message.</param>
        /// <returns>A ResponseMessage representing the reply response that can subsequently be modified and sent.</returns>
        public ResponseMessage CreateReply(bool replyAll)
        {
            this.ThrowIfThisIsNew();

            return new ResponseMessage(
                this,
                replyAll ? ResponseMessageType.ReplyAll : ResponseMessageType.Reply);
        }

        /// <summary>
        /// Creates a forward response to the message.
        /// </summary>
        /// <returns>A ResponseMessage representing the forward response that can subsequently be modified and sent.</returns>
        public ResponseMessage CreateForward()
        {
            this.ThrowIfThisIsNew();

            return new ResponseMessage(this, ResponseMessageType.Forward);
        }

        /// <summary>
        /// Replies to the message. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="bodyPrefix">The prefix to prepend to the original body of the message.</param>
        /// <param name="replyAll">Indicates whether the reply should be sent to all of the original recipients of the message.</param>
        public System.Threading.Tasks.Task Reply(MessageBody bodyPrefix, bool replyAll)
        {
            ResponseMessage responseMessage = this.CreateReply(replyAll);

            responseMessage.BodyPrefix = bodyPrefix;

            return responseMessage.SendAndSaveCopy();
        }

        /// <summary>
        /// Forwards the message. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="bodyPrefix">The prefix to prepend to the original body of the message.</param>
        /// <param name="toRecipients">The recipients to forward the message to.</param>
        public System.Threading.Tasks.Task Forward(MessageBody bodyPrefix, params EmailAddress[] toRecipients)
        {
            return this.Forward(bodyPrefix, (IEnumerable<EmailAddress>)toRecipients);
        }

        /// <summary>
        /// Forwards the message. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="bodyPrefix">The prefix to prepend to the original body of the message.</param>
        /// <param name="toRecipients">The recipients to forward the message to.</param>
        public System.Threading.Tasks.Task Forward(MessageBody bodyPrefix, IEnumerable<EmailAddress> toRecipients)
        {
            ResponseMessage responseMessage = this.CreateForward();

            responseMessage.BodyPrefix = bodyPrefix;
            responseMessage.ToRecipients.AddRange(toRecipients);

            return responseMessage.SendAndSaveCopy();
        }

        /// <summary>
        /// Sends this e-mail message. Calling this method results in at least one call to EWS.
        /// </summary>
        public System.Threading.Tasks.Task Send(CancellationToken token = default(CancellationToken))
        {
            return this.InternalSend(null, MessageDisposition.SendOnly, token);
        }

        /// <summary>
        /// Sends this e-mail message and saves a copy of it in the specified folder. SendAndSaveCopy does not work if the
        /// message has unsaved attachments. In that case, the message must first be saved and then sent. Calling this method
        /// results in a call to EWS.
        /// </summary>
        /// <param name="destinationFolderId">The Id of the folder in which to save the copy.</param>
        public System.Threading.Tasks.Task SendAndSaveCopy(FolderId destinationFolderId, CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParam(destinationFolderId, "destinationFolderId");

            return this.InternalSend(destinationFolderId, MessageDisposition.SendAndSaveCopy, token);
        }

        /// <summary>
        /// Sends this e-mail message and saves a copy of it in the specified folder. SendAndSaveCopy does not work if the
        /// message has unsaved attachments. In that case, the message must first be saved and then sent. Calling this method
        /// results in a call to EWS.
        /// </summary>
        /// <param name="destinationFolderName">The name of the folder in which to save the copy.</param>
        public System.Threading.Tasks.Task SendAndSaveCopy(WellKnownFolderName destinationFolderName, CancellationToken token = default(CancellationToken))
        {
            return this.InternalSend(new FolderId(destinationFolderName), MessageDisposition.SendAndSaveCopy, token);
        }

        /// <summary>
        /// Sends this e-mail message and saves a copy of it in the Sent Items folder. SendAndSaveCopy does not work if the
        /// message has unsaved attachments. In that case, the message must first be saved and then sent. Calling this method
        /// results in a call to EWS.
        /// </summary>
        public System.Threading.Tasks.Task SendAndSaveCopy(CancellationToken token = default(CancellationToken))
        {
            return this.InternalSend(new FolderId(WellKnownFolderName.SentItems), MessageDisposition.SendAndSaveCopy, token);
        }

        /// <summary>
        /// Suppresses the read receipt on the message. Calling this method results in a call to EWS.
        /// </summary>
        public System.Threading.Tasks.Task SuppressReadReceipt(CancellationToken token = default(CancellationToken))
        {
            this.ThrowIfThisIsNew();

            return new SuppressReadReceipt(this).InternalCreate(null, null, token);
        }

        #region Properties

        /// <summary>
        /// Gets the list of To recipients for the e-mail message.
        /// </summary>
        public EmailAddressCollection ToRecipients
        {
            get { return (EmailAddressCollection)this.PropertyBag[EmailMessageSchema.ToRecipients]; }
        }

        /// <summary>
        /// Gets the list of Bcc recipients for the e-mail message.
        /// </summary>
        public EmailAddressCollection BccRecipients
        {
            get { return (EmailAddressCollection)this.PropertyBag[EmailMessageSchema.BccRecipients]; }
        }

        /// <summary>
        /// Gets the Likers associated with the message.
        /// </summary>
        public EmailAddressCollection Likers
        {
            get { return (EmailAddressCollection)this.PropertyBag[EmailMessageSchema.Likers]; }
        }

        /// <summary>
        /// Gets the list of Cc recipients for the e-mail message.
        /// </summary>
        public EmailAddressCollection CcRecipients
        {
            get { return (EmailAddressCollection)this.PropertyBag[EmailMessageSchema.CcRecipients]; }
        }

        /// <summary>
        /// Gets the conversation topic of the e-mail message.
        /// </summary>
        public string ConversationTopic
        {
            get { return (string)this.PropertyBag[EmailMessageSchema.ConversationTopic]; }
        }

        /// <summary>
        /// Gets the conversation index of the e-mail message.
        /// </summary>
        public byte[] ConversationIndex
        {
            get { return (byte[])this.PropertyBag[EmailMessageSchema.ConversationIndex]; }
        }

        /// <summary>
        /// Gets or sets the "on behalf" sender of the e-mail message.
        /// </summary>
        public EmailAddress From
        {
            get { return (EmailAddress)this.PropertyBag[EmailMessageSchema.From]; }
            set { this.PropertyBag[EmailMessageSchema.From] = value; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether this is an associated message.
        /// </summary>
        public new bool IsAssociated
        {
            get { return base.IsAssociated; }

            // The "new" keyword is used to expose the setter only on Message types, because
            // EWS only supports creation of FAI Message types.  IsAssociated is a readonly
            // property of the Item type but it is used by the CreateItem web method for creating
            // associated messages.
            set { this.PropertyBag[EmailMessageSchema.IsAssociated] = value; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether a read receipt is requested for the e-mail message.
        /// </summary>
        public bool IsDeliveryReceiptRequested
        {
            get { return (bool)this.PropertyBag[EmailMessageSchema.IsDeliveryReceiptRequested]; }
            set { this.PropertyBag[EmailMessageSchema.IsDeliveryReceiptRequested] = value; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the e-mail message is read.
        /// </summary>
        public bool IsRead
        {
            get { return (bool)this.PropertyBag[EmailMessageSchema.IsRead]; }
            set { this.PropertyBag[EmailMessageSchema.IsRead] = value; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether a read receipt is requested for the e-mail message.
        /// </summary>
        public bool IsReadReceiptRequested
        {
            get { return (bool)this.PropertyBag[EmailMessageSchema.IsReadReceiptRequested]; }
            set { this.PropertyBag[EmailMessageSchema.IsReadReceiptRequested] = value; }
        }

        /// <summary>
        /// Gets or sets a value indicating whether a response is requested for the e-mail message.
        /// </summary>
        public bool? IsResponseRequested
        {
            get { return (bool?)this.PropertyBag[EmailMessageSchema.IsResponseRequested]; }
            set { this.PropertyBag[EmailMessageSchema.IsResponseRequested] = value; }
        }

        /// <summary>
        /// Gets the Internet Message Id of the e-mail message.
        /// </summary>
        public string InternetMessageId
        {
            get { return (string)this.PropertyBag[EmailMessageSchema.InternetMessageId]; }
        }

        /// <summary>
        /// Gets or sets the references of the e-mail message.
        /// </summary>
        public string References
        {
            get { return (string)this.PropertyBag[EmailMessageSchema.References]; }
            set { this.PropertyBag[EmailMessageSchema.References] = value; }
        }

        /// <summary>
        /// Gets a list of e-mail addresses to which replies should be addressed.
        /// </summary>
        public EmailAddressCollection ReplyTo
        {
            get { return (EmailAddressCollection)this.PropertyBag[EmailMessageSchema.ReplyTo]; }
        }

        /// <summary>
        /// Gets or sets the sender of the e-mail message.
        /// </summary>
        public EmailAddress Sender
        {
            get { return (EmailAddress)this.PropertyBag[EmailMessageSchema.Sender]; }
            set { this.PropertyBag[EmailMessageSchema.Sender] = value; }
        }

        /// <summary>
        /// Gets the ReceivedBy property of the e-mail message.
        /// </summary>
        public EmailAddress ReceivedBy
        {
            get { return (EmailAddress)this.PropertyBag[EmailMessageSchema.ReceivedBy]; }
        }

        /// <summary>
        /// Gets the ReceivedRepresenting property of the e-mail message.
        /// </summary>
        public EmailAddress ReceivedRepresenting
        {
            get { return (EmailAddress)this.PropertyBag[EmailMessageSchema.ReceivedRepresenting]; }
        }

        /// <summary>
        /// Gets the ApprovalRequestData property of the e-mail message.
        /// </summary>
        public ApprovalRequestData ApprovalRequestData
        {
            get { return (ApprovalRequestData)this.PropertyBag[EmailMessageSchema.ApprovalRequestData]; }
        }

        /// <summary>
        /// Gets the VotingInformation property of the e-mail message.
        /// </summary>
        public VotingInformation VotingInformation
        {
            get { return (VotingInformation)this.PropertyBag[EmailMessageSchema.VotingInformation]; }
        }
        #endregion
    }
}