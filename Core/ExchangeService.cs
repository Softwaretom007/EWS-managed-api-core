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

using System.Security.Cryptography;

namespace Microsoft.Exchange.WebServices.Data
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Net;
    using System.Xml;
    using Microsoft.Exchange.WebServices.Autodiscover;
    using Microsoft.Exchange.WebServices.Data.Enumerations;
    using Microsoft.Exchange.WebServices.Data.Groups;
    using System.Threading.Tasks;
    using System.Threading;
    using System.Net.Http;
    using System.Runtime.InteropServices;
    using Microsoft.Exchange.WebServices.NETStandard.Core;

    /// <summary>
    /// Represents a binding to the Exchange Web Services.
    /// </summary>
    public sealed class ExchangeService : ExchangeServiceBase
    {
        #region Fields

        private CultureInfo preferredCulture;
        private ImpersonatedUserId impersonatedUserId;
        private PrivilegedUserId privilegedUserId;
        private ManagementRoles managementRoles;
        private IFileAttachmentContentHandler fileAttachmentContentHandler;
        private UnifiedMessaging unifiedMessaging;
        private bool enableScpLookup = true;
        private readonly IEWSHttpClient ewsHttpClient;

        #endregion

        #region Response object operations

        /// <summary>
        /// Create response object.
        /// </summary>
        /// <param name="responseObject">The response object.</param>
        /// <param name="parentFolderId">The parent folder id.</param>
        /// <param name="messageDisposition">The message disposition.</param>
        /// <returns>The list of items created or modified as a result of the "creation" of the response object.</returns>
        internal async Task<List<Item>> InternalCreateResponseObject(
            ServiceObject responseObject,
            FolderId parentFolderId,
            MessageDisposition? messageDisposition,
            CancellationToken token)
        {
            CreateResponseObjectRequest request = new CreateResponseObjectRequest(this, ServiceErrorHandling.ThrowOnError);

            request.ParentFolderId = parentFolderId;
            request.Items = new ServiceObject[] { responseObject };
            request.MessageDisposition = messageDisposition;

            ServiceResponseCollection<CreateResponseObjectResponse> responses = await request.ExecuteAsync(token).ConfigureAwait(false);

            return responses[0].Items;
        }

        #endregion

        #region Folder operations

        /// <summary>
        /// Creates a folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="folder">The folder.</param>
        /// <param name="parentFolderId">The parent folder id.</param>
        internal System.Threading.Tasks.Task CreateFolder(
            Folder folder,
            FolderId parentFolderId, CancellationToken token)
        {
            CreateFolderRequest request = new CreateFolderRequest(this, ServiceErrorHandling.ThrowOnError);

            request.Folders = new Folder[] { folder };
            request.ParentFolderId = parentFolderId;

            return request.ExecuteAsync(token);
        }

        /// <summary>
        /// Updates a folder.
        /// </summary>
        /// <param name="folder">The folder.</param>
        internal System.Threading.Tasks.Task UpdateFolder(Folder folder, CancellationToken token)
        {
            UpdateFolderRequest request = new UpdateFolderRequest(this, ServiceErrorHandling.ThrowOnError);

            request.Folders.Add(folder);

            return request.ExecuteAsync(token);
        }

        /// <summary>
        /// Copies a folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="folderId">The folder id.</param>
        /// <param name="destinationFolderId">The destination folder id.</param>
        /// <returns>Copy of folder.</returns>
        internal async Task<Folder> CopyFolder(
            FolderId folderId,
            FolderId destinationFolderId,
            CancellationToken token)
        {
            CopyFolderRequest request = new CopyFolderRequest(this, ServiceErrorHandling.ThrowOnError);

            request.DestinationFolderId = destinationFolderId;
            request.FolderIds.Add(folderId);

            ServiceResponseCollection<MoveCopyFolderResponse> responses = await request.ExecuteAsync(token).ConfigureAwait(false);

            return responses[0].Folder;
        }

        /// <summary>
        /// Move a folder.
        /// </summary>
        /// <param name="folderId">The folder id.</param>
        /// <param name="destinationFolderId">The destination folder id.</param>
        /// <returns>Moved folder.</returns>
        internal async Task<Folder> MoveFolder(
            FolderId folderId,
            FolderId destinationFolderId,
            CancellationToken token)
        {
            MoveFolderRequest request = new MoveFolderRequest(this, ServiceErrorHandling.ThrowOnError);

            request.DestinationFolderId = destinationFolderId;
            request.FolderIds.Add(folderId);

            ServiceResponseCollection<MoveCopyFolderResponse> responses = await request.ExecuteAsync(token).ConfigureAwait(false);

            return responses[0].Folder;
        }

        /// <summary>
        /// Finds folders.
        /// </summary>
        /// <param name="parentFolderIds">The parent folder ids.</param>
        /// <param name="searchFilter">The search filter. Available search filter classes
        /// include SearchFilter.IsEqualTo, SearchFilter.ContainsSubstring and 
        /// SearchFilter.SearchFilterCollection</param>
        /// <param name="view">The view controlling the number of folders returned.</param>
        /// <param name="errorHandlingMode">Indicates the type of error handling should be done.</param>
        /// <returns>Collection of service responses.</returns>
        private Task<ServiceResponseCollection<FindFolderResponse>> InternalFindFolders(
            IEnumerable<FolderId> parentFolderIds,
            SearchFilter searchFilter,
            FolderView view,
            ServiceErrorHandling errorHandlingMode,
            CancellationToken token)
        {
            FindFolderRequest request = new FindFolderRequest(this, errorHandlingMode);

            request.ParentFolderIds.AddRange(parentFolderIds);
            request.SearchFilter = searchFilter;
            request.View = view;

            return request.ExecuteAsync(token);
        }

        /// <summary>
        /// Obtains a list of folders by searching the sub-folders of the specified folder.
        /// </summary>
        /// <param name="parentFolderId">The Id of the folder in which to search for folders.</param>
        /// <param name="searchFilter">The search filter. Available search filter classes
        /// include SearchFilter.IsEqualTo, SearchFilter.ContainsSubstring and 
        /// SearchFilter.SearchFilterCollection</param>
        /// <param name="view">The view controlling the number of folders returned.</param>
        /// <returns>An object representing the results of the search operation.</returns>
        public async Task<FindFoldersResults> FindFolders(FolderId parentFolderId, SearchFilter searchFilter, FolderView view, CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParam(parentFolderId, "parentFolderId");
            EwsUtilities.ValidateParam(view, "view");
            EwsUtilities.ValidateParamAllowNull(searchFilter, "searchFilter");

            ServiceResponseCollection<FindFolderResponse> responses = await this.InternalFindFolders(
                new FolderId[] { parentFolderId },
                searchFilter,
                view,
                ServiceErrorHandling.ThrowOnError, token).ConfigureAwait(false);

            return responses[0].Results;
        }

        /// <summary>
        /// Obtains a list of folders by searching the sub-folders of each of the specified folders.
        /// </summary>
        /// <param name="parentFolderIds">The Ids of the folders in which to search for folders.</param>
        /// <param name="searchFilter">The search filter. Available search filter classes
        /// include SearchFilter.IsEqualTo, SearchFilter.ContainsSubstring and 
        /// SearchFilter.SearchFilterCollection</param>
        /// <param name="view">The view controlling the number of folders returned.</param>
        /// <returns>An object representing the results of the search operation.</returns>
        public Task<ServiceResponseCollection<FindFolderResponse>> FindFolders(IEnumerable<FolderId> parentFolderIds, SearchFilter searchFilter, FolderView view,
            CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParam(parentFolderIds, "parentFolderIds");
            EwsUtilities.ValidateParam(view, "view");
            EwsUtilities.ValidateParamAllowNull(searchFilter, "searchFilter");

            return this.InternalFindFolders(
                parentFolderIds,
                searchFilter,
                view,
                ServiceErrorHandling.ReturnErrors,
                token);
        }

        /// <summary>
        /// Obtains a list of folders by searching the sub-folders of the specified folder.
        /// </summary>
        /// <param name="parentFolderId">The Id of the folder in which to search for folders.</param>
        /// <param name="view">The view controlling the number of folders returned.</param>
        /// <returns>An object representing the results of the search operation.</returns>
        public async Task<FindFoldersResults> FindFolders(FolderId parentFolderId, FolderView view, CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParam(parentFolderId, "parentFolderId");
            EwsUtilities.ValidateParam(view, "view");

            ServiceResponseCollection<FindFolderResponse> responses = await this.InternalFindFolders(
                new FolderId[] { parentFolderId },
                null, /* searchFilter */
                view,
                ServiceErrorHandling.ThrowOnError,
                token).ConfigureAwait(false);

            return responses[0].Results;
        }

        /// <summary>
        /// Obtains a list of folders by searching the sub-folders of the specified folder.
        /// </summary>
        /// <param name="parentFolderName">The name of the folder in which to search for folders.</param>
        /// <param name="searchFilter">The search filter. Available search filter classes
        /// include SearchFilter.IsEqualTo, SearchFilter.ContainsSubstring and 
        /// SearchFilter.SearchFilterCollection</param>
        /// <param name="view">The view controlling the number of folders returned.</param>
        /// <returns>An object representing the results of the search operation.</returns>
        public Task<FindFoldersResults> FindFolders(WellKnownFolderName parentFolderName, SearchFilter searchFilter, FolderView view)
        {
            return this.FindFolders(new FolderId(parentFolderName), searchFilter, view);
        }

        /// <summary>
        /// Obtains a list of folders by searching the sub-folders of the specified folder.
        /// </summary>
        /// <param name="parentFolderName">The name of the folder in which to search for folders.</param>
        /// <param name="view">The view controlling the number of folders returned.</param>
        /// <returns>An object representing the results of the search operation.</returns>
        public Task<FindFoldersResults> FindFolders(WellKnownFolderName parentFolderName, FolderView view)
        {
            return this.FindFolders(new FolderId(parentFolderName), view);
        }

        /// <summary>
        /// Load specified properties for a folder.
        /// </summary>
        /// <param name="folder">The folder.</param>
        /// <param name="propertySet">The property set.</param>
        internal Task<ServiceResponseCollection<ServiceResponse>> LoadPropertiesForFolder(
            Folder folder,
            PropertySet propertySet,
            CancellationToken token)
        {
            EwsUtilities.ValidateParam(folder, "folder");
            EwsUtilities.ValidateParam(propertySet, "propertySet");

            GetFolderRequestForLoad request = new GetFolderRequestForLoad(this, ServiceErrorHandling.ThrowOnError);

            request.FolderIds.Add(folder);
            request.PropertySet = propertySet;

            return request.ExecuteAsync(token);
        }

        /// <summary>
        /// Binds to a folder.
        /// </summary>
        /// <param name="folderId">The folder id.</param>
        /// <param name="propertySet">The property set.</param>
        /// <returns>Folder</returns>
        internal async Task<Folder> BindToFolder(FolderId folderId, PropertySet propertySet, CancellationToken token)
        {
            EwsUtilities.ValidateParam(folderId, "folderId");
            EwsUtilities.ValidateParam(propertySet, "propertySet");

            ServiceResponseCollection<GetFolderResponse> responses = await this.InternalBindToFolders(
                new[] { folderId },
                propertySet,
                ServiceErrorHandling.ThrowOnError,
                token
            );

            return responses[0].Folder;
        }

        /// <summary>
        /// Binds to folder.
        /// </summary>
        /// <typeparam name="TFolder">The type of the folder.</typeparam>
        /// <param name="folderId">The folder id.</param>
        /// <param name="propertySet">The property set.</param>
        /// <returns>Folder</returns>
        internal async Task<TFolder> BindToFolder<TFolder>(FolderId folderId, PropertySet propertySet, CancellationToken token)
            where TFolder : Folder
        {
            Folder result = await this.BindToFolder(folderId, propertySet, token);

            if (result is TFolder)
            {
                return (TFolder)result;
            }
            else
            {
                throw new ServiceLocalException(
                    string.Format(
                        Strings.FolderTypeNotCompatible,
                        result.GetType().Name,
                        typeof(TFolder).Name));
            }
        }

        /// <summary>
        /// Binds to multiple folders in a single call to EWS.
        /// </summary>
        /// <param name="folderIds">The Ids of the folders to bind to.</param>
        /// <param name="propertySet">The set of properties to load.</param>
        /// <returns>A ServiceResponseCollection providing results for each of the specified folder Ids.</returns>
        public Task<ServiceResponseCollection<GetFolderResponse>> BindToFolders(
            IEnumerable<FolderId> folderIds,
            PropertySet propertySet,
            CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParamCollection(folderIds, "folderIds");
            EwsUtilities.ValidateParam(propertySet, "propertySet");

            return this.InternalBindToFolders(
                folderIds,
                propertySet,
                ServiceErrorHandling.ReturnErrors,
                token
            );
        }

        /// <summary>
        /// Binds to multiple folders in a single call to EWS.
        /// </summary>
        /// <param name="folderIds">The Ids of the folders to bind to.</param>
        /// <param name="propertySet">The set of properties to load.</param>
        /// <param name="errorHandling">Type of error handling to perform.</param>
        /// <returns>A ServiceResponseCollection providing results for each of the specified folder Ids.</returns>
        private Task<ServiceResponseCollection<GetFolderResponse>> InternalBindToFolders(
            IEnumerable<FolderId> folderIds,
            PropertySet propertySet,
            ServiceErrorHandling errorHandling,
            CancellationToken token)
        {
            GetFolderRequest request = new GetFolderRequest(this, errorHandling);

            request.FolderIds.AddRange(folderIds);
            request.PropertySet = propertySet;

            return request.ExecuteAsync(token);
        }

        /// <summary>
        /// Deletes a folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="folderId">The folder id.</param>
        /// <param name="deleteMode">The delete mode.</param>
        internal Task<ServiceResponseCollection<ServiceResponse>> DeleteFolder(
            FolderId folderId,
            DeleteMode deleteMode,
            CancellationToken token)
        {
            EwsUtilities.ValidateParam(folderId, "folderId");

            DeleteFolderRequest request = new DeleteFolderRequest(this, ServiceErrorHandling.ThrowOnError);

            request.FolderIds.Add(folderId);
            request.DeleteMode = deleteMode;

            return request.ExecuteAsync(token);
        }

        /// <summary>
        /// Empties a folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="folderId">The folder id.</param>
        /// <param name="deleteMode">The delete mode.</param>
        /// <param name="deleteSubFolders">if set to <c>true</c> empty folder should also delete sub folders.</param>
        internal Task<ServiceResponseCollection<ServiceResponse>> EmptyFolder(
            FolderId folderId,
            DeleteMode deleteMode,
            bool deleteSubFolders,
            CancellationToken token)
        {
            EwsUtilities.ValidateParam(folderId, "folderId");

            EmptyFolderRequest request = new EmptyFolderRequest(this, ServiceErrorHandling.ThrowOnError);

            request.FolderIds.Add(folderId);
            request.DeleteMode = deleteMode;
            request.DeleteSubFolders = deleteSubFolders;

            return request.ExecuteAsync(token);
        }

        /// <summary>
        /// Marks all items in folder as read/unread. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="folderId">The folder id.</param>
        /// <param name="readFlag">If true, items marked as read, otherwise unread.</param>
        /// <param name="suppressReadReceipts">If true, suppress read receipts for items.</param>
        internal Task<ServiceResponseCollection<ServiceResponse>> MarkAllItemsAsRead(
            FolderId folderId,
            bool readFlag,
            bool suppressReadReceipts,
            CancellationToken token)
        {
            EwsUtilities.ValidateParam(folderId, "folderId");
            EwsUtilities.ValidateMethodVersion(this, ExchangeVersion.Exchange2013, "MarkAllItemsAsRead");

            MarkAllItemsAsReadRequest request = new MarkAllItemsAsReadRequest(this, ServiceErrorHandling.ThrowOnError);

            request.FolderIds.Add(folderId);
            request.ReadFlag = readFlag;
            request.SuppressReadReceipts = suppressReadReceipts;

            return request.ExecuteAsync(token);
        }

        #endregion

        #region Item operations

        /// <summary>
        /// Creates multiple items in a single EWS call. Supported item classes are EmailMessage, Appointment, Contact, PostItem, Task and Item.
        /// CreateItems does not support items that have unsaved attachments.
        /// </summary>
        /// <param name="items">The items to create.</param>
        /// <param name="parentFolderId">The Id of the folder in which to place the newly created items. If null, items are created in their default folders.</param>
        /// <param name="messageDisposition">Indicates the disposition mode for items of type EmailMessage. Required if items contains at least one EmailMessage instance.</param>
        /// <param name="sendInvitationsMode">Indicates if and how invitations should be sent for items of type Appointment. Required if items contains at least one Appointment instance.</param>
        /// <param name="errorHandling">What type of error handling should be performed.</param>
        /// <returns>A ServiceResponseCollection providing creation results for each of the specified items.</returns>
        private Task<ServiceResponseCollection<ServiceResponse>> InternalCreateItems(
            IEnumerable<Item> items,
            FolderId parentFolderId,
            MessageDisposition? messageDisposition,
            SendInvitationsMode? sendInvitationsMode,
            ServiceErrorHandling errorHandling,
            CancellationToken token)
        {
            CreateItemRequest request = new CreateItemRequest(this, errorHandling);

            request.ParentFolderId = parentFolderId;
            request.Items = items;
            request.MessageDisposition = messageDisposition;
            request.SendInvitationsMode = sendInvitationsMode;

            return request.ExecuteAsync(token);
        }

        /// <summary>
        /// Creates multiple items in a single EWS call. Supported item classes are EmailMessage, Appointment, Contact, PostItem, Task and Item.
        /// CreateItems does not support items that have unsaved attachments.
        /// </summary>
        /// <param name="items">The items to create.</param>
        /// <param name="parentFolderId">The Id of the folder in which to place the newly created items. If null, items are created in their default folders.</param>
        /// <param name="messageDisposition">Indicates the disposition mode for items of type EmailMessage. Required if items contains at least one EmailMessage instance.</param>
        /// <param name="sendInvitationsMode">Indicates if and how invitations should be sent for items of type Appointment. Required if items contains at least one Appointment instance.</param>
        /// <returns>A ServiceResponseCollection providing creation results for each of the specified items.</returns>
        public Task<ServiceResponseCollection<ServiceResponse>> CreateItems(
            IEnumerable<Item> items,
            FolderId parentFolderId,
            MessageDisposition? messageDisposition,
            SendInvitationsMode? sendInvitationsMode,
            CancellationToken token = default(CancellationToken))
        {
            // All items have to be new.
            if (!items.TrueForAll((item) => item.IsNew))
            {
                throw new ServiceValidationException(Strings.CreateItemsDoesNotHandleExistingItems);
            }

            // Make sure that all items do *not* have unprocessed attachments.
            if (!items.TrueForAll((item) => !item.HasUnprocessedAttachmentChanges()))
            {
                throw new ServiceValidationException(Strings.CreateItemsDoesNotAllowAttachments);
            }

            return this.InternalCreateItems(
                items,
                parentFolderId,
                messageDisposition,
                sendInvitationsMode,
                ServiceErrorHandling.ReturnErrors,
                token);
        }

        /// <summary>
        /// Creates an item. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="item">The item to create.</param>
        /// <param name="parentFolderId">The Id of the folder in which to place the newly created item. If null, the item is created in its default folders.</param>
        /// <param name="messageDisposition">Indicates the disposition mode for items of type EmailMessage. Required if item is an EmailMessage instance.</param>
        /// <param name="sendInvitationsMode">Indicates if and how invitations should be sent for item of type Appointment. Required if item is an Appointment instance.</param>
        internal System.Threading.Tasks.Task CreateItem(
            Item item,
            FolderId parentFolderId,
            MessageDisposition? messageDisposition,
            SendInvitationsMode? sendInvitationsMode,
            CancellationToken token)
        {
            return this.InternalCreateItems(
                new Item[] { item },
                parentFolderId,
                messageDisposition,
                sendInvitationsMode,
                ServiceErrorHandling.ThrowOnError,
                token);
        }

        /// <summary>
        /// Updates multiple items in a single EWS call. UpdateItems does not support items that have unsaved attachments.
        /// </summary>
        /// <param name="items">The items to update.</param>
        /// <param name="savedItemsDestinationFolderId">The folder in which to save sent messages, meeting invitations or cancellations. If null, the messages, meeting invitation or cancellations are saved in the Sent Items folder.</param>
        /// <param name="conflictResolution">The conflict resolution mode.</param>
        /// <param name="messageDisposition">Indicates the disposition mode for items of type EmailMessage. Required if items contains at least one EmailMessage instance.</param>
        /// <param name="sendInvitationsOrCancellationsMode">Indicates if and how invitations and/or cancellations should be sent for items of type Appointment. Required if items contains at least one Appointment instance.</param>
        /// <param name="errorHandling">What type of error handling should be performed.</param>
        /// <param name="suppressReadReceipt">Whether to suppress read receipts</param>
        /// <returns>A ServiceResponseCollection providing update results for each of the specified items.</returns>
        private Task<ServiceResponseCollection<UpdateItemResponse>> InternalUpdateItems(
            IEnumerable<Item> items,
            FolderId savedItemsDestinationFolderId,
            ConflictResolutionMode conflictResolution,
            MessageDisposition? messageDisposition,
            SendInvitationsOrCancellationsMode? sendInvitationsOrCancellationsMode,
            ServiceErrorHandling errorHandling,
            bool suppressReadReceipt,
            CancellationToken token)
        {
            UpdateItemRequest request = new UpdateItemRequest(this, errorHandling);

            request.Items.AddRange(items);
            request.SavedItemsDestinationFolder = savedItemsDestinationFolderId;
            request.MessageDisposition = messageDisposition;
            request.ConflictResolutionMode = conflictResolution;
            request.SendInvitationsOrCancellationsMode = sendInvitationsOrCancellationsMode;
            request.SuppressReadReceipts = suppressReadReceipt;

            return request.ExecuteAsync(token);
        }

        /// <summary>
        /// Updates multiple items in a single EWS call. UpdateItems does not support items that have unsaved attachments.
        /// </summary>
        /// <param name="items">The items to update.</param>
        /// <param name="savedItemsDestinationFolderId">The folder in which to save sent messages, meeting invitations or cancellations. If null, the messages, meeting invitation or cancellations are saved in the Sent Items folder.</param>
        /// <param name="conflictResolution">The conflict resolution mode.</param>
        /// <param name="messageDisposition">Indicates the disposition mode for items of type EmailMessage. Required if items contains at least one EmailMessage instance.</param>
        /// <param name="sendInvitationsOrCancellationsMode">Indicates if and how invitations and/or cancellations should be sent for items of type Appointment. Required if items contains at least one Appointment instance.</param>
        /// <returns>A ServiceResponseCollection providing update results for each of the specified items.</returns>
        public Task<ServiceResponseCollection<UpdateItemResponse>> UpdateItems(
            IEnumerable<Item> items,
            FolderId savedItemsDestinationFolderId,
            ConflictResolutionMode conflictResolution,
            MessageDisposition? messageDisposition,
            SendInvitationsOrCancellationsMode? sendInvitationsOrCancellationsMode)
        {
            return this.UpdateItems(items, savedItemsDestinationFolderId, conflictResolution, messageDisposition, sendInvitationsOrCancellationsMode, false);
        }

        /// <summary>
        /// Updates multiple items in a single EWS call. UpdateItems does not support items that have unsaved attachments.
        /// </summary>
        /// <param name="items">The items to update.</param>
        /// <param name="savedItemsDestinationFolderId">The folder in which to save sent messages, meeting invitations or cancellations. If null, the messages, meeting invitation or cancellations are saved in the Sent Items folder.</param>
        /// <param name="conflictResolution">The conflict resolution mode.</param>
        /// <param name="messageDisposition">Indicates the disposition mode for items of type EmailMessage. Required if items contains at least one EmailMessage instance.</param>
        /// <param name="sendInvitationsOrCancellationsMode">Indicates if and how invitations and/or cancellations should be sent for items of type Appointment. Required if items contains at least one Appointment instance.</param>
        /// <param name="suppressReadReceipts">Whether to suppress read receipts</param>
        /// <returns>A ServiceResponseCollection providing update results for each of the specified items.</returns>
        public Task<ServiceResponseCollection<UpdateItemResponse>> UpdateItems(
            IEnumerable<Item> items,
            FolderId savedItemsDestinationFolderId,
            ConflictResolutionMode conflictResolution,
            MessageDisposition? messageDisposition,
            SendInvitationsOrCancellationsMode? sendInvitationsOrCancellationsMode,
            bool suppressReadReceipts,
            CancellationToken token = default(CancellationToken))
        {
            // All items have to exist on the server (!new) and modified (dirty)
            if (!items.TrueForAll((item) => (!item.IsNew && item.IsDirty)))
            {
                throw new ServiceValidationException(Strings.UpdateItemsDoesNotSupportNewOrUnchangedItems);
            }

            // Make sure that all items do *not* have unprocessed attachments.
            if (!items.TrueForAll((item) => !item.HasUnprocessedAttachmentChanges()))
            {
                throw new ServiceValidationException(Strings.UpdateItemsDoesNotAllowAttachments);
            }

            return this.InternalUpdateItems(
                items,
                savedItemsDestinationFolderId,
                conflictResolution,
                messageDisposition,
                sendInvitationsOrCancellationsMode,
                ServiceErrorHandling.ReturnErrors,
                suppressReadReceipts,
                token);
        }

        /// <summary>
        /// Updates an item.
        /// </summary>
        /// <param name="item">The item to update.</param>
        /// <param name="savedItemsDestinationFolderId">The folder in which to save sent messages, meeting invitations or cancellations. If null, the message, meeting invitation or cancellation is saved in the Sent Items folder.</param>
        /// <param name="conflictResolution">The conflict resolution mode.</param>
        /// <param name="messageDisposition">Indicates the disposition mode for an item of type EmailMessage. Required if item is an EmailMessage instance.</param>
        /// <param name="sendInvitationsOrCancellationsMode">Indicates if and how invitations and/or cancellations should be sent for ian tem of type Appointment. Required if item is an Appointment instance.</param>
        /// <returns>Updated item.</returns>
        internal Task<Item> UpdateItem(
            Item item,
            FolderId savedItemsDestinationFolderId,
            ConflictResolutionMode conflictResolution,
            MessageDisposition? messageDisposition,
            SendInvitationsOrCancellationsMode? sendInvitationsOrCancellationsMode,
            CancellationToken token)
        {
            return this.UpdateItem(item, savedItemsDestinationFolderId, conflictResolution, messageDisposition, sendInvitationsOrCancellationsMode, false, token);
        }

        /// <summary>
        /// Updates an item.
        /// </summary>
        /// <param name="item">The item to update.</param>
        /// <param name="savedItemsDestinationFolderId">The folder in which to save sent messages, meeting invitations or cancellations. If null, the message, meeting invitation or cancellation is saved in the Sent Items folder.</param>
        /// <param name="conflictResolution">The conflict resolution mode.</param>
        /// <param name="messageDisposition">Indicates the disposition mode for an item of type EmailMessage. Required if item is an EmailMessage instance.</param>
        /// <param name="sendInvitationsOrCancellationsMode">Indicates if and how invitations and/or cancellations should be sent for ian tem of type Appointment. Required if item is an Appointment instance.</param>
        /// <param name="suppressReadReceipts">Whether to suppress read receipts</param>
        /// <returns>Updated item.</returns>
        internal async Task<Item> UpdateItem(
            Item item,
            FolderId savedItemsDestinationFolderId,
            ConflictResolutionMode conflictResolution,
            MessageDisposition? messageDisposition,
            SendInvitationsOrCancellationsMode? sendInvitationsOrCancellationsMode,
            bool suppressReadReceipts,
            CancellationToken token)
        {
            ServiceResponseCollection<UpdateItemResponse> responses = await this.InternalUpdateItems(
                new Item[] { item },
                savedItemsDestinationFolderId,
                conflictResolution,
                messageDisposition,
                sendInvitationsOrCancellationsMode,
                ServiceErrorHandling.ThrowOnError,
                suppressReadReceipts,
                token).ConfigureAwait(false);

            return responses[0].ReturnedItem;
        }

        /// <summary>
        /// Sends an item.
        /// </summary>
        /// <param name="item">The item.</param>
        /// <param name="savedCopyDestinationFolderId">The saved copy destination folder id.</param>
        internal System.Threading.Tasks.Task SendItem(
            Item item,
            FolderId savedCopyDestinationFolderId,
            CancellationToken token)
        {
            SendItemRequest request = new SendItemRequest(this, ServiceErrorHandling.ThrowOnError);

            request.Items = new Item[] { item };
            request.SavedCopyDestinationFolderId = savedCopyDestinationFolderId;

            return request.ExecuteAsync(token);
        }

        /// <summary>
        /// Copies multiple items in a single call to EWS.
        /// </summary>
        /// <param name="itemIds">The Ids of the items to copy.</param>
        /// <param name="destinationFolderId">The Id of the folder to copy the items to.</param>
        /// <param name="returnNewItemIds">Flag indicating whether service should return new ItemIds or not.</param>
        /// <param name="errorHandling">What type of error handling should be performed.</param>
        /// <returns>A ServiceResponseCollection providing copy results for each of the specified item Ids.</returns>
        private Task<ServiceResponseCollection<MoveCopyItemResponse>> InternalCopyItems(
            IEnumerable<ItemId> itemIds,
            FolderId destinationFolderId,
            bool? returnNewItemIds,
            ServiceErrorHandling errorHandling,
            CancellationToken token)
        {
            CopyItemRequest request = new CopyItemRequest(this, errorHandling);
            request.ItemIds.AddRange(itemIds);
            request.DestinationFolderId = destinationFolderId;
            request.ReturnNewItemIds = returnNewItemIds;

            return request.ExecuteAsync(token);
        }

        /// <summary>
        /// Copies multiple items in a single call to EWS.
        /// </summary>
        /// <param name="itemIds">The Ids of the items to copy.</param>
        /// <param name="destinationFolderId">The Id of the folder to copy the items to.</param>
        /// <returns>A ServiceResponseCollection providing copy results for each of the specified item Ids.</returns>
        public Task<ServiceResponseCollection<MoveCopyItemResponse>> CopyItems(
            IEnumerable<ItemId> itemIds,
            FolderId destinationFolderId,
            CancellationToken token = default(CancellationToken))
        {
            return this.InternalCopyItems(
                itemIds,
                destinationFolderId,
                null,
                ServiceErrorHandling.ReturnErrors,
                token);
        }

        /// <summary>
        /// Copies multiple items in a single call to EWS.
        /// </summary>
        /// <param name="itemIds">The Ids of the items to copy.</param>
        /// <param name="destinationFolderId">The Id of the folder to copy the items to.</param>
        /// <param name="returnNewItemIds">Flag indicating whether service should return new ItemIds or not.</param>
        /// <returns>A ServiceResponseCollection providing copy results for each of the specified item Ids.</returns>
        public Task<ServiceResponseCollection<MoveCopyItemResponse>> CopyItems(
            IEnumerable<ItemId> itemIds,
            FolderId destinationFolderId,
            bool returnNewItemIds,
            CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateMethodVersion(
                this,
                ExchangeVersion.Exchange2010_SP1,
                "CopyItems");

            return this.InternalCopyItems(
                itemIds,
                destinationFolderId,
                returnNewItemIds,
                ServiceErrorHandling.ReturnErrors,
                token);
        }

        /// <summary>
        /// Copies an item. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="itemId">The Id of the item to copy.</param>
        /// <param name="destinationFolderId">The Id of the folder to copy the item to.</param>
        /// <returns>The copy of the item.</returns>
        internal async Task<Item> CopyItem(
            ItemId itemId,
            FolderId destinationFolderId,
            CancellationToken token)
        {
            return (await this.InternalCopyItems(
                new ItemId[] { itemId },
                destinationFolderId,
                null,
                ServiceErrorHandling.ThrowOnError,
                token).ConfigureAwait(false))[0].Item;
        }

        /// <summary>
        /// Moves multiple items in a single call to EWS.
        /// </summary>
        /// <param name="itemIds">The Ids of the items to move.</param>
        /// <param name="destinationFolderId">The Id of the folder to move the items to.</param>
        /// <param name="returnNewItemIds">Flag indicating whether service should return new ItemIds or not.</param>
        /// <param name="errorHandling">What type of error handling should be performed.</param>
        /// <returns>A ServiceResponseCollection providing copy results for each of the specified item Ids.</returns>
        private Task<ServiceResponseCollection<MoveCopyItemResponse>> InternalMoveItems(
            IEnumerable<ItemId> itemIds,
            FolderId destinationFolderId,
            bool? returnNewItemIds,
            ServiceErrorHandling errorHandling,
            CancellationToken token)
        {
            MoveItemRequest request = new MoveItemRequest(this, errorHandling);

            request.ItemIds.AddRange(itemIds);
            request.DestinationFolderId = destinationFolderId;
            request.ReturnNewItemIds = returnNewItemIds;

            return request.ExecuteAsync(token);
        }

        /// <summary>
        /// Moves multiple items in a single call to EWS.
        /// </summary>
        /// <param name="itemIds">The Ids of the items to move.</param>
        /// <param name="destinationFolderId">The Id of the folder to move the items to.</param>
        /// <returns>A ServiceResponseCollection providing copy results for each of the specified item Ids.</returns>
        public Task<ServiceResponseCollection<MoveCopyItemResponse>> MoveItems(
            IEnumerable<ItemId> itemIds,
            FolderId destinationFolderId,
            CancellationToken token = default(CancellationToken))
        {
            return this.InternalMoveItems(
                itemIds,
                destinationFolderId,
                null,
                ServiceErrorHandling.ReturnErrors,
                token);
        }

        /// <summary>
        /// Moves multiple items in a single call to EWS.
        /// </summary>
        /// <param name="itemIds">The Ids of the items to move.</param>
        /// <param name="destinationFolderId">The Id of the folder to move the items to.</param>
        /// <param name="returnNewItemIds">Flag indicating whether service should return new ItemIds or not.</param>
        /// <returns>A ServiceResponseCollection providing copy results for each of the specified item Ids.</returns>
        public Task<ServiceResponseCollection<MoveCopyItemResponse>> MoveItems(
            IEnumerable<ItemId> itemIds,
            FolderId destinationFolderId,
            bool returnNewItemIds,
            CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateMethodVersion(
                this,
                ExchangeVersion.Exchange2010_SP1,
                "MoveItems");

            return this.InternalMoveItems(
                itemIds,
                destinationFolderId,
                returnNewItemIds,
                ServiceErrorHandling.ReturnErrors,
                token);
        }

        /// <summary>
        /// Move an item.
        /// </summary>
        /// <param name="itemId">The Id of the item to move.</param>
        /// <param name="destinationFolderId">The Id of the folder to move the item to.</param>
        /// <returns>The moved item.</returns>
        internal async Task<Item> MoveItem(
            ItemId itemId,
            FolderId destinationFolderId,
            CancellationToken token)
        {
            return (await this.InternalMoveItems(
                new ItemId[] { itemId },
                destinationFolderId,
                null,
                ServiceErrorHandling.ThrowOnError,
                token).ConfigureAwait(false))[0].Item;
        }

        /// <summary>
        /// Archives multiple items in a single call to EWS.
        /// </summary>
        /// <param name="itemIds">The Ids of the items to move.</param>
        /// <param name="sourceFolderId">The Id of the folder in primary corresponding to which items are being archived to.</param>
        /// <returns>A ServiceResponseCollection providing copy results for each of the specified item Ids.</returns>
        public Task<ServiceResponseCollection<ArchiveItemResponse>> ArchiveItems(
            IEnumerable<ItemId> itemIds,
            FolderId sourceFolderId,
            CancellationToken token = default(CancellationToken))
        {
            ArchiveItemRequest request = new ArchiveItemRequest(this, ServiceErrorHandling.ReturnErrors);

            request.Ids.AddRange(itemIds);
            request.SourceFolderId = sourceFolderId;

            return request.ExecuteAsync(token);
        }

        /// <summary>
        /// Finds items.
        /// </summary>
        /// <typeparam name="TItem">The type of the item.</typeparam>
        /// <param name="parentFolderIds">The parent folder ids.</param>
        /// <param name="searchFilter">The search filter. Available search filter classes
        /// include SearchFilter.IsEqualTo, SearchFilter.ContainsSubstring and 
        /// SearchFilter.SearchFilterCollection</param>
        /// <param name="queryString">query string to be used for indexed search.</param>
        /// <param name="view">The view controlling the number of items returned.</param>
        /// <param name="groupBy">The group by.</param>
        /// <param name="errorHandlingMode">Indicates the type of error handling should be done.</param>
        /// <returns>Service response collection.</returns>
        internal Task<ServiceResponseCollection<FindItemResponse<TItem>>> FindItems<TItem>(
            IEnumerable<FolderId> parentFolderIds,
            SearchFilter searchFilter,
            string queryString,
            ViewBase view,
            Grouping groupBy,
            ServiceErrorHandling errorHandlingMode,
            CancellationToken token)
            where TItem : Item
        {
            EwsUtilities.ValidateParamCollection(parentFolderIds, "parentFolderIds");
            EwsUtilities.ValidateParam(view, "view");
            EwsUtilities.ValidateParamAllowNull(groupBy, "groupBy");
            EwsUtilities.ValidateParamAllowNull(queryString, "queryString");
            EwsUtilities.ValidateParamAllowNull(searchFilter, "searchFilter");

            FindItemRequest<TItem> request = new FindItemRequest<TItem>(this, errorHandlingMode);

            request.ParentFolderIds.AddRange(parentFolderIds);
            request.SearchFilter = searchFilter;
            request.QueryString = queryString;
            request.View = view;
            request.GroupBy = groupBy;

            return request.ExecuteAsync(token);
        }

        /// <summary>
        /// Obtains a list of items by searching the contents of a specific folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="parentFolderId">The Id of the folder in which to search for items.</param>
        /// <param name="queryString">the search string to be used for indexed search, if any.</param>
        /// <param name="view">The view controlling the number of items returned.</param>
        /// <returns>An object representing the results of the search operation.</returns>
        public async Task<FindItemsResults<Item>> FindItems(FolderId parentFolderId, string queryString, ViewBase view, CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParamAllowNull(queryString, "queryString");

            ServiceResponseCollection<FindItemResponse<Item>> responses = await this.FindItems<Item>(
                new FolderId[] { parentFolderId },
                null, /* searchFilter */
                queryString,
                view,
                null,   /* groupBy */
                ServiceErrorHandling.ThrowOnError,
                token).ConfigureAwait(false);

            return responses[0].Results;
        }

        /// <summary>
        /// Obtains a list of items by searching the contents of a specific folder. 
        /// Along with conversations, a list of highlight terms are returned.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="parentFolderId">The Id of the folder in which to search for items.</param>
        /// <param name="queryString">the search string to be used for indexed search, if any.</param>
        /// <param name="returnHighlightTerms">Flag indicating if highlight terms should be returned in the response</param>
        /// <param name="view">The view controlling the number of items returned.</param>
        /// <returns>An object representing the results of the search operation.</returns>
        public async Task<FindItemsResults<Item>> FindItems(FolderId parentFolderId, string queryString, bool returnHighlightTerms, ViewBase view,
            CancellationToken token = default(CancellationToken))
        {
            FolderId[] parentFolderIds = new FolderId[] { parentFolderId };

            EwsUtilities.ValidateParamCollection(parentFolderIds, "parentFolderIds");
            EwsUtilities.ValidateParam(view, "view");
            EwsUtilities.ValidateParamAllowNull(queryString, "queryString");
            EwsUtilities.ValidateParamAllowNull(returnHighlightTerms, "returnHighlightTerms");
            EwsUtilities.ValidateMethodVersion(this, ExchangeVersion.Exchange2013, "FindItems");

            FindItemRequest<Item> request = new FindItemRequest<Item>(this, ServiceErrorHandling.ThrowOnError);

            request.ParentFolderIds.AddRange(parentFolderIds);
            request.QueryString = queryString;
            request.ReturnHighlightTerms = returnHighlightTerms;
            request.View = view;

            ServiceResponseCollection<FindItemResponse<Item>> responses = await request.ExecuteAsync(token).ConfigureAwait(false);
            return responses[0].Results;
        }

        /// <summary>
        /// Obtains a list of items by searching the contents of a specific folder. 
        /// Along with conversations, a list of highlight terms are returned.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="parentFolderId">The Id of the folder in which to search for items.</param>
        /// <param name="queryString">the search string to be used for indexed search, if any.</param>
        /// <param name="returnHighlightTerms">Flag indicating if highlight terms should be returned in the response</param>
        /// <param name="view">The view controlling the number of items returned.</param>
        /// <param name="groupBy">The group by clause.</param>
        /// <returns>An object representing the results of the search operation.</returns>
        public async Task<GroupedFindItemsResults<Item>> FindItems(FolderId parentFolderId, string queryString, bool returnHighlightTerms, ViewBase view, Grouping groupBy, CancellationToken token = default(CancellationToken))
        {
            FolderId[] parentFolderIds = new FolderId[] { parentFolderId };

            EwsUtilities.ValidateParamCollection(parentFolderIds, "parentFolderIds");
            EwsUtilities.ValidateParam(view, "view");
            EwsUtilities.ValidateParam(groupBy, "groupBy");
            EwsUtilities.ValidateParamAllowNull(queryString, "queryString");
            EwsUtilities.ValidateParamAllowNull(returnHighlightTerms, "returnHighlightTerms");
            EwsUtilities.ValidateMethodVersion(this, ExchangeVersion.Exchange2013, "FindItems");

            FindItemRequest<Item> request = new FindItemRequest<Item>(this, ServiceErrorHandling.ThrowOnError);

            request.ParentFolderIds.AddRange(parentFolderIds);
            request.QueryString = queryString;
            request.ReturnHighlightTerms = returnHighlightTerms;
            request.View = view;
            request.GroupBy = groupBy;

            ServiceResponseCollection<FindItemResponse<Item>> responses = await request.ExecuteAsync(token).ConfigureAwait(false);
            return responses[0].GroupedFindResults;
        }

        /// <summary>
        /// Obtains a list of items by searching the contents of a specific folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="parentFolderId">The Id of the folder in which to search for items.</param>
        /// <param name="searchFilter">The search filter. Available search filter classes
        /// include SearchFilter.IsEqualTo, SearchFilter.ContainsSubstring and 
        /// SearchFilter.SearchFilterCollection</param>
        /// <param name="view">The view controlling the number of items returned.</param>
        /// <returns>An object representing the results of the search operation.</returns>
        public async Task<FindItemsResults<Item>> FindItems(FolderId parentFolderId, SearchFilter searchFilter, ViewBase view, CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParamAllowNull(searchFilter, "searchFilter");

            ServiceResponseCollection<FindItemResponse<Item>> responses = await this.FindItems<Item>(
                new FolderId[] { parentFolderId },
                searchFilter,
                null, /* queryString */
                view,
                null,   /* groupBy */
                ServiceErrorHandling.ThrowOnError,
                token).ConfigureAwait(false);

            return responses[0].Results;
        }

        /// <summary>
        /// Obtains a list of items by searching the contents of a specific folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="parentFolderId">The Id of the folder in which to search for items.</param>
        /// <param name="view">The view controlling the number of items returned.</param>
        /// <returns>An object representing the results of the search operation.</returns>
        public async Task<FindItemsResults<Item>> FindItems(FolderId parentFolderId, ViewBase view, CancellationToken token = default(CancellationToken))
        {
            ServiceResponseCollection<FindItemResponse<Item>> responses = await this.FindItems<Item>(
                new FolderId[] { parentFolderId },
                null, /* searchFilter */
                null, /* queryString */
                view,
                null, /* groupBy */
                ServiceErrorHandling.ThrowOnError,
                token).ConfigureAwait(false);

            return responses[0].Results;
        }

        /// <summary>
        /// Obtains a list of items by searching the contents of a specific folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="parentFolderName">The name of the folder in which to search for items.</param>
        /// <param name="queryString">query string to be used for indexed search</param>
        /// <param name="view">The view controlling the number of items returned.</param>
        /// <returns>An object representing the results of the search operation.</returns>
        public Task<FindItemsResults<Item>> FindItems(WellKnownFolderName parentFolderName, string queryString, ViewBase view)
        {
            return this.FindItems(new FolderId(parentFolderName), queryString, view);
        }

        /// <summary>
        /// Obtains a list of items by searching the contents of a specific folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="parentFolderName">The name of the folder in which to search for items.</param>
        /// <param name="searchFilter">The search filter. Available search filter classes
        /// include SearchFilter.IsEqualTo, SearchFilter.ContainsSubstring and 
        /// SearchFilter.SearchFilterCollection</param>
        /// <param name="view">The view controlling the number of items returned.</param>
        /// <returns>An object representing the results of the search operation.</returns>
        public Task<FindItemsResults<Item>> FindItems(WellKnownFolderName parentFolderName, SearchFilter searchFilter, ViewBase view)
        {
            return this.FindItems(
                new FolderId(parentFolderName),
                searchFilter,
                view);
        }

        /// <summary>
        /// Obtains a list of items by searching the contents of a specific folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="parentFolderName">The name of the folder in which to search for items.</param>
        /// <param name="view">The view controlling the number of items returned.</param>
        /// <returns>An object representing the results of the search operation.</returns>
        public Task<FindItemsResults<Item>> FindItems(WellKnownFolderName parentFolderName, ViewBase view, CancellationToken token = default(CancellationToken))
        {
            return this.FindItems(
                new FolderId(parentFolderName),
                (SearchFilter)null,
                view,
                token);
        }

        /// <summary>
        /// Obtains a grouped list of items by searching the contents of a specific folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="parentFolderId">The Id of the folder in which to search for items.</param>
        /// <param name="queryString">query string to be used for indexed search</param>
        /// <param name="view">The view controlling the number of items returned.</param>
        /// <param name="groupBy">The group by clause.</param>
        /// <returns>A list of items containing the contents of the specified folder.</returns>
        public async Task<GroupedFindItemsResults<Item>> FindItems(
            FolderId parentFolderId,
            string queryString,
            ViewBase view,
            Grouping groupBy,
            CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParam(groupBy, "groupBy");
            EwsUtilities.ValidateParamAllowNull(queryString, "queryString");

            ServiceResponseCollection<FindItemResponse<Item>> responses = await this.FindItems<Item>(
                new FolderId[] { parentFolderId },
                null, /* searchFilter */
                queryString,
                view,
                groupBy,
                ServiceErrorHandling.ThrowOnError,
                token).ConfigureAwait(false);

            return responses[0].GroupedFindResults;
        }

        /// <summary>
        /// Obtains a grouped list of items by searching the contents of a specific folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="parentFolderId">The Id of the folder in which to search for items.</param>
        /// <param name="searchFilter">The search filter. Available search filter classes
        /// include SearchFilter.IsEqualTo, SearchFilter.ContainsSubstring and 
        /// SearchFilter.SearchFilterCollection</param>
        /// <param name="view">The view controlling the number of items returned.</param>
        /// <param name="groupBy">The group by clause.</param>
        /// <returns>A list of items containing the contents of the specified folder.</returns>
        public async Task<GroupedFindItemsResults<Item>> FindItems(
            FolderId parentFolderId,
            SearchFilter searchFilter,
            ViewBase view,
            Grouping groupBy,
            CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParam(groupBy, "groupBy");
            EwsUtilities.ValidateParamAllowNull(searchFilter, "searchFilter");

            ServiceResponseCollection<FindItemResponse<Item>> responses = await this.FindItems<Item>(
                new FolderId[] { parentFolderId },
                searchFilter,
                null, /* queryString */
                view,
                groupBy,
                ServiceErrorHandling.ThrowOnError,
                token).ConfigureAwait(false);

            return responses[0].GroupedFindResults;
        }

        /// <summary>
        /// Obtains a grouped list of items by searching the contents of a specific folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="parentFolderId">The Id of the folder in which to search for items.</param>
        /// <param name="view">The view controlling the number of items returned.</param>
        /// <param name="groupBy">The group by clause.</param>
        /// <returns>A list of items containing the contents of the specified folder.</returns>
        public async Task<GroupedFindItemsResults<Item>> FindItems(
            FolderId parentFolderId,
            ViewBase view,
            Grouping groupBy,
            CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParam(groupBy, "groupBy");

            ServiceResponseCollection<FindItemResponse<Item>> responses = await this.FindItems<Item>(
                new FolderId[] { parentFolderId },
                null, /* searchFilter */
                null, /* queryString */
                view,
                groupBy,
                ServiceErrorHandling.ThrowOnError,
                token).ConfigureAwait(false);

            return responses[0].GroupedFindResults;
        }

        /// <summary>
        /// Obtains a grouped list of items by searching the contents of a specific folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="parentFolderId">The Id of the folder in which to search for items.</param>
        /// <param name="searchFilter">The search filter. Available search filter classes
        /// include SearchFilter.IsEqualTo, SearchFilter.ContainsSubstring and 
        /// SearchFilter.SearchFilterCollection</param>
        /// <param name="view">The view controlling the number of items returned.</param>
        /// <param name="groupBy">The group by clause.</param>
        /// <typeparam name="TItem">Type of item.</typeparam>
        /// <returns>A list of items containing the contents of the specified folder.</returns>
        internal Task<ServiceResponseCollection<FindItemResponse<TItem>>> FindItems<TItem>(
            FolderId parentFolderId,
            SearchFilter searchFilter,
            ViewBase view,
            Grouping groupBy,
            CancellationToken token)
            where TItem : Item
        {
            return this.FindItems<TItem>(
                new FolderId[] { parentFolderId },
                searchFilter,
                null, /* queryString */
                view,
                groupBy,
                ServiceErrorHandling.ThrowOnError,
                token);
        }

        /// <summary>
        /// Obtains a grouped list of items by searching the contents of a specific folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="parentFolderName">The name of the folder in which to search for items.</param>
        /// <param name="queryString">query string to be used for indexed search</param>
        /// <param name="view">The view controlling the number of items returned.</param>
        /// <param name="groupBy">The group by clause.</param>
        /// <returns>A collection of grouped items representing the contents of the specified.</returns>
        public Task<GroupedFindItemsResults<Item>> FindItems(
            WellKnownFolderName parentFolderName,
            string queryString,
            ViewBase view,
            Grouping groupBy,
            CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParam(groupBy, "groupBy");

            return this.FindItems(
                new FolderId(parentFolderName),
                queryString,
                view,
                groupBy,
                token);
        }

        /// <summary>
        /// Obtains a grouped list of items by searching the contents of a specific folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="parentFolderName">The name of the folder in which to search for items.</param>
        /// <param name="searchFilter">The search filter. Available search filter classes
        /// include SearchFilter.IsEqualTo, SearchFilter.ContainsSubstring and 
        /// SearchFilter.SearchFilterCollection</param>
        /// <param name="view">The view controlling the number of items returned.</param>
        /// <param name="groupBy">The group by clause.</param>
        /// <returns>A collection of grouped items representing the contents of the specified.</returns>
        public Task<GroupedFindItemsResults<Item>> FindItems(
            WellKnownFolderName parentFolderName,
            SearchFilter searchFilter,
            ViewBase view,
            Grouping groupBy,
            CancellationToken token = default(CancellationToken))
        {
            return this.FindItems(
                new FolderId(parentFolderName),
                searchFilter,
                view,
                groupBy,
                token);
        }

        /// <summary>
        /// Obtains a list of appointments by searching the contents of a specific folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="parentFolderId">The id of the calendar folder in which to search for items.</param>
        /// <param name="calendarView">The calendar view controlling the number of appointments returned.</param>
        /// <returns>A collection of appointments representing the contents of the specified folder.</returns>
        public async Task<FindItemsResults<Appointment>> FindAppointments(FolderId parentFolderId, CalendarView calendarView, CancellationToken token = default(CancellationToken))
        {
            ServiceResponseCollection<FindItemResponse<Appointment>> response = await this.FindItems<Appointment>(
                new FolderId[] { parentFolderId },
                null, /* searchFilter */
                null, /* queryString */
                calendarView,
                null, /* groupBy */
                ServiceErrorHandling.ThrowOnError,
                token).ConfigureAwait(false);

            return response[0].Results;
        }

        /// <summary>
        /// Obtains a list of appointments by searching the contents of a specific folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="parentFolderName">The name of the calendar folder in which to search for items.</param>
        /// <param name="calendarView">The calendar view controlling the number of appointments returned.</param>
        /// <returns>A collection of appointments representing the contents of the specified folder.</returns>
        public Task<FindItemsResults<Appointment>> FindAppointments(WellKnownFolderName parentFolderName, CalendarView calendarView, CancellationToken token = default(CancellationToken))
        {
            return this.FindAppointments(new FolderId(parentFolderName), calendarView, token);
        }

        /// <summary>
        /// Loads the properties of multiple items in a single call to EWS.
        /// </summary>
        /// <param name="items">The items to load the properties of.</param>
        /// <param name="propertySet">The set of properties to load.</param>
        /// <returns>A ServiceResponseCollection providing results for each of the specified items.</returns>
        public Task<ServiceResponseCollection<ServiceResponse>> LoadPropertiesForItems(IEnumerable<Item> items, PropertySet propertySet, CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParamCollection(items, "items");
            EwsUtilities.ValidateParam(propertySet, "propertySet");

            return this.InternalLoadPropertiesForItems(
                items,
                propertySet,
                ServiceErrorHandling.ReturnErrors,
                token);
        }

        /// <summary>
        /// Loads the properties of multiple items in a single call to EWS.
        /// </summary>
        /// <param name="items">The items to load the properties of.</param>
        /// <param name="propertySet">The set of properties to load.</param>
        /// <param name="errorHandling">Indicates the type of error handling should be done.</param>
        /// <returns>A ServiceResponseCollection providing results for each of the specified items.</returns>
        internal Task<ServiceResponseCollection<ServiceResponse>> InternalLoadPropertiesForItems(
            IEnumerable<Item> items,
            PropertySet propertySet,
            ServiceErrorHandling errorHandling,
            CancellationToken token)
        {
            GetItemRequestForLoad request = new GetItemRequestForLoad(this, errorHandling);

            request.ItemIds.AddRange(items);
            request.PropertySet = propertySet;

            return request.ExecuteAsync(token);
        }

        /// <summary>
        /// Binds to multiple items in a single call to EWS.
        /// </summary>
        /// <param name="itemIds">The Ids of the items to bind to.</param>
        /// <param name="propertySet">The set of properties to load.</param>
        /// <param name="anchorMailbox">The SmtpAddress of mailbox that hosts all items we need to bind to</param>
        /// <param name="errorHandling">Type of error handling to perform.</param>
        /// <returns>A ServiceResponseCollection providing results for each of the specified item Ids.</returns>
        private Task<ServiceResponseCollection<GetItemResponse>> InternalBindToItems(
            IEnumerable<ItemId> itemIds,
            PropertySet propertySet,
            string anchorMailbox,
            ServiceErrorHandling errorHandling,
            CancellationToken token)
        {
            GetItemRequest request = new GetItemRequest(this, errorHandling);

            request.ItemIds.AddRange(itemIds);
            request.PropertySet = propertySet;
            request.AnchorMailbox = anchorMailbox;

            return request.ExecuteAsync(token);
        }

        /// <summary>
        /// Binds to multiple items in a single call to EWS.
        /// </summary>
        /// <param name="itemIds">The Ids of the items to bind to.</param>
        /// <param name="propertySet">The set of properties to load.</param>
        /// <returns>A ServiceResponseCollection providing results for each of the specified item Ids.</returns>
        public Task<ServiceResponseCollection<GetItemResponse>> BindToItems(IEnumerable<ItemId> itemIds, PropertySet propertySet, CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParamCollection(itemIds, "itemIds");
            EwsUtilities.ValidateParam(propertySet, "propertySet");

            return this.InternalBindToItems(
                itemIds,
                propertySet,
                null, /* anchorMailbox */
                ServiceErrorHandling.ReturnErrors,
                token);
        }

        /// <summary>
        /// Binds to multiple items in a single call to EWS.
        /// </summary>
        /// <param name="itemIds">The Ids of the items to bind to.</param>
        /// <param name="propertySet">The set of properties to load.</param>
        /// <param name="anchorMailbox">The SmtpAddress of mailbox that hosts all items we need to bind to</param>
        /// <returns>A ServiceResponseCollection providing results for each of the specified item Ids.</returns>
        /// <remarks>
        /// This API designed to be used primarily in groups scenarios where we want to set the
        /// anchor mailbox header so that request is routed directly to the group mailbox backend server.
        /// </remarks>
        public Task<ServiceResponseCollection<GetItemResponse>> BindToGroupItems(
            IEnumerable<ItemId> itemIds,
            PropertySet propertySet,
            string anchorMailbox,
            CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParamCollection(itemIds, "itemIds");
            EwsUtilities.ValidateParam(propertySet, "propertySet");
            EwsUtilities.ValidateParam(propertySet, "anchorMailbox");

            return this.InternalBindToItems(
                itemIds,
                propertySet,
                anchorMailbox,
                ServiceErrorHandling.ReturnErrors,
                token);
        }

        /// <summary>
        /// Binds to item.
        /// </summary>
        /// <param name="itemId">The item id.</param>
        /// <param name="propertySet">The property set.</param>
        /// <returns>Item.</returns>
        internal async Task<Item> BindToItem(ItemId itemId, PropertySet propertySet, CancellationToken token)
        {
            EwsUtilities.ValidateParam(itemId, "itemId");
            EwsUtilities.ValidateParam(propertySet, "propertySet");

            ServiceResponseCollection<GetItemResponse> responses = await this.InternalBindToItems(
                new ItemId[] { itemId },
                propertySet,
                null, /* anchorMailbox */
                ServiceErrorHandling.ThrowOnError,
                token).ConfigureAwait(false);

            return responses[0].Item;
        }

        /// <summary>
        /// Binds to item.
        /// </summary>
        /// <typeparam name="TItem">The type of the item.</typeparam>
        /// <param name="itemId">The item id.</param>
        /// <param name="propertySet">The property set.</param>
        /// <returns>Item</returns>
        internal async Task<TItem> BindToItem<TItem>(ItemId itemId, PropertySet propertySet, CancellationToken token)
            where TItem : Item
        {
            Item result = await this.BindToItem(itemId, propertySet, token).ConfigureAwait(false);

            if (result is TItem)
            {
                return (TItem)result;
            }
            else
            {
                throw new ServiceLocalException(
                    string.Format(
                        Strings.ItemTypeNotCompatible,
                        result.GetType().Name,
                        typeof(TItem).Name));
            }
        }

        /// <summary>
        /// Deletes multiple items in a single call to EWS.
        /// </summary>
        /// <param name="itemIds">The Ids of the items to delete.</param>
        /// <param name="deleteMode">The deletion mode.</param>
        /// <param name="sendCancellationsMode">Indicates whether cancellation messages should be sent. Required if any of the item Ids represents an Appointment.</param>
        /// <param name="affectedTaskOccurrences">Indicates which instance of a recurring task should be deleted. Required if any of the item Ids represents a Task.</param>
        /// <param name="errorHandling">Type of error handling to perform.</param>
        /// <param name="suppressReadReceipts">Whether to suppress read receipts</param>
        /// <returns>A ServiceResponseCollection providing deletion results for each of the specified item Ids.</returns>
        private Task<ServiceResponseCollection<ServiceResponse>> InternalDeleteItems(
            IEnumerable<ItemId> itemIds,
            DeleteMode deleteMode,
            SendCancellationsMode? sendCancellationsMode,
            AffectedTaskOccurrence? affectedTaskOccurrences,
            ServiceErrorHandling errorHandling,
            bool suppressReadReceipts,
            CancellationToken token)
        {
            DeleteItemRequest request = new DeleteItemRequest(this, errorHandling);

            request.ItemIds.AddRange(itemIds);
            request.DeleteMode = deleteMode;
            request.SendCancellationsMode = sendCancellationsMode;
            request.AffectedTaskOccurrences = affectedTaskOccurrences;
            request.SuppressReadReceipts = suppressReadReceipts;

            return request.ExecuteAsync(token);
        }

        /// <summary>
        /// Deletes multiple items in a single call to EWS.
        /// </summary>
        /// <param name="itemIds">The Ids of the items to delete.</param>
        /// <param name="deleteMode">The deletion mode.</param>
        /// <param name="sendCancellationsMode">Indicates whether cancellation messages should be sent. Required if any of the item Ids represents an Appointment.</param>
        /// <param name="affectedTaskOccurrences">Indicates which instance of a recurring task should be deleted. Required if any of the item Ids represents a Task.</param>
        /// <returns>A ServiceResponseCollection providing deletion results for each of the specified item Ids.</returns>
        public Task<ServiceResponseCollection<ServiceResponse>> DeleteItems(
            IEnumerable<ItemId> itemIds,
            DeleteMode deleteMode,
            SendCancellationsMode? sendCancellationsMode,
            AffectedTaskOccurrence? affectedTaskOccurrences)
        {
            return this.DeleteItems(itemIds, deleteMode, sendCancellationsMode, affectedTaskOccurrences, false);
        }

        /// <summary>
        /// Deletes multiple items in a single call to EWS.
        /// </summary>
        /// <param name="itemIds">The Ids of the items to delete.</param>
        /// <param name="deleteMode">The deletion mode.</param>
        /// <param name="sendCancellationsMode">Indicates whether cancellation messages should be sent. Required if any of the item Ids represents an Appointment.</param>
        /// <param name="affectedTaskOccurrences">Indicates which instance of a recurring task should be deleted. Required if any of the item Ids represents a Task.</param>
        /// <returns>A ServiceResponseCollection providing deletion results for each of the specified item Ids.</returns>
        /// <param name="suppressReadReceipt">Whether to suppress read receipts</param>
        public Task<ServiceResponseCollection<ServiceResponse>> DeleteItems(
            IEnumerable<ItemId> itemIds,
            DeleteMode deleteMode,
            SendCancellationsMode? sendCancellationsMode,
            AffectedTaskOccurrence? affectedTaskOccurrences,
            bool suppressReadReceipt,
            CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParamCollection(itemIds, "itemIds");

            return this.InternalDeleteItems(
                itemIds,
                deleteMode,
                sendCancellationsMode,
                affectedTaskOccurrences,
                ServiceErrorHandling.ReturnErrors,
                suppressReadReceipt,
                token);
        }

        /// <summary>
        /// Deletes an item. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="itemId">The Id of the item to delete.</param>
        /// <param name="deleteMode">The deletion mode.</param>
        /// <param name="sendCancellationsMode">Indicates whether cancellation messages should be sent. Required if the item Id represents an Appointment.</param>
        /// <param name="affectedTaskOccurrences">Indicates which instance of a recurring task should be deleted. Required if item Id represents a Task.</param>
        internal Task<ServiceResponseCollection<ServiceResponse>> DeleteItem(
            ItemId itemId,
            DeleteMode deleteMode,
            SendCancellationsMode? sendCancellationsMode,
            AffectedTaskOccurrence? affectedTaskOccurrences,
            CancellationToken token)
        {
            return this.DeleteItem(itemId, deleteMode, sendCancellationsMode, affectedTaskOccurrences, false, token);
        }

        /// <summary>
        /// Deletes an item. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="itemId">The Id of the item to delete.</param>
        /// <param name="deleteMode">The deletion mode.</param>
        /// <param name="sendCancellationsMode">Indicates whether cancellation messages should be sent. Required if the item Id represents an Appointment.</param>
        /// <param name="affectedTaskOccurrences">Indicates which instance of a recurring task should be deleted. Required if item Id represents a Task.</param>
        /// <param name="suppressReadReceipts">Whether to suppress read receipts</param>
        internal Task<ServiceResponseCollection<ServiceResponse>> DeleteItem(
            ItemId itemId,
            DeleteMode deleteMode,
            SendCancellationsMode? sendCancellationsMode,
            AffectedTaskOccurrence? affectedTaskOccurrences,
            bool suppressReadReceipts,
            CancellationToken token)
        {
            EwsUtilities.ValidateParam(itemId, "itemId");

            return this.InternalDeleteItems(
                new ItemId[] { itemId },
                deleteMode,
                sendCancellationsMode,
                affectedTaskOccurrences,
                ServiceErrorHandling.ThrowOnError,
                suppressReadReceipts,
                token);
        }

        /// <summary>
        /// Mark items as junk.
        /// </summary>
        /// <param name="itemIds">ItemIds for the items to mark</param>
        /// <param name="isJunk">Whether the items are junk.  If true, senders are add to blocked sender list. If false, senders are removed.</param>
        /// <param name="moveItem">Whether to move the item.  Items are moved to junk folder if isJunk is true, inbox if isJunk is false.</param>
        /// <returns>A ServiceResponseCollection providing itemIds for each of the moved items..</returns>
        public Task<ServiceResponseCollection<MarkAsJunkResponse>> MarkAsJunk(IEnumerable<ItemId> itemIds, bool isJunk, bool moveItem,
            CancellationToken token = default(CancellationToken))
        {
            MarkAsJunkRequest request = new MarkAsJunkRequest(this, ServiceErrorHandling.ReturnErrors);
            request.ItemIds.AddRange(itemIds);
            request.IsJunk = isJunk;
            request.MoveItem = moveItem;
            return request.ExecuteAsync(token);
        }

        #endregion

        #region People operations

        /// <summary>
        /// This method is for search scenarios. Retrieves a set of personas satisfying the specified search conditions.
        /// </summary>
        /// <param name="folderId">Id of the folder being searched</param>
        /// <param name="searchFilter">The search filter. Available search filter classes
        /// include SearchFilter.IsEqualTo, SearchFilter.ContainsSubstring and 
        /// SearchFilter.SearchFilterCollection</param>
        /// <param name="view">The view which defines the number of persona being returned</param>
        /// <param name="queryString">The query string for which the search is being performed</param>
        /// <returns>A collection of personas matching the search conditions</returns>
        public async Task<ICollection<Persona>> FindPeople(FolderId folderId, SearchFilter searchFilter, ViewBase view, string queryString,
            CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParamAllowNull(folderId, "folderId");
            EwsUtilities.ValidateParamAllowNull(searchFilter, "searchFilter");
            EwsUtilities.ValidateParam(view, "view");
            EwsUtilities.ValidateParam(queryString, "queryString");
            EwsUtilities.ValidateMethodVersion(this, ExchangeVersion.Exchange2013_SP1, "FindPeople");

            FindPeopleRequest request = new FindPeopleRequest(this);

            request.FolderId = folderId;
            request.SearchFilter = searchFilter;
            request.View = view;
            request.QueryString = queryString;

            return (await request.Execute(token).ConfigureAwait(false)).Personas;
        }

        /// <summary>
        /// This method is for search scenarios. Retrieves a set of personas satisfying the specified search conditions.
        /// </summary>
        /// <param name="folderName">Name of the folder being searched</param>
        /// <param name="searchFilter">The search filter. Available search filter classes
        /// include SearchFilter.IsEqualTo, SearchFilter.ContainsSubstring and 
        /// SearchFilter.SearchFilterCollection</param>
        /// <param name="view">The view which defines the number of persona being returned</param>
        /// <param name="queryString">The query string for which the search is being performed</param>
        /// <returns>A collection of personas matching the search conditions</returns>
        public Task<ICollection<Persona>> FindPeople(WellKnownFolderName folderName, SearchFilter searchFilter, ViewBase view, string queryString)
        {
            return this.FindPeople(new FolderId(folderName), searchFilter, view, queryString);
        }

        /// <summary>
        /// This method is for browse scenarios. Retrieves a set of personas satisfying the specified browse conditions.
        /// Browse scenariosdon't require query string.
        /// </summary>
        /// <param name="folderId">Id of the folder being browsed</param>
        /// <param name="searchFilter">Search filter</param>
        /// <param name="view">The view which defines paging and the number of persona being returned</param>
        /// <returns>A result object containing resultset for browsing</returns>
        public async Task<FindPeopleResults> FindPeople(FolderId folderId, SearchFilter searchFilter, ViewBase view,
            CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParamAllowNull(folderId, "folderId");
            EwsUtilities.ValidateParamAllowNull(searchFilter, "searchFilter");
            EwsUtilities.ValidateParamAllowNull(view, "view");
            EwsUtilities.ValidateMethodVersion(this, ExchangeVersion.Exchange2013_SP1, "FindPeople");

            FindPeopleRequest request = new FindPeopleRequest(this);

            request.FolderId = folderId;
            request.SearchFilter = searchFilter;
            request.View = view;

            return (await request.Execute(token).ConfigureAwait(false)).Results;
        }

        /// <summary>
        /// This method is for browse scenarios. Retrieves a set of personas satisfying the specified browse conditions.
        /// Browse scenarios don't require query string.
        /// </summary>
        /// <param name="folderName">Name of the folder being browsed</param>
        /// <param name="searchFilter">Search filter</param>
        /// <param name="view">The view which defines paging and the number of personas being returned</param>
        /// <returns>A result object containing resultset for browsing</returns>
        public Task<FindPeopleResults> FindPeople(WellKnownFolderName folderName, SearchFilter searchFilter, ViewBase view)
        {
            return this.FindPeople(new FolderId(folderName), searchFilter, view);
        }

        /// <summary>
        /// Retrieves all people who are relevant to the user
        /// </summary>
        /// <param name="view">The view which defines the number of personas being returned</param>
        /// <returns>A collection of personas matching the query string</returns>
        public Task<IPeopleQueryResults> BrowsePeople(ViewBase view)
        {
            return this.BrowsePeople(view, null);
        }

        /// <summary>
        /// Retrieves all people who are relevant to the user
        /// </summary>
        /// <param name="view">The view which defines the number of personas being returned</param>
        /// <param name="context">The context for this query. See PeopleQueryContextKeys for keys</param>
        /// <returns>A collection of personas matching the query string</returns>
        public Task<IPeopleQueryResults> BrowsePeople(ViewBase view, Dictionary<string, string> context, CancellationToken token = default(CancellationToken))
        {
            return this.PerformPeopleQuery(view, string.Empty, context, null, token);
        }

        /// <summary>
        /// Searches for people who are relevant to the user, automatically determining
        /// the best sources to use.
        /// </summary>
        /// <param name="view">The view which defines the number of personas being returned</param>
        /// <param name="queryString">The query string for which the search is being performed</param>
        /// <returns>A collection of personas matching the query string</returns>
        public Task<IPeopleQueryResults> SearchPeople(ViewBase view, string queryString)
        {
            return this.SearchPeople(view, queryString, null, null);
        }

        /// <summary>
        /// Searches for people who are relevant to the user
        /// </summary>
        /// <param name="view">The view which defines the number of personas being returned</param>
        /// <param name="queryString">The query string for which the search is being performed</param>
        /// <param name="context">The context for this query. See PeopleQueryContextKeys for keys</param>
        /// <param name="queryMode">The scope of the query.</param>
        /// <returns>A collection of personas matching the query string</returns>
        public Task<IPeopleQueryResults> SearchPeople(ViewBase view, string queryString, Dictionary<string, string> context, PeopleQueryMode queryMode, CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParam(queryString, "queryString");

            return this.PerformPeopleQuery(view, queryString, context, queryMode, token);
        }

        /// <summary>
        /// Performs a People Query FindPeople call
        /// </summary>
        /// <param name="view">The view which defines the number of personas being returned</param>
        /// <param name="queryString">The query string for which the search is being performed</param>
        /// <param name="context">The context for this query</param>
        /// <param name="queryMode">The scope of the query.</param>
        /// <returns></returns>
        private async Task<IPeopleQueryResults> PerformPeopleQuery(ViewBase view, string queryString, Dictionary<string, string> context, PeopleQueryMode queryMode, CancellationToken token)
        {
            EwsUtilities.ValidateParam(view, "view");
            EwsUtilities.ValidateMethodVersion(this, ExchangeVersion.Exchange2015, "FindPeople");

            if (context == null)
            {
                context = new Dictionary<string, string>();
            }

            if (queryMode == null)
            {
                queryMode = PeopleQueryMode.Auto;
            }

            FindPeopleRequest request = new FindPeopleRequest(this);
            request.View = view;
            request.QueryString = queryString;
            request.SearchPeopleSuggestionIndex = true;
            request.Context = context;
            request.QueryMode = queryMode;

            FindPeopleResponse response = await request.Execute(token).ConfigureAwait(false);

            PeopleQueryResults results = new PeopleQueryResults();
            results.Personas = response.Personas.ToList();
            results.TransactionId = response.TransactionId;

            return results;
        }

        #endregion

        #region PeopleInsights operations

        /// <summary>
        /// This method is for retreiving people insight for given email addresses
        /// </summary>
        /// <param name="emailAddresses">Specified eamiladdresses to retrieve</param>
        /// <returns>The collection of Person objects containing the insight info</returns>
        public async Task<Collection<Person>> GetPeopleInsights(IEnumerable<string> emailAddresses, CancellationToken token = default(CancellationToken))
        {
            GetPeopleInsightsRequest request = new GetPeopleInsightsRequest(this);
            request.Emailaddresses.AddRange(emailAddresses);

            return (await request.Execute(token).ConfigureAwait(false)).People;
        }

        #endregion
        #region Attachment operations

        /// <summary>
        /// Gets an attachment.
        /// </summary>
        /// <param name="attachments">The attachments.</param>
        /// <param name="bodyType">Type of the body.</param>
        /// <param name="additionalProperties">The additional properties.</param>
        /// <param name="errorHandling">Type of error handling to perform.</param>
        /// <returns>Service response collection.</returns>
        private Task<ServiceResponseCollection<GetAttachmentResponse>> InternalGetAttachments(
            IEnumerable<Attachment> attachments,
            BodyType? bodyType,
            IEnumerable<PropertyDefinitionBase> additionalProperties,
            ServiceErrorHandling errorHandling,
            CancellationToken token)
        {
            GetAttachmentRequest request = new GetAttachmentRequest(this, errorHandling);

            request.Attachments.AddRange(attachments);
            request.BodyType = bodyType;

            if (additionalProperties != null)
            {
                request.AdditionalProperties.AddRange(additionalProperties);
            }

            return request.ExecuteAsync(token);
        }

        /// <summary>
        /// Gets attachments.
        /// </summary>
        /// <param name="attachments">The attachments.</param>
        /// <param name="bodyType">Type of the body.</param>
        /// <param name="additionalProperties">The additional properties.</param>
        /// <returns>Service response collection.</returns>
        public Task<ServiceResponseCollection<GetAttachmentResponse>> GetAttachments(
            Attachment[] attachments,
            BodyType? bodyType,
            IEnumerable<PropertyDefinitionBase> additionalProperties,
            CancellationToken token = default(CancellationToken))
        {
            return this.InternalGetAttachments(
                attachments,
                bodyType,
                additionalProperties,
                ServiceErrorHandling.ReturnErrors,
                token);
        }

        /// <summary>
        /// Gets attachments.
        /// </summary>
        /// <param name="attachmentIds">The attachment ids.</param>
        /// <param name="bodyType">Type of the body.</param>
        /// <param name="additionalProperties">The additional properties.</param>
        /// <returns>Service response collection.</returns>
        public Task<ServiceResponseCollection<GetAttachmentResponse>> GetAttachments(
            string[] attachmentIds,
            BodyType? bodyType,
            IEnumerable<PropertyDefinitionBase> additionalProperties,
            CancellationToken token = default(CancellationToken))
        {
            GetAttachmentRequest request = new GetAttachmentRequest(this, ServiceErrorHandling.ReturnErrors);

            request.AttachmentIds.AddRange(attachmentIds);
            request.BodyType = bodyType;

            if (additionalProperties != null)
            {
                request.AdditionalProperties.AddRange(additionalProperties);
            }

            return request.ExecuteAsync(token);
        }

        /// <summary>
        /// Gets an attachment.
        /// </summary>
        /// <param name="attachment">The attachment.</param>
        /// <param name="bodyType">Type of the body.</param>
        /// <param name="additionalProperties">The additional properties.</param>
        internal Task<ServiceResponseCollection<GetAttachmentResponse>> GetAttachment(
            Attachment attachment,
            BodyType? bodyType,
            IEnumerable<PropertyDefinitionBase> additionalProperties,
            CancellationToken token)
        {
            return this.InternalGetAttachments(
                new Attachment[] { attachment },
                bodyType,
                additionalProperties,
                ServiceErrorHandling.ThrowOnError,
                token);
        }

        /// <summary>
        /// Creates attachments.
        /// </summary>
        /// <param name="parentItemId">The parent item id.</param>
        /// <param name="attachments">The attachments.</param>
        /// <returns>Service response collection.</returns>
        internal Task<ServiceResponseCollection<CreateAttachmentResponse>> CreateAttachments(
            string parentItemId,
            IEnumerable<Attachment> attachments,
            CancellationToken token)
        {
            CreateAttachmentRequest request = new CreateAttachmentRequest(this, ServiceErrorHandling.ReturnErrors);

            request.ParentItemId = parentItemId;
            request.Attachments.AddRange(attachments);

            return request.ExecuteAsync(token);
        }

        /// <summary>
        /// Deletes attachments.
        /// </summary>
        /// <param name="attachments">The attachments.</param>
        /// <returns>Service response collection.</returns>
        internal Task<ServiceResponseCollection<DeleteAttachmentResponse>> DeleteAttachments(IEnumerable<Attachment> attachments, CancellationToken token)
        {
            DeleteAttachmentRequest request = new DeleteAttachmentRequest(this, ServiceErrorHandling.ReturnErrors);

            request.Attachments.AddRange(attachments);

            return request.ExecuteAsync(token);
        }

        #endregion

        #region AD related operations

        /// <summary>
        /// Finds contacts in the user's Contacts folder and the Global Address List (in that order) that have names
        /// that match the one passed as a parameter. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="nameToResolve">The name to resolve.</param>
        /// <returns>A collection of name resolutions whose names match the one passed as a parameter.</returns>
        public Task<NameResolutionCollection> ResolveName(string nameToResolve)
        {
            return this.ResolveName(
                nameToResolve,
                ResolveNameSearchLocation.ContactsThenDirectory,
                false);
        }

        /// <summary>
        /// Finds contacts in the Global Address List and/or in specific contact folders that have names
        /// that match the one passed as a parameter. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="nameToResolve">The name to resolve.</param>
        /// <param name="parentFolderIds">The Ids of the contact folders in which to look for matching contacts.</param>
        /// <param name="searchScope">The scope of the search.</param>
        /// <param name="returnContactDetails">Indicates whether full contact information should be returned for each of the found contacts.</param>
        /// <returns>A collection of name resolutions whose names match the one passed as a parameter.</returns>
        public Task<NameResolutionCollection> ResolveName(
            string nameToResolve,
            IEnumerable<FolderId> parentFolderIds,
            ResolveNameSearchLocation searchScope,
            bool returnContactDetails)
        {
            return ResolveName(
                nameToResolve,
                parentFolderIds,
                searchScope,
                returnContactDetails,
                null);
        }

        /// <summary>
        /// Finds contacts in the Global Address List and/or in specific contact folders that have names
        /// that match the one passed as a parameter. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="nameToResolve">The name to resolve.</param>
        /// <param name="parentFolderIds">The Ids of the contact folders in which to look for matching contacts.</param>
        /// <param name="searchScope">The scope of the search.</param>
        /// <param name="returnContactDetails">Indicates whether full contact information should be returned for each of the found contacts.</param>
        /// <param name="contactDataPropertySet">The property set for the contct details</param>
        /// <returns>A collection of name resolutions whose names match the one passed as a parameter.</returns>
        public async Task<NameResolutionCollection> ResolveName(
            string nameToResolve,
            IEnumerable<FolderId> parentFolderIds,
            ResolveNameSearchLocation searchScope,
            bool returnContactDetails,
            PropertySet contactDataPropertySet,
            CancellationToken token = default(CancellationToken))
        {
            if (contactDataPropertySet != null)
            {
                EwsUtilities.ValidateMethodVersion(this, ExchangeVersion.Exchange2010_SP1, "ResolveName");
            }

            EwsUtilities.ValidateParam(nameToResolve, "nameToResolve");
            if (parentFolderIds != null)
            {
                EwsUtilities.ValidateParamCollection(parentFolderIds, "parentFolderIds");
            }

            ResolveNamesRequest request = new ResolveNamesRequest(this);

            request.NameToResolve = nameToResolve;
            request.ReturnFullContactData = returnContactDetails;
            request.ParentFolderIds.AddRange(parentFolderIds);
            request.SearchLocation = searchScope;
            request.ContactDataPropertySet = contactDataPropertySet;

            return (await request.ExecuteAsync(token).ConfigureAwait(false))[0].Resolutions;
        }

        /// <summary>
        /// Finds contacts in the Global Address List that have names that match the one passed as a parameter.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="nameToResolve">The name to resolve.</param>
        /// <param name="searchScope">The scope of the search.</param>
        /// <param name="returnContactDetails">Indicates whether full contact information should be returned for each of the found contacts.</param>
        /// <param name="contactDataPropertySet">Propety set for contact details</param>
        /// <returns>A collection of name resolutions whose names match the one passed as a parameter.</returns>
        public Task<NameResolutionCollection> ResolveName(
            string nameToResolve,
            ResolveNameSearchLocation searchScope,
            bool returnContactDetails,
            PropertySet contactDataPropertySet)
        {
            return this.ResolveName(
                nameToResolve,
                null,
                searchScope,
                returnContactDetails,
                contactDataPropertySet);
        }

        /// <summary>
        /// Finds contacts in the Global Address List that have names that match the one passed as a parameter.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="nameToResolve">The name to resolve.</param>
        /// <param name="searchScope">The scope of the search.</param>
        /// <param name="returnContactDetails">Indicates whether full contact information should be returned for each of the found contacts.</param>
        /// <returns>A collection of name resolutions whose names match the one passed as a parameter.</returns>
        public Task<NameResolutionCollection> ResolveName(
            string nameToResolve,
            ResolveNameSearchLocation searchScope,
            bool returnContactDetails)
        {
            return this.ResolveName(
                nameToResolve,
                null,
                searchScope,
                returnContactDetails);
        }

        /// <summary>
        /// Expands a group by retrieving a list of its members. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="emailAddress">The e-mail address of the group.</param>
        /// <returns>An ExpandGroupResults containing the members of the group.</returns>
        public async Task<ExpandGroupResults> ExpandGroup(EmailAddress emailAddress, CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParam(emailAddress, "emailAddress");

            ExpandGroupRequest request = new ExpandGroupRequest(this);

            request.EmailAddress = emailAddress;

            return (await request.ExecuteAsync(token).ConfigureAwait(false))[0].Members;
        }

        /// <summary>
        /// Expands a group by retrieving a list of its members. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="groupId">The Id of the group to expand.</param>
        /// <returns>An ExpandGroupResults containing the members of the group.</returns>
        public Task<ExpandGroupResults> ExpandGroup(ItemId groupId)
        {
            EwsUtilities.ValidateParam(groupId, "groupId");

            EmailAddress emailAddress = new EmailAddress();
            emailAddress.Id = groupId;

            return this.ExpandGroup(emailAddress);
        }

        /// <summary>
        /// Expands a group by retrieving a list of its members. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="smtpAddress">The SMTP address of the group to expand.</param>
        /// <returns>An ExpandGroupResults containing the members of the group.</returns>
        public Task<ExpandGroupResults> ExpandGroup(string smtpAddress)
        {
            EwsUtilities.ValidateParam(smtpAddress, "smtpAddress");

            return this.ExpandGroup(new EmailAddress(smtpAddress));
        }

        /// <summary>
        /// Expands a group by retrieving a list of its members. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="address">The SMTP address of the group to expand.</param>
        /// <param name="routingType">The routing type of the address of the group to expand.</param>
        /// <returns>An ExpandGroupResults containing the members of the group.</returns>
        public Task<ExpandGroupResults> ExpandGroup(string address, string routingType)
        {
            EwsUtilities.ValidateParam(address, "address");
            EwsUtilities.ValidateParam(routingType, "routingType");

            EmailAddress emailAddress = new EmailAddress(address);
            emailAddress.RoutingType = routingType;

            return this.ExpandGroup(emailAddress);
        }

        /// <summary>
        /// Get the password expiration date
        /// </summary>
        /// <param name="mailboxSmtpAddress">The e-mail address of the user.</param>
        /// <returns>The password expiration date.</returns>
        public async Task<DateTime?> GetPasswordExpirationDate(string mailboxSmtpAddress, CancellationToken token = default(CancellationToken))
        {
            GetPasswordExpirationDateRequest request = new GetPasswordExpirationDateRequest(this);
            request.MailboxSmtpAddress = mailboxSmtpAddress;

            return (await request.Execute(token).ConfigureAwait(false)).PasswordExpirationDate;
        }
        #endregion

        #region Notification operations

        /// <summary>
        /// Subscribes to pull notifications. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="folderIds">The Ids of the folder to subscribe to.</param>
        /// <param name="timeout">The timeout, in minutes, after which the subscription expires. Timeout must be between 1 and 1440.</param>
        /// <param name="watermark">An optional watermark representing a previously opened subscription.</param>
        /// <param name="eventTypes">The event types to subscribe to.</param>
        /// <returns>A PullSubscription representing the new subscription.</returns>
        public async Task<PullSubscription> SubscribeToPullNotifications(
            IEnumerable<FolderId> folderIds,
            int timeout,
            string watermark,
            CancellationToken token = default(CancellationToken),
            params EventType[] eventTypes)
        {
            EwsUtilities.ValidateParamCollection(folderIds, "folderIds");

            return (await this.BuildSubscribeToPullNotificationsRequest(
                 folderIds,
                 timeout,
                 watermark,
                 eventTypes).ExecuteAsync(token).ConfigureAwait(false))[0].Subscription;
        }

        /// <summary>
        /// Subscribes to pull notifications on all folders in the authenticated user's mailbox. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="timeout">The timeout, in minutes, after which the subscription expires. Timeout must be between 1 and 1440.</param>
        /// <param name="watermark">An optional watermark representing a previously opened subscription.</param>
        /// <param name="eventTypes">The event types to subscribe to.</param>
        /// <returns>A PullSubscription representing the new subscription.</returns>
        public async Task<PullSubscription> SubscribeToPullNotificationsOnAllFolders(
            int timeout,
            string watermark,
            CancellationToken token = default(CancellationToken),
            params EventType[] eventTypes
            )
        {
            EwsUtilities.ValidateMethodVersion(
                this,
                ExchangeVersion.Exchange2010,
                "SubscribeToPullNotificationsOnAllFolders");

            return (await this.BuildSubscribeToPullNotificationsRequest(
                null,
                timeout,
                watermark,
                eventTypes).ExecuteAsync(token).ConfigureAwait(false))[0].Subscription;
        }

        /// <summary>
        /// Builds a request to subscribe to pull notifications in the authenticated user's mailbox. 
        /// </summary>
        /// <param name="folderIds">The Ids of the folder to subscribe to.</param>
        /// <param name="timeout">The timeout, in minutes, after which the subscription expires. Timeout must be between 1 and 1440.</param>
        /// <param name="watermark">An optional watermark representing a previously opened subscription.</param>
        /// <param name="eventTypes">The event types to subscribe to.</param>
        /// <returns>A request to subscribe to pull notifications in the authenticated user's mailbox. </returns>
        private SubscribeToPullNotificationsRequest BuildSubscribeToPullNotificationsRequest(
            IEnumerable<FolderId> folderIds,
            int timeout,
            string watermark,
            EventType[] eventTypes)
        {
            if (timeout < 1 || timeout > 1440)
            {
                throw new ArgumentOutOfRangeException("timeout", Strings.TimeoutMustBeBetween1And1440);
            }

            EwsUtilities.ValidateParamCollection(eventTypes, "eventTypes");

            SubscribeToPullNotificationsRequest request = new SubscribeToPullNotificationsRequest(this);

            if (folderIds != null)
            {
                request.FolderIds.AddRange(folderIds);
            }

            request.Timeout = timeout;
            request.EventTypes.AddRange(eventTypes);
            request.Watermark = watermark;

            return request;
        }

        /// <summary>
        /// Unsubscribes from a subscription. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="subscriptionId">The Id of the pull subscription to unsubscribe from.</param>
        internal System.Threading.Tasks.Task Unsubscribe(string subscriptionId, CancellationToken token)
        {
            return this.BuildUnsubscribeRequest(subscriptionId).ExecuteAsync(token);
        }

        /// <summary>
        /// Buids a request to unsubscribe from a subscription.
        /// </summary>
        /// <param name="subscriptionId">The Id of the subscription for which to get the events.</param>
        /// <returns>A request to unsubscribe from a subscription.</returns>
        private UnsubscribeRequest BuildUnsubscribeRequest(string subscriptionId)
        {
            EwsUtilities.ValidateParam(subscriptionId, "subscriptionId");

            UnsubscribeRequest request = new UnsubscribeRequest(this);

            request.SubscriptionId = subscriptionId;

            return request;
        }

        /// <summary>
        /// Retrieves the latests events associated with a pull subscription. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="subscriptionId">The Id of the pull subscription for which to get the events.</param>
        /// <param name="watermark">The watermark representing the point in time where to start receiving events.</param>
        /// <returns>A GetEventsResults containing a list of events associated with the subscription.</returns>
        internal async Task<GetEventsResults> GetEvents(string subscriptionId, string watermark, CancellationToken token)
        {
            return (await this.BuildGetEventsRequest(subscriptionId, watermark).ExecuteAsync(token).ConfigureAwait(false))[0].Results;
        }

        /// <summary>
        /// Builds an request to retrieve the latests events associated with a pull subscription.
        /// </summary>
        /// <param name="subscriptionId">The Id of the pull subscription for which to get the events.</param>
        /// <param name="watermark">The watermark representing the point in time where to start receiving events.</param>
        /// <returns>An request to retrieve the latests events associated with a pull subscription. </returns>
        private GetEventsRequest BuildGetEventsRequest(
            string subscriptionId,
            string watermark)
        {
            EwsUtilities.ValidateParam(subscriptionId, "subscriptionId");
            EwsUtilities.ValidateParam(watermark, "watermark");

            GetEventsRequest request = new GetEventsRequest(this);

            request.SubscriptionId = subscriptionId;
            request.Watermark = watermark;

            return request;
        }

        /// <summary>
        /// Subscribes to push notifications. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="folderIds">The Ids of the folder to subscribe to.</param>
        /// <param name="url">The URL of the Web Service endpoint the Exchange server should push events to.</param>
        /// <param name="frequency">The frequency, in minutes, at which the Exchange server should contact the Web Service endpoint. Frequency must be between 1 and 1440.</param>
        /// <param name="watermark">An optional watermark representing a previously opened subscription.</param>
        /// <param name="eventTypes">The event types to subscribe to.</param>
        /// <returns>A PushSubscription representing the new subscription.</returns>
        public async Task<PushSubscription> SubscribeToPushNotifications(
            IEnumerable<FolderId> folderIds,
            Uri url,
            int frequency,
            string watermark,
            CancellationToken token = default(CancellationToken),
            params EventType[] eventTypes)
        {
            EwsUtilities.ValidateParamCollection(folderIds, "folderIds");

            return (await this.BuildSubscribeToPushNotificationsRequest(
                folderIds,
                url,
                frequency,
                watermark,
                null,
                null, // AnchorMailbox
                eventTypes).ExecuteAsync(token).ConfigureAwait(false))[0].Subscription;
        }

        /// <summary>
        /// Subscribes to push notifications on all folders in the authenticated user's mailbox. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="url">The URL of the Web Service endpoint the Exchange server should push events to.</param>
        /// <param name="frequency">The frequency, in minutes, at which the Exchange server should contact the Web Service endpoint. Frequency must be between 1 and 1440.</param>
        /// <param name="watermark">An optional watermark representing a previously opened subscription.</param>
        /// <param name="eventTypes">The event types to subscribe to.</param>
        /// <returns>A PushSubscription representing the new subscription.</returns>
        public async Task<PushSubscription> SubscribeToPushNotificationsOnAllFolders(
            Uri url,
            int frequency,
            string watermark,
            CancellationToken token = default(CancellationToken),
            params EventType[] eventTypes)
        {
            EwsUtilities.ValidateMethodVersion(
                this,
                ExchangeVersion.Exchange2010,
                "SubscribeToPushNotificationsOnAllFolders");

            return (await this.BuildSubscribeToPushNotificationsRequest(
                null,
                url,
                frequency,
                watermark,
                null,
                null, // AnchorMailbox
                eventTypes).ExecuteAsync(token).ConfigureAwait(false))[0].Subscription;
        }

        /// <summary>
        /// Subscribes to push notifications. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="folderIds">The Ids of the folder to subscribe to.</param>
        /// <param name="url">The URL of the Web Service endpoint the Exchange server should push events to.</param>
        /// <param name="frequency">The frequency, in minutes, at which the Exchange server should contact the Web Service endpoint. Frequency must be between 1 and 1440.</param>
        /// <param name="watermark">An optional watermark representing a previously opened subscription.</param>
        /// <param name="callerData">Optional caller data that will be returned the call back.</param>
        /// <param name="eventTypes">The event types to subscribe to.</param>
        /// <returns>A PushSubscription representing the new subscription.</returns>
        public async Task<PushSubscription> SubscribeToPushNotifications(
            IEnumerable<FolderId> folderIds,
            Uri url,
            int frequency,
            string watermark,
            string callerData,
            CancellationToken token = default(CancellationToken),
            params EventType[] eventTypes)
        {
            EwsUtilities.ValidateParamCollection(folderIds, "folderIds");

            return (await this.BuildSubscribeToPushNotificationsRequest(
                folderIds,
                url,
                frequency,
                watermark,
                callerData,
                null, // AnchorMailbox
                eventTypes).ExecuteAsync(token).ConfigureAwait(false))[0].Subscription;
        }

        /// <summary>
        /// Subscribes to push notifications on a group mailbox. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="groupMailboxSmtp">The smtpaddress of the group mailbox to subscribe to.</param>
        /// <param name="url">The URL of the Web Service endpoint the Exchange server should push events to.</param>
        /// <param name="frequency">The frequency, in minutes, at which the Exchange server should contact the Web Service endpoint. Frequency must be between 1 and 1440.</param>
        /// <param name="watermark">An optional watermark representing a previously opened subscription.</param>
        /// <param name="callerData">Optional caller data that will be returned the call back.</param>
        /// <param name="eventTypes">The event types to subscribe to.</param>
        /// <returns>A PushSubscription representing the new subscription.</returns>
        public async Task<PushSubscription> SubscribeToGroupPushNotifications(
            string groupMailboxSmtp,
            Uri url,
            int frequency,
            string watermark,
            string callerData,
            CancellationToken token = default(CancellationToken),
            params EventType[] eventTypes)
        {
            var folderIds = new FolderId[] { new FolderId(WellKnownFolderName.Inbox, new Mailbox(groupMailboxSmtp)) };
            return (await this.BuildSubscribeToPushNotificationsRequest(
                folderIds,
                url,
                frequency,
                watermark,
                callerData,
                groupMailboxSmtp, // AnchorMailbox
                eventTypes).ExecuteAsync(token).ConfigureAwait(false))[0].Subscription;
        }

        /// <summary>
        /// Subscribes to push notifications on all folders in the authenticated user's mailbox. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="url">The URL of the Web Service endpoint the Exchange server should push events to.</param>
        /// <param name="frequency">The frequency, in minutes, at which the Exchange server should contact the Web Service endpoint. Frequency must be between 1 and 1440.</param>
        /// <param name="watermark">An optional watermark representing a previously opened subscription.</param>
        /// <param name="callerData">Optional caller data that will be returned the call back.</param>
        /// <param name="eventTypes">The event types to subscribe to.</param>
        /// <returns>A PushSubscription representing the new subscription.</returns>
        public async Task<PushSubscription> SubscribeToPushNotificationsOnAllFolders(
            Uri url,
            int frequency,
            string watermark,
            string callerData,
            CancellationToken token = default(CancellationToken),
            params EventType[] eventTypes)
        {
            EwsUtilities.ValidateMethodVersion(
                this,
                ExchangeVersion.Exchange2010,
                "SubscribeToPushNotificationsOnAllFolders");

            return (await this.BuildSubscribeToPushNotificationsRequest(
                null,
                url,
                frequency,
                watermark,
                callerData,
                null, // AnchorMailbox
                eventTypes).ExecuteAsync(token).ConfigureAwait(false))[0].Subscription;
        }

        /// <summary>
        /// Set a TeamMailbox
        /// </summary>
        /// <param name="emailAddress">TeamMailbox email address</param>
        /// <param name="sharePointSiteUrl">SharePoint site URL</param>
        /// <param name="state">TeamMailbox lifecycle state</param>
        public System.Threading.Tasks.Task SetTeamMailbox(EmailAddress emailAddress, Uri sharePointSiteUrl, TeamMailboxLifecycleState state,
            CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateMethodVersion(this, ExchangeVersion.Exchange2013, "SetTeamMailbox");

            if (emailAddress == null)
            {
                throw new ArgumentNullException("emailAddress");
            }

            if (sharePointSiteUrl == null)
            {
                throw new ArgumentNullException("sharePointSiteUrl");
            }

            SetTeamMailboxRequest request = new SetTeamMailboxRequest(this, emailAddress, sharePointSiteUrl, state);
            return request.Execute(token);
        }

        /// <summary>
        /// Unpin a TeamMailbox
        /// </summary>
        /// <param name="emailAddress">TeamMailbox email address</param>
        public System.Threading.Tasks.Task UnpinTeamMailbox(EmailAddress emailAddress, CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateMethodVersion(this, ExchangeVersion.Exchange2013, "UnpinTeamMailbox");

            if (emailAddress == null)
            {
                throw new ArgumentNullException("emailAddress");
            }

            UnpinTeamMailboxRequest request = new UnpinTeamMailboxRequest(this, emailAddress);
            return request.Execute(token);
        }

        /// <summary>
        /// Builds an request to request to subscribe to push notifications in the authenticated user's mailbox.
        /// </summary>
        /// <param name="folderIds">The Ids of the folder to subscribe to.</param>
        /// <param name="url">The URL of the Web Service endpoint the Exchange server should push events to.</param>
        /// <param name="frequency">The frequency, in minutes, at which the Exchange server should contact the Web Service endpoint. Frequency must be between 1 and 1440.</param>
        /// <param name="watermark">An optional watermark representing a previously opened subscription.</param>
        /// <param name="callerData">Optional caller data that will be returned the call back.</param>
        /// <param name="anchorMailbox">The smtpaddress of the mailbox to subscribe to.</param>
        /// <param name="eventTypes">The event types to subscribe to.</param>
        /// <returns>A request to request to subscribe to push notifications in the authenticated user's mailbox.</returns>
        private SubscribeToPushNotificationsRequest BuildSubscribeToPushNotificationsRequest(
            IEnumerable<FolderId> folderIds,
            Uri url,
            int frequency,
            string watermark,
            string callerData,
            string anchorMailbox,
            EventType[] eventTypes)
        {
            EwsUtilities.ValidateParam(url, "url");

            if (frequency < 1 || frequency > 1440)
            {
                throw new ArgumentOutOfRangeException("frequency", Strings.FrequencyMustBeBetween1And1440);
            }

            EwsUtilities.ValidateParamCollection(eventTypes, "eventTypes");

            SubscribeToPushNotificationsRequest request = new SubscribeToPushNotificationsRequest(this);
            request.AnchorMailbox = anchorMailbox;

            if (folderIds != null)
            {
                request.FolderIds.AddRange(folderIds);
            }

            request.Url = url;
            request.Frequency = frequency;
            request.EventTypes.AddRange(eventTypes);
            request.Watermark = watermark;
            request.CallerData = callerData;

            return request;
        }

        /// <summary>
        /// Subscribes to streaming notifications. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="folderIds">The Ids of the folder to subscribe to.</param>
        /// <param name="eventTypes">The event types to subscribe to.</param>
        /// <returns>A StreamingSubscription representing the new subscription.</returns>
        public async System.Threading.Tasks.Task<StreamingSubscription> SubscribeToStreamingNotifications(
            IEnumerable<FolderId> folderIds,
            CancellationToken token = default(CancellationToken),
            params EventType[] eventTypes)
        {
            EwsUtilities.ValidateMethodVersion(
                this,
                ExchangeVersion.Exchange2010_SP1,
                "SubscribeToStreamingNotifications");

            EwsUtilities.ValidateParamCollection(folderIds, "folderIds");

            return (await this.BuildSubscribeToStreamingNotificationsRequest(folderIds, eventTypes).ExecuteAsync(token).ConfigureAwait(false))[0].Subscription;
        }

        /// <summary>
        /// Subscribes to streaming notifications on all folders in the authenticated user's mailbox. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="eventTypes">The event types to subscribe to.</param>
        /// <returns>A StreamingSubscription representing the new subscription.</returns>
        public async Task<StreamingSubscription> SubscribeToStreamingNotificationsOnAllFolders(
            CancellationToken token = default(CancellationToken),
            params EventType[] eventTypes)
        {
            EwsUtilities.ValidateMethodVersion(
                this,
                ExchangeVersion.Exchange2010_SP1,
                "SubscribeToStreamingNotificationsOnAllFolders");

            return (await this.BuildSubscribeToStreamingNotificationsRequest(null, eventTypes).ExecuteAsync(token).ConfigureAwait(false))[0].Subscription;
        }

        /// <summary>
        /// Builds request to subscribe to streaming notifications in the authenticated user's mailbox. 
        /// </summary>
        /// <param name="folderIds">The Ids of the folder to subscribe to.</param>
        /// <param name="eventTypes">The event types to subscribe to.</param>
        /// <returns>A request to subscribe to streaming notifications in the authenticated user's mailbox. </returns>
        private SubscribeToStreamingNotificationsRequest BuildSubscribeToStreamingNotificationsRequest(
            IEnumerable<FolderId> folderIds,
            EventType[] eventTypes)
        {
            EwsUtilities.ValidateParamCollection(eventTypes, "eventTypes");

            SubscribeToStreamingNotificationsRequest request = new SubscribeToStreamingNotificationsRequest(this);

            if (folderIds != null)
            {
                request.FolderIds.AddRange(folderIds);
            }

            request.EventTypes.AddRange(eventTypes);

            return request;
        }

        #endregion

        #region Synchronization operations

        /// <summary>
        /// Synchronizes the items of a specific folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="syncFolderId">The Id of the folder containing the items to synchronize with.</param>
        /// <param name="propertySet">The set of properties to retrieve for synchronized items.</param>
        /// <param name="ignoredItemIds">The optional list of item Ids that should be ignored.</param>
        /// <param name="maxChangesReturned">The maximum number of changes that should be returned.</param>
        /// <param name="syncScope">The sync scope identifying items to include in the ChangeCollection.</param>
        /// <param name="syncState">The optional sync state representing the point in time when to start the synchronization.</param>
        /// <returns>A ChangeCollection containing a list of changes that occurred in the specified folder.</returns>
        public System.Threading.Tasks.Task<ChangeCollection<ItemChange>> SyncFolderItems(
            FolderId syncFolderId,
            PropertySet propertySet,
            IEnumerable<ItemId> ignoredItemIds,
            int maxChangesReturned,
            SyncFolderItemsScope syncScope,
            string syncState)
        {
            return this.SyncFolderItems(
                syncFolderId,
                propertySet,
                ignoredItemIds,
                maxChangesReturned,
                0, // numberOfDays
                syncScope,
                syncState);
        }

        /// <summary>
        /// Synchronizes the items of a specific folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="syncFolderId">The Id of the folder containing the items to synchronize with.</param>
        /// <param name="propertySet">The set of properties to retrieve for synchronized items.</param>
        /// <param name="ignoredItemIds">The optional list of item Ids that should be ignored.</param>
        /// <param name="maxChangesReturned">The maximum number of changes that should be returned.</param>
        /// <param name="numberOfDays">Limit the changes returned to this many days ago; 0 means no limit.</param>
        /// <param name="syncScope">The sync scope identifying items to include in the ChangeCollection.</param>
        /// <param name="syncState">The optional sync state representing the point in time when to start the synchronization.</param>
        /// <returns>A ChangeCollection containing a list of changes that occurred in the specified folder.</returns>
        public async System.Threading.Tasks.Task<ChangeCollection<ItemChange>> SyncFolderItems(
            FolderId syncFolderId,
            PropertySet propertySet,
            IEnumerable<ItemId> ignoredItemIds,
            int maxChangesReturned,
            int numberOfDays,
            SyncFolderItemsScope syncScope,
            string syncState,
            CancellationToken token = default(CancellationToken))
        {
            return (await this.BuildSyncFolderItemsRequest(
                syncFolderId,
                propertySet,
                ignoredItemIds,
                maxChangesReturned,
                numberOfDays,
                syncScope,
                syncState).ExecuteAsync(token).ConfigureAwait(false))[0].Changes;
        }

        /// <summary>
        /// Builds a request to synchronize the items of a specific folder.
        /// </summary>
        /// <param name="syncFolderId">The Id of the folder containing the items to synchronize with.</param>
        /// <param name="propertySet">The set of properties to retrieve for synchronized items.</param>
        /// <param name="ignoredItemIds">The optional list of item Ids that should be ignored.</param>
        /// <param name="maxChangesReturned">The maximum number of changes that should be returned.</param>
        /// <param name="numberOfDays">Limit the changes returned to this many days ago; 0 means no limit.</param>
        /// <param name="syncScope">The sync scope identifying items to include in the ChangeCollection.</param>
        /// <param name="syncState">The optional sync state representing the point in time when to start the synchronization.</param>
        /// <returns>A request to synchronize the items of a specific folder.</returns>
        private SyncFolderItemsRequest BuildSyncFolderItemsRequest(
            FolderId syncFolderId,
            PropertySet propertySet,
            IEnumerable<ItemId> ignoredItemIds,
            int maxChangesReturned,
            int numberOfDays,
            SyncFolderItemsScope syncScope,
            string syncState)
        {
            EwsUtilities.ValidateParam(syncFolderId, "syncFolderId");
            EwsUtilities.ValidateParam(propertySet, "propertySet");

            SyncFolderItemsRequest request = new SyncFolderItemsRequest(this);

            request.SyncFolderId = syncFolderId;
            request.PropertySet = propertySet;
            if (ignoredItemIds != null)
            {
                request.IgnoredItemIds.AddRange(ignoredItemIds);
            }
            request.MaxChangesReturned = maxChangesReturned;
            request.NumberOfDays = numberOfDays;
            request.SyncScope = syncScope;
            request.SyncState = syncState;

            return request;
        }

        /// <summary>
        /// Synchronizes the sub-folders of a specific folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="syncFolderId">The Id of the folder containing the items to synchronize with. A null value indicates the root folder of the mailbox.</param>
        /// <param name="propertySet">The set of properties to retrieve for synchronized items.</param>
        /// <param name="syncState">The optional sync state representing the point in time when to start the synchronization.</param>
        /// <returns>A ChangeCollection containing a list of changes that occurred in the specified folder.</returns>
        public async Task<ChangeCollection<FolderChange>> SyncFolderHierarchy(
            FolderId syncFolderId,
            PropertySet propertySet,
            string syncState,
            CancellationToken token = default(CancellationToken))
        {
            return (await this.BuildSyncFolderHierarchyRequest(
                syncFolderId,
                propertySet,
                syncState).ExecuteAsync(token).ConfigureAwait(false))[0].Changes;
        }

        /// <summary>
        /// Synchronizes the entire folder hierarchy of the mailbox this Service is connected to. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="propertySet">The set of properties to retrieve for synchronized items.</param>
        /// <param name="syncState">The optional sync state representing the point in time when to start the synchronization.</param>
        /// <returns>A ChangeCollection containing a list of changes that occurred in the specified folder.</returns>
        public Task<ChangeCollection<FolderChange>> SyncFolderHierarchy(PropertySet propertySet, string syncState)
        {
            return this.SyncFolderHierarchy(
                null,
                propertySet,
                syncState);
        }

        /// <summary>
        /// Builds a request to synchronize the specified folder hierarchy of the mailbox this Service is connected to.
        /// </summary>
        /// <param name="syncFolderId">The Id of the folder containing the items to synchronize with. A null value indicates the root folder of the mailbox.</param>
        /// <param name="propertySet">The set of properties to retrieve for synchronized items.</param>
        /// <param name="syncState">The optional sync state representing the point in time when to start the synchronization.</param>
        /// <returns>A request to synchronize the specified folder hierarchy of the mailbox this Service is connected to.</returns>
        private SyncFolderHierarchyRequest BuildSyncFolderHierarchyRequest(
            FolderId syncFolderId,
            PropertySet propertySet,
            string syncState)
        {
            EwsUtilities.ValidateParamAllowNull(syncFolderId, "syncFolderId");  // Null syncFolderId is allowed
            EwsUtilities.ValidateParam(propertySet, "propertySet");

            SyncFolderHierarchyRequest request = new SyncFolderHierarchyRequest(this);

            request.PropertySet = propertySet;
            request.SyncFolderId = syncFolderId;
            request.SyncState = syncState;

            return request;
        }

        #endregion

        #region Availability operations

        /// <summary>
        /// Gets Out of Office (OOF) settings for a specific user. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="smtpAddress">The SMTP address of the user for which to retrieve OOF settings.</param>
        /// <returns>An OofSettings instance containing OOF information for the specified user.</returns>
        public async Task<OofSettings> GetUserOofSettings(string smtpAddress, CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParam(smtpAddress, "smtpAddress");

            GetUserOofSettingsRequest request = new GetUserOofSettingsRequest(this);

            request.SmtpAddress = smtpAddress;

            return (await request.Execute(token).ConfigureAwait(false)).OofSettings;
        }

        /// <summary>
        /// Sets the Out of Office (OOF) settings for a specific mailbox. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="smtpAddress">The SMTP address of the user for which to set OOF settings.</param>
        /// <param name="oofSettings">The OOF settings.</param>
        public System.Threading.Tasks.Task SetUserOofSettings(string smtpAddress, OofSettings oofSettings, CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParam(smtpAddress, "smtpAddress");
            EwsUtilities.ValidateParam(oofSettings, "oofSettings");

            SetUserOofSettingsRequest request = new SetUserOofSettingsRequest(this);

            request.SmtpAddress = smtpAddress;
            request.OofSettings = oofSettings;

            return request.Execute(token);
        }

        /// <summary>
        /// Gets detailed information about the availability of a set of users, rooms, and resources within a
        /// specified time window.
        /// </summary>
        /// <param name="attendees">The attendees for which to retrieve availability information.</param>
        /// <param name="timeWindow">The time window in which to retrieve user availability information.</param>
        /// <param name="requestedData">The requested data (free/busy and/or suggestions).</param>
        /// <param name="options">The options controlling the information returned.</param>
        /// <returns>
        /// The availability information for each user appears in a unique FreeBusyResponse object. The order of users
        /// in the request determines the order of availability data for each user in the response.
        /// </returns>
        public Task<GetUserAvailabilityResults> GetUserAvailability(
            IEnumerable<AttendeeInfo> attendees,
            TimeWindow timeWindow,
            AvailabilityData requestedData,
            AvailabilityOptions options,
            CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParamCollection(attendees, "attendees");
            EwsUtilities.ValidateParam(timeWindow, "timeWindow");
            EwsUtilities.ValidateParam(options, "options");

            GetUserAvailabilityRequest request = new GetUserAvailabilityRequest(this);

            request.Attendees = attendees;
            request.TimeWindow = timeWindow;
            request.RequestedData = requestedData;
            request.Options = options;

            return request.Execute(token);
        }

        /// <summary>
        /// Gets detailed information about the availability of a set of users, rooms, and resources within a
        /// specified time window.
        /// </summary>
        /// <param name="attendees">The attendees for which to retrieve availability information.</param>
        /// <param name="timeWindow">The time window in which to retrieve user availability information.</param>
        /// <param name="requestedData">The requested data (free/busy and/or suggestions).</param>
        /// <returns>
        /// The availability information for each user appears in a unique FreeBusyResponse object. The order of users
        /// in the request determines the order of availability data for each user in the response.
        /// </returns>
        public Task<GetUserAvailabilityResults> GetUserAvailability(
            IEnumerable<AttendeeInfo> attendees,
            TimeWindow timeWindow,
            AvailabilityData requestedData)
        {
            return this.GetUserAvailability(
                attendees,
                timeWindow,
                requestedData,
                new AvailabilityOptions());
        }

        /// <summary>
        /// Retrieves a collection of all room lists in the organization.
        /// </summary>
        /// <returns>An EmailAddressCollection containing all the room lists in the organization.</returns>
        public async Task<EmailAddressCollection> GetRoomLists(CancellationToken token = default(CancellationToken))
        {
            GetRoomListsRequest request = new GetRoomListsRequest(this);

            return (await request.Execute(token).ConfigureAwait(false)).RoomLists;
        }

        /// <summary>
        /// Retrieves a collection of all rooms in the specified room list in the organization.
        /// </summary>
        /// <param name="emailAddress">The e-mail address of the room list.</param>
        /// <returns>A collection of EmailAddress objects representing all the rooms within the specifed room list.</returns>
        public async Task<Collection<EmailAddress>> GetRooms(EmailAddress emailAddress, CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParam(emailAddress, "emailAddress");

            GetRoomsRequest request = new GetRoomsRequest(this);

            request.RoomList = emailAddress;

            return (await request.Execute(token).ConfigureAwait(false)).Rooms;
        }
        #endregion

        #region Conversation
        /// <summary>
        /// Retrieves a collection of all Conversations in the specified Folder.
        /// </summary>
        /// <param name="view">The view controlling the number of conversations returned.</param>
        /// <param name="folderId">The Id of the folder in which to search for conversations.</param>
        /// <returns>Collection of conversations.</returns>
        public async Task<ICollection<Conversation>> FindConversation(ViewBase view, FolderId folderId, CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParam(view, "view");
            EwsUtilities.ValidateParam(folderId, "folderId");
            EwsUtilities.ValidateMethodVersion(
                                            this,
                                            ExchangeVersion.Exchange2010_SP1,
                                            "FindConversation");

            FindConversationRequest request = new FindConversationRequest(this);

            request.View = view;
            request.FolderId = new FolderIdWrapper(folderId);

            return (await request.Execute(token).ConfigureAwait(false)).Conversations;
        }

        /// <summary>
        /// Retrieves a collection of all Conversations in the specified Folder.
        /// </summary>
        /// <param name="view">The view controlling the number of conversations returned.</param>
        /// <param name="folderId">The Id of the folder in which to search for conversations.</param>
        /// <param name="anchorMailbox">The anchorMailbox Smtp address to route the request directly to group mailbox.</param>
        /// <returns>Collection of conversations.</returns>
        /// <remarks>
        /// This API designed to be used primarily in groups scenarios where we want to set the
        /// anchor mailbox header so that request is routed directly to the group mailbox backend server.
        /// </remarks>
        public async Task<Collection<Conversation>> FindGroupConversation(
            ViewBase view,
            FolderId folderId,
            string anchorMailbox,
            CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParam(view, "view");
            EwsUtilities.ValidateParam(folderId, "folderId");
            EwsUtilities.ValidateParam(anchorMailbox, "anchorMailbox");
            EwsUtilities.ValidateMethodVersion(
                                            this,
                                            ExchangeVersion.Exchange2015,
                                            "FindConversation");

            FindConversationRequest request = new FindConversationRequest(this);

            request.View = view;
            request.FolderId = new FolderIdWrapper(folderId);
            request.AnchorMailbox = anchorMailbox;

            return (await request.Execute(token).ConfigureAwait(false)).Conversations;
        }

        /// <summary>
        /// Retrieves a collection of all Conversations in the specified Folder.
        /// </summary>
        /// <param name="view">The view controlling the number of conversations returned.</param>
        /// <param name="folderId">The Id of the folder in which to search for conversations.</param>
        /// <param name="queryString">The query string for which the search is being performed</param>
        /// <returns>Collection of conversations.</returns>
        public async Task<ICollection<Conversation>> FindConversation(ViewBase view, FolderId folderId, string queryString, CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParam(view, "view");
            EwsUtilities.ValidateParamAllowNull(queryString, "queryString");
            EwsUtilities.ValidateParam(folderId, "folderId");
            EwsUtilities.ValidateMethodVersion(
                                            this,
                                            ExchangeVersion.Exchange2013, // This method is only applicable for Exchange2013
                                            "FindConversation");

            FindConversationRequest request = new FindConversationRequest(this);

            request.View = view;
            request.FolderId = new FolderIdWrapper(folderId);
            request.QueryString = queryString;

            return (await request.Execute(token).ConfigureAwait(false)).Conversations;
        }

        /// <summary>
        /// Searches for and retrieves a collection of Conversations in the specified Folder.
        /// Along with conversations, a list of highlight terms are returned.
        /// </summary>
        /// <param name="view">The view controlling the number of conversations returned.</param>
        /// <param name="folderId">The Id of the folder in which to search for conversations.</param>
        /// <param name="queryString">The query string for which the search is being performed</param>
        /// <param name="returnHighlightTerms">Flag indicating if highlight terms should be returned in the response</param>
        /// <returns>FindConversation results.</returns>
        public async Task<FindConversationResults> FindConversation(ViewBase view, FolderId folderId, string queryString, bool returnHighlightTerms, CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParam(view, "view");
            EwsUtilities.ValidateParamAllowNull(queryString, "queryString");
            EwsUtilities.ValidateParam(returnHighlightTerms, "returnHighlightTerms");
            EwsUtilities.ValidateParam(folderId, "folderId");
            EwsUtilities.ValidateMethodVersion(
                                            this,
                                            ExchangeVersion.Exchange2013, // This method is only applicable for Exchange2013
                                            "FindConversation");

            FindConversationRequest request = new FindConversationRequest(this);

            request.View = view;
            request.FolderId = new FolderIdWrapper(folderId);
            request.QueryString = queryString;
            request.ReturnHighlightTerms = returnHighlightTerms;

            return (await request.Execute(token).ConfigureAwait(false)).Results;
        }

        /// <summary>
        /// Searches for and retrieves a collection of Conversations in the specified Folder.
        /// Along with conversations, a list of highlight terms are returned.
        /// </summary>
        /// <param name="view">The view controlling the number of conversations returned.</param>
        /// <param name="folderId">The Id of the folder in which to search for conversations.</param>
        /// <param name="queryString">The query string for which the search is being performed</param>
        /// <param name="returnHighlightTerms">Flag indicating if highlight terms should be returned in the response</param>
        /// <param name="mailboxScope">The mailbox scope to reference.</param>
        /// <returns>FindConversation results.</returns>
        public async Task<FindConversationResults> FindConversation(ViewBase view, FolderId folderId, string queryString, bool returnHighlightTerms, MailboxSearchLocation? mailboxScope, CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParam(view, "view");
            EwsUtilities.ValidateParamAllowNull(queryString, "queryString");
            EwsUtilities.ValidateParam(returnHighlightTerms, "returnHighlightTerms");
            EwsUtilities.ValidateParam(folderId, "folderId");

            EwsUtilities.ValidateMethodVersion(
                                            this,
                                            ExchangeVersion.Exchange2013, // This method is only applicable for Exchange2013
                                            "FindConversation");

            FindConversationRequest request = new FindConversationRequest(this);

            request.View = view;
            request.FolderId = new FolderIdWrapper(folderId);
            request.QueryString = queryString;
            request.ReturnHighlightTerms = returnHighlightTerms;
            request.MailboxScope = mailboxScope;

            return (await request.Execute(token).ConfigureAwait(false)).Results;
        }

        /// <summary>
        /// Gets the items for a set of conversations.
        /// </summary>
        /// <param name="conversations">Conversations with items to load.</param>
        /// <param name="propertySet">The set of properties to load.</param>
        /// <param name="foldersToIgnore">The folders to ignore.</param>
        /// <param name="sortOrder">Sort order of conversation tree nodes.</param>
        /// <param name="mailboxScope">The mailbox scope to reference.</param>
        /// <param name="anchorMailbox">The smtpaddress of the mailbox that hosts the conversations</param>
        /// <param name="maxItemsToReturn">Maximum number of items to return.</param>
        /// <param name="errorHandling">What type of error handling should be performed.</param>
        /// <returns>GetConversationItems response.</returns>
        internal Task<ServiceResponseCollection<GetConversationItemsResponse>> InternalGetConversationItems(
                            IEnumerable<ConversationRequest> conversations,
                            PropertySet propertySet,
                            IEnumerable<FolderId> foldersToIgnore,
                            ConversationSortOrder? sortOrder,
                            MailboxSearchLocation? mailboxScope,
                            int? maxItemsToReturn,
                            string anchorMailbox,
                            ServiceErrorHandling errorHandling,
                            CancellationToken token)
        {
            EwsUtilities.ValidateParam(conversations, "conversations");
            EwsUtilities.ValidateParam(propertySet, "itemProperties");
            EwsUtilities.ValidateParamAllowNull(foldersToIgnore, "foldersToIgnore");
            EwsUtilities.ValidateMethodVersion(
                                            this,
                                            ExchangeVersion.Exchange2013,
                                            "GetConversationItems");

            GetConversationItemsRequest request = new GetConversationItemsRequest(this, errorHandling);
            request.ItemProperties = propertySet;
            request.FoldersToIgnore = new FolderIdCollection(foldersToIgnore);
            request.SortOrder = sortOrder;
            request.MailboxScope = mailboxScope;
            request.MaxItemsToReturn = maxItemsToReturn;
            request.AnchorMailbox = anchorMailbox;
            request.Conversations = conversations.ToList();

            return request.ExecuteAsync(token);
        }

        /// <summary>
        /// Gets the items for a set of conversations.
        /// </summary>
        /// <param name="conversations">Conversations with items to load.</param>
        /// <param name="propertySet">The set of properties to load.</param>
        /// <param name="foldersToIgnore">The folders to ignore.</param>
        /// <param name="sortOrder">Conversation item sort order.</param>
        /// <returns>GetConversationItems response.</returns>
        public Task<ServiceResponseCollection<GetConversationItemsResponse>> GetConversationItems(
                                                IEnumerable<ConversationRequest> conversations,
                                                PropertySet propertySet,
                                                IEnumerable<FolderId> foldersToIgnore,
                                                ConversationSortOrder? sortOrder,
                                                CancellationToken token = default(CancellationToken))
        {
            return this.InternalGetConversationItems(
                                conversations,
                                propertySet,
                                foldersToIgnore,
                                null,               /* sortOrder */
                                null,               /* mailboxScope */
                                null,               /* maxItemsToReturn*/
                                null, /* anchorMailbox */
                                ServiceErrorHandling.ReturnErrors,
                                token);
        }

        /// <summary>
        /// Gets the items for a conversation.
        /// </summary>
        /// <param name="conversationId">The conversation id.</param>
        /// <param name="propertySet">The set of properties to load.</param>
        /// <param name="syncState">The optional sync state representing the point in time when to start the synchronization.</param>
        /// <param name="foldersToIgnore">The folders to ignore.</param>
        /// <param name="sortOrder">Conversation item sort order.</param>
        /// <returns>ConversationResponseType response.</returns>
        public async Task<ConversationResponse> GetConversationItems(
                                                ConversationId conversationId,
                                                PropertySet propertySet,
                                                string syncState,
                                                IEnumerable<FolderId> foldersToIgnore,
                                                ConversationSortOrder? sortOrder,
                                                CancellationToken token = default(CancellationToken))
        {
            List<ConversationRequest> conversations = new List<ConversationRequest>();
            conversations.Add(new ConversationRequest(conversationId, syncState));

            return (await this.InternalGetConversationItems(
                                conversations,
                                propertySet,
                                foldersToIgnore,
                                sortOrder,
                                null,           /* mailboxScope */
                                null,           /* maxItemsToReturn */
                                null, /* anchorMailbox */
                                ServiceErrorHandling.ThrowOnError,
                                token).ConfigureAwait(false))[0].Conversation;
        }

        /// <summary>
        /// Gets the items for a conversation.
        /// </summary>
        /// <param name="conversationId">The conversation id.</param>
        /// <param name="propertySet">The set of properties to load.</param>
        /// <param name="syncState">The optional sync state representing the point in time when to start the synchronization.</param>
        /// <param name="foldersToIgnore">The folders to ignore.</param>
        /// <param name="sortOrder">Conversation item sort order.</param>
        /// <param name="anchorMailbox">The smtp address of the mailbox hosting the conversations</param>
        /// <returns>ConversationResponseType response.</returns>
        /// <remarks>
        /// This API designed to be used primarily in groups scenarios where we want to set the
        /// anchor mailbox header so that request is routed directly to the group mailbox backend server.
        /// </remarks>
        public async Task<ConversationResponse> GetGroupConversationItems(
                                                ConversationId conversationId,
                                                PropertySet propertySet,
                                                string syncState,
                                                IEnumerable<FolderId> foldersToIgnore,
                                                ConversationSortOrder? sortOrder,
                                                string anchorMailbox,
                                                CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParam(anchorMailbox, "anchorMailbox");

            List<ConversationRequest> conversations = new List<ConversationRequest>();
            conversations.Add(new ConversationRequest(conversationId, syncState));

            return (await this.InternalGetConversationItems(
                                conversations,
                                propertySet,
                                foldersToIgnore,
                                sortOrder,
                                null,           /* mailboxScope */
                                null,           /* maxItemsToReturn */
                                anchorMailbox, /* anchorMailbox */
                                ServiceErrorHandling.ThrowOnError,
                                token).ConfigureAwait(false))[0].Conversation;
        }

        /// <summary>
        /// Gets the items for a set of conversations.
        /// </summary>
        /// <param name="conversations">Conversations with items to load.</param>
        /// <param name="propertySet">The set of properties to load.</param>
        /// <param name="foldersToIgnore">The folders to ignore.</param>
        /// <param name="sortOrder">Conversation item sort order.</param>
        /// <param name="mailboxScope">The mailbox scope to reference.</param>
        /// <returns>GetConversationItems response.</returns>
        public Task<ServiceResponseCollection<GetConversationItemsResponse>> GetConversationItems(
                                                IEnumerable<ConversationRequest> conversations,
                                                PropertySet propertySet,
                                                IEnumerable<FolderId> foldersToIgnore,
                                                ConversationSortOrder? sortOrder,
                                                MailboxSearchLocation? mailboxScope,
                                                CancellationToken token = default(CancellationToken))
        {
            return this.InternalGetConversationItems(
                                conversations,
                                propertySet,
                                foldersToIgnore,
                                null,               /* sortOrder */
                                mailboxScope,
                                null,               /* maxItemsToReturn*/
                                null, /* anchorMailbox */
                                ServiceErrorHandling.ReturnErrors,
                                token);
        }

        /// <summary>
        /// Applies ConversationAction on the specified conversation.
        /// </summary>
        /// <param name="actionType">ConversationAction</param>
        /// <param name="conversationIds">The conversation ids.</param>
        /// <param name="processRightAway">True to process at once . This is blocking
        /// and false to let the Assistant process it in the back ground</param>
        /// <param name="categories">Catgories that need to be stamped can be null or empty</param>
        /// <param name="enableAlwaysDelete">True moves every current and future messages in the conversation
        /// to deleted items folder. False stops the alwasy delete action. This is applicable only if
        /// the action is AlwaysDelete</param>
        /// <param name="destinationFolderId">Applicable if the action is AlwaysMove. This moves every current message and future
        /// message in the conversation to the specified folder. Can be null if tis is then it stops
        /// the always move action</param>
        /// <param name="errorHandlingMode">The error handling mode.</param>
        /// <returns></returns>
        private Task<ServiceResponseCollection<ServiceResponse>> ApplyConversationAction(
                ConversationActionType actionType,
                IEnumerable<ConversationId> conversationIds,
                bool processRightAway,
                StringList categories,
                bool enableAlwaysDelete,
                FolderId destinationFolderId,
                ServiceErrorHandling errorHandlingMode,
                CancellationToken token)
        {
            EwsUtilities.Assert(
                actionType == ConversationActionType.AlwaysCategorize ||
                actionType == ConversationActionType.AlwaysMove ||
                actionType == ConversationActionType.AlwaysDelete,
                "ApplyConversationAction",
                "Invalid actionType");

            EwsUtilities.ValidateParam(conversationIds, "conversationId");
            EwsUtilities.ValidateMethodVersion(
                                this,
                                ExchangeVersion.Exchange2010_SP1,
                                "ApplyConversationAction");

            ApplyConversationActionRequest request = new ApplyConversationActionRequest(this, errorHandlingMode);

            foreach (var conversationId in conversationIds)
            {
                ConversationAction action = new ConversationAction();

                action.Action = actionType;
                action.ConversationId = conversationId;
                action.ProcessRightAway = processRightAway;
                action.Categories = categories;
                action.EnableAlwaysDelete = enableAlwaysDelete;
                action.DestinationFolderId = destinationFolderId != null ? new FolderIdWrapper(destinationFolderId) : null;

                request.ConversationActions.Add(action);
            }

            return request.ExecuteAsync(token);
        }

        /// <summary>
        /// Applies one time conversation action on items in specified folder inside
        /// the conversation.
        /// </summary>
        /// <param name="actionType">The action.</param>
        /// <param name="idTimePairs">The id time pairs.</param>
        /// <param name="contextFolderId">The context folder id.</param>
        /// <param name="destinationFolderId">The destination folder id.</param>
        /// <param name="deleteType">Type of the delete.</param>
        /// <param name="isRead">The is read.</param>
        /// <param name="retentionPolicyType">Retention policy type.</param>
        /// <param name="retentionPolicyTagId">Retention policy tag id.  Null will clear the policy.</param>
        /// <param name="flag">Flag status.</param>
        /// <param name="suppressReadReceipts">Suppress read receipts flag.</param>
        /// <param name="errorHandlingMode">The error handling mode.</param>
        /// <returns></returns>
        private Task<ServiceResponseCollection<ServiceResponse>> ApplyConversationOneTimeAction(
            ConversationActionType actionType,
            IEnumerable<KeyValuePair<ConversationId, DateTime?>> idTimePairs,
            FolderId contextFolderId,
            FolderId destinationFolderId,
            DeleteMode? deleteType,
            bool? isRead,
            RetentionType? retentionPolicyType,
            Guid? retentionPolicyTagId,
            Flag flag,
            bool? suppressReadReceipts,
            ServiceErrorHandling errorHandlingMode,
            CancellationToken token)
        {
            EwsUtilities.Assert(
                actionType == ConversationActionType.Move ||
                actionType == ConversationActionType.Delete ||
                actionType == ConversationActionType.SetReadState ||
                actionType == ConversationActionType.SetRetentionPolicy ||
                actionType == ConversationActionType.Copy ||
                actionType == ConversationActionType.Flag,
                "ApplyConversationOneTimeAction",
                "Invalid actionType");

            EwsUtilities.ValidateParamCollection(idTimePairs, "idTimePairs");
            EwsUtilities.ValidateMethodVersion(
                                this,
                                ExchangeVersion.Exchange2010_SP1,
                                "ApplyConversationAction");

            ApplyConversationActionRequest request = new ApplyConversationActionRequest(this, errorHandlingMode);

            foreach (var idTimePair in idTimePairs)
            {
                ConversationAction action = new ConversationAction();

                action.Action = actionType;
                action.ConversationId = idTimePair.Key;
                action.ContextFolderId = contextFolderId != null ? new FolderIdWrapper(contextFolderId) : null;
                action.DestinationFolderId = destinationFolderId != null ? new FolderIdWrapper(destinationFolderId) : null;
                action.ConversationLastSyncTime = idTimePair.Value;
                action.IsRead = isRead;
                action.DeleteType = deleteType;
                action.RetentionPolicyType = retentionPolicyType;
                action.RetentionPolicyTagId = retentionPolicyTagId;
                action.Flag = flag;
                action.SuppressReadReceipts = suppressReadReceipts;

                request.ConversationActions.Add(action);
            }

            return request.ExecuteAsync(token);
        }

        /// <summary>
        /// Sets up a conversation so that any item received within that conversation is always categorized.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="conversationId">The id of the conversation.</param>
        /// <param name="categories">The categories that should be stamped on items in the conversation.</param>
        /// <param name="processSynchronously">Indicates whether the method should return only once enabling this rule and stamping existing items
        /// in the conversation is completely done. If processSynchronously is false, the method returns immediately.</param>
        /// <returns></returns>
        public Task<ServiceResponseCollection<ServiceResponse>> EnableAlwaysCategorizeItemsInConversations(
            IEnumerable<ConversationId> conversationId,
            IEnumerable<String> categories,
            bool processSynchronously,
            CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParamCollection(categories, "categories");
            return this.ApplyConversationAction(
                        ConversationActionType.AlwaysCategorize,
                        conversationId,
                        processSynchronously,
                        new StringList(categories),
                        false,
                        null,
                        ServiceErrorHandling.ReturnErrors,
                        token);
        }

        /// <summary>
        /// Sets up a conversation so that any item received within that conversation is no longer categorized.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="conversationId">The id of the conversation.</param>
        /// <param name="processSynchronously">Indicates whether the method should return only once disabling this rule and removing the categories from existing items
        /// in the conversation is completely done. If processSynchronously is false, the method returns immediately.</param>
        /// <returns></returns>
        public Task<ServiceResponseCollection<ServiceResponse>> DisableAlwaysCategorizeItemsInConversations(
            IEnumerable<ConversationId> conversationId,
            bool processSynchronously,
            CancellationToken token = default(CancellationToken))
        {
            return this.ApplyConversationAction(
                        ConversationActionType.AlwaysCategorize,
                        conversationId,
                        processSynchronously,
                        null,
                        false,
                        null,
                        ServiceErrorHandling.ReturnErrors,
                        token);
        }

        /// <summary>
        /// Sets up a conversation so that any item received within that conversation is always moved to Deleted Items folder.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="conversationId">The id of the conversation.</param>
        /// <param name="processSynchronously">Indicates whether the method should return only once enabling this rule and deleting existing items
        /// in the conversation is completely done. If processSynchronously is false, the method returns immediately.</param>
        /// <returns></returns>
        public Task<ServiceResponseCollection<ServiceResponse>> EnableAlwaysDeleteItemsInConversations(
            IEnumerable<ConversationId> conversationId,
            bool processSynchronously,
            CancellationToken token = default(CancellationToken))
        {
            return this.ApplyConversationAction(
                        ConversationActionType.AlwaysDelete,
                        conversationId,
                        processSynchronously,
                        null,
                        true,
                        null,
                        ServiceErrorHandling.ReturnErrors,
                        token);
        }

        /// <summary>
        /// Sets up a conversation so that any item received within that conversation is no longer moved to Deleted Items folder.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="conversationId">The id of the conversation.</param>
        /// <param name="processSynchronously">Indicates whether the method should return only once disabling this rule and restoring the items
        /// in the conversation is completely done. If processSynchronously is false, the method returns immediately.</param>
        /// <returns></returns>
        public Task<ServiceResponseCollection<ServiceResponse>> DisableAlwaysDeleteItemsInConversations(
            IEnumerable<ConversationId> conversationId,
            bool processSynchronously,
            CancellationToken token = default(CancellationToken))
        {
            return this.ApplyConversationAction(
                        ConversationActionType.AlwaysDelete,
                        conversationId,
                        processSynchronously,
                        null,
                        false,
                        null,
                        ServiceErrorHandling.ReturnErrors,
                        token);
        }

        /// <summary>
        /// Sets up a conversation so that any item received within that conversation is always moved to a specific folder.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="conversationId">The id of the conversation.</param>
        /// <param name="destinationFolderId">The Id of the folder to which conversation items should be moved.</param>
        /// <param name="processSynchronously">Indicates whether the method should return only once enabling this rule and moving existing items
        /// in the conversation is completely done. If processSynchronously is false, the method returns immediately.</param>
        /// <returns></returns>
        public Task<ServiceResponseCollection<ServiceResponse>> EnableAlwaysMoveItemsInConversations(
            IEnumerable<ConversationId> conversationId,
            FolderId destinationFolderId,
            bool processSynchronously,
            CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParam(destinationFolderId, "destinationFolderId");
            return this.ApplyConversationAction(
                        ConversationActionType.AlwaysMove,
                        conversationId,
                        processSynchronously,
                        null,
                        false,
                        destinationFolderId,
                        ServiceErrorHandling.ReturnErrors,
                        token);
        }

        /// <summary>
        /// Sets up a conversation so that any item received within that conversation is no longer moved to a specific folder.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="conversationIds">The conversation ids.</param>
        /// <param name="processSynchronously">Indicates whether the method should return only once disabling this rule is completely done.
        /// If processSynchronously is false, the method returns immediately.</param>
        /// <returns></returns>
        public Task<ServiceResponseCollection<ServiceResponse>> DisableAlwaysMoveItemsInConversations(
            IEnumerable<ConversationId> conversationIds,
            bool processSynchronously,
            CancellationToken token = default(CancellationToken))
        {
            return this.ApplyConversationAction(
                        ConversationActionType.AlwaysMove,
                        conversationIds,
                        processSynchronously,
                        null,
                        false,
                        null,
                        ServiceErrorHandling.ReturnErrors,
                        token);
        }

        /// <summary>
        /// Moves the items in the specified conversation to the specified destination folder.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="idLastSyncTimePairs">The pairs of Id of conversation whose
        /// items should be moved and the dateTime conversation was last synced
        /// (Items received after that dateTime will not be moved).</param>
        /// <param name="contextFolderId">The Id of the folder that contains the conversation.</param>
        /// <param name="destinationFolderId">The Id of the destination folder.</param>
        /// <returns></returns>
        public Task<ServiceResponseCollection<ServiceResponse>> MoveItemsInConversations(
            IEnumerable<KeyValuePair<ConversationId, DateTime?>> idLastSyncTimePairs,
            FolderId contextFolderId,
            FolderId destinationFolderId,
            CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParam(destinationFolderId, "destinationFolderId");
            return this.ApplyConversationOneTimeAction(
                ConversationActionType.Move,
                idLastSyncTimePairs,
                contextFolderId,
                destinationFolderId,
                null,
                null,
                null,
                null,
                null,
                null,
                ServiceErrorHandling.ReturnErrors,
                token);
        }

        /// <summary>
        /// Copies the items in the specified conversation to the specified destination folder.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="idLastSyncTimePairs">The pairs of Id of conversation whose
        /// items should be copied and the date and time conversation was last synced
        /// (Items received after that date will not be copied).</param>
        /// <param name="contextFolderId">The context folder id.</param>
        /// <param name="destinationFolderId">The destination folder id.</param>
        /// <returns></returns>
        public Task<ServiceResponseCollection<ServiceResponse>> CopyItemsInConversations(
            IEnumerable<KeyValuePair<ConversationId, DateTime?>> idLastSyncTimePairs,
            FolderId contextFolderId,
            FolderId destinationFolderId,
            CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParam(destinationFolderId, "destinationFolderId");
            return this.ApplyConversationOneTimeAction(
                ConversationActionType.Copy,
                idLastSyncTimePairs,
                contextFolderId,
                destinationFolderId,
                null,
                null,
                null,
                null,
                null,
                null,
                ServiceErrorHandling.ReturnErrors,
                token);
        }

        /// <summary>
        /// Deletes the items in the specified conversation. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="idLastSyncTimePairs">The pairs of Id of conversation whose
        /// items should be deleted and the date and time conversation was last synced
        /// (Items received after that date will not be deleted).</param>
        /// <param name="contextFolderId">The Id of the folder that contains the conversation.</param>
        /// <param name="deleteMode">The deletion mode.</param>
        /// <returns></returns>
        public Task<ServiceResponseCollection<ServiceResponse>> DeleteItemsInConversations(
            IEnumerable<KeyValuePair<ConversationId, DateTime?>> idLastSyncTimePairs,
            FolderId contextFolderId,
            DeleteMode deleteMode,
            CancellationToken token = default(CancellationToken))
        {
            return this.ApplyConversationOneTimeAction(
                ConversationActionType.Delete,
                idLastSyncTimePairs,
                contextFolderId,
                null,
                deleteMode,
                null,
                null,
                null,
                null,
                null,
                ServiceErrorHandling.ReturnErrors,
                token);
        }

        /// <summary>
        /// Sets the read state for items in conversation. Calling this method would
        /// result in call to EWS.
        /// </summary>
        /// <param name="idLastSyncTimePairs">The pairs of Id of conversation whose
        /// items should have their read state set and the date and time conversation
        /// was last synced (Items received after that date will not have their read
        /// state set).</param>
        /// <param name="contextFolderId">The Id of the folder that contains the conversation.</param>
        /// <param name="isRead">if set to <c>true</c>, conversation items are marked as read; otherwise they are marked as unread.</param>
        public Task<ServiceResponseCollection<ServiceResponse>> SetReadStateForItemsInConversations(
            IEnumerable<KeyValuePair<ConversationId, DateTime?>> idLastSyncTimePairs,
            FolderId contextFolderId,
            bool isRead,
            CancellationToken token = default(CancellationToken))
        {
            return this.ApplyConversationOneTimeAction(
                ConversationActionType.SetReadState,
                idLastSyncTimePairs,
                contextFolderId,
                null,
                null,
                isRead,
                null,
                null,
                null,
                null,
                ServiceErrorHandling.ReturnErrors,
                token);
        }

        /// <summary>
        /// Sets the read state for items in conversation. Calling this method would
        /// result in call to EWS.
        /// </summary>
        /// <param name="idLastSyncTimePairs">The pairs of Id of conversation whose
        /// items should have their read state set and the date and time conversation
        /// was last synced (Items received after that date will not have their read
        /// state set).</param>
        /// <param name="contextFolderId">The Id of the folder that contains the conversation.</param>
        /// <param name="isRead">if set to <c>true</c>, conversation items are marked as read; otherwise they are marked as unread.</param>
        /// <param name="suppressReadReceipts">if set to <c>true</c> read receipts are suppressed.</param>
        public Task<ServiceResponseCollection<ServiceResponse>> SetReadStateForItemsInConversations(
            IEnumerable<KeyValuePair<ConversationId, DateTime?>> idLastSyncTimePairs,
            FolderId contextFolderId,
            bool isRead,
            bool suppressReadReceipts,
            CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateMethodVersion(this, ExchangeVersion.Exchange2013, "SetReadStateForItemsInConversations");

            return this.ApplyConversationOneTimeAction(
                ConversationActionType.SetReadState,
                idLastSyncTimePairs,
                contextFolderId,
                null,
                null,
                isRead,
                null,
                null,
                null,
                suppressReadReceipts,
                ServiceErrorHandling.ReturnErrors,
                token);
        }

        /// <summary>
        /// Sets the retention policy for items in conversation. Calling this method would
        /// result in call to EWS.
        /// </summary>
        /// <param name="idLastSyncTimePairs">The pairs of Id of conversation whose
        /// items should have their retention policy set and the date and time conversation
        /// was last synced (Items received after that date will not have their retention
        /// policy set).</param>
        /// <param name="contextFolderId">The Id of the folder that contains the conversation.</param>
        /// <param name="retentionPolicyType">Retention policy type.</param>
        /// <param name="retentionPolicyTagId">Retention policy tag id.  Null will clear the policy.</param>
        public Task<ServiceResponseCollection<ServiceResponse>> SetRetentionPolicyForItemsInConversations(
            IEnumerable<KeyValuePair<ConversationId, DateTime?>> idLastSyncTimePairs,
            FolderId contextFolderId,
            RetentionType retentionPolicyType,
            Guid? retentionPolicyTagId,
            CancellationToken token = default(CancellationToken))
        {
            return this.ApplyConversationOneTimeAction(
                ConversationActionType.SetRetentionPolicy,
                idLastSyncTimePairs,
                contextFolderId,
                null,
                null,
                null,
                retentionPolicyType,
                retentionPolicyTagId,
                null,
                null,
                ServiceErrorHandling.ReturnErrors,
                token);
        }

        /// <summary>
        /// Sets flag status for items in conversation. Calling this method would result in call to EWS.
        /// </summary>
        /// <param name="idLastSyncTimePairs">The pairs of Id of conversation whose
        /// items should have their read state set and the date and time conversation
        /// was last synced (Items received after that date will not have their read
        /// state set).</param>
        /// <param name="contextFolderId">The Id of the folder that contains the conversation.</param>
        /// <param name="flagStatus">Flag status to apply to conversation items.</param>
        public Task<ServiceResponseCollection<ServiceResponse>> SetFlagStatusForItemsInConversations(
            IEnumerable<KeyValuePair<ConversationId, DateTime?>> idLastSyncTimePairs,
            FolderId contextFolderId,
            Flag flagStatus,
            CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateMethodVersion(this, ExchangeVersion.Exchange2013, "SetFlagStatusForItemsInConversations");

            return this.ApplyConversationOneTimeAction(
                ConversationActionType.Flag,
                idLastSyncTimePairs,
                contextFolderId,
                null,
                null,
                null,
                null,
                null,
                flagStatus,
                null,
                ServiceErrorHandling.ReturnErrors,
                token);
        }

        #endregion

        #region Id conversion operations

        /// <summary>
        /// Converts multiple Ids from one format to another in a single call to EWS.
        /// </summary>
        /// <param name="ids">The Ids to convert.</param>
        /// <param name="destinationFormat">The destination format.</param>
        /// <param name="errorHandling">Type of error handling to perform.</param>
        /// <returns>A ServiceResponseCollection providing conversion results for each specified Ids.</returns>
        private Task<ServiceResponseCollection<ConvertIdResponse>> InternalConvertIds(
            IEnumerable<AlternateIdBase> ids,
            IdFormat destinationFormat,
            ServiceErrorHandling errorHandling,
            CancellationToken token)
        {
            EwsUtilities.ValidateParamCollection(ids, "ids");

            ConvertIdRequest request = new ConvertIdRequest(this, errorHandling);

            request.Ids.AddRange(ids);
            request.DestinationFormat = destinationFormat;

            return request.ExecuteAsync(token);
        }

        /// <summary>
        /// Converts multiple Ids from one format to another in a single call to EWS.
        /// </summary>
        /// <param name="ids">The Ids to convert.</param>
        /// <param name="destinationFormat">The destination format.</param>
        /// <returns>A ServiceResponseCollection providing conversion results for each specified Ids.</returns>
        public Task<ServiceResponseCollection<ConvertIdResponse>> ConvertIds(IEnumerable<AlternateIdBase> ids, IdFormat destinationFormat, CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParamCollection(ids, "ids");

            return this.InternalConvertIds(
                ids,
                destinationFormat,
                ServiceErrorHandling.ReturnErrors,
                token);
        }

        /// <summary>
        /// Converts Id from one format to another in a single call to EWS.
        /// </summary>
        /// <param name="id">The Id to convert.</param>
        /// <param name="destinationFormat">The destination format.</param>
        /// <returns>The converted Id.</returns>
        public async Task<AlternateIdBase> ConvertId(AlternateIdBase id, IdFormat destinationFormat, CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParam(id, "id");

            ServiceResponseCollection<ConvertIdResponse> responses = await this.InternalConvertIds(
                new AlternateIdBase[] { id },
                destinationFormat,
                ServiceErrorHandling.ThrowOnError,
                token);

            return responses[0].ConvertedId;
        }

        #endregion

        #region Delegate management operations

        /// <summary>
        /// Adds delegates to a specific mailbox. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="mailbox">The mailbox to add delegates to.</param>
        /// <param name="meetingRequestsDeliveryScope">Indicates how meeting requests should be sent to delegates.</param>
        /// <param name="delegateUsers">The delegate users to add.</param>
        /// <returns>A collection of DelegateUserResponse objects providing the results of the operation.</returns>
        public Task<Collection<DelegateUserResponse>> AddDelegates(
            Mailbox mailbox,
            MeetingRequestsDeliveryScope? meetingRequestsDeliveryScope,
            params DelegateUser[] delegateUsers)
        {
            return this.AddDelegates(
                mailbox,
                meetingRequestsDeliveryScope,
                (IEnumerable<DelegateUser>)delegateUsers);
        }

        /// <summary>
        /// Adds delegates to a specific mailbox. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="mailbox">The mailbox to add delegates to.</param>
        /// <param name="meetingRequestsDeliveryScope">Indicates how meeting requests should be sent to delegates.</param>
        /// <param name="delegateUsers">The delegate users to add.</param>
        /// <returns>A collection of DelegateUserResponse objects providing the results of the operation.</returns>
        public async Task<Collection<DelegateUserResponse>> AddDelegates(
            Mailbox mailbox,
            MeetingRequestsDeliveryScope? meetingRequestsDeliveryScope,
            IEnumerable<DelegateUser> delegateUsers,
            CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParam(mailbox, "mailbox");
            EwsUtilities.ValidateParamCollection(delegateUsers, "delegateUsers");

            AddDelegateRequest request = new AddDelegateRequest(this);

            request.Mailbox = mailbox;
            request.DelegateUsers.AddRange(delegateUsers);
            request.MeetingRequestsDeliveryScope = meetingRequestsDeliveryScope;

            DelegateManagementResponse response = await request.Execute(token).ConfigureAwait(false);
            return response.DelegateUserResponses;
        }

        /// <summary>
        /// Updates delegates on a specific mailbox. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="mailbox">The mailbox to update delegates on.</param>
        /// <param name="meetingRequestsDeliveryScope">Indicates how meeting requests should be sent to delegates.</param>
        /// <param name="delegateUsers">The delegate users to update.</param>
        /// <returns>A collection of DelegateUserResponse objects providing the results of the operation.</returns>
        public Task<Collection<DelegateUserResponse>> UpdateDelegates(
            Mailbox mailbox,
            MeetingRequestsDeliveryScope? meetingRequestsDeliveryScope,
            CancellationToken token = default(CancellationToken),
            params DelegateUser[] delegateUsers)
        {
            return this.UpdateDelegates(
                mailbox,
                meetingRequestsDeliveryScope,
                (IEnumerable<DelegateUser>)delegateUsers);
        }

        /// <summary>
        /// Updates delegates on a specific mailbox. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="mailbox">The mailbox to update delegates on.</param>
        /// <param name="meetingRequestsDeliveryScope">Indicates how meeting requests should be sent to delegates.</param>
        /// <param name="delegateUsers">The delegate users to update.</param>
        /// <returns>A collection of DelegateUserResponse objects providing the results of the operation.</returns>
        public async Task<Collection<DelegateUserResponse>> UpdateDelegates(
            Mailbox mailbox,
            MeetingRequestsDeliveryScope? meetingRequestsDeliveryScope,
            IEnumerable<DelegateUser> delegateUsers,
            CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParam(mailbox, "mailbox");
            EwsUtilities.ValidateParamCollection(delegateUsers, "delegateUsers");

            UpdateDelegateRequest request = new UpdateDelegateRequest(this);

            request.Mailbox = mailbox;
            request.DelegateUsers.AddRange(delegateUsers);
            request.MeetingRequestsDeliveryScope = meetingRequestsDeliveryScope;

            DelegateManagementResponse response = await request.Execute(token).ConfigureAwait(false);
            return response.DelegateUserResponses;
        }

        /// <summary>
        /// Removes delegates on a specific mailbox. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="mailbox">The mailbox to remove delegates from.</param>
        /// <param name="userIds">The Ids of the delegate users to remove.</param>
        /// <returns>A collection of DelegateUserResponse objects providing the results of the operation.</returns>
        public Task<Collection<DelegateUserResponse>> RemoveDelegates(Mailbox mailbox, CancellationToken token = default(CancellationToken), params UserId[] userIds)
        {
            return this.RemoveDelegates(mailbox, (IEnumerable<UserId>)userIds, token);
        }

        /// <summary>
        /// Removes delegates on a specific mailbox. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="mailbox">The mailbox to remove delegates from.</param>
        /// <param name="userIds">The Ids of the delegate users to remove.</param>
        /// <returns>A collection of DelegateUserResponse objects providing the results of the operation.</returns>
        public async Task<Collection<DelegateUserResponse>> RemoveDelegates(Mailbox mailbox, IEnumerable<UserId> userIds, CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParam(mailbox, "mailbox");
            EwsUtilities.ValidateParamCollection(userIds, "userIds");

            RemoveDelegateRequest request = new RemoveDelegateRequest(this);

            request.Mailbox = mailbox;
            request.UserIds.AddRange(userIds);

            DelegateManagementResponse response = await request.Execute(token).ConfigureAwait(false);
            return response.DelegateUserResponses;
        }

        /// <summary>
        /// Retrieves the delegates of a specific mailbox. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="mailbox">The mailbox to retrieve the delegates of.</param>
        /// <param name="includePermissions">Indicates whether detailed permissions should be returned fro each delegate.</param>
        /// <param name="userIds">The optional Ids of the delegate users to retrieve.</param>
        /// <returns>A GetDelegateResponse providing the results of the operation.</returns>
        public Task<DelegateInformation> GetDelegates(
            Mailbox mailbox,
            bool includePermissions,
            CancellationToken token = default(CancellationToken),
            params UserId[] userIds)
        {
            return this.GetDelegates(
                mailbox,
                includePermissions,
                (IEnumerable<UserId>)userIds);
        }

        /// <summary>
        /// Retrieves the delegates of a specific mailbox. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="mailbox">The mailbox to retrieve the delegates of.</param>
        /// <param name="includePermissions">Indicates whether detailed permissions should be returned fro each delegate.</param>
        /// <param name="userIds">The optional Ids of the delegate users to retrieve.</param>
        /// <returns>A GetDelegateResponse providing the results of the operation.</returns>
        public async Task<DelegateInformation> GetDelegates(
            Mailbox mailbox,
            bool includePermissions,
            IEnumerable<UserId> userIds,
            CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParam(mailbox, "mailbox");

            GetDelegateRequest request = new GetDelegateRequest(this);

            request.Mailbox = mailbox;
            request.UserIds.AddRange(userIds);
            request.IncludePermissions = includePermissions;

            GetDelegateResponse response = await request.Execute(token).ConfigureAwait(false);
            DelegateInformation delegateInformation = new DelegateInformation(
                response.DelegateUserResponses,
                response.MeetingRequestsDeliveryScope);

            return delegateInformation;
        }

        #endregion

        #region UserConfiguration operations

        /// <summary>
        /// Creates a UserConfiguration.
        /// </summary>
        /// <param name="userConfiguration">The UserConfiguration.</param>
        internal System.Threading.Tasks.Task CreateUserConfiguration(UserConfiguration userConfiguration, CancellationToken token)
        {
            EwsUtilities.ValidateParam(userConfiguration, "userConfiguration");

            CreateUserConfigurationRequest request = new CreateUserConfigurationRequest(this);

            request.UserConfiguration = userConfiguration;

            return request.ExecuteAsync(token);
        }

        /// <summary>
        /// Deletes a UserConfiguration.
        /// </summary>
        /// <param name="name">Name of the UserConfiguration to retrieve.</param>
        /// <param name="parentFolderId">Id of the folder containing the UserConfiguration.</param>
        internal System.Threading.Tasks.Task DeleteUserConfiguration(string name, FolderId parentFolderId, CancellationToken token)
        {
            EwsUtilities.ValidateParam(name, "name");
            EwsUtilities.ValidateParam(parentFolderId, "parentFolderId");

            DeleteUserConfigurationRequest request = new DeleteUserConfigurationRequest(this);

            request.Name = name;
            request.ParentFolderId = parentFolderId;

            return request.ExecuteAsync(token);
        }

        /// <summary>
        /// Gets a UserConfiguration.
        /// </summary>
        /// <param name="name">Name of the UserConfiguration to retrieve.</param>
        /// <param name="parentFolderId">Id of the folder containing the UserConfiguration.</param>
        /// <param name="properties">Properties to retrieve.</param>
        /// <returns>A UserConfiguration.</returns>
        internal async Task<UserConfiguration> GetUserConfiguration(
            string name,
            FolderId parentFolderId,
            UserConfigurationProperties properties,
            CancellationToken token)
        {
            EwsUtilities.ValidateParam(name, "name");
            EwsUtilities.ValidateParam(parentFolderId, "parentFolderId");

            GetUserConfigurationRequest request = new GetUserConfigurationRequest(this);

            request.Name = name;
            request.ParentFolderId = parentFolderId;
            request.Properties = properties;

            return (await request.ExecuteAsync(token).ConfigureAwait(false))[0].UserConfiguration;
        }

        /// <summary>
        /// Loads the properties of the specified userConfiguration.
        /// </summary>
        /// <param name="userConfiguration">The userConfiguration containing properties to load.</param>
        /// <param name="properties">Properties to retrieve.</param>
        internal System.Threading.Tasks.Task LoadPropertiesForUserConfiguration(UserConfiguration userConfiguration, UserConfigurationProperties properties,
            CancellationToken token)
        {
            EwsUtilities.Assert(
                userConfiguration != null,
                "ExchangeService.LoadPropertiesForUserConfiguration",
                "userConfiguration is null");

            GetUserConfigurationRequest request = new GetUserConfigurationRequest(this);

            request.UserConfiguration = userConfiguration;
            request.Properties = properties;

            return request.ExecuteAsync(token);
        }

        /// <summary>
        /// Updates a UserConfiguration.
        /// </summary>
        /// <param name="userConfiguration">The UserConfiguration.</param>
        internal System.Threading.Tasks.Task UpdateUserConfiguration(UserConfiguration userConfiguration, CancellationToken token)
        {
            EwsUtilities.ValidateParam(userConfiguration, "userConfiguration");

            UpdateUserConfigurationRequest request = new UpdateUserConfigurationRequest(this);

            request.UserConfiguration = userConfiguration;

            return request.ExecuteAsync(token);
        }

        #endregion

        #region InboxRule operations
        /// <summary>
        /// Retrieves inbox rules of the authenticated user.
        /// </summary>
        /// <returns>A RuleCollection object containing the authenticated user's inbox rules.</returns>
        public async Task<RuleCollection> GetInboxRules(CancellationToken token = default(CancellationToken))
        {
            GetInboxRulesRequest request = new GetInboxRulesRequest(this);

            return (await request.Execute(token).ConfigureAwait(false)).Rules;
        }

        /// <summary>
        /// Retrieves the inbox rules of the specified user.
        /// </summary>
        /// <param name="mailboxSmtpAddress">The SMTP address of the user whose inbox rules should be retrieved.</param>
        /// <returns>A RuleCollection object containing the inbox rules of the specified user.</returns>
        public async Task<RuleCollection> GetInboxRules(string mailboxSmtpAddress, CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParam(mailboxSmtpAddress, "MailboxSmtpAddress");

            GetInboxRulesRequest request = new GetInboxRulesRequest(this);
            request.MailboxSmtpAddress = mailboxSmtpAddress;

            return (await request.Execute(token).ConfigureAwait(false)).Rules;
        }

        /// <summary>
        /// Updates the authenticated user's inbox rules by applying the specified operations.
        /// </summary>
        /// <param name="operations">The operations that should be applied to the user's inbox rules.</param>
        /// <param name="removeOutlookRuleBlob">Indicate whether or not to remove Outlook Rule Blob.</param>
        public System.Threading.Tasks.Task UpdateInboxRules(
            IEnumerable<RuleOperation> operations,
            bool removeOutlookRuleBlob,
            CancellationToken token = default(CancellationToken))
        {
            UpdateInboxRulesRequest request = new UpdateInboxRulesRequest(this);
            request.InboxRuleOperations = operations;
            request.RemoveOutlookRuleBlob = removeOutlookRuleBlob;
            return request.Execute(token);
        }

        /// <summary>
        /// Update the specified user's inbox rules by applying the specified operations.
        /// </summary>
        /// <param name="operations">The operations that should be applied to the user's inbox rules.</param>
        /// <param name="removeOutlookRuleBlob">Indicate whether or not to remove Outlook Rule Blob.</param>
        /// <param name="mailboxSmtpAddress">The SMTP address of the user whose inbox rules should be updated.</param>
        public System.Threading.Tasks.Task UpdateInboxRules(
            IEnumerable<RuleOperation> operations,
            bool removeOutlookRuleBlob,
            string mailboxSmtpAddress,
            CancellationToken token = default(CancellationToken))
        {
            UpdateInboxRulesRequest request = new UpdateInboxRulesRequest(this);
            request.InboxRuleOperations = operations;
            request.RemoveOutlookRuleBlob = removeOutlookRuleBlob;
            request.MailboxSmtpAddress = mailboxSmtpAddress;
            return request.Execute(token);
        }
        #endregion

        #region eDiscovery/Compliance operations

        /// <summary>
        /// Get dicovery search configuration
        /// </summary>
        /// <param name="searchId">Search Id</param>
        /// <param name="expandGroupMembership">True if want to expand group membership</param>
        /// <param name="inPlaceHoldConfigurationOnly">True if only want the inplacehold configuration</param>
        /// <returns>Service response object</returns>
        public Task<GetDiscoverySearchConfigurationResponse> GetDiscoverySearchConfiguration(string searchId, bool expandGroupMembership, bool inPlaceHoldConfigurationOnly, CancellationToken token = default(CancellationToken))
        {
            GetDiscoverySearchConfigurationRequest request = new GetDiscoverySearchConfigurationRequest(this);
            request.SearchId = searchId;
            request.ExpandGroupMembership = expandGroupMembership;
            request.InPlaceHoldConfigurationOnly = inPlaceHoldConfigurationOnly;

            return request.Execute(token);
        }

        /// <summary>
        /// Get searchable mailboxes
        /// </summary>
        /// <param name="searchFilter">Search filter</param>
        /// <param name="expandGroupMembership">True if want to expand group membership</param>
        /// <returns>Service response object</returns>
        public Task<GetSearchableMailboxesResponse> GetSearchableMailboxes(string searchFilter, bool expandGroupMembership, CancellationToken token = default(CancellationToken))
        {
            GetSearchableMailboxesRequest request = new GetSearchableMailboxesRequest(this);
            request.SearchFilter = searchFilter;
            request.ExpandGroupMembership = expandGroupMembership;

            return request.Execute(token);
        }

        /// <summary>
        /// Search mailboxes
        /// </summary>
        /// <param name="mailboxQueries">Collection of query and mailboxes</param>
        /// <param name="resultType">Search result type</param>
        /// <returns>Collection of search mailboxes response object</returns>
        public Task<ServiceResponseCollection<SearchMailboxesResponse>> SearchMailboxes(IEnumerable<MailboxQuery> mailboxQueries, SearchResultType resultType, CancellationToken token = default(CancellationToken))
        {
            SearchMailboxesRequest request = new SearchMailboxesRequest(this, ServiceErrorHandling.ReturnErrors);
            if (mailboxQueries != null)
            {
                request.SearchQueries.AddRange(mailboxQueries);
            }

            request.ResultType = resultType;

            return request.ExecuteAsync(token);
        }

        /// <summary>
        /// Search mailboxes
        /// </summary>
        /// <param name="mailboxQueries">Collection of query and mailboxes</param>
        /// <param name="resultType">Search result type</param>
        /// <param name="sortByProperty">Sort by property name</param>
        /// <param name="sortOrder">Sort order</param>
        /// <param name="pageSize">Page size</param>
        /// <param name="pageDirection">Page navigation direction</param>
        /// <param name="pageItemReference">Item reference used for paging</param>
        /// <returns>Collection of search mailboxes response object</returns>
        public Task<ServiceResponseCollection<SearchMailboxesResponse>> SearchMailboxes(
            IEnumerable<MailboxQuery> mailboxQueries,
            SearchResultType resultType,
            string sortByProperty,
            SortDirection sortOrder,
            int pageSize,
            SearchPageDirection pageDirection,
            string pageItemReference,
            CancellationToken token = default(CancellationToken))
        {
            SearchMailboxesRequest request = new SearchMailboxesRequest(this, ServiceErrorHandling.ReturnErrors);
            if (mailboxQueries != null)
            {
                request.SearchQueries.AddRange(mailboxQueries);
            }

            request.ResultType = resultType;
            request.SortByProperty = sortByProperty;
            request.SortOrder = sortOrder;
            request.PageSize = pageSize;
            request.PageDirection = pageDirection;
            request.PageItemReference = pageItemReference;

            return request.ExecuteAsync(token);
        }

        /// <summary>
        /// Search mailboxes
        /// </summary>
        /// <param name="searchParameters">Search mailboxes parameters</param>
        /// <returns>Collection of search mailboxes response object</returns>
        public Task<ServiceResponseCollection<SearchMailboxesResponse>> SearchMailboxes(SearchMailboxesParameters searchParameters, CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParam(searchParameters, "searchParameters");
            EwsUtilities.ValidateParam(searchParameters.SearchQueries, "searchParameters.SearchQueries");

            SearchMailboxesRequest request = this.CreateSearchMailboxesRequest(searchParameters);
            return request.ExecuteAsync(token);
        }

        /// <summary>
        /// Set hold on mailboxes
        /// </summary>
        /// <param name="holdId">Hold id</param>
        /// <param name="actionType">Action type</param>
        /// <param name="query">Query string</param>
        /// <param name="mailboxes">Collection of mailboxes</param>
        /// <returns>Service response object</returns>
        public Task<SetHoldOnMailboxesResponse> SetHoldOnMailboxes(string holdId, HoldAction actionType, string query, string[] mailboxes, CancellationToken token = default(CancellationToken))
        {
            SetHoldOnMailboxesRequest request = new SetHoldOnMailboxesRequest(this);
            request.HoldId = holdId;
            request.ActionType = actionType;
            request.Query = query;
            request.Mailboxes = mailboxes;
            request.InPlaceHoldIdentity = null;

            return request.Execute(token);
        }

        /// <summary>
        /// Set hold on mailboxes
        /// </summary>
        /// <param name="holdId">Hold id</param>
        /// <param name="actionType">Action type</param>
        /// <param name="query">Query string</param>
        /// <param name="inPlaceHoldIdentity">in-place hold identity</param>
        /// <returns>Service response object</returns>
        public Task<SetHoldOnMailboxesResponse> SetHoldOnMailboxes(string holdId, HoldAction actionType, string query, string inPlaceHoldIdentity)
        {
            return this.SetHoldOnMailboxes(holdId, actionType, query, inPlaceHoldIdentity, null);
        }

        /// <summary>
        /// Set hold on mailboxes
        /// </summary>
        /// <param name="holdId">Hold id</param>
        /// <param name="actionType">Action type</param>
        /// <param name="query">Query string</param>
        /// <param name="inPlaceHoldIdentity">in-place hold identity</param>
        /// <param name="itemHoldPeriod">item hold period</param>
        /// <returns>Service response object</returns>
        public Task<SetHoldOnMailboxesResponse> SetHoldOnMailboxes(string holdId, HoldAction actionType, string query, string inPlaceHoldIdentity, string itemHoldPeriod, CancellationToken token = default(CancellationToken))
        {
            SetHoldOnMailboxesRequest request = new SetHoldOnMailboxesRequest(this);
            request.HoldId = holdId;
            request.ActionType = actionType;
            request.Query = query;
            request.Mailboxes = null;
            request.InPlaceHoldIdentity = inPlaceHoldIdentity;
            request.ItemHoldPeriod = itemHoldPeriod;

            return request.Execute(token);
        }

        /// <summary>
        /// Set hold on mailboxes
        /// </summary>
        /// <param name="parameters">Set hold parameters</param>
        /// <returns>Service response object</returns>
        public Task<SetHoldOnMailboxesResponse> SetHoldOnMailboxes(SetHoldOnMailboxesParameters parameters, CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParam(parameters, "parameters");

            SetHoldOnMailboxesRequest request = new SetHoldOnMailboxesRequest(this);
            request.HoldId = parameters.HoldId;
            request.ActionType = parameters.ActionType;
            request.Query = parameters.Query;
            request.Mailboxes = parameters.Mailboxes;
            request.Language = parameters.Language;
            request.InPlaceHoldIdentity = parameters.InPlaceHoldIdentity;

            return request.Execute(token);
        }

        /// <summary>
        /// Get hold on mailboxes
        /// </summary>
        /// <param name="holdId">Hold id</param>
        /// <returns>Service response object</returns>
        public Task<GetHoldOnMailboxesResponse> GetHoldOnMailboxes(string holdId, CancellationToken token = default(CancellationToken))
        {
            GetHoldOnMailboxesRequest request = new GetHoldOnMailboxesRequest(this);
            request.HoldId = holdId;

            return request.Execute(token);
        }

        /// <summary>
        /// Get non indexable item details
        /// </summary>
        /// <param name="mailboxes">Array of mailbox legacy DN</param>
        /// <returns>Service response object</returns>
        public Task<GetNonIndexableItemDetailsResponse> GetNonIndexableItemDetails(string[] mailboxes)
        {
            return this.GetNonIndexableItemDetails(mailboxes, null, null, null);
        }

        /// <summary>
        /// Get non indexable item details
        /// </summary>
        /// <param name="mailboxes">Array of mailbox legacy DN</param>
        /// <param name="pageSize">The page size</param>
        /// <param name="pageItemReference">Page item reference</param>
        /// <param name="pageDirection">Page direction</param>
        /// <returns>Service response object</returns>
        public Task<GetNonIndexableItemDetailsResponse> GetNonIndexableItemDetails(string[] mailboxes, int? pageSize, string pageItemReference, SearchPageDirection? pageDirection)
        {
            GetNonIndexableItemDetailsParameters parameters = new GetNonIndexableItemDetailsParameters
            {
                Mailboxes = mailboxes,
                PageSize = pageSize,
                PageItemReference = pageItemReference,
                PageDirection = pageDirection,
                SearchArchiveOnly = false,
            };

            return GetNonIndexableItemDetails(parameters);
        }

        /// <summary>
        /// Get non indexable item details
        /// </summary>
        /// <param name="parameters">Get non indexable item details parameters</param>
        /// <returns>Service response object</returns>
        public Task<GetNonIndexableItemDetailsResponse> GetNonIndexableItemDetails(GetNonIndexableItemDetailsParameters parameters, CancellationToken token = default(CancellationToken))
        {
            GetNonIndexableItemDetailsRequest request = this.CreateGetNonIndexableItemDetailsRequest(parameters);

            return request.Execute(token);
        }

        /// <summary>
        /// Get non indexable item statistics
        /// </summary>
        /// <param name="mailboxes">Array of mailbox legacy DN</param>
        /// <returns>Service response object</returns>
        public Task<GetNonIndexableItemStatisticsResponse> GetNonIndexableItemStatistics(string[] mailboxes)
        {
            GetNonIndexableItemStatisticsParameters parameters = new GetNonIndexableItemStatisticsParameters
            {
                Mailboxes = mailboxes,
                SearchArchiveOnly = false,
            };

            return this.GetNonIndexableItemStatistics(parameters);
        }

        /// <summary>
        /// Get non indexable item statistics
        /// </summary>
        /// <param name="parameters">Get non indexable item statistics parameters</param>
        /// <returns>Service response object</returns>
        public Task<GetNonIndexableItemStatisticsResponse> GetNonIndexableItemStatistics(GetNonIndexableItemStatisticsParameters parameters, CancellationToken token = default(CancellationToken))
        {
            GetNonIndexableItemStatisticsRequest request = this.CreateGetNonIndexableItemStatisticsRequest(parameters);

            return request.Execute(token);
        }

        /// <summary>
        /// Create get non indexable item details request
        /// </summary>
        /// <param name="parameters">Get non indexable item details parameters</param>
        /// <returns>GetNonIndexableItemDetails request</returns>
        private GetNonIndexableItemDetailsRequest CreateGetNonIndexableItemDetailsRequest(GetNonIndexableItemDetailsParameters parameters)
        {
            EwsUtilities.ValidateParam(parameters, "parameters");
            EwsUtilities.ValidateParam(parameters.Mailboxes, "parameters.Mailboxes");

            GetNonIndexableItemDetailsRequest request = new GetNonIndexableItemDetailsRequest(this);
            request.Mailboxes = parameters.Mailboxes;
            request.PageSize = parameters.PageSize;
            request.PageItemReference = parameters.PageItemReference;
            request.PageDirection = parameters.PageDirection;
            request.SearchArchiveOnly = parameters.SearchArchiveOnly;

            return request;
        }

        /// <summary>
        /// Create get non indexable item statistics request
        /// </summary>
        /// <param name="parameters">Get non indexable item statistics parameters</param>
        /// <returns>Service response object</returns>
        private GetNonIndexableItemStatisticsRequest CreateGetNonIndexableItemStatisticsRequest(GetNonIndexableItemStatisticsParameters parameters)
        {
            EwsUtilities.ValidateParam(parameters, "parameters");
            EwsUtilities.ValidateParam(parameters.Mailboxes, "parameters.Mailboxes");

            GetNonIndexableItemStatisticsRequest request = new GetNonIndexableItemStatisticsRequest(this);
            request.Mailboxes = parameters.Mailboxes;
            request.SearchArchiveOnly = parameters.SearchArchiveOnly;

            return request;
        }

        /// <summary>
        /// Creates SearchMailboxesRequest from SearchMailboxesParameters
        /// </summary>
        /// <param name="searchParameters">search parameters</param>
        /// <returns>request object</returns>
        private SearchMailboxesRequest CreateSearchMailboxesRequest(SearchMailboxesParameters searchParameters)
        {
            SearchMailboxesRequest request = new SearchMailboxesRequest(this, ServiceErrorHandling.ReturnErrors);
            request.SearchQueries.AddRange(searchParameters.SearchQueries);
            request.ResultType = searchParameters.ResultType;
            request.PreviewItemResponseShape = searchParameters.PreviewItemResponseShape;
            request.SortByProperty = searchParameters.SortBy;
            request.SortOrder = searchParameters.SortOrder;
            request.Language = searchParameters.Language;
            request.PerformDeduplication = searchParameters.PerformDeduplication;
            request.PageSize = searchParameters.PageSize;
            request.PageDirection = searchParameters.PageDirection;
            request.PageItemReference = searchParameters.PageItemReference;

            return request;
        }
        #endregion

        #region MRM operations

        /// <summary>
        /// Get user retention policy tags.
        /// </summary>
        /// <returns>Service response object.</returns>
        public Task<GetUserRetentionPolicyTagsResponse> GetUserRetentionPolicyTags(CancellationToken token = default(CancellationToken))
        {
            GetUserRetentionPolicyTagsRequest request = new GetUserRetentionPolicyTagsRequest(this);

            return request.Execute(token);
        }

        #endregion

        #region Autodiscover

        /// <summary>
        /// Default implementation of AutodiscoverRedirectionUrlValidationCallback.
        /// Always returns true indicating that the URL can be used.
        /// </summary>
        /// <param name="redirectionUrl">The redirection URL.</param>
        /// <returns>Returns true.</returns>
        private bool DefaultAutodiscoverRedirectionUrlValidationCallback(string redirectionUrl)
        {
            throw new AutodiscoverLocalException(string.Format(Strings.AutodiscoverRedirectBlocked, redirectionUrl));
        }

        /// <summary>
        /// Initializes the Url property to the Exchange Web Services URL for the specified e-mail address by
        /// calling the Autodiscover service.
        /// </summary>
        /// <param name="emailAddress">The email address to use.</param>
        public System.Threading.Tasks.Task AutodiscoverUrl(string emailAddress)
        {
            return this.AutodiscoverUrl(emailAddress, this.DefaultAutodiscoverRedirectionUrlValidationCallback);
        }

        /// <summary>
        /// Initializes the Url property to the Exchange Web Services URL for the specified e-mail address by
        /// calling the Autodiscover service.
        /// </summary>
        /// <param name="emailAddress">The email address to use.</param>
        /// <param name="validateRedirectionUrlCallback">The callback used to validate redirection URL.</param>
        public async System.Threading.Tasks.Task AutodiscoverUrl(string emailAddress, AutodiscoverRedirectionUrlValidationCallback validateRedirectionUrlCallback)
        {
            Uri exchangeServiceUrl;

            if (this.RequestedServerVersion > ExchangeVersion.Exchange2007_SP1)
            {
                try
                {
                    exchangeServiceUrl = await this.GetAutodiscoverUrl(
                        emailAddress,
                        this.RequestedServerVersion,
                        validateRedirectionUrlCallback);

                    config.AdjustServiceUriFromCredentials(exchangeServiceUrl);
                    return;
                }
                catch (AutodiscoverLocalException ex)
                {
                    this.TraceMessage(
                        TraceFlags.AutodiscoverResponse,
                        string.Format("Autodiscover service call failed with error '{0}'. Will try legacy service", ex.Message));
                }
                catch (ServiceRemoteException ex)
                {
                    // Special case: if the caller's account is locked we want to return this exception, not continue.
                    if (ex is AccountIsLockedException)
                    {
                        throw;
                    }

                    this.TraceMessage(
                        TraceFlags.AutodiscoverResponse,
                        string.Format("Autodiscover service call failed with error '{0}'. Will try legacy service", ex.Message));
                }
            }

            // Try legacy Autodiscover provider
            exchangeServiceUrl = await this.GetAutodiscoverUrl(
                emailAddress,
                ExchangeVersion.Exchange2007_SP1,
                validateRedirectionUrlCallback);

            config.AdjustServiceUriFromCredentials(exchangeServiceUrl);
        }

        

        /// <summary>
        /// Gets the EWS URL from Autodiscover.
        /// </summary>
        /// <param name="emailAddress">The email address.</param>
        /// <param name="requestedServerVersion">Exchange version.</param>
        /// <param name="validateRedirectionUrlCallback">The validate redirection URL callback.</param>
        /// <returns>Ews URL</returns>
        private async Task<Uri> GetAutodiscoverUrl(
            string emailAddress,
            ExchangeVersion requestedServerVersion,
            AutodiscoverRedirectionUrlValidationCallback validateRedirectionUrlCallback)
        {
            AutodiscoverService autodiscoverService = new AutodiscoverService(this, requestedServerVersion)
            {
                RedirectionUrlValidationCallback = validateRedirectionUrlCallback,
                EnableScpLookup = this.EnableScpLookup
            };

            GetUserSettingsResponse response = await autodiscoverService.GetUserSettings(
                emailAddress,
                UserSettingName.InternalEwsUrl,
                UserSettingName.ExternalEwsUrl);

            switch (response.ErrorCode)
            {
                case AutodiscoverErrorCode.NoError:
                    return this.GetEwsUrlFromResponse(response, autodiscoverService.IsExternal.GetValueOrDefault(true));

                case AutodiscoverErrorCode.InvalidUser:
                    throw new ServiceRemoteException(
                        string.Format(Strings.InvalidUser, emailAddress));

                case AutodiscoverErrorCode.InvalidRequest:
                    throw new ServiceRemoteException(
                        string.Format(Strings.InvalidAutodiscoverRequest, response.ErrorMessage));

                default:
                    this.TraceMessage(
                        TraceFlags.AutodiscoverConfiguration,
                        string.Format("No EWS Url returned for user {0}, error code is {1}", emailAddress, response.ErrorCode));

                    throw new ServiceRemoteException(response.ErrorMessage);
            }
        }

        /// <summary>
        /// Gets the EWS URL from Autodiscover GetUserSettings response.
        /// </summary>
        /// <param name="response">The response.</param>
        /// <param name="isExternal">If true, Autodiscover call was made externally.</param>
        /// <returns>EWS URL.</returns>
        private Uri GetEwsUrlFromResponse(GetUserSettingsResponse response, bool isExternal)
        {
            string uriString;

            // Figure out which URL to use: Internal or External.
            // AutoDiscover may not return an external protocol. First try external, then internal.
            // Either protocol may be returned without a configured URL.
            if ((isExternal &&
                response.TryGetSettingValue<string>(UserSettingName.ExternalEwsUrl, out uriString)) &&
                !string.IsNullOrEmpty(uriString))
            {
                return new Uri(uriString);
            }
            else if ((response.TryGetSettingValue<string>(UserSettingName.InternalEwsUrl, out uriString) ||
                     response.TryGetSettingValue<string>(UserSettingName.ExternalEwsUrl, out uriString)) &&
                     !string.IsNullOrEmpty(uriString))
            {
                return new Uri(uriString);
            }

            // If Autodiscover doesn't return an internal or external EWS URL, throw an exception.
            throw new AutodiscoverLocalException(Strings.AutodiscoverDidNotReturnEwsUrl);
        }

        #endregion

        #region ClientAccessTokens

        /// <summary>
        /// GetClientAccessToken
        /// </summary>
        /// <param name="idAndTypes">Id and Types</param>
        /// <returns>A ServiceResponseCollection providing token results for each of the specified id and types.</returns>
        public Task<ServiceResponseCollection<GetClientAccessTokenResponse>> GetClientAccessToken(IEnumerable<KeyValuePair<string, ClientAccessTokenType>> idAndTypes)
        {
            GetClientAccessTokenRequest request = new GetClientAccessTokenRequest(this, ServiceErrorHandling.ReturnErrors);
            List<ClientAccessTokenRequest> requestList = new List<ClientAccessTokenRequest>();
            foreach (KeyValuePair<string, ClientAccessTokenType> idAndType in idAndTypes)
            {
                ClientAccessTokenRequest clientAccessTokenRequest = new ClientAccessTokenRequest(idAndType.Key, idAndType.Value);
                requestList.Add(clientAccessTokenRequest);
            }

            return this.GetClientAccessToken(requestList.ToArray());
        }

        /// <summary>
        /// GetClientAccessToken
        /// </summary>
        /// <param name="tokenRequests">Token requests array</param>
        /// <returns>A ServiceResponseCollection providing token results for each of the specified id and types.</returns>
        public Task<ServiceResponseCollection<GetClientAccessTokenResponse>> GetClientAccessToken(ClientAccessTokenRequest[] tokenRequests, CancellationToken token = default(CancellationToken))
        {
            GetClientAccessTokenRequest request = new GetClientAccessTokenRequest(this, ServiceErrorHandling.ReturnErrors);
            request.TokenRequests = tokenRequests;
            return request.ExecuteAsync(token);
        }

        #endregion

        #region Client Extensibility

        /// <summary>
        /// Get the app manifests.
        /// </summary>
        /// <returns>Collection of manifests</returns>
        public async Task<Collection<XmlDocument>> GetAppManifests(CancellationToken token = default(CancellationToken))
        {
            GetAppManifestsRequest request = new GetAppManifestsRequest(this);
            return (await request.Execute(token).ConfigureAwait(false)).Manifests;
        }

        /// <summary>
        /// Get the app manifests.  Works with Exchange 2013 SP1 or later EWS.
        /// </summary>
        /// <param name="apiVersionSupported">The api version supported by the client.</param>
        /// <param name="schemaVersionSupported">The schema version supported by the client.</param>
        /// <returns>Collection of manifests</returns>
        public async Task<Collection<ClientApp>> GetAppManifests(string apiVersionSupported, string schemaVersionSupported, CancellationToken token = default(CancellationToken))
        {
            GetAppManifestsRequest request = new GetAppManifestsRequest(this);
            request.ApiVersionSupported = apiVersionSupported;
            request.SchemaVersionSupported = schemaVersionSupported;

            return (await request.Execute(token).ConfigureAwait(false)).Apps;
        }

        /// <summary>
        /// Install App. 
        /// </summary>
        /// <param name="manifestStream">The manifest's plain text XML stream. 
        /// Notice: Stream has state. If you want this function read from the expected position of the stream,
        /// please make sure set read position by manifestStream.Position = expectedPosition.
        /// Be aware read manifestStream.Lengh puts stream's Position at stream end.
        /// If you retrieve manifestStream.Lengh before call this function, nothing will be read.
        /// When this function succeeds, manifestStream is closed. This is by EWS design to 
        /// release resource in timely manner. </param>
        /// <remarks>Exception will be thrown for errors. </remarks>
        public System.Threading.Tasks.Task InstallApp(Stream manifestStream, CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParam(manifestStream, "manifestStream");

            return this.InternalInstallApp(manifestStream, null, null, false, token);
        }

        /// <summary>
        /// Install App. 
        /// </summary>
        /// <param name="manifestStream">The manifest's plain text XML stream. 
        /// Notice: Stream has state. If you want this function read from the expected position of the stream,
        /// please make sure set read position by manifestStream.Position = expectedPosition.
        /// Be aware read manifestStream.Lengh puts stream's Position at stream end.
        /// If you retrieve manifestStream.Lengh before call this function, nothing will be read.
        /// When this function succeeds, manifestStream is closed. This is by EWS design to 
        /// release resource in timely manner. </param>
        /// <param name="marketplaceAssetId">The asset id of the addin in marketplace</param>
        /// <param name="marketplaceContentMarket">The target market for content</param>
        /// <param name="sendWelcomeEmail">Whether to send welcome email for the addin</param>
        /// <returns>True if the app was not already installed. False if it was not installed. Null if it is not a user mailbox.</returns>
        /// <remarks>Exception will be thrown for errors. </remarks>
        internal async Task<bool?> InternalInstallApp(Stream manifestStream, string marketplaceAssetId, string marketplaceContentMarket, bool sendWelcomeEmail, CancellationToken token)
        {
            EwsUtilities.ValidateParam(manifestStream, "manifestStream");

            InstallAppRequest request = new InstallAppRequest(this, manifestStream, marketplaceAssetId, marketplaceContentMarket, false);

            InstallAppResponse response = await request.Execute(token).ConfigureAwait(false);

            return response.WasFirstInstall;
        }

        /// <summary>
        /// Uninstall app. 
        /// </summary>
        /// <param name="id">App ID</param>
        /// <remarks>Exception will be thrown for errors. </remarks>
        public System.Threading.Tasks.Task UninstallApp(string id, CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParam(id, "id");

            UninstallAppRequest request = new UninstallAppRequest(this, id);

            return request.Execute(token);
        }

        /// <summary>
        /// Disable App.
        /// </summary>
        /// <param name="id">App ID</param>
        /// <param name="disableReason">Disable reason</param>
        /// <remarks>Exception will be thrown for errors. </remarks>
        public System.Threading.Tasks.Task DisableApp(string id, DisableReasonType disableReason, CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParam(id, "id");
            EwsUtilities.ValidateParam(disableReason, "disableReason");

            DisableAppRequest request = new DisableAppRequest(this, id, disableReason);

            return request.Execute(token);
        }

        /// <summary>
        /// Sets the consent state of an extension.
        /// </summary>
        /// <param name="id">Extension id.</param>
        /// <param name="state">Sets the consent state of an extension.</param>
        /// <remarks>Exception will be thrown for errors. </remarks>
        public System.Threading.Tasks.Task RegisterConsent(string id, ConsentState state, CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParam(id, "id");
            EwsUtilities.ValidateParam(state, "state");

            RegisterConsentRequest request = new RegisterConsentRequest(this, id, state);

            return request.Execute(token);
        }

        /// <summary>
        /// Get App Marketplace Url.
        /// </summary>
        /// <remarks>Exception will be thrown for errors. </remarks>
        public Task<string> GetAppMarketplaceUrl()
        {
            return GetAppMarketplaceUrl(null, null);
        }

        /// <summary>
        /// Get App Marketplace Url.  Works with Exchange 2013 SP1 or later EWS.
        /// </summary>
        /// <param name="apiVersionSupported">The api version supported by the client.</param>
        /// <param name="schemaVersionSupported">The schema version supported by the client.</param>
        /// <remarks>Exception will be thrown for errors. </remarks>
        public async Task<string> GetAppMarketplaceUrl(string apiVersionSupported, string schemaVersionSupported, CancellationToken token = default(CancellationToken))
        {
            GetAppMarketplaceUrlRequest request = new GetAppMarketplaceUrlRequest(this);
            request.ApiVersionSupported = apiVersionSupported;
            request.SchemaVersionSupported = schemaVersionSupported;

            return (await request.Execute(token).ConfigureAwait(false)).AppMarketplaceUrl;
        }

        /// <summary>
        /// Get the client extension data. This method is used in server-to-server calls to retrieve ORG extensions for
        /// admin powershell/UMC access and user's powershell/UMC access as well as user's activation for OWA/Outlook.
        /// This is expected to never be used or called directly from user client.
        /// </summary>
        /// <param name="requestedExtensionIds">An array of requested extension IDs to return.</param>
        /// <param name="shouldReturnEnabledOnly">Whether enabled extension only should be returned, e.g. for user's
        /// OWA/Outlook activation scenario.</param>
        /// <param name="isUserScope">Whether it's called from admin or user scope</param>
        /// <param name="userId">Specifies optional (if called with user scope) user identity. This will allow to do proper
        /// filtering in cases where admin installs an extension for specific users only</param>
        /// <param name="userEnabledExtensionIds">Optional list of org extension IDs which user enabled. This is necessary for
        /// proper result filtering on the server end. E.g. if admin installed N extensions but didn't enable them, it does not
        /// make sense to return manifests for those which user never enabled either. Used only when asked
        /// for enabled extension only (activation scenario).</param>
        /// <param name="userDisabledExtensionIds">Optional list of org extension IDs which user disabled. This is necessary for
        /// proper result filtering on the server end. E.g. if admin installed N optional extensions and enabled them, it does
        /// not make sense to retrieve manifests for extensions which user disabled for him or herself. Used only when asked
        /// for enabled extension only (activation scenario).</param>
        /// <param name="isDebug">Optional flag to indicate whether it is debug mode. 
        /// If it is, org master table in arbitration mailbox will be returned for debugging purpose.</param>
        /// <returns>Collection of ClientExtension objects</returns>
        public Task<GetClientExtensionResponse> GetClientExtension(
            StringList requestedExtensionIds,
            bool shouldReturnEnabledOnly,
            bool isUserScope,
            string userId,
            StringList userEnabledExtensionIds,
            StringList userDisabledExtensionIds,
            bool isDebug,
            CancellationToken token = default(CancellationToken))
        {
            GetClientExtensionRequest request = new GetClientExtensionRequest(
                this,
                requestedExtensionIds,
                shouldReturnEnabledOnly,
                isUserScope,
                userId,
                userEnabledExtensionIds,
                userDisabledExtensionIds,
                isDebug);

            return request.Execute(token);
        }

        /// <summary>
        /// Get the OME (i.e. Office Message Encryption) configuration data. This method is used in server-to-server calls to retrieve OME configuration
        /// </summary>
        /// <returns>OME Configuration response object</returns>
        public Task<GetOMEConfigurationResponse> GetOMEConfiguration(CancellationToken token = default(CancellationToken))
        {
            GetOMEConfigurationRequest request = new GetOMEConfigurationRequest(this);

            return request.Execute(token);
        }

        /// <summary>
        /// Set the OME (i.e. Office Message Encryption) configuration data. This method is used in server-to-server calls to set encryption configuration
        /// </summary>
        /// <param name="xml">The xml</param>
        public System.Threading.Tasks.Task SetOMEConfiguration(string xml, CancellationToken token = default(CancellationToken))
        {
            SetOMEConfigurationRequest request = new SetOMEConfigurationRequest(this, xml);

            return request.Execute(token);
        }

        /// <summary>
        /// Set the client extension data. This method is used in server-to-server calls to install/uninstall/configure ORG
        /// extensions to support admin's management of ORG extensions via powershell/UMC.
        /// </summary>
        /// <param name="actions">List of actions to execute.</param>
        public System.Threading.Tasks.Task SetClientExtension(List<SetClientExtensionAction> actions, CancellationToken token = default(CancellationToken))
        {
            SetClientExtensionRequest request = new SetClientExtensionRequest(this, actions);

            return request.ExecuteAsync(token);
        }

        #endregion

        #region Groups
        /// <summary>
        /// Gets the list of unified groups associated with the user
        /// </summary>
        /// <param name="requestedUnifiedGroupsSets">The Requested Unified Groups Sets</param>
        /// <param name="userSmtpAddress">The smtp address of accessing user.</param>
        /// <returns>UserUnified groups.</returns>
        public Task<Collection<UnifiedGroupsSet>> GetUserUnifiedGroups(
                            IEnumerable<RequestedUnifiedGroupsSet> requestedUnifiedGroupsSets,
                            string userSmtpAddress,
                            CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParam(requestedUnifiedGroupsSets, "requestedUnifiedGroupsSets");
            EwsUtilities.ValidateParam(userSmtpAddress, "userSmtpAddress");

            return this.GetUserUnifiedGroupsInternal(requestedUnifiedGroupsSets, userSmtpAddress, token);
        }

        /// <summary>
        /// Gets the list of unified groups associated with the user
        /// </summary>
        /// <param name="requestedUnifiedGroupsSets">The Requested Unified Groups Sets</param>
        /// <returns>UserUnified groups.</returns>
        public Task<Collection<UnifiedGroupsSet>> GetUserUnifiedGroups(IEnumerable<RequestedUnifiedGroupsSet> requestedUnifiedGroupsSets, CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParam(requestedUnifiedGroupsSets, "requestedUnifiedGroupsSets");
            return this.GetUserUnifiedGroupsInternal(requestedUnifiedGroupsSets, null, token);
        }

        /// <summary>
        /// Gets the list of unified groups associated with the user
        /// </summary>
        /// <param name="requestedUnifiedGroupsSets">The Requested Unified Groups Sets</param>
        /// <param name="userSmtpAddress">The smtp address of accessing user.</param>
        /// <returns>UserUnified groups.</returns>
        private async Task<Collection<UnifiedGroupsSet>> GetUserUnifiedGroupsInternal(
                            IEnumerable<RequestedUnifiedGroupsSet> requestedUnifiedGroupsSets,
                            string userSmtpAddress, CancellationToken token)
        {
            GetUserUnifiedGroupsRequest request = new GetUserUnifiedGroupsRequest(this);

            if (!string.IsNullOrEmpty(userSmtpAddress))
            {
                request.UserSmtpAddress = userSmtpAddress;
            }

            if (requestedUnifiedGroupsSets != null)
            {
                request.RequestedUnifiedGroupsSets = requestedUnifiedGroupsSets;
            }

            return (await request.Execute(token).ConfigureAwait(false)).GroupsSets;
        }

        /// <summary>
        /// Gets the UnifiedGroupsUnseenCount for the group specfied 
        /// </summary>
        /// <param name="groupMailboxSmtpAddress">The smtpaddress of group for which unseendata is desired</param>
        /// <param name="lastVisitedTimeUtc">The LastVisitedTimeUtc of group for which unseendata is desired</param>
        /// <returns>UnifiedGroupsUnseenCount</returns>
        public async Task<int> GetUnifiedGroupUnseenCount(string groupMailboxSmtpAddress, DateTime lastVisitedTimeUtc, CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParam(groupMailboxSmtpAddress, "groupMailboxSmtpAddress");

            GetUnifiedGroupUnseenCountRequest request = new GetUnifiedGroupUnseenCountRequest(
                this, lastVisitedTimeUtc, UnifiedGroupIdentityType.SmtpAddress, groupMailboxSmtpAddress);

            request.AnchorMailbox = groupMailboxSmtpAddress;

            return (await request.Execute(token).ConfigureAwait(false)).UnseenCount;
        }

        /// <summary>
        /// Sets the LastVisitedTime for the group specfied 
        /// </summary>
        /// <param name="groupMailboxSmtpAddress">The smtpaddress of group for which unseendata is desired</param>
        /// <param name="lastVisitedTimeUtc">The LastVisitedTimeUtc of group for which unseendata is desired</param>
        public System.Threading.Tasks.Task SetUnifiedGroupLastVisitedTime(string groupMailboxSmtpAddress, DateTime lastVisitedTimeUtc,
            CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParam(groupMailboxSmtpAddress, "groupMailboxSmtpAddress");

            SetUnifiedGroupLastVisitedTimeRequest request = new SetUnifiedGroupLastVisitedTimeRequest(this, lastVisitedTimeUtc, UnifiedGroupIdentityType.SmtpAddress, groupMailboxSmtpAddress);

            return request.Execute(token);
        }

        #endregion

        #region Diagnostic Method -- Only used by test

        /// <summary>
        /// Executes the diagnostic method.
        /// </summary>
        /// <param name="verb">The verb.</param>
        /// <param name="parameter">The parameter.</param>
        /// <returns></returns>
        internal async Task<XmlDocument> ExecuteDiagnosticMethod(string verb, XmlNode parameter, CancellationToken token)
        {
            ExecuteDiagnosticMethodRequest request = new ExecuteDiagnosticMethodRequest(this);
            request.Verb = verb;
            request.Parameter = parameter;

            return (await request.ExecuteAsync(token).ConfigureAwait(false))[0].ReturnValue;
        }
        #endregion

        #region Validation

        /// <summary>
        /// Validates this instance.
        /// </summary>
        internal override void Validate()
        {
            base.Validate();

            if (this.Url == null)
            {
                throw new ServiceLocalException(Strings.ServiceUrlMustBeSet);
            }

            if (this.PrivilegedUserId != null && this.ImpersonatedUserId != null)
            {
                throw new ServiceLocalException(Strings.CannotSetBothImpersonatedAndPrivilegedUser);
            }

            // only one of PrivilegedUserId|ImpersonatedUserId|ManagementRoles can be set.
        }


        #endregion

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="ExchangeService"/> class, targeting
        /// the latest supported version of EWS and scoped to the system's current time zone.
        /// </summary>
        public ExchangeService(IEWSHttpClient ewsHttpClient, IEWSStaticConfig iConfig) : base(iConfig.GetInstance())
        {
            this.ewsHttpClient = ewsHttpClient;
        }

        

        #endregion

        #region Utilities
        /// <summary>
        /// Creates an HttpWebRequest instance and initializes it with the appropriate parameters,
        /// based on the configuration of this service object.
        /// </summary>
        /// <param name="methodName">Name of the method.</param>
        /// <returns>
        /// An initialized instance of HttpWebRequest.
        /// </returns>
        internal IEwsHttpWebRequest PrepareHttpWebRequest(string methodName)
        {
            IEwsHttpWebRequest request = new EwsHttpWebRequest(this.ewsHttpClient.GetInstance());
            try
            {
                request.AdditionalMessageHeaders.AddRange(this.config.additionalMessageHeaders);
                request.ContentType = "text/xml; charset=utf-8";
                request.Accept = "text/xml";
                request.Method = "POST";
                request.UserAgent = this.UserAgent;
                request.ConnectionGroupName = this.config.connectionGroupName;

                lock (this.httpResponseHeaders)
                {
                    this.httpResponseHeaders.Clear();
                }

                return request;
            }
            catch (Exception)
            {
                request.Dispose();
                throw;
            }

        }

        /// <summary>
        /// Processes an HTTP error response.
        /// </summary>
        /// <param name="httpWebResponse">The HTTP web response.</param>
        /// <param name="webException">The web exception.</param>
        internal override void ProcessHttpErrorResponse(IEwsHttpWebResponse httpWebResponse, EwsHttpClientException webException)
        {
            this.InternalProcessHttpErrorResponse(
                httpWebResponse,
                webException,
                TraceFlags.EwsResponseHttpHeaders,
                TraceFlags.EwsResponse);
        }

        #endregion

        #region Properties

        /// <summary>
        /// Gets or sets the URL of the Exchange Web Services. 
        /// </summary>
        public Uri Url =>  this.config.ServerUrl;

        /// <summary>
        /// Gets or sets the Id of the user that EWS should impersonate. 
        /// </summary>
        public ImpersonatedUserId ImpersonatedUserId
        {
            get { return this.impersonatedUserId; }
            set { this.impersonatedUserId = value; }
        }

        /// <summary>
        /// Gets or sets the Id of the user that EWS should open his/her mailbox with privileged logon type. 
        /// </summary>
        internal PrivilegedUserId PrivilegedUserId
        {
            get { return this.privilegedUserId; }
            set { this.privilegedUserId = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public ManagementRoles ManagementRoles
        {
            get { return this.managementRoles; }
            set { this.managementRoles = value; }
        }

        /// <summary>
        /// Gets or sets the preferred culture for messages returned by the Exchange Web Services.
        /// </summary>
        public CultureInfo PreferredCulture
        {
            get { return this.preferredCulture; }
            set { this.preferredCulture = value; }
        }

        /// <summary>
        /// Gets or sets the DateTime precision for DateTime values returned from Exchange Web Services.
        /// </summary>
        public DateTimePrecision DateTimePrecision { get; } = DateTimePrecision.Default;

        /// <summary>
        /// Gets or sets a file attachment content handler.
        /// </summary>
        public IFileAttachmentContentHandler FileAttachmentContentHandler
        {
            get { return this.fileAttachmentContentHandler; }
            set { this.fileAttachmentContentHandler = value; }
        }

        /// <summary>
        /// Gets the time zone this service is scoped to.
        /// </summary>
        public new TimeZoneInfo TimeZone
        {
            get { return base.TimeZone; }
        }

        /// <summary>
        /// Provides access to the Unified Messaging functionalities.
        /// </summary>
        public UnifiedMessaging UnifiedMessaging
        {
            get
            {
                if (this.unifiedMessaging == null)
                {
                    this.unifiedMessaging = new UnifiedMessaging(this);
                }

                return this.unifiedMessaging;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the AutodiscoverUrl method should perform SCP (Service Connection Point) record lookup when determining
        /// the Autodiscover service URL.
        /// </summary>
        public bool EnableScpLookup
        {
            get { return this.enableScpLookup; }
            set { this.enableScpLookup = value; }
        }

        /// <summary>
        /// Exchange 2007 compatibility mode flag. (Off by default)
        /// </summary>
        private bool exchange2007CompatibilityMode;

        /// <summary>
        /// Gets or sets a value indicating whether Exchange2007 compatibility mode is enabled.
        /// </summary>
        /// <remarks>
        /// In order to support E12 servers, the Exchange2007CompatibilityMode property can be used 
        /// to indicate that we should use "Exchange2007" as the server version string rather than 
        /// Exchange2007_SP1.
        /// </remarks>
        internal bool Exchange2007CompatibilityMode
        {
            get { return this.exchange2007CompatibilityMode; }
            set { this.exchange2007CompatibilityMode = value; }
        }

        #endregion
    }
}