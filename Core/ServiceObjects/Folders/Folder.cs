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
    /// Represents a generic folder.
    /// </summary>
    [ServiceObjectDefinition(XmlElementNames.Folder)]
    public class Folder : ServiceObject
    {
        /// <summary>
        /// Initializes an unsaved local instance of <see cref="Folder"/>. To bind to an existing folder, use Folder.Bind() instead.
        /// </summary>
        /// <param name="service">EWS service to which this object belongs.</param>
        public Folder(ExchangeService service)
            : base(service)
        {
        }

        /// <summary>
        /// Binds to an existing folder, whatever its actual type is, and loads the specified set of properties.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="service">The service to use to bind to the folder.</param>
        /// <param name="id">The Id of the folder to bind to.</param>
        /// <param name="propertySet">The set of properties to load.</param>
        /// <returns>A Folder instance representing the folder corresponding to the specified Id.</returns>
        public static Task<Folder> Bind(
            ExchangeService service,
            FolderId id,
            PropertySet propertySet,
            CancellationToken token = default(CancellationToken))
        {
            return service.BindToFolder<Folder>(id, propertySet, token);
        }

        /// <summary>
        /// Binds to an existing folder, whatever its actual type is, and loads its first class properties.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="service">The service to use to bind to the folder.</param>
        /// <param name="id">The Id of the folder to bind to.</param>
        /// <returns>A Folder instance representing the folder corresponding to the specified Id.</returns>
        public static Task<Folder> Bind(ExchangeService service, FolderId id, CancellationToken token = default(CancellationToken))
        {
            return Folder.Bind(
                service,
                id,
                PropertySet.FirstClassProperties, 
                token);
        }

        /// <summary>
        /// Binds to an existing folder, whatever its actual type is, and loads the specified set of properties.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="service">The service to use to bind to the folder.</param>
        /// <param name="name">The name of the folder to bind to.</param>
        /// <param name="propertySet">The set of properties to load.</param>
        /// <returns>A Folder instance representing the folder with the specified name.</returns>
        public static Task<Folder> Bind(
            ExchangeService service,
            WellKnownFolderName name,
            PropertySet propertySet,
            CancellationToken token = default(CancellationToken))
        {
            return Folder.Bind(
                service,
                new FolderId(name),
                propertySet,
                token);
        }

        /// <summary>
        /// Binds to an existing folder, whatever its actual type is, and loads its first class properties.
        /// Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="service">The service to use to bind to the folder.</param>
        /// <param name="name">The name of the folder to bind to.</param>
        /// <returns>A Folder instance representing the folder with the specified name.</returns>
        public static Task<Folder> Bind(ExchangeService service, WellKnownFolderName name, CancellationToken token = default(CancellationToken))
        {
            return Folder.Bind(
                service,
                new FolderId(name),
                PropertySet.FirstClassProperties,
                token);
        }

        /// <summary>
        /// Validates this instance.
        /// </summary>
        internal override void Validate()
        {
            base.Validate();

            // Validate folder permissions
            if (this.PropertyBag.Contains(FolderSchema.Permissions))
            {
                this.Permissions.Validate();
            }
        }

        /// <summary>
        /// Internal method to return the schema associated with this type of object.
        /// </summary>
        /// <returns>The schema associated with this type of object.</returns>
        internal override ServiceObjectSchema GetSchema()
        {
            return FolderSchema.Instance;
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
        /// Gets the name of the change XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetChangeXmlElementName()
        {
            return XmlElementNames.FolderChange;
        }

        /// <summary>
        /// Gets the name of the set field XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetSetFieldXmlElementName()
        {
            return XmlElementNames.SetFolderField;
        }

        /// <summary>
        /// Gets the name of the delete field XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetDeleteFieldXmlElementName()
        {
            return XmlElementNames.DeleteFolderField;
        }

        /// <summary>
        /// Loads the specified set of properties on the object.
        /// </summary>
        /// <param name="propertySet">The properties to load.</param>
        internal override Task<ServiceResponseCollection<ServiceResponse>> InternalLoad(PropertySet propertySet, CancellationToken token)
        {
            this.ThrowIfThisIsNew();

            return this.Service.LoadPropertiesForFolder(this, propertySet, token);
        }

        /// <summary>
        /// Deletes the object.
        /// </summary>
        /// <param name="deleteMode">The deletion mode.</param>
        /// <param name="sendCancellationsMode">Indicates whether meeting cancellation messages should be sent.</param>
        /// <param name="affectedTaskOccurrences">Indicate which occurrence of a recurring task should be deleted.</param>
        internal override Task<ServiceResponseCollection<ServiceResponse>> InternalDelete(
            DeleteMode deleteMode,
            SendCancellationsMode? sendCancellationsMode,
            AffectedTaskOccurrence? affectedTaskOccurrences,
            CancellationToken token)
        {
            this.ThrowIfThisIsNew();

            return this.Service.DeleteFolder( this.Id, deleteMode, token);
        }

        /// <summary>
        /// Deletes the folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="deleteMode">Deletion mode.</param>
        public Task<ServiceResponseCollection<ServiceResponse>> Delete(DeleteMode deleteMode, CancellationToken token = default(CancellationToken))
        {
            return this.InternalDelete(deleteMode, null, null, token);
        }

        /// <summary>
        /// Empties the folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="deleteMode">The deletion mode.</param>
        /// <param name="deleteSubFolders">Indicates whether sub-folders should also be deleted.</param>
        public Task<ServiceResponseCollection<ServiceResponse>> Empty(
            DeleteMode deleteMode,
            bool deleteSubFolders,
            CancellationToken token = default(CancellationToken))
        {
            this.ThrowIfThisIsNew();
            return this.Service.EmptyFolder(
                this.Id,
                deleteMode,
                deleteSubFolders, 
                token);
        }

        /// <summary>
        /// Marks all items in folder as read. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="suppressReadReceipts">If true, suppress sending read receipts for items.</param>
        public Task<ServiceResponseCollection<ServiceResponse>> MarkAllItemsAsRead(bool suppressReadReceipts, CancellationToken token = default(CancellationToken))
        {
            this.ThrowIfThisIsNew();
            return this.Service.MarkAllItemsAsRead(
                this.Id,
                true,
                suppressReadReceipts,
                token);
        }

        /// <summary>
        /// Marks all items in folder as read. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="suppressReadReceipts">If true, suppress sending read receipts for items.</param>
        public Task<ServiceResponseCollection<ServiceResponse>> MarkAllItemsAsUnread(bool suppressReadReceipts, CancellationToken token = default(CancellationToken))
        {
            this.ThrowIfThisIsNew();
            return this.Service.MarkAllItemsAsRead(
                this.Id,
                false,
                suppressReadReceipts,
                token);
        }

        /// <summary>
        /// Saves this folder in a specific folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="parentFolderId">The Id of the folder in which to save this folder.</param>
        public async System.Threading.Tasks.Task Save(FolderId parentFolderId, CancellationToken token = default(CancellationToken))
        {
            this.ThrowIfThisIsNotNew();

            EwsUtilities.ValidateParam(parentFolderId, "parentFolderId");

            if (this.IsDirty)
            {
                await this.Service.CreateFolder(this, parentFolderId, token);
            }
        }

        /// <summary>
        /// Saves this folder in a specific folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="parentFolderName">The name of the folder in which to save this folder.</param>
        public System.Threading.Tasks.Task Save(WellKnownFolderName parentFolderName)
        {
            return this.Save(new FolderId(parentFolderName));
        }

        /// <summary>
        /// Applies the local changes that have been made to this folder. Calling this method results in a call to EWS.
        /// </summary>
        public async System.Threading.Tasks.Task Update(CancellationToken token = default(CancellationToken))
        {
            if (this.IsDirty)
            {
                if (this.PropertyBag.GetIsUpdateCallNecessary())
                {
                    await this.Service.UpdateFolder(this, token);
                }
            }
        }

        /// <summary>
        /// Copies this folder into a specific folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="destinationFolderId">The Id of the folder in which to copy this folder.</param>
        /// <returns>A Folder representing the copy of this folder.</returns>
        public Task<Folder> Copy(FolderId destinationFolderId, CancellationToken token = default(CancellationToken))
        {
            this.ThrowIfThisIsNew();

            EwsUtilities.ValidateParam(destinationFolderId, "destinationFolderId");

            return this.Service.CopyFolder(this.Id, destinationFolderId, token);
        }

        /// <summary>
        /// Copies this folder into the specified folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="destinationFolderName">The name of the folder in which to copy this folder.</param>
        /// <returns>A Folder representing the copy of this folder.</returns>
        public Task<Folder> Copy(WellKnownFolderName destinationFolderName)
        {
            return this.Copy(new FolderId(destinationFolderName));
        }

        /// <summary>
        /// Moves this folder to a specific folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="destinationFolderId">The Id of the folder in which to move this folder.</param>
        /// <returns>A new folder representing this folder in its new location. After Move completes, this folder does not exist anymore.</returns>
        public Task<Folder> Move(FolderId destinationFolderId, CancellationToken token = default(CancellationToken))
        {
            this.ThrowIfThisIsNew();

            EwsUtilities.ValidateParam(destinationFolderId, "destinationFolderId");

            return this.Service.MoveFolder(this.Id, destinationFolderId, token);
        }

        /// <summary>
        /// Moves this folder to the specified folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="destinationFolderName">The name of the folder in which to move this folder.</param>
        /// <returns>A new folder representing this folder in its new location. After Move completes, this folder does not exist anymore.</returns>
        public Task<Folder> Move(WellKnownFolderName destinationFolderName)
        {
            return this.Move(new FolderId(destinationFolderName));
        }

        /// <summary>
        /// Find items.
        /// </summary>
        /// <typeparam name="TItem">The type of the item.</typeparam>
        /// <param name="queryString">query string to be used for indexed search</param>
        /// <param name="view">The view controlling the number of items returned.</param>
        /// <param name="groupBy">The group by.</param>
        /// <returns>FindItems response collection.</returns>
        internal Task<ServiceResponseCollection<FindItemResponse<TItem>>> InternalFindItems<TItem>(
            string queryString,
            ViewBase view,
            Grouping groupBy,
            CancellationToken token)
            where TItem : Item
        {
            this.ThrowIfThisIsNew();

            return this.Service.FindItems<TItem>(
                new FolderId[] { this.Id },
                null, /* searchFilter */
                queryString,
                view,
                groupBy,
                ServiceErrorHandling.ThrowOnError,
                token);
        }

        /// <summary>
        /// Find items.
        /// </summary>
        /// <typeparam name="TItem">The type of the item.</typeparam>
        /// <param name="searchFilter">The search filter. Available search filter classes
        /// include SearchFilter.IsEqualTo, SearchFilter.ContainsSubstring and 
        /// SearchFilter.SearchFilterCollection</param>
        /// <param name="view">The view controlling the number of items returned.</param>
        /// <param name="groupBy">The group by.</param>
        /// <returns>FindItems response collection.</returns>
        internal Task<ServiceResponseCollection<FindItemResponse<TItem>>> InternalFindItems<TItem>(
            SearchFilter searchFilter,
            ViewBase view,
            Grouping groupBy,
            CancellationToken token)
            where TItem : Item
        {
            this.ThrowIfThisIsNew();

            return this.Service.FindItems<TItem>(
                new FolderId[] { this.Id },
                searchFilter,
                null, /* queryString */
                view,
                groupBy,
                ServiceErrorHandling.ThrowOnError,
                token);
        }

        /// <summary>
        /// Obtains a list of items by searching the contents of this folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="searchFilter">The search filter. Available search filter classes
        /// include SearchFilter.IsEqualTo, SearchFilter.ContainsSubstring and 
        /// SearchFilter.SearchFilterCollection</param>
        /// <param name="view">The view controlling the number of items returned.</param>
        /// <returns>An object representing the results of the search operation.</returns>
        public async Task<FindItemsResults<Item>> FindItems(SearchFilter searchFilter, ItemView view, CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParamAllowNull(searchFilter, "searchFilter");

            ServiceResponseCollection<FindItemResponse<Item>> responses = await this.InternalFindItems<Item>(
                searchFilter,
                view, 
                null /* groupBy */,
                token).ConfigureAwait(false);

            return responses[0].Results;
        }

        /// <summary>
        /// Obtains a list of items by searching the contents of this folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="queryString">query string to be used for indexed search</param>
        /// <param name="view">The view controlling the number of items returned.</param>
        /// <returns>An object representing the results of the search operation.</returns>
        public async Task<FindItemsResults<Item>> FindItems(string queryString, ItemView view, CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParamAllowNull(queryString, "queryString");

            ServiceResponseCollection<FindItemResponse<Item>> responses = await this.InternalFindItems<Item>(queryString, view, null /* groupBy */, token).ConfigureAwait(false);

            return responses[0].Results;
        }

        /// <summary>
        /// Obtains a list of items by searching the contents of this folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="view">The view controlling the number of items returned.</param>
        /// <returns>An object representing the results of the search operation.</returns>
        public async Task<FindItemsResults<Item>> FindItems(ItemView view, CancellationToken token = default(CancellationToken))
        {
            ServiceResponseCollection<FindItemResponse<Item>> responses = await this.InternalFindItems<Item>(
                (SearchFilter)null,
                view,
                null /* groupBy */ , token).ConfigureAwait(false);

            return responses[0].Results;
        }

        /// <summary>
        /// Obtains a grouped list of items by searching the contents of this folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="searchFilter">The search filter. Available search filter classes
        /// include SearchFilter.IsEqualTo, SearchFilter.ContainsSubstring and 
        /// SearchFilter.SearchFilterCollection</param>
        /// <param name="view">The view controlling the number of items returned.</param>
        /// <param name="groupBy">The grouping criteria.</param>
        /// <returns>A collection of grouped items representing the contents of this folder.</returns>
        public async Task<GroupedFindItemsResults<Item>> FindItems(SearchFilter searchFilter, ItemView view, Grouping groupBy, CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParam(groupBy, "groupBy");
            EwsUtilities.ValidateParamAllowNull(searchFilter, "searchFilter");

            ServiceResponseCollection<FindItemResponse<Item>> responses = await this.InternalFindItems<Item>(
                searchFilter,
                view, 
                groupBy,
                token).ConfigureAwait(false);

            return responses[0].GroupedFindResults;
        }

        /// <summary>
        /// Obtains a grouped list of items by searching the contents of this folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="queryString">query string to be used for indexed search</param>
        /// <param name="view">The view controlling the number of items returned.</param>
        /// <param name="groupBy">The grouping criteria.</param>
        /// <returns>A collection of grouped items representing the contents of this folder.</returns>
        public async Task<GroupedFindItemsResults<Item>> FindItems(string queryString, ItemView view, Grouping groupBy, CancellationToken token = default(CancellationToken))
        {
            EwsUtilities.ValidateParam(groupBy, "groupBy");

            ServiceResponseCollection<FindItemResponse<Item>> responses = await this.InternalFindItems<Item>(queryString, view, groupBy, token).ConfigureAwait(false);

            return responses[0].GroupedFindResults;
        }

        /// <summary>
        /// Obtains a list of folders by searching the sub-folders of this folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="view">The view controlling the number of folders returned.</param>
        /// <returns>An object representing the results of the search operation.</returns>
        public Task<FindFoldersResults> FindFolders(FolderView view)
        {
            this.ThrowIfThisIsNew();

            return this.Service.FindFolders(this.Id, view);
        }

        /// <summary>
        /// Obtains a list of folders by searching the sub-folders of this folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="searchFilter">The search filter. Available search filter classes
        /// include SearchFilter.IsEqualTo, SearchFilter.ContainsSubstring and 
        /// SearchFilter.SearchFilterCollection</param>
        /// <param name="view">The view controlling the number of folders returned.</param>
        /// <returns>An object representing the results of the search operation.</returns>
        public Task<FindFoldersResults> FindFolders(SearchFilter searchFilter, FolderView view)
        {
            this.ThrowIfThisIsNew();

            return this.Service.FindFolders(this.Id, searchFilter, view);
        }

        /// <summary>
        /// Obtains a grouped list of items by searching the contents of this folder. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="view">The view controlling the number of items returned.</param>
        /// <param name="groupBy">The grouping criteria.</param>
        /// <returns>A collection of grouped items representing the contents of this folder.</returns>
        public Task<GroupedFindItemsResults<Item>> FindItems(ItemView view, Grouping groupBy)
        {
            EwsUtilities.ValidateParam(groupBy, "groupBy");

            return this.FindItems(
                (SearchFilter)null,
                view,
                groupBy);
        }

        /// <summary>
        /// Get the property definition for the Id property.
        /// </summary>
        /// <returns>A PropertyDefinition instance.</returns>
        internal override PropertyDefinition GetIdPropertyDefinition()
        {
            return FolderSchema.Id;
        }

        /// <summary>
        /// Sets the extended property.
        /// </summary>
        /// <param name="extendedPropertyDefinition">The extended property definition.</param>
        /// <param name="value">The value.</param>
        public void SetExtendedProperty(ExtendedPropertyDefinition extendedPropertyDefinition, object value)
        {
            this.ExtendedProperties.SetExtendedProperty(extendedPropertyDefinition, value);
        }

        /// <summary>
        /// Removes an extended property.
        /// </summary>
        /// <param name="extendedPropertyDefinition">The extended property definition.</param>
        /// <returns>True if property was removed.</returns>
        public bool RemoveExtendedProperty(ExtendedPropertyDefinition extendedPropertyDefinition)
        {
            return this.ExtendedProperties.RemoveExtendedProperty(extendedPropertyDefinition);
        }

        /// <summary>
        /// Gets a list of extended properties defined on this object.
        /// </summary>
        /// <returns>Extended properties collection.</returns>
        internal override ExtendedPropertyCollection GetExtendedProperties()
        {
            return this.ExtendedProperties;
        }

        #region Properties

        /// <summary>
        /// Gets the Id of the folder.
        /// </summary>
        public FolderId Id
        {
            get { return (FolderId)this.PropertyBag[this.GetIdPropertyDefinition()]; }
        }

        /// <summary>
        /// Gets the Id of this folder's parent folder.
        /// </summary>
        public FolderId ParentFolderId
        {
            get { return (FolderId)this.PropertyBag[FolderSchema.ParentFolderId]; }
        }

        /// <summary>
        /// Gets the number of child folders this folder has.
        /// </summary>
        public int ChildFolderCount
        {
            get { return (int)this.PropertyBag[FolderSchema.ChildFolderCount]; }
        }

        /// <summary>
        /// Gets or sets the display name of the folder.
        /// </summary>
        public string DisplayName
        {
            get { return (string)this.PropertyBag[FolderSchema.DisplayName]; }
            set { this.PropertyBag[FolderSchema.DisplayName] = value; }
        }

        /// <summary>
        /// Gets or sets the custom class name of this folder.
        /// </summary>
        public string FolderClass
        {
            get { return (string)this.PropertyBag[FolderSchema.FolderClass]; }
            set { this.PropertyBag[FolderSchema.FolderClass] = value; }
        }

        /// <summary>
        /// Gets the total number of items contained in the folder.
        /// </summary>
        public int TotalCount
        {
            get { return (int)this.PropertyBag[FolderSchema.TotalCount]; }
        }

        /// <summary>
        /// Gets a list of extended properties associated with the folder.
        /// </summary>
        public ExtendedPropertyCollection ExtendedProperties
        {
            get { return (ExtendedPropertyCollection)this.PropertyBag[ServiceObjectSchema.ExtendedProperties]; }
        }

        /// <summary>
        /// Gets the Email Lifecycle Management (ELC) information associated with the folder.
        /// </summary>
        public ManagedFolderInformation ManagedFolderInformation
        {
            get { return (ManagedFolderInformation)this.PropertyBag[FolderSchema.ManagedFolderInformation]; }
        }

        /// <summary>
        /// Gets a value indicating the effective rights the current authenticated user has on the folder.
        /// </summary>
        public EffectiveRights EffectiveRights
        {
            get { return (EffectiveRights)this.PropertyBag[FolderSchema.EffectiveRights]; }
        }

        /// <summary>
        /// Gets a list of permissions for the folder.
        /// </summary>
        public FolderPermissionCollection Permissions
        {
            get { return (FolderPermissionCollection)this.PropertyBag[FolderSchema.Permissions]; }
        }

        /// <summary>
        /// Gets the number of unread items in the folder.
        /// </summary>
        public int UnreadCount
        {
            get { return (int)this.PropertyBag[FolderSchema.UnreadCount]; }
        }

        /// <summary>
        /// Gets or sets the policy tag.
        /// </summary>
        public PolicyTag PolicyTag
        {
            get { return (PolicyTag)this.PropertyBag[FolderSchema.PolicyTag]; }
            set { this.PropertyBag[FolderSchema.PolicyTag] = value; }
        }

        /// <summary>
        /// Gets or sets the archive tag.
        /// </summary>
        public ArchiveTag ArchiveTag
        {
            get { return (ArchiveTag)this.PropertyBag[FolderSchema.ArchiveTag]; }
            set { this.PropertyBag[FolderSchema.ArchiveTag] = value; }
        }

        /// <summary>
        /// Gets the well known name of this folder, if any, as a string.
        /// </summary>
        /// <value>The well known name of this folder as a string, or null if this folder isn't a well known folder.</value>
        public string WellKnownFolderNameAsString
        {
            get { return (string)this.PropertyBag[FolderSchema.WellKnownFolderName]; }
        }

        /// <summary>
        /// Gets the well known name of this folder, if any.
        /// </summary>
        /// <value>The well known name of this folder, or null if this folder isn't a well known folder.</value>
        public WellKnownFolderName? WellKnownFolderName
        {
            get
            {
                WellKnownFolderName result;

                if (EwsUtilities.TryParse<WellKnownFolderName>(this.WellKnownFolderNameAsString, out result))
                {
                    return result;
                }
                else
                {
                    return null;
                }
            }
        }

        #endregion
    }
}