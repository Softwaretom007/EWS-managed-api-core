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
    using System;
    using System.Collections.Generic;
    using System.Text;
    using System.Threading;
    using System.Threading.Tasks;

    /// <summary>
    /// Represents a GetAppManifests request.
    /// </summary>
    internal sealed class GetAppManifestsRequest : SimpleServiceRequestBase
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="GetAppManifestsRequest"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        internal GetAppManifestsRequest(ExchangeService service)
            : base(service)
        {
        }

        /// <summary>
        /// Gets or sets the api version supported by the client.
        /// This tells Exchange service which app manifests should be returned based on the api version.
        /// </summary>
        /// <value>The Api version supported.</value>
        internal string ApiVersionSupported
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets the Schema version supported by the client.
        /// This tells Exchange service which app manifests should be returned based on the schema version.
        /// </summary>
        /// <value>The schema version supported.</value>
        internal string SchemaVersionSupported
        {
            get;
            set;
        }

        /// <summary>
        /// Validate request.
        /// </summary>
        internal override void Validate()
        {
            base.Validate();
            EwsUtilities.ValidateNonBlankStringParamAllowNull(this.ApiVersionSupported, "ApiVersionSupported");
            EwsUtilities.ValidateNonBlankStringParamAllowNull(this.SchemaVersionSupported, "SchemaVersionSupported");
        }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetXmlElementName()
        {
            return XmlElementNames.GetAppManifestsRequest;
        }

        /// <summary>
        /// Writes XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
        {
            if (!string.IsNullOrEmpty(this.ApiVersionSupported))
            {
                writer.WriteElementValue(XmlNamespace.Messages, "ApiVersionSupported", this.ApiVersionSupported);
            }

            if (!string.IsNullOrEmpty(this.SchemaVersionSupported))
            {
                writer.WriteElementValue(XmlNamespace.Messages, "SchemaVersionSupported", this.SchemaVersionSupported);
            }
        }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetResponseXmlElementName()
        {
            return XmlElementNames.GetAppManifestsResponse;
        }

        /// <summary>
        /// Parses the response.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>Response object.</returns>
        internal override object ParseResponse(EwsServiceXmlReader reader)
        {
            GetAppManifestsResponse response = new GetAppManifestsResponse();
            response.LoadFromXml(reader, XmlElementNames.GetAppManifestsResponse);
            return response;
        }

        /// <summary>
        /// Gets the request version.
        /// </summary>
        /// <returns>Earliest Exchange version in which this request is supported.</returns>
        internal override ExchangeVersion GetMinimumRequiredServerVersion()
        {
            return ExchangeVersion.Exchange2013;
        }

        /// <summary>
        /// Executes this request.
        /// </summary>
        /// <returns>Service response.</returns>
        internal async Task<GetAppManifestsResponse> Execute(CancellationToken token)
        {
            GetAppManifestsResponse serviceResponse = (GetAppManifestsResponse)await this.InternalExecuteAsync(token).ConfigureAwait(false);
            serviceResponse.ThrowIfNecessary();
            return serviceResponse;
        }
    }
}