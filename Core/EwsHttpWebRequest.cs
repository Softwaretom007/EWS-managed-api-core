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
    using Microsoft.Exchange.WebServices.NETStandard.Core;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Net;
    using System.Net.Http;
    using System.Net.Http.Headers;
    using System.Net.Security;
    using System.Security.Cryptography.X509Certificates;
    using System.Text;
    using System.Threading;
    using System.Threading.Tasks;

    /// <summary>
    /// Represents an implementation of the IEwsHttpWebRequest interface that uses HttpWebRequest.
    /// </summary>
    internal class EwsHttpWebRequest : IEwsHttpWebRequest
    {
        /// <summary>
        /// Underlying HttpWebRequest.
        /// </summary>
        readonly HttpClient _httpClient;

        /// <summary>
        /// Initializes a new instance of the <see cref="EwsHttpWebRequest"/> class.
        /// </summary>
        /// <param name="uri">The URI.</param>
        public EwsHttpWebRequest(EWSHttpClient ewsClient)
        {
            _httpClient = ewsClient.Client;
        }

        #region IEwsHttpWebRequest Members

        /// <summary>
        /// Aborts this instance.
        /// </summary>
        public void Abort()
        {
            _httpClient.CancelPendingRequests();
        }

        /// <summary>
        /// Gets a <see cref="T:System.IO.Stream"/> object to use to write request data.
        /// </summary>
        /// <returns>
        /// A <see cref="T:System.IO.Stream"/> to use to write request data.
        /// </returns>
        public string Content { get; set; }

        /// <summary>
        /// Returns a response from an Internet resource.
        /// </summary>
        /// <returns>
        /// A <see cref="T:System.Net.HttpWebResponse"/> that contains the response from the Internet resource.
        /// </returns>
        public async Task<IEwsHttpWebResponse> GetResponse(CancellationToken token)
        {
            var message = new HttpRequestMessage(new HttpMethod(Method), RequestUri);
            message.Content = new StringContent(Content);
            if (!string.IsNullOrEmpty(ContentType))
            {
                message.Content.Headers.ContentType = null;
                message.Content.Headers.TryAddWithoutValidation("Content-Type", ContentType);
            }

            if (!string.IsNullOrEmpty(UserAgent))
            {
                message.Headers.UserAgent.Clear();
                message.Headers.UserAgent.ParseAdd(UserAgent);
            }

            if (!string.IsNullOrEmpty(Accept))
            {
                message.Headers.Accept.Clear();
                message.Headers.Accept.ParseAdd(Accept);
            }
            this.AdditionalMessageHeaders.ForEach(x => message.Headers.TryAddWithoutValidation(x.Key, x.Value));

            HttpResponseMessage response = null;
            try
            {
                response = await _httpClient.SendAsync(message, token);
            }
            catch (Exception exception)
            {
                throw new EwsHttpClientException(exception);
            }
            if (!response.IsSuccessStatusCode)
                throw new EwsHttpClientException(response);

            return new EwsHttpWebResponse(response);
        }

        public void Dispose()
        {
            //_httpClient.Dispose();
        }

        /// <summary>
        /// Gets or sets the value of the Accept HTTP header.
        /// </summary>
        /// <returns>The value of the Accept HTTP header. The default value is null.</returns>
        public string Accept
        {
            get;
            set;
        }



        ///// <summary>
        ///// Gets or sets the client certificates.
        ///// </summary>
        ///// <value></value>
        ///// <returns>The collection of X509 client certificates.</returns>
        //public X509CertificateCollection ClientCertificates
        //{
        //    get;
        //    set;
        //}

        /// <summary>
        /// Gets or sets the value of the Content-type HTTP header.
        /// </summary>
        /// <returns>The value of the Content-type HTTP header. The default value is null.</returns>
        public string ContentType
        {
            get;
            set;
        }


        /// <summary>
        /// Specifies a collection of the name/value pairs that make up the HTTP headers.
        /// </summary>
        /// <returns>A <see cref="T:System.Net.WebHeaderCollection"/> that contains the name/value pairs that make up the headers for the HTTP request.</returns>
        string IEwsHttpWebRequest.Headers
        {
            get { return EwsUtilities.FormatHttpHeaders(this._httpClient.DefaultRequestHeaders); }
        }

        public List<KeyValuePair<string, IEnumerable<string>>> AdditionalMessageHeaders { get; set; } = new List<KeyValuePair<string, IEnumerable<string>>>();



        /// <summary>
        /// Gets or sets the method for the request.
        /// </summary>
        /// <returns>The request method to use to contact the Internet resource. The default value is GET.</returns>
        /// <exception cref="T:System.ArgumentException">No method is supplied.-or- The method string contains invalid characters. </exception>
        public string Method
        {
            get;
            set;
        }


        /// <summary>
        /// Gets the original Uniform Resource Identifier (URI) of the request.
        /// </summary>
        /// <returns>A <see cref="T:System.Uri"/> that contains the URI of the Internet resource passed to the <see cref="M:System.Net.WebRequest.Create(System.String)"/> method.</returns>
        public Uri RequestUri
        {
            get;
        }

        /// <summary>
        /// Gets or sets the value of the User-agent HTTP header.
        /// </summary>
        /// <returns>The value of the User-agent HTTP header. The default value is null.The value for this property is stored in <see cref="T:System.Net.WebHeaderCollection"/>. If WebHeaderCollection is set, the property value is lost.</returns>
        public string UserAgent
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets the name of the connection group for the request. 
        /// </summary>
        public string ConnectionGroupName
        {
            get;
            set;
        }
       
        #endregion
    }
}