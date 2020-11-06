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
    using System.IO;
    using System.Xml;
    using System.Xml.XPath;

    /// <summary>
    /// Factory methods to safely instantiate XXE vulnerable object.
    /// </summary>
    internal class SafeXmlFactory
    {
        #region Members
        /// <summary>
        /// Safe xml reader settings.
        /// </summary>
        private static XmlReaderSettings defaultSettings = new XmlReaderSettings()
        {
            Async = true,
            DtdProcessing = DtdProcessing.Prohibit,
        };
        #endregion

        #region XmlTextReader
        /// <summary>
        /// Initializes a new instance of the XmlTextReader class with the specified stream.
        /// </summary>
        /// <param name="stream">The stream containing the XML data to read.</param>
        /// <returns>A new instance of the XmlTextReader class.</returns>
        public static XmlReader CreateSafeXmlTextReader(Stream stream)
        {
            XmlReader xtr = XmlReader.Create(stream, new XmlReaderSettings() { Async = true, DtdProcessing = DtdProcessing.Ignore, CheckCharacters = false });
            return xtr;
        }

        /// <summary>
        /// Initializes a new instance of the XmlTextReader class with the specified file.
        /// </summary>
        /// <param name="url">The URL for the file containing the XML data. The BaseURI is set to this value.</param>
        /// <returns>A new instance of the XmlTextReader class.</returns>
        public static XmlReader CreateSafeXmlTextReader(string url)
        {
            XmlReader xtr = XmlReader.Create(url, new XmlReaderSettings() { Async = true, DtdProcessing = DtdProcessing.Ignore, CheckCharacters = false });
            return xtr;
        }

        /// <summary>
        /// Initializes a new instance of the XmlTextReader class with the specified TextReader.
        /// </summary>
        /// <param name="input">The TextReader containing the XML data to read.</param>
        /// <returns>A new instance of the XmlTextReader class.</returns>
        public static XmlReader CreateSafeXmlTextReader(TextReader input)
        {
            XmlReader xtr = XmlReader.Create(input, new XmlReaderSettings() { Async = true, DtdProcessing = DtdProcessing.Ignore, CheckCharacters = false });
            return xtr;
        }

        /// <summary>
        /// Initializes a new instance of the XmlTextReader class with the specified stream and XmlNameTable.
        /// </summary>
        /// <param name="input">The stream containing the XML data to read.</param>
        /// <param name="nt">The XmlNameTable to use.</param>
        /// <returns>A new instance of the XmlTextReader class.</returns>
        public static XmlReader CreateSafeXmlTextReader(Stream input, XmlNameTable nt)
        {
            XmlReader xtr = XmlReader.Create(input, new XmlReaderSettings() { Async = true, NameTable = nt, DtdProcessing = DtdProcessing.Ignore, CheckCharacters = false });
            return xtr;
        }

        /// <summary>
        /// Initializes a new instance of the XmlTextReader class with the specified file and XmlNameTable.
        /// </summary>
        /// <param name="url">The URL for the file containing the XML data to read.</param>
        /// <param name="nt">The XmlNameTable to use.</param>
        /// <returns>A new instance of the XmlTextReader class.</returns>
        public static XmlReader CreateSafeXmlTextReader(string url, XmlNameTable nt)
        {
            XmlReader xtr = XmlReader.Create(url, new XmlReaderSettings() { Async = true, NameTable = nt, DtdProcessing = DtdProcessing.Ignore, CheckCharacters = false });
            return xtr;
        }

        /// <summary>
        /// Initializes a new instance of the XmlTextReader class with the specified TextReader.
        /// </summary>
        /// <param name="input">The TextReader containing the XML data to read.</param>
        /// <param name="nt">The XmlNameTable to use.</param>
        /// <returns>A new instance of the XmlTextReader class.</returns>
        public static XmlReader CreateSafeXmlTextReader(TextReader input, XmlNameTable nt)
        {
            XmlReader xtr = XmlReader.Create(input, new XmlReaderSettings() { Async = true, NameTable = nt, DtdProcessing = DtdProcessing.Ignore, CheckCharacters = false });
            return xtr;
        }
        #endregion

        #region XPathDocument
        /// <summary>
        /// Initializes a new instance of the XPathDocument class from the XML data in the specified Stream object.
        /// </summary>
        /// <param name="stream">The Stream object that contains the XML data.</param>
        /// <returns>A new instance of the XPathDocument class.</returns>
        public static XPathDocument CreateXPathDocument(Stream stream)
        {
            using (XmlReader xr = XmlReader.Create(stream, SafeXmlFactory.defaultSettings))
            {
                return CreateXPathDocument(xr);
            }
        }

        /// <summary>
        /// Initializes a new instance of the XPathDocument class from the XML data in the specified file.
        /// </summary>
        /// <param name="uri">The path of the file that contains the XML data.</param>
        /// <returns>A new instance of the XPathDocument class.</returns>
        public static XPathDocument CreateXPathDocument(string uri)
        {
            using (XmlReader xr = XmlReader.Create(uri, SafeXmlFactory.defaultSettings))
            {
                return CreateXPathDocument(xr);
            }
        }

        /// <summary>
        /// Initializes a new instance of the XPathDocument class from the XML data that is contained in the specified TextReader object.
        /// </summary>
        /// <param name="textReader">The TextReader object that contains the XML data.</param>
        /// <returns>A new instance of the XPathDocument class.</returns>
        public static XPathDocument CreateXPathDocument(TextReader textReader)
        {
            using (XmlReader xr = XmlReader.Create(textReader, SafeXmlFactory.defaultSettings))
            {
                return CreateXPathDocument(xr);
            }
        }

        /// <summary>
        /// Initializes a new instance of the XPathDocument class from the XML data that is contained in the specified XmlReader object.
        /// </summary>
        /// <param name="reader">The XmlReader object that contains the XML data.</param>
        /// <returns>A new instance of the XPathDocument class.</returns>
        public static XPathDocument CreateXPathDocument(XmlReader reader)
        {
            // we need to check to see if the reader is configured properly
            if (reader.Settings != null)
            {
                if (reader.Settings.DtdProcessing != DtdProcessing.Prohibit)
                {
                    throw new XmlDtdException();
                }
            }

            return new XPathDocument(reader);
        }

        /// <summary>
        /// Initializes a new instance of the XPathDocument class from the XML data in the file specified with the white space handling specified.
        /// </summary>
        /// <param name="uri">The path of the file that contains the XML data.</param>
        /// <param name="space">An XmlSpace object.</param>
        /// <returns>A new instance of the XPathDocument class.</returns>
        public static XPathDocument CreateXPathDocument(string uri, XmlSpace space)
        {
            using (XmlReader xr = XmlReader.Create(uri, SafeXmlFactory.defaultSettings))
            {
                return CreateXPathDocument(xr, space);
            }
        }

        /// <summary>
        /// Initializes a new instance of the XPathDocument class from the XML data that is contained in the specified XmlReader object with the specified white space handling.
        /// </summary>
        /// <param name="reader">The XmlReader object that contains the XML data.</param>
        /// <param name="space">An XmlSpace object.</param>
        /// <returns>A new instance of the XPathDocument class.</returns>
        public static XPathDocument CreateXPathDocument(XmlReader reader, XmlSpace space)
        {
            // we need to check to see if the reader is configured properly
            if (reader.Settings != null)
            {
                if (reader.Settings.DtdProcessing != DtdProcessing.Prohibit)
                {
                    throw new XmlDtdException();
                }
            }

            return new XPathDocument(reader, space);
        }
        #endregion
    }
}