using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Security.Policy;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace Microsoft.Exchange.WebServices.NETStandard.Core
{
    public interface IEWSStaticConfig
    {
        public EWSStaticConfig GetInstance();
    }
    public class EWSStaticConfig : IEWSStaticConfig
    {
        /// <summary>
        /// Default UserAgent
        /// </summary>
        private static string defaultUserAgent = "ExchangeServicesClient/" + EwsUtilities.BuildVersion;

       
        public TraceFlags traceFlags = TraceFlags.None;
        public ExchangeCredentials credentials { get; private set; }

        public int timeout = 100000;
        public bool preAuthenticate = true;
        public string userAgent = defaultUserAgent;
        public bool acceptGzipEncoding = true;
        public bool keepAlive = true;
        public string connectionGroupName;
        public DateTimePrecision dateTimePrecision = DateTimePrecision.Default;
        public string clientRequestId;
        public bool returnClientRequestId;
        public IWebProxy webProxy;
        public CookieContainer cookieContainer { get; private set; }
        public ITraceListener traceListener = new EwsTraceListener();
        public bool sendClientLatencies = true;
        public string targetServerVersion { get; private set; }

        /// <summary>
        /// Set "credentials" to null for enabling
        /// </summary>
        public bool useDefaultCredentials { get { return credentials == null; } }
        /// <summary>
        /// Set "traceflags" for enabling
        /// </summary>
        public bool traceEnabled { get { return traceFlags != TraceFlags.None && traceListener != null; } }

        private Uri serverUrl = null;
        public Uri ServerUrl { get => serverUrl; set => serverUrl = value; }

        public bool needSignature => this.credentials != null && this.credentials.NeedSignature;

        public ExchangeVersion requestedServerVersion;
        public TimeZoneInfo timeZone = TimeZoneInfo.Local;
        public List<KeyValuePair<string, IEnumerable<string>>> additionalMessageHeaders = new List<KeyValuePair<string, IEnumerable<string>>>();


        public EWSStaticConfig(Uri serverUrl, ExchangeCredentials credentials, TimeZoneInfo timeZoneInfo, ExchangeVersion version = ExchangeVersion.Exchange2013_SP1, CookieContainer cookieContainer = null )
        {
            requestedServerVersion = version;
            timeZone = timeZoneInfo ?? TimeZoneInfo.Local;
            this.cookieContainer = cookieContainer;
            this.credentials = credentials;

            this.AdjustServiceUriFromCredentials(serverUrl);
            if ((serverUrl.Scheme != "http") && (serverUrl.Scheme != "https")) // Verify that the protocol is something that we can handle
            {
                throw new ServiceLocalException(string.Format(Strings.UnsupportedWebProtocol, serverUrl.Scheme));
            }

            if (!this.useDefaultCredentials)
            {
                if (this.credentials == null)
                {
                    throw new ServiceLocalException(Strings.CredentialsRequired);
                }
                credentials.AdditionalHeaders(serverUrl, additionalMessageHeaders);
            }
        }


        public HttpClientHandler GenerateHttpClientHandler()
        {
            Console.WriteLine("********** Creating  HttpClientHandler  ************");
            var httpClientHandler = new HttpClientHandler()
            {
                AutomaticDecompression = DecompressionMethods.Deflate | DecompressionMethods.GZip,
                PreAuthenticate = this.preAuthenticate,
                AllowAutoRedirect = true,
                UseCookies = false,
                //CookieContainer = this.cookieContainer ?? new CookieContainer(),
                UseDefaultCredentials = this.useDefaultCredentials,
                MaxAutomaticRedirections = 1,
                CheckCertificateRevocationList = false,
                MaxConnectionsPerServer = 200
            };
            if (this.webProxy != null)
            {
                httpClientHandler.Proxy = this.webProxy;
            }
            httpClientHandler.UseProxy = (this.webProxy != null);
            httpClientHandler.ServerCertificateCustomValidationCallback = (sender, cert, chain, sslPolicyErrors) => { return true; };
            
            if(this.credentials != null)
            {
                var configCreds = this.credentials.ConfigureCredentials();
                CredentialCache credentialCache = new CredentialCache();
                credentialCache.Add(serverUrl, "NTLM", configCreds as NetworkCredential);
                credentialCache.Add(serverUrl, "Digest", configCreds as NetworkCredential);
                credentialCache.Add(serverUrl, "Basic", configCreds as NetworkCredential);
                httpClientHandler.Credentials = credentialCache;
            }

            return httpClientHandler;
        }

        public SocketsHttpHandler GenerateSocketHandler()
        {
            Console.WriteLine("********** Creating  HttpHandler  ************");
            var socketHttpHandler = new SocketsHttpHandler()
            {
                AutomaticDecompression = DecompressionMethods.Deflate | DecompressionMethods.GZip,
                PreAuthenticate = this.preAuthenticate,
                AllowAutoRedirect = true,
                UseCookies = false,
                //CookieContainer = this.cookieContainer ?? new CookieContainer(),
                MaxAutomaticRedirections = 1,
                MaxConnectionsPerServer = 200,
                PooledConnectionIdleTimeout = TimeSpan.FromHours(1),
                KeepAlivePingDelay = Timeout.InfiniteTimeSpan,
                KeepAlivePingPolicy = HttpKeepAlivePingPolicy.Always
            };
            if (this.webProxy != null)
            {
                socketHttpHandler.Proxy = this.webProxy;
            }
            socketHttpHandler.UseProxy = (this.webProxy != null);
            socketHttpHandler.SslOptions.RemoteCertificateValidationCallback  = (sender, cert, chain, sslPolicyErrors) => { return true; };

            if (this.credentials != null)
            {
                var configCreds = this.credentials.ConfigureCredentials();
                CredentialCache credentialCache = new CredentialCache();
                credentialCache.Add(serverUrl, "NTLM", configCreds as NetworkCredential);
                credentialCache.Add(serverUrl, "Digest", configCreds as NetworkCredential);
                credentialCache.Add(serverUrl, "Basic", configCreds as NetworkCredential);

                socketHttpHandler.Credentials = credentialCache;
                credentials.AdditionalHeaders(serverUrl, additionalMessageHeaders);
            }

            return socketHttpHandler;
        }

        public void SignStream(MemoryStream stream)
        {
            this.credentials.Sign(stream);
        }

        public void EmitExtraSoapHeaderNamespaceAliases_IfRequired(XmlWriter writer)
        {
            if (this.credentials != null)
            {
                this.credentials.EmitExtraSoapHeaderNamespaceAliases(writer);
            }
        }
        public void SerializeWSSecurityHeaders_IfRequired(XmlWriter writer)
        {
            if (this.credentials != null)
            {
                this.credentials.SerializeWSSecurityHeaders(writer);
            }
        }

        public static ExchangeCredentials AdjustLinuxAuthentication(Uri url, ExchangeCredentials serviceCredentials)
        {
            if (!(serviceCredentials is WebCredentials))// Nothing to adjust
                return serviceCredentials;

            var networkCredentials = ((WebCredentials)serviceCredentials).Credentials as NetworkCredential;
            if (networkCredentials != null)
            {
                CredentialCache credentialCache = new CredentialCache();
                credentialCache.Add(url, "NTLM", networkCredentials);
                credentialCache.Add(url, "Digest", networkCredentials);
                credentialCache.Add(url, "Basic", networkCredentials);

                serviceCredentials = credentialCache;
            }
            return serviceCredentials;
        }

        /// <summary>
        /// Adjusts the service URI based on the current type of credentials.
        /// </summary>
        /// <remarks>
        /// Autodiscover will always return the "plain" EWS endpoint URL but if the client
        /// is using WindowsLive credentials, ExchangeService needs to use the WS-Security endpoint.
        /// </remarks>
        /// <param name="uri">The URI.</param>
        /// <returns>Adjusted URL.</returns>
        public void AdjustServiceUriFromCredentials(Uri uri)
        {
            this.serverUrl = (this.credentials != null) ? this.credentials.AdjustUrl(uri) : uri;
        }

        /// <summary>
        /// Validates a new-style version string.
        /// This validation is not as strict as server-side validation.
        /// </summary>
        /// <param name="version"> the version string </param>
        /// <remarks>
        /// The target version string has a required part and an optional part.
        /// The required part is two integers separated by a dot, major.minor
        /// The optional part is a minimum required version, minimum=major.minor
        /// Examples:
        ///   X-EWS-TargetVersion: 2.4
        ///   X-EWS_TargetVersion: 2.9; minimum=2.4
        /// </remarks>
        internal static void ValidateTargetVersion(string version)
        {
            const char ParameterSeparator = ';';
            const string LegacyVersionPrefix = "Exchange20";
            const char ParameterValueSeparator = '=';
            const string ParameterName = "minimum";

            if (String.IsNullOrEmpty(version))
            {
                throw new ArgumentException("Target version must not be empty.");
            }

            string[] parts = version.Trim().Split(ParameterSeparator);
            switch (parts.Length)
            {
                case 1:
                    // Validate the header value. We allow X.Y or Exchange20XX.
                    string part1 = parts[0].Trim();
                    if (parts[0].StartsWith(LegacyVersionPrefix))
                    {
                        // Close enough; misses corner cases like "Exchange2001". Server will do complete validation.
                    }
                    else if (IsMajorMinor(part1))
                    {
                        // Also close enough; misses corner cases like ".5".
                    }
                    else
                    {
                        throw new ArgumentException("Target version must match X.Y or Exchange20XX.");
                    }

                    break;

                case 2:
                    // Validate the optional minimum version parameter, "minimum=X.Y"
                    string part2 = parts[1].Trim();
                    string[] minParts = part2.Split(ParameterValueSeparator);
                    if (minParts.Length == 2 &&
                        minParts[0].Trim().Equals(ParameterName, StringComparison.OrdinalIgnoreCase) &&
                        IsMajorMinor(minParts[1].Trim()))
                    {
                        goto case 1;
                    }

                    throw new ArgumentException("Target version must match X.Y or Exchange20XX.");

                default:
                    throw new ArgumentException("Target version should have the form.");
            }
        }

        private static bool IsMajorMinor(string versionPart)
        {
            const char MajorMinorSeparator = '.';

            string[] parts = versionPart.Split(MajorMinorSeparator);
            if (parts.Length != 2)
            {
                return false;
            }

            foreach (string s in parts)
            {
                foreach (char c in s)
                {
                    if (!Char.IsDigit(c))
                    {
                        return false;
                    }
                }
            }

            return true;
        }

        public EWSStaticConfig GetInstance()
        {
            return this;
        }
    }
}
