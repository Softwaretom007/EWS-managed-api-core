using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft.Exchange.WebServices.NETStandard.Core
{
    public interface IEWSHttpClient
    {
        public EWSHttpClient GetInstance();
    }

    public class EWSHttpClient : IEWSHttpClient
    {
        private readonly IEWSStaticConfig iConfig;

        public HttpClient Client { get; private set; }


        public EWSHttpClient(HttpClient httpClient, IEWSStaticConfig iConfig)
        {
            var config = iConfig.GetInstance();
            Console.WriteLine("********** Creating EWSHttpClient ************");
            httpClient.Timeout = TimeSpan.FromMilliseconds(config.timeout);
            httpClient.DefaultRequestHeaders.ConnectionClose = !config.keepAlive;

            if (config.acceptGzipEncoding)
            {
                httpClient.DefaultRequestHeaders.AcceptEncoding.ParseAdd("gzip,deflate");
            }

            if (!string.IsNullOrEmpty(config.clientRequestId))
            {
                httpClient.DefaultRequestHeaders.TryAddWithoutValidation("client-request-id", config.clientRequestId);
                if (config.returnClientRequestId)
                {
                    httpClient.DefaultRequestHeaders.TryAddWithoutValidation("return-client-request-id", "true");
                }
            }
            httpClient.BaseAddress = config.ServerUrl;

            if (!String.IsNullOrEmpty(config.targetServerVersion))
            {
                httpClient.DefaultRequestHeaders.TryAddWithoutValidation( ExchangeServiceBase.TargetServerVersionHeaderName, config.targetServerVersion);
            }
            httpClient.DefaultRequestVersion = new Version(2, 0);
            httpClient.DefaultRequestHeaders.ExpectContinue = true;

            Client = httpClient;
            this.iConfig = iConfig;
        }

        public EWSHttpClient GetInstance()
        {
            return this;
        }
    }
}
