using Azure;
using Azure.AI.TextAnalytics;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace Autoreply.TextAnalysis
{
    public class TextAnalysisClient
    {
        private readonly TextAnalyticsClient _client;

        public TextAnalysisClient(string apiKey, Uri endpointUri)
        {
            _client = new TextAnalyticsClient(endpointUri, new AzureKeyCredential(apiKey));
        }

        public async Task<KeyPhraseCollection> ExtractKeyPhrasesAsync(string document, CancellationToken cancellationToken = default)
        {
            var result = await _client.ExtractKeyPhrasesAsync(document, "en", cancellationToken: cancellationToken).ConfigureAwait(false);
            return result.Value;
        }
    }
}
