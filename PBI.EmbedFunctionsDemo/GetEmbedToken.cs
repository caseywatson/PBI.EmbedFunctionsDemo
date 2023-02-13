using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Identity.Client;
using Microsoft.PowerBI.Api;
using Microsoft.PowerBI.Api.Models;
using Microsoft.Rest;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using static System.Environment;

namespace PBI.EmbedFunctionsDemo
{
    public static class GetEmbedToken
    {
        public static class EnvironmentVariables
        {
            public const string AadClientId = "AadClientId";
            public const string AadTenantId = "AadTenantId";
            public const string AadClientSecret = "AadClientSecret";
        }

        [FunctionName("GetEmbedToken")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", Route = "pbi/token/{workspaceId}/{reportId}/{accountId}")] HttpRequest req,
            Guid workspaceId, Guid reportId, string accountId, ILogger log)
        {
            const string accountViewerRole = "Account Viewer";

            var pbiClient = await GetPbiClient();
            var pbiReport = await pbiClient.Reports.GetReportInGroupAsync(workspaceId, reportId);

            var rlsIdentity = new EffectiveIdentity(accountId,
                datasets: new List<string> { pbiReport.DatasetId },  
                roles: new List<string> { accountViewerRole });

            var embedTokenRequest = new GenerateTokenRequestV2(
                datasets: new List<GenerateTokenRequestV2Dataset> { new GenerateTokenRequestV2Dataset(pbiReport.DatasetId) },
                reports: new List<GenerateTokenRequestV2Report> { new GenerateTokenRequestV2Report(pbiReport.Id) },
                targetWorkspaces: new List<GenerateTokenRequestV2TargetWorkspace> { new GenerateTokenRequestV2TargetWorkspace(workspaceId) },
                identities: new List<EffectiveIdentity> { rlsIdentity });

            var embedToken = await pbiClient.EmbedToken.GenerateTokenAsync(embedTokenRequest);

            return new OkObjectResult(embedToken);
        }

        private static async Task<PowerBIClient> GetPbiClient()
        {
            var pbiApiBearerToken = await GetPbiApiBearerToken();

            return new PowerBIClient(new TokenCredentials(pbiApiBearerToken, "Bearer"));
        }
        private static async Task<string> GetPbiApiBearerToken()
        {
            const string pbiApiScope = "https://analysis.windows.net/powerbi/api/.default";

            var aadClientId = GetEnvironmentVariable(EnvironmentVariables.AadClientId)
                ?? throw new InvalidOperationException($"[{EnvironmentVariables.AadClientId}] not configured.");

            var aadTenantId = GetEnvironmentVariable(EnvironmentVariables.AadTenantId)
                ?? throw new InvalidOperationException($"[{EnvironmentVariables.AadTenantId}] not configured.");

            var aadClientSecret = GetEnvironmentVariable(EnvironmentVariables.AadClientSecret)
                ?? throw new InvalidOperationException($"[{EnvironmentVariables.AadClientSecret}] not configured.");

            var aadAuthority = $"https://login.microsoftonline.com/{aadTenantId}/";

            var aadClientApp = ConfidentialClientApplicationBuilder
                .Create(aadClientId)
                .WithClientSecret(aadClientSecret)
                .WithAuthority(aadAuthority)
                .Build();

            var aadAuthResult = await aadClientApp.AcquireTokenForClient(new[] { pbiApiScope }).ExecuteAsync();

            return aadAuthResult.AccessToken;
        }
    }
}
