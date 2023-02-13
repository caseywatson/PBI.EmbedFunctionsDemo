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
        // This role needs to be defined on the applicable Power BI dataset.
        // Learn more at https://learn.microsoft.com/power-bi/enterprise/service-admin-rls#define-roles-and-rules-in-power-bi-desktop.

        public const string AccountViewerRole = "Account Viewer";

        // These environment variables are required to authenticate to Azure Active Directory and obtain a bearer token to call the Power BI REST API.
        // Learn more at https://learn.microsoft.com/power-bi/developer/embedded/register-app?tabs=customers.

        public static class EnvironmentVariableNames
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
            try
            {
                var pbiClient = await GetPbiClient();
                var pbiReport = await pbiClient.Reports.GetReportInGroupAsync(workspaceId, reportId);

                var rlsIdentity = new EffectiveIdentity(accountId,
                    datasets: new List<string> { pbiReport.DatasetId },
                    roles: new List<string> { AccountViewerRole });

                var embedTokenRequest = new GenerateTokenRequestV2(
                    datasets: new List<GenerateTokenRequestV2Dataset> { new GenerateTokenRequestV2Dataset(pbiReport.DatasetId) },
                    reports: new List<GenerateTokenRequestV2Report> { new GenerateTokenRequestV2Report(pbiReport.Id) },
                    targetWorkspaces: new List<GenerateTokenRequestV2TargetWorkspace> { new GenerateTokenRequestV2TargetWorkspace(workspaceId) },
                    identities: new List<EffectiveIdentity> { rlsIdentity });

                var embedToken = await pbiClient.EmbedToken.GenerateTokenAsync(embedTokenRequest);

                log.LogInformation(
                    $"Embed token succesfully generated for workspace [{workspaceId}] " +
                    $"report [{reportId}] account [{accountId}] [{AccountViewerRole}].");

                return new OkObjectResult(embedToken);
            }
            catch (Exception ex)
            {
                log.LogError(
                    $"An error occurred while trying to obtain an embed token for worksapce [{workspaceId}] " +
                    $"report [{reportId}] account [{accountId}] [{AccountViewerRole}]: [{ex.Message}].");

                return new StatusCodeResult(StatusCodes.Status500InternalServerError);
            }
        }

        private static async Task<PowerBIClient> GetPbiClient()
        {
            var pbiApiBearerToken = await GetPbiApiBearerToken();

            return new PowerBIClient(new TokenCredentials(pbiApiBearerToken, "Bearer"));
        }

        private static async Task<string> GetPbiApiBearerToken()
        {
            const string pbiApiScope = "https://analysis.windows.net/powerbi/api/.default";

            var aadClientId = GetEnvironmentVariable(EnvironmentVariableNames.AadClientId)
                ?? throw new InvalidOperationException($"[{EnvironmentVariableNames.AadClientId}] not configured.");

            var aadTenantId = GetEnvironmentVariable(EnvironmentVariableNames.AadTenantId)
                ?? throw new InvalidOperationException($"[{EnvironmentVariableNames.AadTenantId}] not configured.");

            var aadClientSecret = GetEnvironmentVariable(EnvironmentVariableNames.AadClientSecret)
                ?? throw new InvalidOperationException($"[{EnvironmentVariableNames.AadClientSecret}] not configured.");

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
