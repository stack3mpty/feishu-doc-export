using feishu_doc_export.Dtos;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WebApiClientCore.Extensions.OAuths;
using WebApiClientCore.Extensions.OAuths.TokenProviders;

namespace feishu_doc_export.HttpApi
{
    public class FeiShuTokenProvider : TokenProvider
    {
        private readonly IFeiShuHttpApi _feiShuHttpApi;

        public FeiShuTokenProvider(IServiceProvider services) : base(services)
        {
            _feiShuHttpApi = services.GetService<IFeiShuHttpApi>();
        }

        protected override async Task<TokenResult> RefreshTokenAsync(IServiceProvider serviceProvider, string refresh_token)
        {
            return await RequestTokenAsync(serviceProvider);
        }

        protected override async Task<TokenResult> RequestTokenAsync(IServiceProvider serviceProvider)
        {
            if (string.Equals(GlobalConfig.AuthMode, "user", StringComparison.OrdinalIgnoreCase))
            {
                return BuildUserTokenResultFromConfig();
            }

            return await RequestTenantTokenAsync();
        }

        private async Task<TokenResult> RequestTenantTokenAsync()
        {
            var requestData = RequestData.CreateAccessToken(GlobalConfig.AppId, GlobalConfig.AppSecret);
            var tokenUrl = FeiShuConsts.BuildOpenApiUrl("/open-apis/auth/v3/tenant_access_token/internal", GlobalConfig.Platform);
            var result = await _feiShuHttpApi.GetTenantAccessToken(tokenUrl, requestData);

            return new TokenResult
            {
                Access_token = result.TenantAccessToken,
                Refresh_token = result.TenantAccessToken,
                Expires_in = result.Expire
            };
        }

        private static TokenResult BuildUserTokenResultFromConfig()
        {
            // user模式默认直接使用命令行传入的user_access_token，避免影响现有tenant模式流程。
            return new TokenResult
            {
                Access_token = GlobalConfig.UserAccessToken,
                Refresh_token = string.IsNullOrWhiteSpace(GlobalConfig.UserRefreshToken)
                    ? GlobalConfig.UserAccessToken
                    : GlobalConfig.UserRefreshToken,
                Expires_in = 86400
            };
        }
    }
}
