using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(RiskComplianceApp.Startup))]
namespace RiskComplianceApp
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
