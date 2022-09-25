using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(excel2vcard.Startup))]
namespace excel2vcard
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
