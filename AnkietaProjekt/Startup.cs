using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(AnkietaProjekt.Startup))]
namespace AnkietaProjekt
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
