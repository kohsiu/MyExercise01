using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(MyExercise01.Startup))]
namespace MyExercise01
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
