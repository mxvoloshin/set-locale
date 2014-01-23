using SetLocale.Client.Web.Controllers;
using SetLocale.Client.Web.Services;

namespace SetLocale.Client.Web.Test.Builders
{
    public class AdminControllerBuilder
    {
        private IFormsAuthenticationService _formAuthenticationService;
        private IUserService _userService;
        private IAppService _appService;
        private IWordService _wordService;

        public AdminControllerBuilder()
        {
            _formAuthenticationService = null;
            _userService = null;
            _appService = null;
            _wordService = null;
        }

        internal AdminControllerBuilder WithFormsAuthenticationService(IFormsAuthenticationService formAuthenticationService)
        {
            _formAuthenticationService = formAuthenticationService;
            return this;
        }

        internal AdminControllerBuilder WithUserService(IUserService userService)
        {
            _userService = userService;
            return this;
        }

        internal AdminControllerBuilder WithAppService(IAppService appService)
        {
            _appService = appService;
            return this;
        }

        internal AdminControllerBuilder WithWordService(IWordService wordService)
        {
            _wordService = wordService;
            return this;
        }

        internal AdminController Build()
        {
            return new AdminController(_userService, _wordService, _formAuthenticationService, _appService);
        }
    }
}
