using System.Threading;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;

namespace MKWTranslationTextMaker
{
    public static class GoogleAuth
    {
        public static UserCredential Login(string googleClientId, string googleClientSecret, string[] scopes)
        {
            ClientSecrets secrets = new ClientSecrets()
            {
                ClientId = googleClientId,
                ClientSecret = googleClientSecret
            };

            return GoogleWebAuthorizationBroker.AuthorizeAsync(secrets, scopes, user: "user", CancellationToken.None).Result;
        }
    }
}