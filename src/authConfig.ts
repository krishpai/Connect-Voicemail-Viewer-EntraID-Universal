import { LogLevel, type Configuration } from "@azure/msal-browser";

/**
 * Configuration object to be passed to MSAL instance on creation.
 * For a full list of MSAL.js configuration parameters, visit:
 * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/configuration.md
 */
const API_SCOPE = import.meta.env.VITE_API_SCOPE;

export const msalConfig: Configuration = {
  auth: {
    clientId: import.meta.env.VITE_APP_CLIENT_ID,
    authority: import.meta.env.VITE_APP_AUTHORITY,
    redirectUri: import.meta.env.VITE_REDIRECT_URI,
    postLogoutRedirectUri: import.meta.env.VITE_REDIRECT_URI,

  },
  cache: {
    cacheLocation: "localStorage", // This configures where your cache will be stored
  },
  system: {
    allowRedirectInIframe: true,
    loggerOptions: {
      // Verbose while confirming the popup flow works on msal-browser@4.
      // Turn back down to Error once confirmed - Verbose with
      // piiLoggingEnabled logs token claims.
      logLevel: LogLevel.Verbose,
      piiLoggingEnabled: true, // Only use this during local debugging!
      loggerCallback: (level, message, containsPii) => {
        if (containsPii) {
          //return;
        }
        switch (level) {
          case LogLevel.Error:
            console.error(message);
            return;
          case LogLevel.Info:
            console.info(message);
            return;
          case LogLevel.Verbose:
            console.debug(message);
            return;
          case LogLevel.Warning:
            console.warn(message);
            return;
          default:
            return;
        }
      },
    },
  },
};

/**
 * Scopes you add here will be prompted for user consent during sign-in.
 * By default, MSAL.js will add OIDC scopes (openid, profile, email) to any login request.
 * For more information about OIDC scopes, visit:
 * https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-permissions-and-consent#openid-connect-scopes
 */
export const loginRequest = {
  scopes: ["User.Read"],
};

// Scopes for Connect backend API access token
export const apiRequest = {
  scopes: [`${API_SCOPE}`],
};

/**
 * Add here the scopes to request when obtaining an access token for MS Graph API. For more information, see:
 * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/resources-and-scopes.md
 */
export const graphConfig = {
  graphMeEndpoint: "https://graph.microsoft.com/v1.0/me",
};

//clientId: "e85ae0cf-008d-4bcf-bb7e-7f1d975eaf5e", //Ticket clinic production client
//authority: "https://login.microsoftonline.com/44d9a3b3-17c3-4c76-9026-41222eb1b4fd", //Ticket clinic production tenant