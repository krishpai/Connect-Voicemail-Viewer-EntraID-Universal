import React from "react";
import { createRoot } from "react-dom/client";
import * as msal from "@azure/msal-browser";
import { MsalProvider } from "@azure/msal-react";
import { msalConfig } from "./authConfig.ts";

import App from "./App.tsx";
import "bootstrap/dist/css/bootstrap.min.css";
import "./index.css";
import { LicenseInfo } from '@mui/x-license';


LicenseInfo.setLicenseKey('e0d9bb8070ce0054c9d9ecb6e82cb58fTz0wLEU9MzI0NzIxNDQwMDAwMDAsUz1wcmVtaXVtLExNPXBlcnBldHVhbCxLVj0y');
/**
 * MSAL should be instantiated outside of the component tree to prevent it
 * from being re-instantiated on re-renders.
 *
 * What it exposes to child components: Looking at the source, MsalProvider tracks inProgress (interaction status)
 * and accounts state, and listens to MSAL events to keep those in sync — so hooks like useMsal, useAccount, and components
 * like AuthenticatedTemplate / UnauthenticatedTemplate all work automatically underneath it.
 */

/*
* msalInstance --> fully configured authentication engine that knows how to log users in, store tokens, 
* refresh tokens silently, and call APIs
*/
const msalInstance = new msal.PublicClientApplication(msalConfig);
const container = document.getElementById("root") as HTMLElement;
const root = createRoot(container);
/**
 * initialize() must be called and awaited before ANY other MSAL API on this instance - 
 * including getActiveAccount() / getAllAccounts(). Calling one of those first throws uninitialized_public_client_application. 
 * This only bites once there's a cached account to iterate (i.e. after a successful sign-in has persisted one to localStorage) 
 * - with an empty cache the internal account loop never runs, so the throwing code path is never reached. 
 * That is why this was silent on every earlier test and only appeared on reload, once a real session existed.
 */
msalInstance.initialize().then(() => {
  console.log("After init:", msalInstance.getActiveAccount());
  root.render(
    <React.StrictMode>
      <MsalProvider instance={msalInstance}>
        <App />
      </MsalProvider>
    </React.StrictMode>,
  );
});

/**
 * All components underneath MsalProvider will have access to the PublicClientApplication instance
 * via context as well as all hooks and components provided by @azure/msal-react.
 * */
