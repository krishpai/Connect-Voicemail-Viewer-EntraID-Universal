import { useCallback } from "react";
import { useMsal } from "@azure/msal-react";
import type {
  AuthenticationResult,
  RedirectRequest,
  SilentRequest,
} from "@azure/msal-browser";

/**
 * The app runs in two hosting contexts and they need different interactive
 * flows:
 *
 *   Standalone browser tab  -> redirect APIs (loginRedirect)
 *   Amazon Connect Agent Workspace iframe -> popup APIs
 *
 * Redirect cannot be used inside the workspace iframe: Entra sets
 * frame-ancestors on its login page, and MSAL itself throws redirect_in_iframe.
 */
const isIframe = window.self !== window.top;

/**
 * True when this document is the popup window MSAL opened. Any interactive
 * MSAL call from here throws block_nested_popups, so callers must never get
 * that far.
 */
const isMsalPopup =
  !!window.opener && window.opener !== window && window.name.startsWith("msal.");

/**
 * Popups land on a blank page instead of the SPA root, so React never boots a
 * second copy of the app inside the popup window. Register this exact URL in
 * Entra as a Single-page application redirect URI.
 */
const POPUP_REDIRECT_URI = `${window.location.origin}/blank.html`;

/**
 * Errors that mean "silent will never work, the user has to interact".
 */
const INTERACTION_ERROR_CODES = new Set([
  "interaction_required",
  "login_required",
  "consent_required",
]);

export interface AcquireTokenOptions {
  /**
   * Set false when the call is NOT triggered by a click. Browsers only allow
   * window.open while a user gesture is active, so an unattended popup is
   * blocked - and a redirect would throw away unsaved state. With this off the
   * hook stays silent and throws instead of interacting, letting the caller
   * decide what to show.
   *
   * Defaults to true.
   */
  allowInteraction?: boolean;
}

function isMsalError(e: unknown): e is { errorCode: string } {
  if (typeof e === "object" && e !== null && "errorCode" in e) {
    const record = e as Record<string, unknown>;
    return typeof record.errorCode === "string";
  }
  return false;
}

/*
      acquireTokenRedirect is meant for:
        - getting additional scopes
        - after the user is already signed in
        - using an existing valid account

      But after a "timed_out":

        - the user is not signed in anymore
        - the account is stale
        - the session cookie is gone
        - MSAL cannot complete the redirect flow cleanly

        So acquireTokenRedirect is the wrong tool for the job.

        //A "timed_out" error means the account is stale, acquireTokenRedirect will not work and a fresh login is needed
        //Microsoft's guidance is: if silent SSO times out, treat the account as invalid and force a clean login.
*/

export function useAcquireTokenWithRecovery() {
  const { instance } = useMsal();

  const acquireTokenWithRecovery = useCallback(
    async (
      request: SilentRequest & RedirectRequest,
      options: AcquireTokenOptions = {}
    ): Promise<AuthenticationResult | undefined> => {
      const { allowInteraction = true } = options;

      if (isMsalPopup) {
        // Nothing should be calling MSAL from inside the MSAL popup.
        throw new Error("Token acquisition attempted inside the MSAL popup window.");
      }

      /**
       * Fresh sign in for the embedded app. loginPopup needs the sign-in
       * scopes alongside the resource scope.
       */
      const loginWithPopup = async (): Promise<AuthenticationResult> => {
        const result = await instance.loginPopup({
          ...request,
          scopes: Array.from(new Set(["openid", "profile", ...(request.scopes ?? [])])),
          redirectUri: POPUP_REDIRECT_URI,
        });
        instance.setActiveAccount(result.account);
        return result;
      };

      const accounts = instance.getAllAccounts();

      if (accounts.length === 0) {
        if (!allowInteraction) {
          // Best effort: depends on the IdP session cookie being readable in a
          // third-party context, which the workspace iframe often prevents.
          const ssoResult = await instance.ssoSilent({ ...request });
          instance.setActiveAccount(ssoResult.account);
          return ssoResult;
        }

        if (isIframe) {
          return loginWithPopup();
        }

        await instance.loginRedirect({
          ...request,
          prompt: "login",
        });
        return undefined;
      }

      const account = accounts[0];

      try {
        return await instance.acquireTokenSilent({
          ...request,
          account,
        });
      } catch (e: unknown) {
        if (isMsalError(e)) {
          const staleAccount = e.errorCode === "timed_out";
          const needsInteraction = INTERACTION_ERROR_CODES.has(e.errorCode);

          if (staleAccount || needsInteraction) {
            if (!allowInteraction) throw e;

            if (isIframe) {
              // A stale account cannot be repaired with acquireToken - sign in fresh.
              if (staleAccount) return loginWithPopup();

              const result = await instance.acquireTokenPopup({
                ...request,
                account,
                redirectUri: POPUP_REDIRECT_URI,
              });
              instance.setActiveAccount(result.account);
              return result;
            }

            if (staleAccount) {
              await instance.loginRedirect({
                ...request,
                prompt: "login",
              });
              return undefined;
            }

            await instance.loginRedirect(request);
            return undefined;
          }
        }
        throw e;
      }
    },
    [instance]
  );

  return acquireTokenWithRecovery;
}
