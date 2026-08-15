import { useEffect, useState, useCallback, useRef } from "react";
import { MsalAuthenticationTemplate, useMsal } from "@azure/msal-react";
import { AmazonConnectApp } from '@amazon-connect/app';
import { AgentClient } from "@amazon-connect/contact";
import { VoiceClient, type CreateOutboundCallResult } from "@amazon-connect/voice";
import { ContactClient } from "@amazon-connect/contact";
import { PageLayout } from "./components/PageLayout";
import { SearchBox } from "./components/SearchBox";
import { SearchResultsView } from "./components/SearchResultsView";
import Divider from '@mui/material/Divider';
import Box from '@mui/material/Box';
import Button from '@mui/material/Button';
import Typography from '@mui/material/Typography';
import { InteractionType, BrowserAuthError } from "@azure/msal-browser";
import { apiRequest } from "./authConfig";
import { useAcquireTokenWithRecovery } from "./hooks/useAcquireTokenWithRecovery";

import "./App.css";

const API_ENDPOINT_ENTRA_AUTH = import.meta.env.VITE_API_URL_ENTRA_AUTH;
const API_ENDPOINT_CONNECT_AUTH = import.meta.env.VITE_API_URL_CONNECT_AUTH;
const API_SCOPE = import.meta.env.VITE_API_SCOPE;
const isIframe = window.self !== window.top; // Immediate check

// Helper to check if we are in an MSAL "hidden" frame
const isMsalInternalFrame = window.location.hash.includes("code=") ||
  window.location.hash.includes("error=") ||
  window.name.includes("msal");

/**
 * True when this document is the popup window MSAL itself opened. MSAL names
 * those windows "msal.<id>" and they always have an opener. The app must not
 * boot its normal logic there - the popup exists only so MSAL can read the
 * auth response, and it closes itself moments later.
 */
const isMsalPopup =
  !!window.opener && window.opener !== window && window.name.startsWith("msal.");

/**
 * Auth states for the embedded (Agent Workspace) path.
 * - checking:        a silent token attempt or popup is in flight
 * - authenticated:   we hold a usable access token for the backend API
 * - signin-required: the agent must click Sign in (popup needs a user gesture)
 */
type EmbeddedAuthStatus = "checking" | "authenticated" | "signin-required";

function App() {
  const { instance, accounts } = useMsal();

  // SDK & Clients State
  const [sdkInitialized, setSdkInitialized] = useState<boolean>(false);
  const [voiceClient, setVoiceClient] = useState<VoiceClient | null>(null);
  const [, setAgentClient] = useState<AgentClient | null>(null);
  const [, setConnectProvider] = useState<AmazonConnectApp | null>(null);
  const [contactClient, setContactClient] = useState<ContactClient | null>(null);

  // Business State
  const [region, setRegion] = useState("");
  const [tier, setTier] = useState("");

  const [userName, setUserName] = useState<string | null | undefined>("");
  const [canDeleteVM, setCanDeleteVM] = useState<string | null | undefined>("Y");

  const [searchResult, setSearchResult] = useState("");
  const [loading, setLoading] = useState<boolean>(false);
  const [connectUserId, setConnectUserId] = useState<string | null>(null);
  const [, setContactId] = useState<string | null>(null);

  // Embedded (iframe) MSAL state
  const [embeddedAuthStatus, setEmbeddedAuthStatus] =
    useState<EmbeddedAuthStatus>("checking");
  const [authError, setAuthError] = useState<string | null>(null);

  // Refs to prevent double-init or stale closures
  const sdkStarted = useRef(false);
  const silentAuthAttempted = useRef(false);

  const acquireTokenWithRecovery = useAcquireTokenWithRecovery();

  /**
   * Interactive sign in for the embedded app. Wired to a button so the popup
   * is opened while the user gesture is still active.
   */
  const signInWithPopup = useCallback(async () => {
    setAuthError(null);
    setEmbeddedAuthStatus("checking");

    try {
      const result = await acquireTokenWithRecovery({ ...apiRequest }, { allowInteraction: true });
      if (!result?.accessToken) throw new Error("No access token returned.");
      setEmbeddedAuthStatus("authenticated");
    } catch (error) {
      console.error("Popup sign in failed:", error);
      setEmbeddedAuthStatus("signin-required");

      if (error instanceof BrowserAuthError) {
        if (error.errorCode === "popup_window_error" || error.errorCode === "empty_window_error") {
          setAuthError("The browser blocked the sign in window. Allow pop-ups for this site and try again.");
          return;
        }
        if (error.errorCode === "user_cancelled") {
          setAuthError("Sign in was cancelled.");
          return;
        }
        if (error.errorCode === "block_nested_popups") {
          setAuthError("Sign in could not start. Reload the app and try again.");
          return;
        }
      }
      setAuthError("Sign in did not complete. Try again.");
    }
  }, [acquireTokenWithRecovery]);

  /**
   * Fetches the user region from the backend API for standalone app.
   */
  const getUserInfo_Entra = useCallback(async (username: string) => {

    const apiUrl = `${API_ENDPOINT_ENTRA_AUTH}?function_code=get_region_of_user&AgentUserName=${encodeURIComponent(username)}`;

    try {
      setLoading(true);

      const authResult = await acquireTokenWithRecovery({ ...apiRequest });

      // 2. Guard against missing tokens
      if (!authResult?.accessToken) {
        throw new Error("Failed to acquire a valid access token.");
      }

      const response = await fetch(apiUrl, {
        method: "GET",
        headers:
        {
          Authorization: `Bearer ${authResult.accessToken}`,
          "Content-Type": "application/json",
        },
      });

      if (!response.ok) {
        throw new Error(`API error: ${response.status} ${response.statusText}`);
      }

      const data = await response.json();

      if (data?.success && data?.found) {
        setRegion(data.region);
        setTier(data.tier);
        setCanDeleteVM('Y');


        console.log("User region identified:", data.region);
        console.log("User tier identified:", data.tier);
        console.log("User VM delete status identified:", data.canDeleteVM);

      }
    }
    catch (error) {
      console.error("Failed to fetch user region:", error);
    }
    finally {
      setLoading(false);
    }
    // Include all stable dependencies
  }, [acquireTokenWithRecovery]);

  /**
   * Fetches the user info from the backend API for the iframe embedded app.
   * Sends the MSAL access token as a bearer token. Silent only - this runs on
   * mount, with no user gesture to open a popup with.
   * Region comes from the routing profile, so it is not overwritten here.
   */
  const getUserInfo_Connect = useCallback(async (connectUserId: string | null) => {
    console.log("*********** in getUserInfo_Connect");
    const apiUrl = `${API_ENDPOINT_CONNECT_AUTH}?function_code=get_user_info&AgentUserId=${connectUserId}`;
    console.log('apiUrl: ', apiUrl)
    try {
      setLoading(true);

      const authResult = await acquireTokenWithRecovery({ ...apiRequest }, { allowInteraction: false });
      if (!authResult?.accessToken) {
        throw new Error("Failed to acquire a valid access token.");
      }

      const response = await fetch(apiUrl, {
        method: "GET",
        headers: {
          Authorization: `Bearer ${authResult.accessToken}`,
          "Content-Type": "application/json",
        },
      });

      if (response.status === 401 || response.status === 403) {
        setEmbeddedAuthStatus("signin-required");
        setAuthError("Your session expired. Sign in again to continue.");
        return;
      }

      if (!response.ok) {
        throw new Error(`API error: ${response.status} ${response.statusText}`);
      }

      const data = await response.json();
      if (data?.success && data?.found) {
        //setRegion(data.region);
        setTier(data.tier);
        setUserName(data.userName);
        setCanDeleteVM(data.canDeleteVM);

        console.log("User name identified:", data.userName);
        console.log("User tier identified:", data.tier);
        console.log("User VM delete status identified:", data.canDeleteVM);

      }
      else {
        throw new Error(`API error: ${response.status} ${response.statusText}`);
      }
    }
    catch (error) {
      console.log('error: ', error)
      setRegion("ALL");
      setUserName("Unknown user");
    }
    finally {
      setLoading(false);
    }
  }, [acquireTokenWithRecovery])

  useEffect(() => {

    // Never run app logic inside the MSAL popup window.
    if (isMsalPopup) return;

    // 1. Standalone logic
    if (!isIframe && accounts.length > 0) {
      console.info("In Standalone logic");
      instance.setActiveAccount(accounts[0]);
      const username = accounts[0].idTokenClaims?.preferred_username;
      setUserName(username ?? "Unknown User");
      if (!username) {
        console.warn("No preferred_username found in claims.");
        return;
      }
      getUserInfo_Entra(username);
    }

    if (isMsalInternalFrame) return;

    // 2. Iframe / Amazon Connect logic
    if (isIframe && !sdkStarted.current) {
      console.info("In Iframe logic");
      sdkStarted.current = true;

      const amazonConnectApp = AmazonConnectApp.init({
        onCreate: async (event) => {
          setSdkInitialized(true); // Handshake complete
          console.log('************ App initialized with context:', event.context);

          // Create an Agent Client using the provider
          const agentClient = new AgentClient({ provider: amazonConnectApp.provider });
          const voiceClient = new VoiceClient({ provider: amazonConnectApp.provider });
          const contactClient = new ContactClient({ provider: amazonConnectApp.provider });

          setAgentClient(agentClient);
          setVoiceClient(voiceClient);
          setContactClient(contactClient);

          const agentARN = await agentClient.getARN();
          const agentRP = await agentClient.getRoutingProfile();
          const agentRegion = agentRP.name.split("_")[1];

          // Extract user ID from ARN
          // ARN format: arn:aws:connect:region:account:instance/instance-id/agent/user-id
          const userIdMatch = agentARN.match(/\/agent\/(.+)$/);
          const connectUserId = userIdMatch ? userIdMatch[1] : null;

          console.log("User ID:", connectUserId);
          console.log("Agent ARN:", agentARN);
          console.log("agentRP:", agentRP);
          console.log("agentRP name:", agentRP.name);
          console.log("User region:", agentRegion);

          setConnectUserId(connectUserId);
          setRegion(agentRegion);

          if (event.context.scope && "contactId" in event.context.scope) {
            setContactId(event.context.scope.contactId);
          }
        },
        onDestroy: async (event) => {
          console.log('App being destroyed:', event);
        },
      });

      // Save the provider to state so you can use it globally in the app
      setConnectProvider(amazonConnectApp.provider);
    };
  }, [accounts, instance, getUserInfo_Entra, accounts.length]);

  /**
   * 3. Embedded auth bootstrap: try silently once, then hand off to the
   *    Sign in button if interaction is needed.
   */
  useEffect(() => {
    if (!isIframe || isMsalInternalFrame || isMsalPopup) return;
    if (silentAuthAttempted.current) return;
    silentAuthAttempted.current = true;

    let cancelled = false;

    (async () => {
      try {
        const result = await acquireTokenWithRecovery({ ...apiRequest }, { allowInteraction: false });
        if (cancelled) return;
        setEmbeddedAuthStatus(result?.accessToken ? "authenticated" : "signin-required");
      } catch (error) {
        console.info("Silent sign in unavailable, interactive sign in required:", error);
        if (!cancelled) setEmbeddedAuthStatus("signin-required");
      }
    })();

    return () => { cancelled = true; };
  }, [acquireTokenWithRecovery]);

  /**
   * 4. Load the agent's profile only once both the SDK handshake and MSAL
   *    sign in have completed.
   */
  useEffect(() => {
    if (!isIframe || isMsalPopup) return;
    if (embeddedAuthStatus !== "authenticated") return;
    if (!connectUserId) return;

    getUserInfo_Connect(connectUserId);
  }, [embeddedAuthStatus, connectUserId, getUserInfo_Connect]);


  const makeOutboundCall = useCallback(async (phoneNumber: string, relatedContactid: string) => {
    console.log("phoneNumber: " + phoneNumber)
    if (!contactClient || !voiceClient) return;
    try {
      const contacts = await contactClient.listContacts();
      console.log(`Active contacts: ${contacts?.length}`);
      const isBusy = contacts?.some(c => c.type === 'voice'); // Check specifically for voice

      if (isBusy) {
        console.log("Agent busy on an existing call, cannot initiate new call");
        return;
      }

      console.log("Calling  " + phoneNumber)
      const outboundCallResult: CreateOutboundCallResult = await voiceClient.createOutboundCall(phoneNumber, { relatedContactId: relatedContactid });
      console.log("Related contactId : " + relatedContactid);
      console.log("outboundCallResult.contactId : " + outboundCallResult.contactId);
    }
    catch (error) {
      console.error("Outbound call failed:", error);
    }
  }, [contactClient, voiceClient]);

  // Render nothing inside the MSAL popup window.
  if (isMsalPopup) {
    return null;
  }

  // If we are in an iframe but the SDK hasn't finished its handshake yet,
  // we show a neutral loading screen to prevent the MSAL Redirect from firing.
  if (isIframe && !sdkInitialized) {
    return <p>Connecting to Agent Workspace...</p>;
  }

  if (isIframe && embeddedAuthStatus === "checking") {
    return <p>Signing you in...</p>;
  }

  if (isIframe && embeddedAuthStatus === "signin-required") {
    return (
      <Box sx={{ display: "flex", flexDirection: "column", alignItems: "flex-start", gap: 1.5, p: 3 }}>
        <Typography variant="body1">Sign in to view voice messages.</Typography>
        {authError && (
          <Typography variant="body2" color="error">{authError}</Typography>
        )}
        <Button variant="contained" onClick={signInWithPopup}>Sign in</Button>
      </Box>
    );
  }

  // Main UI Fragment to keep code DRY
  const renderMainContent = () => (
    <PageLayout userName={userName ?? "User"} region={region}>
      {loading ? (
        <p>Loading preferences...</p>
      ) : (
        <>
          <SearchBox userName={userName ?? "User"} region={region} tier={tier} entraAuth={!isIframe} onSearchResultChange={setSearchResult} />
          <Divider sx={{ my: 0.5, border: "1px solid", borderColor: "primary.dark" }} />
          {searchResult && (<SearchResultsView searchResult={searchResult} entraAuth={!isIframe} canDeleteVM={canDeleteVM} onDialNumberClicked={makeOutboundCall} />)}
        </>
      )}
    </PageLayout>
  );

  return (
    <>
      {isIframe ? (renderMainContent())
        : (
          <MsalAuthenticationTemplate interactionType={InteractionType.Redirect}
            authenticationRequest={{ scopes: ["openid", "profile", `${API_SCOPE}`], }}
            errorComponent={({ error }) => <pre>Error: {error?.errorMessage}</pre>}
            loadingComponent={() => <span>Launching Login redirect...</span>}>
            {accounts.length && renderMainContent()}
          </MsalAuthenticationTemplate>
        )}
    </>
  );
}

export default App;
