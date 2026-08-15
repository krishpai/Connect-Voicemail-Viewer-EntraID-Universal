import React, { useState, useEffect, useCallback } from "react";
import { apiRequest } from "../authConfig";
import { DateRangeSelector } from "./DateRangeSelector";
import { VMCategory } from "./VMCategory";
import { Box, Stack, Typography, Button, FormControl, RadioGroup, FormControlLabel, Radio } from "@mui/material";
import { BrowserAuthError } from "@azure/msal-browser";
import { useAcquireTokenWithRecovery } from "../hooks/useAcquireTokenWithRecovery";

const API_ENDPOINT_ENTRA_AUTH = import.meta.env.VITE_API_URL_ENTRA_AUTH;
const API_ENDPOINT_CONNECT_AUTH = import.meta.env.VITE_API_URL_CONNECT_AUTH;

interface SearchBoxProps {
  userName: string;
  region: string;
  tier: string;
  entraAuth: boolean;
  onSearchResultChange: (value: string) => void;
}

export const SearchBox: React.FC<SearchBoxProps> = ({ userName, region, tier, entraAuth, onSearchResultChange }) => {

  const [vmCategory, setVMCategory] = useState<string>("ALL");
  const [startDate, setStartDate] = useState<string>("");
  const [endDate, setEndDate] = useState<string>("");
  const [searchFailedNoMessages, setSearchFailedNoMessages] = useState<boolean>(false);
  const [searchFailedServerOverloaded, setSearchFailedServerOverloaded] = useState<boolean>(false);
  const [authFailed, setAuthFailed] = useState<string | null>(null);
  const [queryType, setQueryType] = useState<string>("New");
  const [loading, setLoading] = useState<boolean>(false);


  const acquireTokenWithRecovery = useAcquireTokenWithRecovery();

  /**
   * Turns an MSAL failure into a message the agent can act on.
   */
  const describeAuthError = useCallback((error: unknown): string => {
    if (error instanceof BrowserAuthError) {
      if (error.errorCode === "popup_window_error" || error.errorCode === "empty_window_error") {
        return "The browser blocked the sign in window. Allow pop-ups for this site, then try again.";
      }
      if (error.errorCode === "user_cancelled") {
        return "Sign in was cancelled.";
      }
      if (error.errorCode === "block_nested_popups") {
        return "Sign in could not start. Reload the app and try again.";
      }
    }
    return "Sign in did not complete. Try again.";
  }, []);

  const searchClicked = async () => {
    setLoading(true);
    setSearchFailedNoMessages(false);
    setSearchFailedServerOverloaded(false);
    setAuthFailed(null);

    const endpoint = entraAuth ? API_ENDPOINT_ENTRA_AUTH : API_ENDPOINT_CONNECT_AUTH;
    const apiUrl = `${endpoint}?function_code=fetch_voice_messages&userName=${encodeURIComponent(userName)}&vmx3_region=${vmCategory}&user_tier=${tier}&start_date=${startDate}&end_date=${endDate}&query_type=${queryType}`;

    console.log("apiUrl: " + apiUrl);

    let accessToken: string;

    try {
      // Interaction is allowed here: this runs inside the click handler, so a
      // popup still has the user gesture it needs. The hook picks popup or
      // redirect based on whether we are embedded in Agent Workspace.
      const authResult = await acquireTokenWithRecovery({ ...apiRequest }, { allowInteraction: true });
      accessToken = authResult?.accessToken ?? "";
    } catch (error) {
      console.error("Token acquisition failed:", error);
      setAuthFailed(describeAuthError(error));
      onSearchResultChange("");
      setLoading(false);
      return;
    }

    if (!accessToken) {
      // In the standalone tab the hook may have started a redirect, in which
      // case the page is already navigating away.
      setAuthFailed("Sign in did not complete. Try again.");
      onSearchResultChange("");
      setLoading(false);
      return;
    }

    try {
      const response = await fetch(apiUrl, {
        headers: { Authorization: `Bearer ${accessToken}` }
      });

      if (response.status === 401 || response.status === 403) {
        setAuthFailed("Your session expired. Select Retrieve Messages to sign in again.");
        onSearchResultChange("");
        return;
      }

      if (!response.ok) {
        setSearchFailedServerOverloaded(true);
        throw new Error(`API error: ${response.status} ${response.statusText}`);
      }

      const data = await response.json();

      if (data.success && data.matched_objects_count > 0) {
        onSearchResultChange(JSON.stringify(data));
      }
      else {
        setSearchFailedNoMessages(true);
        onSearchResultChange("");
      }
    }
    catch (e) {
      console.log(e);
      onSearchResultChange("");
    }
    finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    if (region) {
      setVMCategory(region);
    }
  }, [region]);

  return (
    <Box
      sx={{
        width: "100%",
        maxWidth: "1000px", // Limits the spread on ultra-wide monitors
        margin: "0 auto",   // Centers the entire component on the screen
        p: 3
      }}
    >

      <Stack
        direction={{ xs: "column", md: "row" }}
        spacing={4}
        alignItems="flex-start"
        justifyContent="center"
        sx={{ width: "100%", mb: 2 }}
      >
        <DateRangeSelector
          onStartDateChange={(val) => setStartDate(val)}
          onEndDateChange={(val) => setEndDate(val)}
        />

        {(tier === "SUPERUSER") && (<VMCategory
          vmCategory={vmCategory}
          onVMCategoryChange={(val) => setVMCategory(val)}
        />)
        }

      </Stack>

      {/* Bottom Section: Action Button & Feedback 
      */}
      <Box
        sx={{
          display: "flex",
          flexDirection: "column",
          alignItems: "center",
          textAlign: "center"
        }}
      >
        <FormControl sx={{ mb: 1, alignItems: "center" }}>
          <RadioGroup
            row
            aria-labelledby="query-type-label"
            name="queryType"
            value={queryType}
            onChange={(e) => setQueryType(e.target.value)}
          >
            <FormControlLabel value="New" control={<Radio />} label="New" />
            <FormControlLabel value="All" control={<Radio />} label="All" />
            <FormControlLabel value="Deleted" control={<Radio />} label="Deleted" />
          </RadioGroup>
        </FormControl>
        <Button
          variant="contained"
          size="large"
          onClick={searchClicked}
          disabled={loading}
          sx={{ minWidth: "150px", borderRadius: "8px" }}
        >
          {loading ? "Fetching..." : "Retrieve Messages"}
        </Button>

        {loading && (
          <Typography sx={{ mt: 2, color: "text.secondary", fontStyle: "italic" }}>
            Please wait, communicating with server...
          </Typography>
        )}

        {!loading && authFailed && (
          <Typography color="error" sx={{ mt: 2, fontWeight: 500 }}>
            {authFailed}
          </Typography>
        )}

        {!loading && searchFailedNoMessages && (
          <Typography color="error" sx={{ mt: 2, fontWeight: 500 }}>
            No voice messages found for the selected criteria.
          </Typography>
        )}
        {!loading && searchFailedServerOverloaded && (
          <Typography color="error" sx={{ mt: 2, fontWeight: 500 }}>
            Search timed out. Narrow the date range or select one region.
          </Typography>
        )}
      </Box>
    </Box>
  );
};
