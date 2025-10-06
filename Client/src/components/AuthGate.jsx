import React, { useEffect, useMemo, useState } from "react";
import { useMsal } from "@azure/msal-react";
import {
  Alert,
  Avatar,
  Box,
  Button,
  Card,
  CardContent,
  CircularProgress,
  Stack,
  Typography,
} from "@mui/material";
import LoginIcon from "@mui/icons-material/Login";
import LogoutIcon from "@mui/icons-material/Logout";
import { loginRequest, msalConfig } from "../authConfig";

const parseEnvList = (value) =>
  String(value || "")
    .split(/[,;\s]+/)
    .map((item) => item.trim().toLowerCase())
    .filter(Boolean);

const allowedEmails = parseEnvList(import.meta.env.VITE_ALLOWED_USERS);
const allowedDomains = parseEnvList(import.meta.env.VITE_ALLOWED_EMAIL_DOMAINS).map((domain) =>
  domain.startsWith("@") ? domain.slice(1) : domain
);

const hasRestrictions = allowedEmails.length > 0 || allowedDomains.length > 0;

const isEmailAllowed = (email) => {
  if (!email) return !hasRestrictions;
  const lower = email.toLowerCase();
  if (!hasRestrictions) return true;
  if (allowedEmails.includes(lower)) return true;
  return allowedDomains.some((domain) => lower.endsWith(`@${domain}`));
};

function AuthGate({ children }) {
  const { instance, accounts, inProgress } = useMsal();
  const [authError, setAuthError] = useState(null);

  const account = useMemo(() => {
    if (accounts.length === 0) return instance.getActiveAccount();
    const active = instance.getActiveAccount();
    return active ?? accounts[0];
  }, [accounts, instance]);

  useEffect(() => {
    if (account) {
      instance.setActiveAccount(account);
    }
  }, [account, instance]);

  const email = account?.username || account?.idTokenClaims?.preferred_username;
  const allowed = isEmailAllowed(email);

  const handleLogin = async () => {
    setAuthError(null);
    try {
      const response = await instance.loginPopup(loginRequest);
      if (response?.account) {
        instance.setActiveAccount(response.account);
      }
    } catch (error) {
      setAuthError(error?.message || "Login failed. Please try again.");
    }
  };

  const handleLogout = async () => {
    setAuthError(null);
    try {
      await instance.logoutPopup({
        postLogoutRedirectUri: msalConfig.auth.postLogoutRedirectUri,
      });
    } catch (error) {
      setAuthError(error?.message || "Logout failed. Please close the window.");
    }
  };

  if (inProgress === "login" || inProgress === "handleRedirect") {
    return (
      <Box sx={{ display: "grid", placeItems: "center", minHeight: "100vh" }}>
        <Stack spacing={2} alignItems="center">
          <CircularProgress color="primary" />
          <Typography variant="body1">Signing you in...</Typography>
        </Stack>
      </Box>
    );
  }

  if (!account) {
    return (
      <Box sx={{ display: "grid", placeItems: "center", minHeight: "100vh", px: 2 }}>
        <Card sx={{ maxWidth: 420, width: "100%" }}>
          <CardContent>
            <Stack spacing={3} alignItems="stretch">
              <Stack spacing={1}>
                <Typography variant="h5" component="h1" textAlign="center">
                  Sign in to continue
                </Typography>
                <Typography variant="body2" color="text.secondary" textAlign="center">
                  Use your Microsoft work account to access the invoice dashboard.
                </Typography>
              </Stack>
              {hasRestrictions && (
                <Alert severity="info">
                  Access is limited to approved users. Contact an administrator if you need help.
                </Alert>
              )}
              {authError && <Alert severity="error">{authError}</Alert>}
              <Button
                variant="contained"
                startIcon={<LoginIcon />}
                onClick={handleLogin}
                size="large"
              >
                Sign in with Microsoft
              </Button>
            </Stack>
          </CardContent>
        </Card>
      </Box>
    );
  }

  if (!allowed) {
    return (
      <Box sx={{ display: "grid", placeItems: "center", minHeight: "100vh", px: 2 }}>
        <Card sx={{ maxWidth: 420, width: "100%" }}>
          <CardContent>
            <Stack spacing={3} alignItems="stretch">
              <Stack spacing={1} alignItems="center">
                <Avatar sx={{ bgcolor: "error.main", width: 64, height: 64, fontSize: 24 }}>
                  !
                </Avatar>
                <Typography variant="h6" textAlign="center">
                  Access denied
                </Typography>
                <Typography variant="body2" color="text.secondary" textAlign="center">
                  {email ? `${email} is not on the approved list.` : "Your account is not approved for this app."}
                </Typography>
              </Stack>
              {authError && <Alert severity="error">{authError}</Alert>}
              <Button
                variant="outlined"
                color="error"
                startIcon={<LogoutIcon />}
                onClick={handleLogout}
              >
                Sign out
              </Button>
            </Stack>
          </CardContent>
        </Card>
      </Box>
    );
  }

  return <>{children}</>;
}

export default AuthGate;
