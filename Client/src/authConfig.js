import { PublicClientApplication } from "@azure/msal-browser";

const clientId = (import.meta.env.VITE_AZURE_CLIENT_ID || "").trim();
const tenantId = (import.meta.env.VITE_AZURE_TENANT_ID || "").trim();
const authorityEnv = (import.meta.env.VITE_AZURE_AUTHORITY || "").trim();

const authConfig = {
  clientId,
  redirectUri: (import.meta.env.VITE_AZURE_REDIRECT_URI || window.location.origin).trim(),
  postLogoutRedirectUri: (import.meta.env.VITE_AZURE_POST_LOGOUT_REDIRECT_URI || window.location.origin).trim(),
};

if (authorityEnv) {
  authConfig.authority = authorityEnv;
} else if (tenantId) {
  authConfig.authority = `https://login.microsoftonline.com/${tenantId}`;
}

const cacheLocation = (import.meta.env.VITE_MSAL_CACHE_LOCATION || "localStorage").trim();
const cacheConfig = {
  cacheLocation,
  storeAuthStateInCookie: String(import.meta.env.VITE_MSAL_STORE_AUTH_STATE_IN_COOKIE || "false").toLowerCase() === "true",
};

export const msalConfig = {
  auth: authConfig,
  cache: cacheConfig,
};

const scopesEnv = (import.meta.env.VITE_AZURE_SCOPES || "");
const scopes = scopesEnv
  .split(/[,;\s]+/)
  .map((scope) => scope.trim())
  .filter(Boolean);

export const loginRequest = {
  scopes: scopes.length > 0 ? scopes : ["User.Read"],
};

export const msalInstance = new PublicClientApplication(msalConfig);
