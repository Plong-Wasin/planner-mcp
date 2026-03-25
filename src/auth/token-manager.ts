import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import { TokenResponse, TokenData } from "./types.js";
import { log } from "../utils/logger.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const tokenFilePath = path.join(__dirname, "../../.access-token.txt");
const DEFAULT_CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e";

let accessToken: string | null = null;
let refreshToken: string | null = null;
let tokenExpiresAt: number | null = null;

// Try to read the stored access token and refresh token
try {
  if (fs.existsSync(tokenFilePath)) {
    const tokenData = fs.readFileSync(tokenFilePath, "utf8");
    try {
      // Try to parse as JSON first (new format)
      const parsedToken = JSON.parse(tokenData);
      accessToken = parsedToken.accessToken || parsedToken.token; // Support both old and new format
      refreshToken = parsedToken.refreshToken || null;
      tokenExpiresAt = parsedToken.expiresAt || null;
    } catch (parseError) {
      // Fall back to using the raw token (old format)
      accessToken = tokenData;
    }
  }
} catch (error) {
  console.error("Error reading access token file:", error);
}

// Alternatively, check if token is in environment variables
if (!accessToken && process.env.GRAPH_ACCESS_TOKEN) {
  accessToken = process.env.GRAPH_ACCESS_TOKEN;
}

export function isTokenExpired(): boolean {
  if (!tokenExpiresAt) return false; // If we don't have expiration info, assume it's valid
  return Date.now() >= (tokenExpiresAt - 5 * 60 * 1000); // Expire 5 minutes early to be safe
}

export async function refreshAccessToken(): Promise<boolean> {
  if (!refreshToken) {
    log("INFO", "No refresh token available");
    return false;
  }

  try {
    log("INFO", "Attempting to refresh access token");

    const tokenResponse = await fetch(
      `https://login.microsoftonline.com/common/oauth2/v2.0/token`,
      {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({
          client_id: DEFAULT_CLIENT_ID,
          grant_type: "refresh_token",
          refresh_token: refreshToken,
        }),
      },
    );

    const tokenData = (await tokenResponse.json()) as TokenResponse;

    if (tokenData.access_token) {
      accessToken = tokenData.access_token;
      refreshToken = tokenData.refresh_token || refreshToken; // Update refresh token if a new one is provided
      tokenExpiresAt = Date.now() + (60 * 60 * 1000); // Set expiration to 1 hour from now (default)

      // Save updated tokens to file
      try {
        fs.writeFileSync(
          tokenFilePath,
          JSON.stringify({
            accessToken: accessToken,
            refreshToken: refreshToken,
            expiresAt: tokenExpiresAt,
          }),
        );
      } catch (saveError) {
        console.error("Warning: Could not save refreshed token to file:", saveError);
      }

      log("INFO", "Access token refreshed successfully");
      return true;
    }

    if (tokenData.error === "invalid_grant" || tokenData.error === "refresh_token_expired") {
      log("ERROR", "Refresh token expired or invalid", { error: tokenData.error });
      // Clear tokens
      accessToken = null;
      refreshToken = null;
      tokenExpiresAt = null;
      try {
        fs.unlinkSync(tokenFilePath);
      } catch (e) {
        // Ignore error
      }
      return false;
    }

    log("ERROR", "Failed to refresh token", { error: tokenData.error });
    return false;
  } catch (error) {
    log("ERROR", "Error refreshing token", { error: (error as Error).message });
    return false;
  }
}

export async function loadTokenFromFile(): Promise<boolean> {
  if (accessToken) return true; // Already loaded

  try {
    if (fs.existsSync(tokenFilePath)) {
      const tokenData = fs.readFileSync(tokenFilePath, "utf8");
      try {
        const parsedToken = JSON.parse(tokenData);
        accessToken = parsedToken.accessToken || parsedToken.token;
        refreshToken = parsedToken.refreshToken || null;
        tokenExpiresAt = parsedToken.expiresAt || null;
        log("INFO", "Loaded token from file");
        return true;
      } catch (parseError) {
        accessToken = tokenData;
        log("INFO", "Loaded token from file (legacy format)");
        return true;
      }
    }
  } catch (error) {
    log("ERROR", "Failed to load token from file", { error: (error as Error).message });
  }

  return false;
}

export function getAccessToken(): string | null {
  return accessToken;
}

export function getTokenData(): TokenData {
  return {
    accessToken,
    refreshToken,
    tokenExpiresAt,
  };
}

export function setAccessToken(token: string, refreshTok?: string | null, expiresAt?: number | null): void {
  accessToken = token;
  refreshToken = refreshTok || null;
  tokenExpiresAt = expiresAt || null;

  // Save tokens to file
  try {
    fs.writeFileSync(
      tokenFilePath,
      JSON.stringify({
        accessToken: accessToken,
        refreshToken: refreshToken,
        expiresAt: tokenExpiresAt,
      }),
    );
    log("INFO", "Tokens saved to file successfully");
  } catch (saveError) {
    console.error("Warning: Could not save tokens to file:", saveError);
  }
}

export function clearTokens(): void {
  accessToken = null;
  refreshToken = null;
  tokenExpiresAt = null;
  try {
    fs.unlinkSync(tokenFilePath);
  } catch (e) {
    // Ignore error
  }
}
