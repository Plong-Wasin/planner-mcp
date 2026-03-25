import { DeviceCodeResponse, TokenResponse, DeviceCodeInfo } from "./types.js";
import { setAccessToken, clearTokens, getAccessToken, isTokenExpired, loadTokenFromFile } from "./token-manager.js";
import { log } from "../utils/logger.js";
import { handleGraphError } from "../utils/error-handler.js";

const SCOPES = ["Group.Read.All", "Tasks.ReadWrite", "GroupMember.Read.All", "Directory.Read.All", "offline_access"];
const DEFAULT_CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e";

let deviceCodeInfo: DeviceCodeInfo | null = null;

export async function startDeviceCodeFlow(clientId?: string, tenantId?: string, force?: boolean): Promise<{
  success: boolean;
  message: string;
}> {
  try {
    // Check if we already have a valid token (unless force is true)
    if (!force) {
      const token = getAccessToken();
      if (token && !isTokenExpired()) {
        return {
          success: true,
          message: "✅ Already authenticated with a valid access token. No need to re-authenticate. Use force=true to re-authenticate.",
        };
      }

      // Try to load from file if not in memory
      if (!token) {
        const loaded = await loadTokenFromFile();
        if (loaded && !isTokenExpired()) {
          return {
            success: true,
            message: "✅ Already authenticated with a valid access token (loaded from storage). No need to re-authenticate. Use force=true to re-authenticate.",
          };
        }
      }
    }

    const client = clientId || DEFAULT_CLIENT_ID;
    const tenant = tenantId || "common";

    // Initiate device code flow
    const deviceCodeResponse = await fetch(
      `https://login.microsoftonline.com/${tenant}/oauth2/v2.0/devicecode`,
      {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({
          client_id: client,
          scope: SCOPES.join(" "),
        }),
      },
    );

    if (!deviceCodeResponse.ok) {
      throw new Error("Failed to initiate device code flow");
    }

    const deviceCode = (await deviceCodeResponse.json()) as DeviceCodeResponse;

    // Store device code info for polling
    deviceCodeInfo = {
      deviceCode: deviceCode.device_code,
      clientId: client,
      tenantId: tenant,
      interval: deviceCode.interval || 5,
      expiresAt: Date.now() + deviceCode.expires_in * 1000,
    };

    // Return instructions immediately without waiting
    return {
      success: true,
      message: `Authentication Required!

1. Visit this URL: ${deviceCode.verification_uri}
2. Enter this code: ${deviceCode.user_code}
3. Sign in with your Microsoft account

After completing the authentication, use the "auth_poll" tool to check if authentication is complete.

⏱️ Code expires in ${Math.floor(deviceCode.expires_in / 60)} minutes`,
    };
  } catch (error) {
    return {
      success: false,
      message: `Failed to start authentication: ${handleGraphError(error)}`,
    };
  }
}

export async function pollForToken(): Promise<{
  success: boolean;
  message: string;
}> {
  try {
    if (!deviceCodeInfo) {
      return {
        success: false,
        message: "No authentication in progress. Please call auth_start first.",
      };
    }

    // Check if device code has expired
    if (Date.now() > deviceCodeInfo.expiresAt) {
      deviceCodeInfo = null;
      return {
        success: false,
        message: "❌ Authentication code has expired. Please call auth_start again.",
      };
    }

    // Poll once for token
    const tokenResponse = await fetch(
      `https://login.microsoftonline.com/${deviceCodeInfo.tenantId}/oauth2/v2.0/token`,
      {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({
          client_id: deviceCodeInfo.clientId,
          grant_type: "urn:ietf:params:oauth:grant-type:device_code",
          device_code: deviceCodeInfo.deviceCode,
        }),
      },
    );

    const tokenData = (await tokenResponse.json()) as TokenResponse;

    if (tokenData.access_token) {
      const tokenExpiresAt = Date.now() + (60 * 60 * 1000); // Set expiration to 1 hour from now
      setAccessToken(tokenData.access_token, tokenData.refresh_token || null, tokenExpiresAt);
      deviceCodeInfo = null; // Clear the device code info

      log("INFO", "Authentication successful", {
        hasRefreshToken: !!tokenData.refresh_token,
        expiresAt: new Date(tokenExpiresAt).toISOString(),
      });

      return {
        success: true,
        message: "✅ Authentication successful! You can now use all Planner tools.",
      };
    }

    if (tokenData.error === "authorization_pending") {
      const remainingSeconds = Math.floor(
        (deviceCodeInfo.expiresAt - Date.now()) / 1000,
      );
      const remainingMinutes = Math.floor(remainingSeconds / 60);
      return {
        success: false,
        message: `⏳ Waiting for authentication... Please complete the sign-in on the Microsoft website and call auth_poll again.

Time remaining: ${remainingMinutes}m ${remainingSeconds % 60}s`,
      };
    }

    if (tokenData.error === "authorization_declined") {
      deviceCodeInfo = null;
      return {
        success: false,
        message: "❌ Authentication was declined. Please call auth_start to try again.",
      };
    }

    if (tokenData.error === "expired_token") {
      deviceCodeInfo = null;
      return {
        success: false,
        message: "❌ Authentication code has expired. Please call auth_start again.",
      };
    }

    // Unknown error
    return {
      success: false,
      message: `⚠️ Unexpected error: ${tokenData.error || "Unknown error"}. ${tokenData.error_description || ""}`,
    };
  } catch (error) {
    return {
      success: false,
      message: `Authentication check failed: ${handleGraphError(error)}`,
    };
  }
}

export function getDeviceCodeInfo(): DeviceCodeInfo | null {
  return deviceCodeInfo;
}
