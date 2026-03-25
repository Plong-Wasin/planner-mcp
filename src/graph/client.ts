import { Client } from "@microsoft/microsoft-graph-client";
import { getAccessToken, isTokenExpired, refreshAccessToken, loadTokenFromFile } from "../auth/token-manager.js";
import { log } from "../utils/logger.js";

export async function getGraphClient(): Promise<Client> {
  // If no token in memory, try loading from file
  if (!getAccessToken()) {
    await loadTokenFromFile();
  }

  // Still no token after trying to load from file
  if (!getAccessToken()) {
    throw new Error("Not authenticated. Please run authentication first or set GRAPH_ACCESS_TOKEN environment variable.");
  }

  if (isTokenExpired()) {
    log("INFO", "Access token expired, attempting to refresh");
    const refreshed = await refreshAccessToken();
    if (!refreshed || !getAccessToken()) {
      throw new Error("Access token expired and refresh failed. Please re-authenticate.");
    }
  }

  const client = Client.init({
    authProvider: (done) => {
      done(null, getAccessToken()!);
    },
  });

  return client;
}
