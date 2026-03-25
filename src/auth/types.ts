export interface DeviceCodeResponse {
  device_code: string;
  user_code: string;
  verification_uri: string;
  expires_in: number;
  interval: number;
  message?: string;
}

export interface TokenResponse {
  access_token?: string;
  refresh_token?: string;
  error?: string;
  error_description?: string;
}

export interface DeviceCodeInfo {
  deviceCode: string;
  clientId: string;
  tenantId: string;
  interval: number;
  expiresAt: number;
}

export interface TokenData {
  accessToken: string | null;
  refreshToken: string | null;
  tokenExpiresAt: number | null;
}
