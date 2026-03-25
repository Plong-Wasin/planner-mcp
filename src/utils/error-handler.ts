export function handleGraphError(error: any): string {
  if (error?.code) {
    return `Graph API Error (${error.code}): ${error.message || "Unknown error"}`;
  }
  return `Error: ${error?.message || String(error)}`;
}
