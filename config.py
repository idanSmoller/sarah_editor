# ─────────────────────────────────────────────────────────────────────────────
# Azure AD configuration — fill in your Client ID below to enable OneDrive.
# ─────────────────────────────────────────────────────────────────────────────
# How to get a Client ID (one-time setup, ~5 minutes):
#
#   1. Go to https://portal.azure.com and sign in with any Microsoft account.
#   2. Open "Azure Active Directory" → "App registrations" → "New registration".
#   3. Name: anything (e.g. "Video Editor").
#   4. Supported account types: "Accounts in any organizational directory
#      and personal Microsoft accounts".
#   5. Redirect URI: choose "Public client / native" → http://localhost
#   6. Click Register.
#   7. Copy the "Application (client) ID" shown on the overview page.
#   8. Go to "API permissions" → "Add a permission" → Microsoft Graph
#      → Delegated → add:  Files.Read   and   Files.ReadWrite
#
# Paste your ID below, replacing the placeholder string.
# ─────────────────────────────────────────────────────────────────────────────

AZURE_CLIENT_ID = "YOUR_CLIENT_ID_HERE"
