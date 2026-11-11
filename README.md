# TeamImp

TeamImp is a macOS desktop companion for Microsoft Teams and Microsoft 365 school assignments.

## Views

- Meldingen → Chats: browse chats and read recent messages.
- Meldingen → Vermeldingen: find recent messages that @mention the signed-in user.
- Taken → Openstaand / Verstreken / Voltooid: group assignments by due date and submission state.

## Microsoft sign-in (one-time app setup)

TeamImp signs in through the normal Microsoft page in the system browser. The user selects their school account, signs in with Microsoft, and is returned automatically to TeamImp. The app uses OAuth 2.0 Authorization Code with PKCE and a temporary localhost callback. No password or client secret is stored by TeamImp.

A public Microsoft Entra app registration is required for any third-party app that requests Microsoft Graph access. This is a one-time developer configuration, not a key that users enter in the app:

1. Create an app registration in Microsoft Entra.
2. Set the supported account type to **Accounts in any organizational directory** (multitenant).
3. Under **Authentication → Add a platform → Mobile and desktop applications**, add the system-browser redirect URI **http://localhost**.
4. Add these delegated Microsoft Graph permissions: `User.Read`, `EduAssignments.ReadBasic`, `Chat.ReadBasic`, and `Chat.Read`.
5. Set the repository Actions variable `MICROSOFT_CLIENT_ID` to the registration's **Application (client) ID**. It is a public identifier, not a secret. GitHub Actions injects it into the macOS build; it is never requested from the end user.

For channel @mentions, TeamImp requests additional delegated permissions when the user selects **Toegang aanvragen**:

- `Team.ReadBasic.All`
- `Channel.ReadBasic.All`
- `ChannelMessage.Read.All`

`ChannelMessage.Read.All` is privileged and may require school administrator consent. Assignment and chat access can still work without it.

## Token security

The authorization-code exchange and token refresh run in the native Rust backend. The refresh token is stored in the macOS Keychain; access tokens remain in memory. Sign out deletes the stored refresh token.

## Assignment status

Assignments come from the Microsoft Graph `education/me/assignments` endpoint. TeamImp checks the current student's submission records where available. An assignment is classified as completed when a submission is marked submitted/released/returned; otherwise it is overdue when its due date is in the past, and open when it is not overdue. If the school tenant limits submission reads, the due-date classification is used as a fallback.

## Sync

The app refreshes data every five minutes while open, and the refresh button triggers a manual sync. This desktop-only build polls Microsoft Graph. Graph change-notification subscriptions require a publicly reachable webhook service and are not part of this local DMG.

## Build locally

Requirements: macOS, Bun, Rust, and the Tauri v2 prerequisites.

Set `VITE_MICROSOFT_CLIENT_ID` to the app registration's public Application (client) ID, then run:

```bash
bun install
bash scripts/generate-macos-icon.sh
bunx tauri dev
bunx tauri build --target aarch64-apple-darwin --bundles dmg
```

GitHub Actions builds the Apple Silicon DMG on every push to `main` and stores it as the `TeamImp-macOS-Apple-Silicon` workflow artifact.
