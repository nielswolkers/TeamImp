# TeamImp

TeamImp is a macOS desktop companion for Microsoft Teams and Microsoft 365 school assignments.

## Views

- Meldingen → Chats: browse chats and read recent messages.
- Meldingen → Vermeldingen: find recent messages that @mention the signed-in user.
- Taken → Openstaand / Verstreken / Voltooid: group assignments by due date and submission state.

## Microsoft Entra setup

Create an app registration in the Microsoft Entra admin center.

1. Choose accounts in this organizational directory only (single tenant) or accounts in any organizational directory (multi-tenant).
2. Under Authentication, enable Allow public client flows.
3. Add the delegated Microsoft Graph permissions below.
4. Copy the Application (client) ID into TeamImp → Settings.

Core delegated permissions requested at sign-in:

- User.Read
- EduAssignments.ReadBasic
- Chat.ReadBasic
- Chat.Read
- offline_access, openid, profile

For channel @mentions, TeamImp offers a separate optional consent flow for:

- Team.ReadBasic.All
- Channel.ReadBasic.All
- ChannelMessage.Read.All

ChannelMessage.Read.All is a privileged permission and commonly requires school administrator consent. If the school tenant does not allow it, chat reading and assignments can still work; channel mentions remain unavailable.

## Authentication and privacy

TeamImp uses the OAuth 2.0 device authorization grant for a public desktop client. No client secret is embedded in the app. Microsoft refresh tokens are stored in the macOS Keychain through the Rust/Tauri backend. Access tokens remain in memory. Sign out removes the stored refresh token.

## Assignment status

Assignments come from the education/me/assignments Graph endpoint. TeamImp checks the current student's submission records where available. An assignment is classified as completed when a submission is marked submitted/released/returned; otherwise it is overdue when its due date is in the past, and open when it is not overdue. If the school tenant limits submission reads, the due-date classification is used as a fallback.

## Sync

The app refreshes data every five minutes while open, and the refresh button triggers a manual sync. This desktop-only build polls Microsoft Graph. Graph change-notification subscriptions require a publicly reachable webhook service and are not part of this local DMG.

## Build locally

Requirements: macOS, Bun, Rust, and the Tauri v2 prerequisites.

Run bun install, then bash scripts/generate-macos-icon.sh, then bunx tauri dev.

Create the Apple Silicon DMG with:
bunx tauri build --target aarch64-apple-darwin --bundles dmg

GitHub Actions runs the same DMG build on every push to main and stores the installer as the TeamImp-macOS-Apple-Silicon workflow artifact.
