import { invoke } from "@tauri-apps/api/core";

const AUTHORITY = "https://login.microsoftonline.com/organizations/oauth2/v2.0";
const GRAPH = "https://graph.microsoft.com/v1.0";
const CLIENT_ID_KEY = "teamimp.microsoft.clientId";
const CHANNEL_ACCESS_KEY = "teamimp.microsoft.channelAccess";

export const BASE_SCOPES = [
  "openid", "profile", "offline_access", "User.Read",
  "EduAssignments.ReadBasic", "Chat.ReadBasic", "Chat.Read"
];
export const CHANNEL_SCOPES = [
  "Team.ReadBasic.All", "Channel.ReadBasic.All", "ChannelMessage.Read.All"
];

export type DeviceChallenge = {
  device_code: string; user_code: string; verification_uri: string;
  message: string; interval: number; expires_in: number;
};
export type Assignment = {
  id: string; displayName?: string; classId?: string; className?: string;
  dueDateTime?: string; assignedDateTime?: string;
  instructions?: { content?: string; contentType?: string };
  webUrl?: string; status?: string; submissionState?: string;
};
export type TeamChat = {
  id: string; topic?: string; chatType?: string;
  lastMessagePreview?: { body?: { content?: string }; createdDateTime?: string; from?: { user?: { displayName?: string } } };
  members?: { userId?: string; displayName?: string }[];
};
export type ChatMessage = {
  id: string; createdDateTime?: string; lastModifiedDateTime?: string;
  from?: { user?: { id?: string; displayName?: string; userIdentityType?: string } };
  body?: { content?: string; contentType?: string };
  mentions?: { id?: number; mentionText?: string; mentioned?: { user?: { id?: string; displayName?: string } } }[];
  webUrl?: string;
};
export type GraphUser = { id: string; displayName: string; userPrincipalName?: string };

let accessToken: string | null = null;
let tokenExpiresAt = 0;

export function getClientId() {
  return localStorage.getItem(CLIENT_ID_KEY) || import.meta.env.VITE_MICROSOFT_CLIENT_ID || "";
}
export function saveClientId(value: string) {
  localStorage.setItem(CLIENT_ID_KEY, value.trim());
}
export function hasChannelAccess() {
  return localStorage.getItem(CHANNEL_ACCESS_KEY) === "true";
}
export function setChannelAccess(enabled: boolean) {
  localStorage.setItem(CHANNEL_ACCESS_KEY, enabled ? "true" : "false");
}

async function tokenRequest(body: URLSearchParams) {
  const response = await fetch(AUTHORITY + "/token", {
    method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body
  });
  const payload = await response.json();
  if (!response.ok) throw new Error(payload.error_description || payload.error || "Microsoft login failed.");
  return payload as { access_token: string; refresh_token?: string; expires_in: number; scope?: string };
}

export async function requestDeviceCode(extraScopes: string[] = []): Promise<DeviceChallenge> {
  const clientId = getClientId();
  if (!clientId) throw new Error("Vul eerst je Microsoft Entra Client ID in bij Instellingen.");
  const response = await fetch(AUTHORITY + "/devicecode", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ client_id: clientId, scope: [...BASE_SCOPES, ...extraScopes].join(" ") })
  });
  const payload = await response.json();
  if (!response.ok) throw new Error(payload.error_description || payload.error || "Device-login kan niet worden gestart.");
  return payload as DeviceChallenge;
}

export async function finishDeviceCodeLogin(challenge: DeviceChallenge, extraScopes: string[] = []) {
  const deadline = Date.now() + challenge.expires_in * 1000;
  let interval = Math.max(challenge.interval || 5, 5);
  while (Date.now() < deadline) {
    await new Promise((resolve) => window.setTimeout(resolve, interval * 1000));
    const response = await fetch(AUTHORITY + "/token", {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        grant_type: "urn:ietf:params:oauth:grant-type:device_code",
        client_id: getClientId(),
        device_code: challenge.device_code
      })
    });
    const payload = await response.json();
    if (response.ok && payload.access_token) {
      accessToken = payload.access_token;
      tokenExpiresAt = Date.now() + Number(payload.expires_in || 3600) * 1000;
      if (payload.refresh_token) await invoke("set_microsoft_refresh_token", { token: payload.refresh_token });
      if (extraScopes.length) setChannelAccess(true);
      return;
    }
    if (payload.error === "authorization_pending") continue;
    if (payload.error === "slow_down") { interval += 5; continue; }
    if (payload.error === "authorization_declined") throw new Error("De Microsoft-aanmelding is geweigerd.");
    if (payload.error === "expired_token") throw new Error("De aanmeldcode is verlopen. Start de login opnieuw.");
    throw new Error(payload.error_description || payload.error || "Microsoft-aanmelding mislukt.");
  }
  throw new Error("De aanmeldcode is verlopen. Start de login opnieuw.");
}

async function refreshAccessToken() {
  const refreshToken = await invoke<string | null>("get_microsoft_refresh_token");
  if (!refreshToken) throw new Error("Niet aangemeld bij Microsoft.");
  const scopes = [...BASE_SCOPES, ...(hasChannelAccess() ? CHANNEL_SCOPES : [])];
  const result = await tokenRequest(new URLSearchParams({
    client_id: getClientId(), grant_type: "refresh_token",
    refresh_token: refreshToken, scope: scopes.join(" ")
  }));
  accessToken = result.access_token;
  tokenExpiresAt = Date.now() + Number(result.expires_in || 3600) * 1000;
  if (result.refresh_token) await invoke("set_microsoft_refresh_token", { token: result.refresh_token });
  return accessToken;
}

async function getAccessToken() {
  if (accessToken && Date.now() < tokenExpiresAt - 60_000) return accessToken;
  try {
    return await refreshAccessToken();
  } catch (error) {
    if (hasChannelAccess()) {
      setChannelAccess(false);
      return refreshAccessToken();
    }
    throw error;
  }
}

async function graph<T>(path: string): Promise<T> {
  const token = await getAccessToken();
  const response = await fetch(GRAPH + path, {
    headers: { Authorization: "Bearer " + token, Accept: "application/json" }
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    const code = payload?.error?.code || response.status;
    const message = payload?.error?.message || response.statusText;
    throw new Error("Microsoft Graph (" + code + "): " + message);
  }
  return payload as T;
}

async function collection<T>(path: string, pageLimit = 5): Promise<T[]> {
  const all: T[] = [];
  let next: string | undefined = GRAPH + path;
  let pages = 0;
  while (next && pages < pageLimit) {
    const token = await getAccessToken();
    const response = await fetch(next, { headers: { Authorization: "Bearer " + token, Accept: "application/json" } });
    const payload = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error("Microsoft Graph: " + (payload?.error?.message || response.statusText));
    all.push(...((payload.value || []) as T[]));
    next = payload["@odata.nextLink"];
    pages += 1;
  }
  return all;
}

export async function getMe() {
  return graph<GraphUser>("/me?$select=id,displayName,userPrincipalName");
}

export async function getAssignments(): Promise<Assignment[]> {
  const assignments = await collection<Assignment>("/education/me/assignments?$top=100&$orderby=dueDateTime", 8);
  return Promise.all(assignments.map(async (assignment) => {
    if (!assignment.classId || !assignment.id) return assignment;
    try {
      const submissions = await collection<{ state?: string; submittedDateTime?: string }>(
        "/education/classes/" + encodeURIComponent(assignment.classId) + "/assignments/" +
        encodeURIComponent(assignment.id) + "/submissions?$top=20", 1
      );
      const submitted = submissions.some((item) =>
        ["submitted", "released", "returned"].includes((item.state || "").toLowerCase()) || Boolean(item.submittedDateTime)
      );
      return { ...assignment, submissionState: submitted ? "submitted" : (submissions[0]?.state || "working") };
    } catch {
      return assignment;
    }
  }));
}

export async function getChats(): Promise<TeamChat[]> {
  return collection<TeamChat>("/me/chats?$top=50&$expand=members,lastMessagePreview", 8);
}
export async function getChatMessages(chatId: string): Promise<ChatMessage[]> {
  return collection<ChatMessage>("/me/chats/" + encodeURIComponent(chatId) + "/messages?$top=50", 3);
}

export async function getMentions(chats: TeamChat[], userId: string): Promise<ChatMessage[]> {
  const chatResults = await Promise.all(chats.slice(0, 20).map(async (chat) => {
    try { return await getChatMessages(chat.id); } catch { return []; }
  }));
  const chatMentions = chatResults.flat().filter((message) =>
    message.mentions?.some((mention) => mention.mentioned?.user?.id === userId)
  );
  if (!hasChannelAccess()) return chatMentions;

  try {
    const teams = await collection<{ id: string; displayName: string }>("/me/joinedTeams?$top=50", 3);
    const channelLists = await Promise.all(teams.slice(0, 12).map(async (team) => {
      try {
        const channels = await collection<{ id: string; displayName: string }>(
          "/teams/" + encodeURIComponent(team.id) + "/channels?$top=30", 2
        );
        return channels.map((channel) => ({ ...channel, teamId: team.id }));
      } catch { return []; }
    }));
    const channels = channelLists.flat().slice(0, 40);
    const channelMessages = await Promise.all(channels.map(async (channel) => {
      try {
        return await collection<ChatMessage>(
          "/teams/" + encodeURIComponent(channel.teamId) + "/channels/" + encodeURIComponent(channel.id) + "/messages?$top=30", 1
        );
      } catch { return []; }
    }));
    const channelMentions = channelMessages.flat().filter((message) =>
      message.mentions?.some((mention) => mention.mentioned?.user?.id === userId)
    );
    return [...chatMentions, ...channelMentions].sort((a, b) =>
      new Date(b.createdDateTime || "").getTime() - new Date(a.createdDateTime || "").getTime()
    );
  } catch {
    return chatMentions;
  }
}

export async function signOut() {
  accessToken = null;
  tokenExpiresAt = 0;
  setChannelAccess(false);
  await invoke("delete_microsoft_refresh_token");
}

export function plainText(html?: string) {
  if (!html) return "";
  const element = document.createElement("div");
  element.innerHTML = html;
  return element.textContent || element.innerText || "";
}

export function assignmentBucket(assignment: Assignment): "open" | "overdue" | "completed" {
  if (assignment.submissionState === "submitted") return "completed";
  if (assignment.dueDateTime && new Date(assignment.dueDateTime).getTime() < Date.now()) return "overdue";
  return "open";
}

export function formatDate(value?: string) {
  if (!value) return "Geen deadline";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "Geen deadline";
  return new Intl.DateTimeFormat("nl-NL", {
    weekday: "short", day: "numeric", month: "short", hour: "2-digit", minute: "2-digit"
  }).format(date);
}
export function openMicrosoftSignInUrl() { return "https://microsoft.com/devicelogin"; }
