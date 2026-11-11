import { useEffect, useMemo, useState, type ReactNode } from "react";
import { invoke } from "@tauri-apps/api/core";
import { listen } from "@tauri-apps/api/event";
import {
  AtSign, Bell, BookOpen, CheckCircle2, ChevronRight, CircleAlert, Clock3,
  Inbox, LoaderCircle, LogOut, MessageCircle, RefreshCw, Search, ShieldCheck, Sparkles
} from "lucide-react";
import {
  Assignment, ChatMessage, TeamChat, assignmentBucket, formatDate, getAssignments, getChats,
  getClientId, getMe, getMentions, hasChannelAccess, plainText, signOut, CHANNEL_SCOPES,
  getChatMessages, setAccessToken, setChannelAccess
} from "./lib/microsoft";

type MainView = "notifications" | "tasks";
type NoticeView = "chats" | "mentions";
type TaskView = "open" | "overdue" | "completed";

function App() {
  const [mainView, setMainView] = useState<MainView>("tasks");
  const [noticeView, setNoticeView] = useState<NoticeView>("chats");
  const [taskView, setTaskView] = useState<TaskView>("open");
  const [connected, setConnected] = useState(false);
  const [profileName, setProfileName] = useState("");
  const [profileId, setProfileId] = useState("");
  const [assignments, setAssignments] = useState<Assignment[]>([]);
  const [chats, setChats] = useState<TeamChat[]>([]);
  const [messages, setMessages] = useState<ChatMessage[]>([]);
  const [mentions, setMentions] = useState<ChatMessage[]>([]);
  const [selectedChat, setSelectedChat] = useState("");
  const [query, setQuery] = useState("");
  const [busy, setBusy] = useState(false);
  const [loadingChat, setLoadingChat] = useState(false);
  const [error, setError] = useState("");
  const [loginStatus, setLoginStatus] = useState("");
  const [channelPermission, setChannelPermission] = useState(hasChannelAccess());
  const [lastSynced, setLastSynced] = useState<Date | null>(null);

  const selectedChatItem = chats.find((chat) => chat.id === selectedChat);

  async function loadWorkspace() {
    setBusy(true);
    setError("");
    try {
      const me = await getMe();
      setProfileName(me.displayName || me.userPrincipalName || "Microsoft-account");
      setProfileId(me.id);
      const [assignmentResult, chatResult] = await Promise.allSettled([getAssignments(), getChats()]);
      const assignmentData = assignmentResult.status === "fulfilled" ? assignmentResult.value : [];
      const chatData = chatResult.status === "fulfilled" ? chatResult.value : [];
      setAssignments(assignmentData);
      setChats(chatData);
      setConnected(true);
      setLastSynced(new Date());
      setSelectedChat((current) => current || chatData[0]?.id || "");
      const warnings: string[] = [];
      if (assignmentResult.status === "rejected") warnings.push("Taken niet beschikbaar: " + String(assignmentResult.reason));
      if (chatResult.status === "rejected") warnings.push("Chats niet beschikbaar: " + String(chatResult.reason));
      setError(warnings.join(" · "));
      const mentionData = await getMentions(chatData, me.id);
      setMentions(mentionData);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Synchroniseren is mislukt.");
      setConnected(false);
    } finally {
      setBusy(false);
    }
  }

  useEffect(() => {
    let active = true;
    getMe().then(() => {
      if (active) void loadWorkspace();
    }).catch(() => undefined);
    return () => { active = false; };
  }, []);

  useEffect(() => {
    if (!connected || !selectedChat) return;
    let active = true;
    setLoadingChat(true);
    getChatMessages(selectedChat)
      .then((data) => {
        if (active) setMessages(data.sort((a, b) =>
          new Date(a.createdDateTime || "").getTime() - new Date(b.createdDateTime || "").getTime()
        ));
      })
      .catch((err) => { if (active) setError(err instanceof Error ? err.message : "Chatberichten ophalen mislukt."); })
      .finally(() => { if (active) setLoadingChat(false); });
    return () => { active = false; };
  }, [connected, selectedChat]);

  useEffect(() => {
    if (!connected) return;
    const interval = window.setInterval(() => { void loadWorkspace(); }, 5 * 60 * 1000);
    return () => window.clearInterval(interval);
  }, [connected]);

  async function beginLogin(extraScopes: string[] = []) {
    setError("");
    setLoginStatus("Microsoft-aanmelding openen…");
    try {
      let removeSuccess: (() => void) | undefined;
      let removeError: (() => void) | undefined;
      removeSuccess = await listen<{ access_token: string; expires_in: number }>("microsoft-auth-success", async (event) => {
        removeSuccess?.();
        removeError?.();
        setAccessToken(event.payload.access_token, event.payload.expires_in);
        if (extraScopes.length > 0) {
          setChannelAccess(true);
          setChannelPermission(true);
        }
        setLoginStatus("");
        await loadWorkspace();
      });
      removeError = await listen<string>("microsoft-auth-error", (event) => {
        removeSuccess?.();
        removeError?.();
        setLoginStatus("");
        setError(event.payload);
      });
      await invoke("start_microsoft_login", {
        clientId: getClientId(),
        extraScopes
      });
      setLoginStatus("Log in in het Microsoft-venster. Je keert daarna automatisch terug naar TeamImp.");
    } catch (err) {
      setLoginStatus("");
      setError(err instanceof Error ? err.message : "Aanmelden mislukt.");
    }
  }

  async function doSignOut() {
    await signOut();
    setConnected(false);
    setAssignments([]);
    setChats([]);
    setMessages([]);
    setMentions([]);
    setProfileName("");
    setProfileId("");
    setError("");
  }

  const visibleAssignments = useMemo(() => assignments
    .filter((item) => assignmentBucket(item) === taskView)
    .filter((item) => (item.displayName || "").toLowerCase().includes(query.toLowerCase()))
    .sort((a, b) => new Date(a.dueDateTime || "2999-01-01").getTime() - new Date(b.dueDateTime || "2999-01-01").getTime()), [assignments, taskView, query]);

  const filteredChats = chats.filter((chat) =>
    (chat.topic || chat.lastMessagePreview?.from?.user?.displayName || chat.chatType || "Chat")
      .toLowerCase().includes(query.toLowerCase())
  );
  const counts = {
    open: assignments.filter((item) => assignmentBucket(item) === "open").length,
    overdue: assignments.filter((item) => assignmentBucket(item) === "overdue").length,
    completed: assignments.filter((item) => assignmentBucket(item) === "completed").length
  };
  const taskTabs: { id: TaskView; label: string; count: number }[] = [
    { id: "open", label: "Openstaand", count: counts.open },
    { id: "overdue", label: "Verstreken", count: counts.overdue },
    { id: "completed", label: "Voltooid", count: counts.completed }
  ];

  return (
    <div className="app-shell">
      <aside className="rail">
        <div className="brand-mark">T</div><div className="rail-divider" />
        <button className={"rail-button " + (mainView === "tasks" ? "active" : "")} onClick={() => setMainView("tasks")} title="Taken"><BookOpen size={19} /></button>
        <button className={"rail-button " + (mainView === "notifications" ? "active" : "")} onClick={() => setMainView("notifications")} title="Meldingen"><Bell size={19} /></button>
        <div className="rail-spacer" />
        {connected && <button className="rail-button signout" onClick={() => void doSignOut()} title="Afmelden"><LogOut size={18} /></button>}
        <div className="rail-avatar">{profileName ? profileName.slice(0, 1).toUpperCase() : "T"}</div>
      </aside>
      <main className="main-area">
        <header className="topbar">
          <div><div className="eyebrow">MICROSOFT TEAMS</div><h1>TeamImp</h1><p className="subtitle">Alles wat je voor school moet weten, op één plek.</p></div>
          <div className="top-actions">
            <div className={"connection-pill " + (connected ? "connected" : "")}><span className="status-dot" />{connected ? "Verbonden" : "Niet verbonden"}</div>
            <button className="icon-button" title="Synchroniseren" onClick={() => void loadWorkspace()} disabled={!connected || busy}><RefreshCw size={17} className={busy ? "spin" : ""} /></button>
          </div>
        </header>

        {!connected ? (
          <section className="connect-panel">
            <div className="connect-orb"><MessageCircle size={28} /></div>
            <div className="eyebrow">JOUW SCHOOLACCOUNT</div>
            <h2>Verbind met Microsoft Teams</h2>
            <p className="muted copy-width">Meld je aan met het Microsoft 365-account van school om je opdrachten, chats en vermeldingen op te halen.</p>
            <button className="primary-button" onClick={() => void beginLogin()}><span>Aanmelden met Microsoft</span><ChevronRight size={17} /></button>
            <p className="muted small copy-width">Je logt rechtstreeks in bij Microsoft. Je wachtwoord wordt niet door TeamImp opgeslagen.</p>
            {loginStatus && <div className="login-status"><LoaderCircle size={16} className="spin" />{loginStatus}</div>}

          </section>
        ) : (
          <>
            <nav className="main-tabs">
              <button className={mainView === "notifications" ? "selected" : ""} onClick={() => setMainView("notifications")}><Bell size={16} />Meldingen{mentions.length > 0 && <span className="tab-count">{mentions.length}</span>}</button>
              <button className={mainView === "tasks" ? "selected" : ""} onClick={() => setMainView("tasks")}><BookOpen size={16} />Taken{counts.open > 0 && <span className="tab-count">{counts.open}</span>}</button>
            </nav>
            <div className="content-toolbar">
              <div className="subtabs">
                {mainView === "tasks" ? taskTabs.map((tab) => (
                  <button key={tab.id} className={taskView === tab.id ? "subtab selected" : "subtab"} onClick={() => setTaskView(tab.id)}>
                    {tab.label}<span className="subtab-count">{tab.count}</span>
                  </button>
                )) : (
                  <>
                    <button className={noticeView === "chats" ? "subtab selected" : "subtab"} onClick={() => setNoticeView("chats")}><MessageCircle size={15} />Chats</button>
                    <button className={noticeView === "mentions" ? "subtab selected" : "subtab"} onClick={() => setNoticeView("mentions")}><AtSign size={15} />Vermeldingen<span className="subtab-count">{mentions.length}</span></button>
                  </>
                )}
              </div>
              <div className="search-box"><Search size={15} /><input value={query} onChange={(event) => setQuery(event.target.value)} placeholder={mainView === "tasks" ? "Zoek opdracht…" : "Zoek chat of bericht…"} /></div>
            </div>

            {mainView === "tasks" && <section className="task-view">
              <div className="section-heading"><div><div className="eyebrow">OPDRACHTEN</div><h2>{taskTabs.find((tab) => tab.id === taskView)?.label}</h2></div><span className="muted small">{visibleAssignments.length} opdrachten</span></div>
              {error && <div className="error-banner"><CircleAlert size={16} />{error}</div>}
              {busy && assignments.length === 0 ? <Loading /> : visibleAssignments.length === 0 ? <EmptyState icon={<Inbox size={23} />} title="Geen opdrachten gevonden" body="Er staan geen opdrachten in deze categorie. Synchroniseer om de laatste status op te halen." /> : (
                <div className="assignment-list">{visibleAssignments.map((assignment) => (
                  <article className="assignment-card" key={assignment.id}>
                    <div className={"assignment-icon " + taskView}>{taskView === "completed" ? <CheckCircle2 size={19} /> : taskView === "overdue" ? <Clock3 size={19} /> : <BookOpen size={19} />}</div>
                    <div className="assignment-main">
                      <h3>{assignment.displayName || "Opdracht zonder titel"}</h3>
                      <p>{assignment.className || "Microsoft Teams"}{assignment.submissionState === "submitted" ? " · Ingeleverd" : ""}</p>
                      {assignment.instructions?.content && <div className="assignment-description">{plainText(assignment.instructions.content).slice(0, 180)}</div>}
                    </div>
                    <div className="assignment-due"><span>{taskView === "completed" ? "Status" : "Deadline"}</span><strong>{taskView === "completed" ? "Ingeleverd" : formatDate(assignment.dueDateTime)}</strong></div>
                  </article>
                ))}</div>
              )}
            </section>}

            {mainView === "notifications" && noticeView === "chats" && (
              <section className="chat-layout">
                <div className="chat-list-panel">
                  <div className="panel-heading"><div><div className="eyebrow">TEAMS</div><h2>Chats</h2></div><span className="muted small">{filteredChats.length}</span></div>
                  <div className="chat-list">
                    {filteredChats.map((chat) => (
                      <button key={chat.id} className={"chat-list-item " + (selectedChat === chat.id ? "selected" : "")} onClick={() => setSelectedChat(chat.id)}>
                        <div className="chat-avatar">{(chat.topic || chat.lastMessagePreview?.from?.user?.displayName || "T").slice(0, 1).toUpperCase()}</div>
                        <div className="chat-info"><strong>{chat.topic || chat.members?.map((member) => member.displayName).filter(Boolean).join(", ") || (chat.chatType === "oneOnOne" ? "Persoonlijke chat" : "Groepschat")}</strong><span>{plainText(chat.lastMessagePreview?.body?.content).slice(0, 52) || chat.chatType || "Microsoft Teams-chat"}</span></div>
                      </button>
                    ))}
                    {filteredChats.length === 0 && <p className="muted empty-inline">Geen chats gevonden.</p>}
                  </div>
                </div>
                <div className="message-panel">
                  {selectedChatItem ? <>
                    <div className="message-header"><div className="chat-avatar large">{(selectedChatItem.topic || "T").slice(0, 1).toUpperCase()}</div><div><h3>{selectedChatItem.topic || "Teams-chat"}</h3><span className="muted small">{selectedChatItem.chatType || "Gesprek"}</span></div></div>
                    <div className="message-stream">
                      {loadingChat ? <Loading /> : messages.length === 0 ? <EmptyState icon={<MessageCircle size={22} />} title="Nog geen berichten" body="Deze chat bevat geen berichten die TeamImp kan ophalen." /> : messages.map((message) => (
                        <article className={"message-bubble " + (message.from?.user?.id === profileId ? "own" : "")} key={message.id}>
                          <div className="message-author">{message.from?.user?.displayName || "Teams-gebruiker"}<span>{message.createdDateTime ? formatDate(message.createdDateTime) : ""}</span></div>
                          <div className="message-text">{plainText(message.body?.content)}</div>
                          {message.mentions && message.mentions.length > 0 && <div className="mention-tag"><AtSign size={12} /> Vermelding</div>}
                        </article>
                      ))}
                    </div>
                  </> : <EmptyState icon={<MessageCircle size={22} />} title="Selecteer een chat" body="Kies links een gesprek om de berichten te bekijken." />}
                </div>
              </section>
            )}

            {mainView === "notifications" && noticeView === "mentions" && (
              <section className="task-view">
                <div className="section-heading"><div><div className="eyebrow">DIRECTE AANDACHT</div><h2>Vermeldingen</h2></div><span className="muted small">{mentions.length} gevonden</span></div>
                {!channelPermission && <div className="permission-banner"><ShieldCheck size={17} /><div><strong>Chatvermeldingen actief</strong><span>Voor vermeldingen in teamkanalen is aanvullende toestemming van je school nodig.</span></div><button className="outline-button" onClick={() => void beginLogin(CHANNEL_SCOPES)}>Toegang aanvragen</button></div>}
                {mentions.length === 0 ? <EmptyState icon={<AtSign size={23} />} title="Geen vermeldingen gevonden" body="Er zijn geen recente @vermeldingen gevonden in de chats die TeamImp kon uitlezen." /> : (
                  <div className="mention-list">{mentions.map((message) => (
                    <article className="mention-card" key={message.id}><div className="mention-symbol"><AtSign size={18} /></div><div><strong>{message.from?.user?.displayName || "Teams-gebruiker"}</strong><p>{plainText(message.body?.content)}</p><span>{formatDate(message.createdDateTime)}</span></div></article>
                  ))}</div>
                )}
              </section>
            )}
            <footer className="footer-line"><span><Sparkles size={13} />{lastSynced ? "Laatst gesynchroniseerd " + new Intl.DateTimeFormat("nl-NL", { hour: "2-digit", minute: "2-digit" }).format(lastSynced) : "TeamImp"}</span><span>Automatische synchronisatie elke 5 minuten</span></footer>
          </>
        )}
      </main>

      {error && !connected && <div className="toast-error"><CircleAlert size={16} />{error}<button onClick={() => setError("")}><X size={14} /></button></div>}
    </div>
  );
}

function Loading() {
  return <div className="loading-state"><LoaderCircle className="spin" size={22} /><span>Gegevens ophalen…</span></div>;
}
function EmptyState({ icon, title, body }: { icon: ReactNode; title: string; body: string }) {
  return <div className="empty-state"><div className="empty-icon">{icon}</div><h3>{title}</h3><p>{body}</p></div>;
}
export default App;
