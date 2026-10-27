/* CASA Voice Log — record a memo, mail it FROM the PM's own mailbox TO casa@.
 *
 * CASA (inside the office) reads casa@, transcribes the audio, and posts it
 * word for word to the job's Procore daily log under the sender's account.
 * This page never talks to CASA directly: it only signs in with Microsoft and
 * sends mail through Graph, so it works from any phone, anywhere.
 *
 * Every memo is saved on the phone (IndexedDB) BEFORE the first send attempt.
 * The queue is retried on open, when the phone comes back online, when the
 * app returns to the screen, and every 30 s while anything waits. A memo
 * leaves the queue only after Microsoft accepts the mail.
 *
 * Mail shape (CASA's contract, backend/casa/voice_log.py):
 *   memo:  subject = job number only; body has "Recorded: <ISO>" and
 *          "Memo: <key>"; the audio attached.
 *   undo:  subject = "Voice log undo <key>"; body "UNDO".
 *
 * Job picker: CASA leaves the PM's active Procore jobs as JSON in a hidden
 * "CASA Voice Log" folder of their own mailbox (every 3 h); the app reads it
 * here and keeps a copy for offline use. Typing any number still works.
 */

"use strict";

const CONFIG = {
  clientId: "816a91fd-d257-4463-bef9-4199530c4aaf",          // CASA Sign-in
  tenantId: "2664842a-08ac-4c7b-92dc-82f97b30ac15",          // Cooper
  casaMailbox: "casa@cples.com",
  scopes: ["User.Read", "Mail.Send", "Mail.ReadWrite"],
  maxSeconds: 15 * 60,
  // sendMail takes the whole request in one call (4 MB limit, base64 adds a
  // third), so bigger memos go draft -> upload session -> send.
  directLimit: 2.8 * 1024 * 1024,
  chunk: 3 * 1024 * 1024,
  undoHours: 24,
  jobFolder: "CASA Voice Log",
  jobListMaxAgeMs: 60 * 60 * 1000,
  jobListShown: 50,
};

const GRAPH = "https://graph.microsoft.com/v1.0";
const JOB_RE = /^\d{5,6}$/;
const $ = (id) => document.getElementById(id);

/* ------------------------------------------------------------ storage */

const store = {
  get(key, fallback) {
    try { return JSON.parse(localStorage.getItem(key)) ?? fallback; }
    catch { return fallback; }
  },
  set(key, value) {
    try { localStorage.setItem(key, JSON.stringify(value)); } catch { /* private mode */ }
  },
};

// The memo queue lives in IndexedDB: it holds audio blobs, survives the app
// being closed, and survives the Microsoft sign-in redirect.
const queue = (() => {
  let dbp;
  function db() {
    dbp ??= new Promise((resolve, reject) => {
      const req = indexedDB.open("casa-voice-log", 1);
      req.onupgradeneeded = () => req.result.createObjectStore("outbox", { keyPath: "key" });
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
    return dbp;
  }
  async function tx(mode, fn) {
    const d = await db();
    return new Promise((resolve, reject) => {
      const t = d.transaction("outbox", mode);
      const req = fn(t.objectStore("outbox"));
      t.oncomplete = () => resolve(req?.result);
      t.onerror = () => reject(t.error);
    });
  }
  return {
    put: (item) => tx("readwrite", (s) => s.put(item)),
    remove: (key) => tx("readwrite", (s) => s.delete(key)),
    all: () => tx("readonly", (s) => s.getAll()),
  };
})();

// Recent memos shown on screen (no audio, just what the PM needs to see).
const recent = {
  list: () => store.get("recent", []),
  save(items) { store.set("recent", items.slice(0, 10)); render.recent(); },
  upsert(entry) {
    const items = recent.list().filter((r) => r.key !== entry.key);
    items.unshift(entry);
    recent.save(items);
  },
  patch(key, fields) {
    recent.save(recent.list().map((r) => (r.key === key ? { ...r, ...fields } : r)));
  },
};

/* ------------------------------------------------------------ sign-in */

const msalApp = new msal.PublicClientApplication({
  auth: {
    clientId: CONFIG.clientId,
    authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`,
    redirectUri: new URL("./", location.href).href,
    navigateToLoginRequestUrl: false,
  },
  cache: { cacheLocation: "localStorage" },
});
let account = null;

class NeedsSignIn extends Error {}

async function token(scopes) {
  if (!account) throw new NeedsSignIn();
  try {
    return (await msalApp.acquireTokenSilent({ scopes, account })).accessToken;
  } catch (e) {
    if (e instanceof msal.InteractionRequiredAuthError) throw new NeedsSignIn();
    throw e;   // usually no signal; the queue retries
  }
}

function signIn() {
  // Redirect, not popup: popups break in home-screen apps. Anything recorded
  // is already in the queue, so leaving the page loses nothing.
  if (account) msalApp.loginRedirect({ scopes: CONFIG.scopes, loginHint: account.username });
  else msalApp.loginRedirect({ scopes: CONFIG.scopes, prompt: "select_account" });
}

// The sign-in expires after 24 h and iPhones block the silent renewal, so the
// app goes to Microsoft by itself: it comes straight back if Microsoft still
// knows the phone, or asks only for the password. Never while a recording is
// unsent, never offline, and at most once in AUTO_GAP_MS (no redirect loops).
const AUTO_GAP_MS = 3 * 60 * 1000;

function autoSignIn() {
  const busy = rec.media?.state === "recording" || !!rec.blob;
  const recent = Date.now() - (store.get("autoSignInAt", 0) || 0) < AUTO_GAP_MS;
  if (!account || busy || recent || !navigator.onLine || document.hidden) {
    showSignInPrompt();
    return;
  }
  store.set("autoSignInAt", Date.now());
  status("Signing you in again…", "");
  signIn();
}

async function checkSignIn() {
  if (!account || !navigator.onLine) return;
  try {
    await token(CONFIG.scopes);
    $("signin-view").hidden = true;
  } catch (e) {
    if (e instanceof NeedsSignIn) autoSignIn();
  }
}

/* ------------------------------------------------------------ graph */

async function graph(path, { method = "POST", body, tok, headers = {} }) {
  const res = await fetch(GRAPH + path, {
    method,
    headers: { Authorization: `Bearer ${tok}`, "Content-Type": "application/json", ...headers },
    body: body === undefined ? undefined : JSON.stringify(body),
  });
  if (!res.ok) {
    let detail = "";
    try { detail = (await res.json()).error?.message || ""; } catch { /* no body */ }
    const err = new Error(`Microsoft said ${res.status}${detail ? `: ${detail}` : ""}`);
    err.status = res.status;
    throw err;
  }
  return res.status === 204 || res.status === 202 ? null : res.json();
}

function toBase64(blob) {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(String(r.result).split(",")[1]);
    r.onerror = () => reject(r.error);
    r.readAsDataURL(blob);
  });
}

function message(item) {
  return {
    subject: item.kind === "undo" ? `Voice log undo ${item.memoKey}` : item.job,
    body: {
      contentType: "Text",
      content: item.kind === "undo"
        ? "UNDO\n\nSent from CASA Voice Log.\n"
        : `Recorded: ${item.recordedAt}\nMemo: ${item.key}\n\nSent from CASA Voice Log.\n`,
    },
    toRecipients: [{ emailAddress: { address: CONFIG.casaMailbox } }],
  };
}

async function sendItem(item) {
  if (item.kind === "undo" || item.blob.size <= CONFIG.directLimit) {
    const msg = message(item);
    if (item.kind === "memo") {
      msg.attachments = [{
        "@odata.type": "#microsoft.graph.fileAttachment",
        name: item.fileName,
        contentType: item.mime,
        contentBytes: await toBase64(item.blob),
      }];
    }
    await graph("/me/sendMail", { body: { message: msg, saveToSentItems: true },
                                  tok: await token(["Mail.Send"]) });
    return;
  }
  // Long memo: draft, upload the audio in chunks, then send the draft.
  const tok = await token(["Mail.ReadWrite", "Mail.Send"]);
  if (item.draftId) {            // a half-done earlier try; start clean
    await graph(`/me/messages/${item.draftId}`, { method: "DELETE", tok }).catch(() => {});
  }
  const draft = await graph("/me/messages", { body: message(item), tok });
  item.draftId = draft.id;
  await queue.put(item);
  const session = await graph(`/me/messages/${draft.id}/attachments/createUploadSession`, {
    body: { AttachmentItem: { attachmentType: "file", name: item.fileName,
                              size: item.blob.size, contentType: item.mime } },
    tok,
  });
  for (let start = 0; start < item.blob.size; start += CONFIG.chunk) {
    const end = Math.min(start + CONFIG.chunk, item.blob.size);
    const res = await fetch(session.uploadUrl, {
      method: "PUT",
      headers: { "Content-Type": "application/octet-stream",
                 "Content-Range": `bytes ${start}-${end - 1}/${item.blob.size}` },
      body: item.blob.slice(start, end),
    });
    if (!res.ok) throw new Error(`Audio upload stopped (${res.status})`);
  }
  await graph(`/me/messages/${draft.id}/send`, { tok });
}

/* ------------------------------------------------------------ queue flush */

let flushing = false;
let retryTimer = null;

async function flush() {
  if (flushing) return;
  flushing = true;
  try {
    const items = (await queue.all()).sort((a, b) => a.queuedAt - b.queuedAt);
    for (const item of items) {
      try {
        await sendItem(item);
        await queue.remove(item.key);
        if (item.kind === "undo") recent.patch(item.memoKey, { state: "undo-sent" });
        else recent.patch(item.key, { state: "sent", sentAt: Date.now() });
        status(item.kind === "undo" ? "Undo sent. CASA will email you when it is removed."
                                    : `Memo for ${item.job} sent to CASA.`, "ok");
      } catch (e) {
        if (e instanceof NeedsSignIn) {
          status("Sign in again to send your saved memos.", "warn");
          autoSignIn();
          break;
        }
        const offline = !navigator.onLine || e instanceof TypeError;
        status(offline ? "No signal. Your memo is saved and will send when the phone is back online."
                       : `Not sent yet: ${e.message}. Saved; trying again.`, "warn");
        break;   // keep the order; try again later
      }
    }
  } finally {
    flushing = false;
    const left = (await queue.all()).length;
    clearTimeout(retryTimer);
    if (left) retryTimer = setTimeout(flush, 30000);
    render.recent();
  }
}

/* ------------------------------------------------------------ recorder */

const rec = { media: null, stream: null, chunks: [], blob: null, mime: "",
              started: 0, tick: null, wake: null };

function pickMime() {
  const options = ["audio/mp4", "audio/webm;codecs=opus", "audio/webm", "audio/ogg;codecs=opus"];
  return options.find((m) => window.MediaRecorder && MediaRecorder.isTypeSupported(m)) || "";
}

function extFor(mime) {
  if (mime.includes("mp4")) return "m4a";
  if (mime.includes("webm")) return "webm";
  if (mime.includes("ogg")) return "ogg";
  return "m4a";
}

async function startRecording() {
  try {
    rec.stream = await navigator.mediaDevices.getUserMedia({ audio: true });
  } catch {
    status("Allow the microphone for this app, then try again.", "err");
    return;
  }
  const mime = pickMime();
  rec.media = new MediaRecorder(rec.stream, {
    ...(mime ? { mimeType: mime } : {}), audioBitsPerSecond: 32000,
  });
  rec.chunks = [];
  rec.media.ondataavailable = (e) => { if (e.data.size) rec.chunks.push(e.data); };
  rec.media.onstop = finishRecording;
  rec.media.start(1000);
  rec.started = Date.now();
  rec.tick = setInterval(updateTimer, 250);
  try { rec.wake = await navigator.wakeLock?.request("screen"); } catch { /* not supported */ }
  setMode("recording");
}

function stopRecording() {
  if (rec.media?.state === "recording") rec.media.stop();
}

function finishRecording() {
  clearInterval(rec.tick);
  rec.stream?.getTracks().forEach((t) => t.stop());
  rec.wake?.release?.().catch(() => {});
  rec.wake = null;
  rec.mime = (rec.media.mimeType || pickMime() || "audio/mp4").split(";")[0];
  rec.blob = new Blob(rec.chunks, { type: rec.mime });
  rec.seconds = Math.round((Date.now() - rec.started) / 1000);
  $("player").src = URL.createObjectURL(rec.blob);
  setMode("review");
}

function discardRecording() {
  if ($("player").src) URL.revokeObjectURL($("player").src);
  $("player").removeAttribute("src");
  rec.blob = null;
  setMode("idle");
}

function updateTimer() {
  const s = Math.floor((Date.now() - rec.started) / 1000);
  $("timer").textContent = clock(s);
  if (s >= CONFIG.maxSeconds) stopRecording();
}

/* ------------------------------------------------------------ send / undo */

function isoWithOffset(d) {
  const pad = (n) => String(n).padStart(2, "0");
  const off = -d.getTimezoneOffset();
  const sign = off >= 0 ? "+" : "-";
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T` +
         `${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}` +
         `${sign}${pad(Math.floor(Math.abs(off) / 60))}:${pad(Math.abs(off) % 60)}`;
}

function newKey() {
  return (crypto.randomUUID ? crypto.randomUUID() : String(Date.now()) + Math.random())
    .replace(/[^0-9a-f]/gi, "").toLowerCase().slice(0, 20);
}

async function sendMemo() {
  const job = $("job").value.trim();
  if (!JOB_RE.test(job)) {
    $("job").classList.add("bad");
    $("job").focus();
    status("Type the 5 or 6 digit job number first.", "err");
    return;
  }
  const started = new Date(rec.started);
  const key = newKey();
  const item = {
    key, kind: "memo", job, mime: rec.mime, blob: rec.blob,
    fileName: `Voice memo ${job} ${isoWithOffset(started).slice(0, 16).replace(":", "")}.${extFor(rec.mime)}`,
    recordedAt: isoWithOffset(started), queuedAt: Date.now(),
  };
  await queue.put(item);              // saved on the phone before any send
  recent.upsert({ key, job, recordedAt: started.getTime(), seconds: rec.seconds, state: "queued" });
  rememberJob(job);
  discardRecording();
  status("Saved. Sending…", "");
  flush();
}

async function undo(key) {
  const entry = recent.list().find((r) => r.key === key);
  if (!entry) return;
  const waiting = (await queue.all()).find((i) => i.key === key);
  if (waiting) {                      // never left the phone: just drop it
    await queue.remove(key);
    recent.patch(key, { state: "cancelled" });
    status("Memo deleted before it was sent.", "ok");
    return;
  }
  await queue.put({ key: `undo-${key}`, memoKey: key, kind: "undo", queuedAt: Date.now() });
  recent.patch(key, { state: "undo-queued" });
  status("Undo saved. Sending…", "");
  flush();
}

/* ------------------------------------------------------------ jobs */

function rememberJob(job) {
  store.set("jobs", [job, ...store.get("jobs", []).filter((j) => j !== job)].slice(0, 6));
  render.jobs();
}

const jobList = {
  all: () => store.get("jobList", { at: 0, jobs: [] }).jobs || [],
  find: (code) => jobList.all().find((j) => j.code === code),

  async refresh() {
    if (!account || !navigator.onLine) return;
    if (Date.now() - store.get("jobList", { at: 0 }).at < CONFIG.jobListMaxAgeMs) return;
    const tok = await token(["Mail.ReadWrite"]);
    const folders = await graph(
      "/me/mailFolders?includeHiddenFolders=true&$top=250&$select=id,displayName",
      { method: "GET", tok });
    const folder = folders.value.find((f) => f.displayName === CONFIG.jobFolder);
    if (!folder) return;                     // CASA has not written one yet
    const msgs = await graph(`/me/mailFolders/${folder.id}/messages?$top=1&$select=body`, {
      method: "GET", tok, headers: { Prefer: 'outlook.body-content-type="text"' } });
    const text = msgs.value[0]?.body?.content;
    if (!text) return;
    const jobs = (JSON.parse(text.trim()).jobs || []).filter((j) => JOB_RE.test(j.code));
    store.set("jobList", { at: Date.now(), jobs });
    render.jobs();
  },

  // Code prefix first, then any label containing the text.
  matches(text) {
    const q = text.trim().toLowerCase();
    const all = jobList.all();
    if (!q) return all;
    const starts = all.filter((j) => j.code.startsWith(q));
    return [...starts, ...all.filter((j) => !starts.includes(j) && j.label.toLowerCase().includes(q))];
  },
};

const picker = {
  open: false,

  show() {
    const items = jobList.matches($("job").value).slice(0, CONFIG.jobListShown);
    if (!jobList.all().length) return picker.hide();
    const list = $("job-list");
    list.replaceChildren(...(items.length ? items.map((j) => {
      const li = document.createElement("li");
      li.setAttribute("role", "option");
      const code = document.createElement("b");
      code.textContent = j.code;
      li.append(code, j.label.slice(j.code.length).trim());
      if (j.code === $("job").value.trim()) li.className = "on";
      // pointerdown, not click: the input's blur would close the list first
      li.addEventListener("pointerdown", (e) => { e.preventDefault(); picker.pick(j.code); });
      return li;
    }) : [Object.assign(document.createElement("li"),
                        { className: "none", textContent: "No job in your list matches. You can still send it." })]));
    list.hidden = false;
    picker.open = true;
    $("job").setAttribute("aria-expanded", "true");
  },

  hide() {
    $("job-list").hidden = true;
    picker.open = false;
    $("job").setAttribute("aria-expanded", "false");
  },

  pick(code) {
    $("job").value = code;
    $("job").classList.remove("bad");
    picker.hide();
    $("job").blur();
    render.jobs();
  },
};

/* ------------------------------------------------------------ render */

function clock(s) { return `${Math.floor(s / 60)}:${String(s % 60).padStart(2, "0")}`; }

function status(text, kind) {
  const el = $("status");
  el.textContent = text;
  el.className = `status ${kind || ""}`;
}

function setMode(mode) {
  const btn = $("rec");
  btn.classList.toggle("recording", mode === "recording");
  btn.disabled = mode === "review";
  btn.setAttribute("aria-label", mode === "recording" ? "Stop recording" : "Start recording");
  $("review").hidden = mode !== "review";
  $("rec-hint").textContent = { idle: "Tap to record", recording: "Recording. Tap to stop.",
                                review: "Play it back, then send or re-record." }[mode];
  if (mode === "idle") $("timer").textContent = "0:00";
}

const STATE_TEXT = {
  queued: ["Saved on phone, waiting to send", "warn"],
  sent: ["Sent to CASA", "ok"],
  "undo-queued": ["Undo saved, waiting to send", "warn"],
  "undo-sent": ["Undo sent", ""],
  cancelled: ["Deleted before sending", ""],
};

const render = {
  jobs() {
    const current = $("job").value.trim();
    $("recent-jobs").replaceChildren(...store.get("jobs", []).map((j) => {
      const b = document.createElement("button");
      b.type = "button";
      b.className = `chip${j === current ? " on" : ""}`;
      b.textContent = j;
      b.onclick = () => { $("job").value = j; $("job").classList.remove("bad"); render.jobs(); };
      return b;
    }));
    // the name under the box: confirms the pick, flags a number not in the list
    const known = jobList.find(current);
    const name = $("job-name");
    name.textContent = known ? known.label.slice(known.code.length).trim()
      : (JOB_RE.test(current) && jobList.all().length ? "Not in your Procore job list" : "");
    name.className = `job-name ${known ? "ok" : name.textContent ? "warn" : ""}`;
    $("job-toggle").hidden = !jobList.all().length;
  },
  recent() {
    const items = recent.list();
    $("recent-card").hidden = !items.length;
    $("recent").replaceChildren(...items.map((r) => {
      const li = document.createElement("li");
      const what = document.createElement("div");
      what.className = "what";
      const when = new Date(r.recordedAt).toLocaleString([], { month: "numeric", day: "numeric",
                                                              hour: "numeric", minute: "2-digit" });
      const [text, cls] = STATE_TEXT[r.state] || [r.state, ""];
      what.innerHTML = `<div class="job-no"></div><div class="when"></div><div class="state ${cls}"></div>`;
      what.children[0].textContent = `Job ${r.job}`;
      what.children[1].textContent = `${when} · ${clock(r.seconds || 0)}`;
      what.children[2].textContent = text;
      li.append(what);
      const canUndo = (r.state === "queued" ||
                       (r.state === "sent" && Date.now() - r.sentAt < CONFIG.undoHours * 3600e3));
      if (canUndo) {
        const b = document.createElement("button");
        b.type = "button";
        b.className = "undo";
        b.textContent = "Undo";
        b.onclick = () => {
          if (!b.classList.contains("confirm")) {     // 2 taps: no accidental undo
            b.classList.add("confirm");
            b.textContent = "Tap to confirm";
            setTimeout(() => { b.classList.remove("confirm"); b.textContent = "Undo"; }, 4000);
            return;
          }
          undo(r.key);
        };
        li.append(b);
      }
      return li;
    }));
  },
};

function showSignInPrompt() {
  $("signin-view").hidden = false;
}

/* ------------------------------------------------------------ start */

async function main() {
  if ("serviceWorker" in navigator) navigator.serviceWorker.register("sw.js").catch(() => {});

  try {
    await msalApp.initialize();
    const result = await msalApp.handleRedirectPromise();
    account = result?.account || msalApp.getAllAccounts()[0] || null;
  } catch (e) {
    status(`Sign-in problem: ${e.message}`, "err");
  }

  $("sign-in").onclick = signIn;
  $("sign-out").onclick = () => msalApp.logoutRedirect({ account });
  $("rec").onclick = () => (rec.media?.state === "recording" ? stopRecording() : startRecording());
  $("rerecord").onclick = () => { discardRecording(); startRecording(); };
  $("send").onclick = sendMemo;
  $("job").oninput = () => { $("job").classList.remove("bad"); render.jobs(); picker.show(); };
  $("job").onfocus = () => picker.show();
  $("job").onblur = () => setTimeout(picker.hide, 150);
  $("job-toggle").addEventListener("pointerdown", (e) => {
    e.preventDefault();                       // keep the keyboard down
    if (picker.open) picker.hide(); else picker.show();
  });
  document.addEventListener("pointerdown", (e) => {
    if (picker.open && !e.target.closest(".combo")) picker.hide();
  });

  if (!window.MediaRecorder) status("This browser cannot record audio. Update the phone's software.", "err");

  $("signin-view").hidden = !!account;
  $("record-view").hidden = !account;
  $("who").hidden = !account;
  if (account) $("who-name").textContent = account.name || account.username;
  $("job").value = store.get("jobs", [])[0] || "";
  render.jobs();
  render.recent();
  setMode("idle");

  const wake = async () => {
    await checkSignIn();
    flush();
    jobList.refresh().catch(() => {});      // the saved copy keeps working
  };
  window.addEventListener("online", wake);
  document.addEventListener("visibilitychange", () => { if (!document.hidden) wake(); });
  if (account) wake();
}

main();
