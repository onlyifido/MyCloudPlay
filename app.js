function validTrack(item) {
  return item && typeof item.id === "string" && typeof item.name === "string";
}
function validFolder(value) {
  return value && Array.isArray(value.value) && value.value.every(validTrack);
}
function readRecentTracks() {
  return readStoredJSON(
    "recent_tracks",
    [],
    (value) => Array.isArray(value) && value.every(validTrack),
  );
}
function escapeHTML(value) {
  return String(value ?? "").replace(
    /[&<>"']/g,
    (char) =>
      ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[
        char
      ],
  );
}
function dismissNotice() {
  document.getElementById("app_notice").classList.add("hidden");
}
function notify(message, actionLabel, action) {
  const notice = document.getElementById("app_notice");
  document.getElementById("notice_text").textContent = message;
  const button = document.getElementById("notice_action");
  button.classList.toggle("hidden", !action);
  button.textContent = actionLabel || "";
  button.onclick = action
    ? () => {
        dismissNotice();
        action();
      }
    : null;
  notice.classList.remove("hidden");
}
function friendlyError(error) {
  if (error.status === 401)
    return "Your session has expired. Connect OneDrive again.";
  if (error.status === 403)
    return "OneDrive did not allow this action. Check your account permissions.";
  if (error.status === 404)
    return "This file or folder is no longer available.";
  if (error.status === 409)
    return "A file or folder with this name already exists.";
  if (error.status === 429)
    return "OneDrive is busy. Please try again shortly.";
  if (error.name === "AbortError" || error.name === "TimeoutError")
    return "The request timed out. Please try again.";
  return "Could not connect to OneDrive. Check your connection and try again.";
}
async function graphRequest(path, options = {}) {
  const url = new URL(
    path.startsWith("https:")
      ? path
      : `https://graph.microsoft.com/v1.0/${path}`,
  );
  if (url.origin !== "https://graph.microsoft.com")
    throw new Error("Unexpected API destination");
  const token = await getToken();
  if (!token) {
    const error = new Error("Sign in required");
    error.status = 401;
    throw error;
  }
  const response = await fetch(url.href, {
    ...options,
    signal: options.signal || AbortSignal.timeout(30000),
    headers: {
      "Content-Type": "application/json",
      ...options.headers,
      Authorization: `Bearer ${token}`,
    },
  });
  if (!response.ok) {
    const error = new Error("OneDrive request failed");
    error.status = response.status;
    throw error;
  }
  return response.status === 204 ? null : response.json();
}
async function fetchAllChildren(folderId) {
  let next = `me/drive/items/${encodeURIComponent(folderId)}/children`;
  const value = [],
    visited = new Set();
  while (next) {
    if (visited.has(next)) throw new Error("Repeated OneDrive page");
    visited.add(next);
    const page = await graphRequest(next);
    if (!validFolder(page)) throw new Error("Invalid OneDrive response");
    value.push(...page.value);
    next = page["@odata.nextLink"];
  }
  return { value };
}
const audio = document.getElementById("audio_player"),
  progressBar = document.getElementById("progress_bar"),
  currentTimeEl = document.getElementById("current_time"),
  durationTimeEl = document.getElementById("duration_time"),
  fileInfoEl = document.getElementById("file_info"),
  extBadge = document.getElementById("file_ext_badge");
const msalConfig = {
  auth: {
    clientId: "8a7108a1-d00c-4f7c-bbed-f4a4df25b81a",
    authority: "https://login.microsoftonline.com/common",
    redirectUri: window.location.origin + "/",
  },
  cache: { cacheLocation: "localStorage" },
};
const msalInstance =
  typeof msal !== "undefined"
    ? new msal.PublicClientApplication(msalConfig)
    : null;
let accountId = "",
  currentFolderAudios = [],
  playingQueue = [],
  currentIndex = -1;
let isShuffle = storage.getItem("isShuffle") === "true";
let repeatMode = Number(storage.getItem("repeatMode") || 0);
if (![0, 1, 2].includes(repeatMode)) repeatMode = 0;
audio.loop = repeatMode === 2;
let currentFolderId = "root",
  pathHistory = [{ id: "root", name: "My Drive" }],
  wakeLock = null;
let folderCache = readStoredJSON(
  "folder_cache",
  {},
  (value) =>
    value &&
    !Array.isArray(value) &&
    typeof value === "object" &&
    Object.values(value).every(validFolder),
);
let urlCache = {};

function formatBytes(bytes, decimals = 2) {
  if (!+bytes) return "0 Bytes";
  const k = 1024,
    dm = decimals < 0 ? 0 : decimals,
    sizes = ["Bytes", "KB", "MB", "GB", "TB"],
    i = Math.floor(Math.log(bytes) / Math.log(k));
  return `${parseFloat((bytes / Math.pow(k, i)).toFixed(dm))} ${sizes[i]}`;
}
function splitName(name) {
  if (!name) return { base: "", ext: "" };
  const lastDot = name.lastIndexOf(".");
  if (lastDot === -1 || name.startsWith(".")) return { base: name, ext: "" };
  return {
    base: name.substring(0, lastDot),
    ext: name.substring(lastDot + 1).toUpperCase(),
  };
}
const modalFocus = new Map();
let activeModal = null;
function openModal(id) {
  const modal = document.getElementById(id);
  if (!modal) return;
  if (activeModal && activeModal !== modal) closeModal(activeModal.id);
  modalFocus.set(id, document.activeElement);
  activeModal = modal;
  modal.classList.remove("hidden", "opacity-0");
  modal.classList.add("flex");
  modal.firstElementChild.classList.remove("scale-95");
  document.getElementById("app_header").inert = true;
  document.getElementById("app_main").inert = true;
  const focus =
    modal.querySelector('input, textarea, button, [tabindex="0"]') || modal;
  focus.focus();
}
function closeModal(id) {
  const modal = document.getElementById(id);
  if (!modal || modal.classList.contains("hidden")) return;
  modal.classList.add("hidden", "opacity-0");
  modal.classList.remove("flex");
  if (activeModal === modal) {
    activeModal = null;
    document.getElementById("app_header").inert = false;
    document.getElementById("app_main").inert = false;
    const previous = modalFocus.get(id);
    if (previous?.isConnected && !previous.closest(".hidden")) previous.focus();
  }
}
window.addEventListener("click", (event) => {
  if (
    !event.target.closest(".menu-btn") &&
    !event.target.closest(".dropdown-content")
  ) {
    document
      .querySelectorAll(".dropdown-content")
      .forEach((menu) => menu.classList.add("hidden"));
    document
      .querySelectorAll('[aria-expanded="true"]')
      .forEach((button) => button.setAttribute("aria-expanded", "false"));
  }
});

const sampleTracks = [
  {
    id: "sample_1",
    name: "Sample Track 1 - Ambient.mp3",
    url: "https://www.soundhelix.com/examples/mp3/SoundHelix-Song-1.mp3",
    size: 6000000,
  },
  {
    id: "sample_2",
    name: "Sample Track 2 - Upbeat.mp3",
    url: "https://www.soundhelix.com/examples/mp3/SoundHelix-Song-2.mp3",
    size: 7000000,
  },
  {
    id: "sample_3",
    name: "Sample Track 3 - Electronic.mp3",
    url: "https://www.soundhelix.com/examples/mp3/SoundHelix-Song-3.mp3",
    size: 5000000,
  },
];

function loadSampleTracks() {
  const container = document.getElementById("file_list_container");
  currentFolderId = "samples";
  pathHistory = [{ id: "samples", name: "Sample tracks" }];
  renderBreadcrumbs();
  container.innerHTML = `<div class="sample-intro"><p>Try a sample, or <button onclick="signIn()">connect OneDrive</button> to play your own files.</p></div>`;
  currentFolderAudios = sampleTracks;
  document.getElementById("shuffle_folder_btn").style.display = "flex";

  sampleTracks.forEach((a, i) => {
    const { base } = splitName(a.name);
    const el = document.createElement("div");
    el.id = `track_row_${a.id}`;
    el.className =
      "flex items-center justify-between p-3 sm:p-3.5 mb-1.5 cursor-pointer rounded-2xl transition group w-full border border-transparent md:hover:bg-gray-50 md:dark:hover:bg-gray-800/50";
    const info = document.createElement("button");
    info.type = "button";
    info.className =
      "track-button flex items-center gap-3 sm:gap-4 flex-1 min-w-0 pr-4";
    info.onclick = () => {
      if (a.id === playingQueue[currentIndex]?.id && audio.src)
        togglePlayPause();
      else {
        playingQueue = [...currentFolderAudios];
        currentIndex = i;
        playTrack(i);
      }
    };
    info.innerHTML = `<div id="track_icon_${escapeHTML(a.id)}" class="w-10 h-10 sm:w-11 sm:h-11 shrink-0 rounded-xl flex items-center justify-center bg-gray-100 dark:bg-gray-900 text-gray-400 font-bold text-sm transition">${i + 1}</div><p id="track_text_${escapeHTML(a.id)}" class="text-sm sm:text-[15px] font-medium truncate text-gray-800 dark:text-gray-200 transition">${escapeHTML(base)}</p>`;

    el.append(info);
    container.appendChild(el);
  });
  updateActiveTrackUI();
}

async function initializeAccount() {
  try {
    const response = msalInstance
      ? await msalInstance.handleRedirectPromise()
      : null;
    const account = response?.account || msalInstance?.getAllAccounts()[0];
    if (account) {
      accountId = account.homeAccountId;
      document.getElementById("login_button").style.display = "none";
      document.getElementById("logout_button").style.display = "inline-block";
      document.getElementById("action_buttons").style.display = "flex";
      await fetchMyFolders("root", "My Drive");
    } else {
      loadSampleTracks();
    }
  } catch {
    loadSampleTracks();
    notify(
      "Sign-in could not be completed. You can still listen to samples.",
      "Try again",
      signIn,
    );
  }
  renderRecentUI();
}
async function getToken() {
  if (!msalInstance) {
    notify(
      "OneDrive sign-in is unavailable. Check your connection and reload.",
    );
    return null;
  }
  const account =
    msalInstance.getAccountByHomeId(accountId) ||
    msalInstance.getAllAccounts()[0];
  if (!account) {
    notify("Connect OneDrive to access your files.", "Connect", signIn);
    return null;
  }
  try {
    return (
      await msalInstance.acquireTokenSilent({
        scopes: ["Files.ReadWrite.All"],
        account,
      })
    ).accessToken;
  } catch (error) {
    if (error instanceof msal.InteractionRequiredAuthError) {
      notify(
        "Your session needs attention. Connect OneDrive again.",
        "Connect",
        signIn,
      );
    } else
      notify(
        "Could not sign in to OneDrive. Check your connection.",
        "Try again",
        signIn,
      );
    return null;
  }
}
async function signIn() {
  if (!msalInstance) {
    notify(
      "OneDrive sign-in could not load. Check your connection and reload.",
    );
    return;
  }
  try {
    await msalInstance.loginRedirect({ scopes: ["Files.ReadWrite.All"] });
  } catch {
    notify("Sign-in could not be started. Please try again.");
  }
}
async function signOut() {
  ++playbackGeneration;
  playbackController?.abort();
  audio.pause();
  audio.removeAttribute("src");
  audio.load();
  storage.removeItem("folder_cache");
  storage.removeItem("recent_tracks");
  folderCache = {};
  urlCache = {};
  try {
    await msalInstance?.logoutRedirect({
      account: msalInstance.getAccountByHomeId(accountId),
    });
  } catch {
    notify("Sign-out could not be completed. Please try again.");
  }
}

let playbackGeneration = 0;
let playbackController = null;
let playbackRetried = false;
let isLoadingTrack = false;
let playbackFinished = false;
function playbackFailed() {
  isLoadingTrack = false;
  document.getElementById("player_card").classList.remove("shimmer-active");
  setStatus("PLAYBACK ERROR", "err");
  notify(
    "This track could not be played. Check your connection or try another file.",
    "Retry",
    () => playTrack(currentIndex),
  );
}
async function playTrack(index, autoPlay = true, retry = false) {
  if (index < 0 || index >= playingQueue.length) return;
  const generation = ++playbackGeneration;
  playbackController?.abort();
  playbackController = new AbortController();
  playbackRetried = retry;
  const track = playingQueue[index],
    { base, ext } = splitName(track.name);
  isLoadingTrack = true;
  playbackFinished = false;
  audio.pause();
  audio.removeAttribute("src");
  audio.load();
  currentIndex = index;
  audio.loop = repeatMode === 2;
  progressBar.value = 0;
  progressBar.disabled = true;
  currentTimeEl.textContent = "0:00";
  durationTimeEl.textContent = "0:00";
  updateActiveTrackUI();
  updateMarqueeText(base);
  extBadge.textContent = ext;
  extBadge.classList.toggle("hidden", !ext);
  dismissNotice();
  setStatus("BUFFERING", "load");
  document.getElementById("player_card").classList.add("shimmer-active");
  const setupMediaSession = () => {
    if ("mediaSession" in navigator) {
      navigator.mediaSession.metadata = new MediaMetadata({
        title: base,
        artist: "MyCloudPlay",
        artwork: [
          {
            src: "data:image/svg+xml;charset=utf-8,%3Csvg xmlns=%22http://www.w3.org/2000/svg%22 viewBox=%220 0 512 512%22%3E%3Cdefs%3E%3ClinearGradient id=%22grad%22 x1=%220%25%22 y1=%22100%25%22 x2=%22100%25%22 y2=%220%25%22%3E%3Cstop offset=%220%25%22 style=%22stop-color:%230078d4;stop-opacity:1%22 /%3E%3Cstop offset=%22100%25%22 style=%22stop-color:%2360a5fa;stop-opacity:1%22 /%3E%3C/linearGradient%3E%3C/defs%3E%3Crect width=%22512%22 height=%22512%22 fill=%22url(%23grad)%22 /%3E%3Cpath d=%22M200 370V130l210-40v230%22 fill=%22none%22 stroke=%22%23ffffff%22 stroke-width=%2232%22 stroke-linecap=%22round%22 stroke-linejoin=%22round%22/%3E%3Ccircle cx=%22150%22 cy=%22370%22 r=%2250%22 fill=%22none%22 stroke=%22%23ffffff%22 stroke-width=%2232%22/%3E%3Ccircle cx=%22360%22 cy=%22330%22 r=%2250%22 fill=%22none%22 stroke=%22%23ffffff%22 stroke-width=%2232%22/%3E%3C/svg%3E",
            sizes: "512x512",
            type: "image/svg+xml",
          },
        ],
      });

      navigator.mediaSession.setActionHandler("play", () => {
        if (audio.paused) togglePlayPause();
      });
      navigator.mediaSession.setActionHandler("pause", () => {
        audio.pause();
      });
      navigator.mediaSession.setActionHandler("previoustrack", () => {
        playPrevTrack();
      });
      navigator.mediaSession.setActionHandler("nexttrack", () => {
        playNextTrack(true);
      });
    }
  };

  try {
    let streamUrl = track.url || urlCache[track.id];
    if (!streamUrl) {
      const data = await graphRequest(
        `me/drive/items/${encodeURIComponent(track.id)}`,
        { signal: playbackController.signal },
      );
      streamUrl = data["@microsoft.graph.downloadUrl"];
    }
    if (generation !== playbackGeneration) return;
    if (
      !streamUrl ||
      !["https:", "blob:"].includes(new URL(streamUrl, location.href).protocol)
    )
      throw new Error("Invalid stream URL");
    if (!track.url) urlCache[track.id] = streamUrl;
    audio.src = streamUrl;
    audio.load();
    updateActiveTrackUI();
    setupMediaSession();
    if (autoPlay) await audio.play();
    if (generation !== playbackGeneration) return;
    isLoadingTrack = false;
    saveRecentTrack(track);
    document.getElementById("player_card").classList.remove("shimmer-active");
    if (!autoPlay) setStatus("READY", "standby");
    document
      .getElementById(`track_row_${track.id}`)
      ?.scrollIntoView({ behavior: "smooth", block: "nearest" });
  } catch (error) {
    if (generation !== playbackGeneration || error.name === "AbortError")
      return;
    if (error.name === "NotAllowedError") {
      isLoadingTrack = false;
      document.getElementById("player_card").classList.remove("shimmer-active");
      setStatus("READY", "standby");
      notify("Press play to start this track.");
    } else if (error.status === 401 || error.status === 403) {
      isLoadingTrack = false;
      document.getElementById("player_card").classList.remove("shimmer-active");
      setStatus("SIGN IN REQUIRED", "err");
      notify(friendlyError(error), "Connect", signIn);
    } else if (!playbackRetried && !track.url) {
      delete urlCache[track.id];
      return playTrack(index, autoPlay, true);
    } else playbackFailed();
  }
}
audio.onerror = () => {
  if (!audio.getAttribute("src") || currentIndex < 0) return;
  if (!playbackRetried && !playingQueue[currentIndex]?.url) {
    delete urlCache[playingQueue[currentIndex].id];
    playTrack(currentIndex, true, true);
  } else playbackFailed();
};

function showTrackInfo() {
  if (currentIndex < 0 || !playingQueue[currentIndex]) return;
  const track = playingQueue[currentIndex];
  const content = document.getElementById("track_info_content");

  const size = track.size ? formatBytes(track.size) : "Unknown";
  const ext = splitName(track.name).ext || "Unknown";
  const audioData = track.audio || {};

  const bitrate = audioData.bitrate
    ? `${Math.round(audioData.bitrate / 1000)} kbps`
    : "Unknown";
  const title = audioData.title || splitName(track.name).base;
  const artist = audioData.artist || "Unknown Artist";
  const album = audioData.album || "Unknown Album";

  content.innerHTML = `
                <div class="flex flex-col py-2 border-b border-gray-150 dark:border-gray-700/50">
                    <span class="text-[10px] text-accent  font-bold uppercase tracking-widest mb-1">File Name</span>
                    <span class="text-gray-900 dark:text-gray-100 font-medium break-all leading-snug">${escapeHTML(track.name)}</span>
                </div>
                <div class="flex justify-between items-center py-2.5 border-b border-gray-150 dark:border-gray-700/50">
                    <span class="text-[10px] text-gray-500 font-bold uppercase tracking-widest">Title</span>
                    <span class="text-gray-900 dark:text-gray-100 font-bold text-right truncate max-w-[65%]">${escapeHTML(title)}</span>
                </div>
                <div class="flex justify-between items-center py-2.5 border-b border-gray-150 dark:border-gray-700/50">
                    <span class="text-[10px] text-gray-500 font-bold uppercase tracking-widest">Artist</span>
                    <span class="text-gray-900 dark:text-gray-100 font-bold text-right truncate max-w-[65%]">${escapeHTML(artist)}</span>
                </div>
                <div class="flex justify-between items-center py-2.5 border-b border-gray-150 dark:border-gray-700/50">
                    <span class="text-[10px] text-gray-500 font-bold uppercase tracking-widest">Album</span>
                    <span class="text-gray-900 dark:text-gray-100 font-bold text-right truncate max-w-[65%]">${escapeHTML(album)}</span>
                </div>
                <div class="flex justify-between items-center py-2.5 border-b border-gray-150 dark:border-gray-700/50">
                    <span class="text-[10px] text-gray-500 font-bold uppercase tracking-widest">Format / Size</span>
                    <span class="text-gray-900 dark:text-gray-100 font-bold">${escapeHTML(ext)} / ${escapeHTML(size)}</span>
                </div>
                <div class="flex justify-between items-center py-2.5">
                    <span class="text-[10px] text-gray-500 font-bold uppercase tracking-widest">Audio Bitrate</span>
                    <span class="text-gray-900 dark:text-gray-100 font-bold">${escapeHTML(bitrate)}</span>
                </div>
            `;
  openModal("track_info_modal");
}

function createFolder() {
  document.getElementById("main_action_drop").classList.add("hidden");
  const input = document.getElementById("create_folder_input");
  input.value = "";
  openModal("create_folder_modal");
  input.focus();
  const button = document.getElementById("create_folder_confirm_btn");
  button.onclick = async () => {
    const name = input.value.trim();
    if (!name) {
      input.focus();
      return;
    }
    const folderId = currentFolderId;
    button.disabled = true;
    try {
      await graphRequest(
        `me/drive/items/${encodeURIComponent(folderId)}/children`,
        { method: "POST", body: JSON.stringify({ name, folder: {} }) },
      );
      delete folderCache[folderId];
      closeModal("create_folder_modal");
      if (currentFolderId === folderId)
        await fetchMyFolders(folderId, pathHistory.at(-1).name, false);
    } catch (error) {
      notify(friendlyError(error));
    } finally {
      button.disabled = false;
    }
  };
}

let folderGeneration = 0;
async function fetchMyFolders(folderId, folderName, isNewNavigation = true) {
  if (folderId === "samples") {
    loadSampleTracks();
    return;
  }
  const generation = ++folderGeneration;
  const container = document.getElementById("file_list_container");
  currentFolderId = folderId;
  if (isNewNavigation && folderId !== "root") {
    if (pathHistory.at(-1)?.id !== folderId)
      pathHistory.push({ id: folderId, name: folderName });
  } else if (folderId === "root")
    pathHistory = [{ id: "root", name: "My Drive" }];
  renderBreadcrumbs();
  const cached = folderCache[folderId];
  if (cached) renderItems(cached);
  else
    container.innerHTML =
      '<div class="empty-state" role="status">Loading files…</div>';
  try {
    const data = await fetchAllChildren(folderId);
    folderCache[folderId] = data;
    storage.setItem("folder_cache", JSON.stringify(folderCache));
    if (generation === folderGeneration && currentFolderId === folderId)
      renderItems(data);
  } catch (error) {
    if (generation !== folderGeneration || currentFolderId !== folderId) return;
    if (!cached) {
      container.replaceChildren();
      const state = document.createElement("div");
      state.className = "empty-state";
      state.textContent = friendlyError(error);
      const retry = document.createElement("button");
      retry.textContent = "Try again";
      retry.onclick = () => fetchMyFolders(folderId, folderName, false);
      state.append(retry);
      container.append(state);
    } else
      notify(
        "Showing saved files. Could not refresh this folder.",
        "Retry",
        () => fetchMyFolders(folderId, folderName, false),
      );
  }

  function renderItems(data) {
    container.innerHTML = "";
    const items = [...data.value].sort((a, b) => a.name.localeCompare(b.name));
    const folders = items.filter((i) => i.folder),
      audios = items.filter(
        (i) => i.file && i.file.mimeType?.includes("audio"),
      );
    currentFolderAudios = audios.map((track) => ({
      ...track,
      folderId,
      folderName: pathHistory.at(-1).name,
    }));
    document.getElementById("shuffle_folder_btn").style.display = audios.length
      ? "flex"
      : "none";

    folders.forEach((f) => {
      const el = document.createElement("div");
      el.className =
        "flex items-center justify-between p-3 sm:p-4 mb-2 border border-gray-100 dark:border-gray-800/80 rounded-[1.2rem] bg-gray-50/50 dark:bg-gray-900/30 group cursor-pointer w-full transition md:hover:bg-gray-100 md:dark:hover:bg-gray-800/50";
      const info = document.createElement("button");
      info.type = "button";
      info.className =
        "track-button flex items-center gap-3 sm:gap-4 flex-1 min-w-0 pr-2";
      info.onclick = () => fetchMyFolders(f.id, f.name);
      info.innerHTML = `<div class="w-10 h-10 sm:w-12 sm:h-12 shrink-0 rounded-xl bg-accent-soft  flex items-center justify-center text-accent"><svg class="w-5 h-5 sm:w-6 sm:h-6" fill="currentColor" viewBox="0 0 24 24"><path d="M10 4H4c-1.1 0-1.99.9-1.99 2L2 18c0 1.1.9 2 2 2h16c1.1 0 2-.9 2-2V8c0-1.1-.9-2-2-2h-8l-2-2z"/></svg></div><span id="folder_text_${escapeHTML(f.id)}" class="text-sm sm:text-base font-semibold truncate text-gray-800 dark:text-gray-200">${escapeHTML(f.name)}</span>`;
      el.append(info, createActionBtn(f.id, f.name, true));
      container.appendChild(el);
    });
    audios.forEach((a, i) => {
      const { base } = splitName(a.name);
      const el = document.createElement("div");
      el.id = `track_row_${a.id}`;
      el.className =
        "flex items-center justify-between p-3 sm:p-3.5 mb-1.5 cursor-pointer rounded-2xl transition group w-full border border-transparent md:hover:bg-gray-50 md:dark:hover:bg-gray-800/50";
      const info = document.createElement("button");
      info.type = "button";
      info.className =
        "track-button flex items-center gap-3 sm:gap-4 flex-1 min-w-0 pr-4";
      info.onclick = () => {
        if (a.id === playingQueue[currentIndex]?.id && audio.src)
          togglePlayPause();
        else {
          playingQueue = [...currentFolderAudios];
          currentIndex = i;
          playTrack(i);
        }
      };
      info.innerHTML = `<div id="track_icon_${escapeHTML(a.id)}" class="w-10 h-10 sm:w-11 sm:h-11 shrink-0 rounded-xl flex items-center justify-center bg-gray-100 dark:bg-gray-900 text-gray-400 font-bold text-sm transition">${i + 1}</div><p id="track_text_${escapeHTML(a.id)}" class="text-sm sm:text-[15px] font-medium truncate text-gray-800 dark:text-gray-200 transition">${escapeHTML(base)}</p>`;
      el.append(info, createActionBtn(a.id, a.name, false));
      container.appendChild(el);
    });
    if (!folders.length && !audios.length)
      container.innerHTML =
        '<div class="empty-state">No audio files or folders here.</div>';
    updateActiveTrackUI();
  }
}
function createActionBtn(id, name, isFolder) {
  const wrap = document.createElement("div");
  wrap.className = "relative flex items-center shrink-0";
  const toggle = document.createElement("button");
  toggle.className = "menu-btn w-10 h-10 text-gray-500 rounded-full";
  toggle.textContent = "⋮";
  toggle.setAttribute("aria-label", `Actions for ${name}`);
  toggle.setAttribute("aria-expanded", "false");
  toggle.setAttribute("aria-controls", `drop-${id}`);
  toggle.onclick = (event) => toggleMenu(event, `drop-${id}`);
  const menu = document.createElement("div");
  menu.id = `drop-${id}`;
  menu.className =
    "dropdown-content absolute right-0 top-10 w-32 rounded-xl border py-1.5 z-50 hidden";
  for (const [label, action] of [
    ["Rename", () => requestRename(id, name, isFolder)],
    ["Delete", () => deleteItem(id, name, isFolder)],
  ]) {
    const button = document.createElement("button");
    button.textContent = label;
    button.className = `block px-4 py-2 text-sm ${label === "Delete" ? "text-red-600" : "text-gray-500"}`;
    button.onclick = (event) => {
      event.stopPropagation();
      menu.classList.add("hidden");
      toggle.setAttribute("aria-expanded", "false");
      action();
    };
    menu.append(button);
  }
  wrap.append(toggle, menu);
  return wrap;
}
function requestRename(id, fullName, isFolder) {
  const { base, ext } = splitName(fullName),
    textarea = document.getElementById("rename_textarea");
  textarea.value = isFolder ? fullName : base;
  openModal("rename_modal");
  textarea.focus();
  const button = document.getElementById("rename_confirm_btn");
  button.onclick = async () => {
    const name = textarea.value.trim();
    if (!name) {
      textarea.focus();
      return;
    }
    if (name === (isFolder ? fullName : base)) {
      closeModal("rename_modal");
      return;
    }
    const finalName = isFolder
      ? name
      : `${name}${ext ? fullName.slice(fullName.lastIndexOf(".")) : ""}`;
    const folderId = currentFolderId;
    button.disabled = true;
    try {
      await graphRequest(`me/drive/items/${encodeURIComponent(id)}`, {
        method: "PATCH",
        body: JSON.stringify({ name: finalName }),
      });
      delete folderCache[folderId];
      playingQueue.forEach((track) => {
        if (track.id === id) track.name = finalName;
      });
      if (playingQueue[currentIndex]?.id === id)
        updateMarqueeText(splitName(finalName).base);
      const recent = readRecentTracks().map((track) =>
        track.id === id ? { ...track, name: finalName } : track,
      );
      storage.setItem("recent_tracks", JSON.stringify(recent));
      renderRecentUI();
      closeModal("rename_modal");
      if (currentFolderId === folderId)
        await fetchMyFolders(folderId, pathHistory.at(-1).name, false);
    } catch (error) {
      notify(friendlyError(error));
    } finally {
      button.disabled = false;
    }
  };
}

function renderBreadcrumbs() {
  const container = document.getElementById("breadcrumb_container");
  container.innerHTML = "";
  if (pathHistory.length > 1) {
    const btn = document.createElement("button");
    btn.setAttribute("aria-label", "Parent folder");
    btn.className =
      "mr-2 p-1 text-gray-400 md:hover:text-accent  transition active:scale-90 flex items-center justify-center rounded-full md:hover:bg-gray-100 md:dark:hover:bg-gray-800 shrink-0";
    btn.innerHTML =
      '<svg class="w-5 h-5 sm:w-6 sm:h-6" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M15 19l-7-7 7-7"></path></svg>';
    btn.onclick = () => {
      if (pathHistory.length <= 1) return;
      pathHistory.pop();
      fetchMyFolders(
        pathHistory[pathHistory.length - 1].id,
        pathHistory[pathHistory.length - 1].name,
        false,
      );
    };
    container.appendChild(btn);
  }
  if (pathHistory.length <= 3) {
    pathHistory.forEach((s, i) => {
      const span = document.createElement(
        i === pathHistory.length - 1 ? "span" : "button",
      );
      span.innerText = s.name;
      span.className =
        i === pathHistory.length - 1
          ? "font-semibold text-gray-900 dark:text-white truncate max-w-[150px] sm:max-w-[250px] shrink-0"
          : "breadcrumb-item transition md:hover:text-accent truncate max-w-[100px] sm:max-w-[150px] shrink-0";
      if (i !== pathHistory.length - 1)
        span.onclick = () => {
          pathHistory = pathHistory.slice(0, i + 1);
          fetchMyFolders(s.id, s.name, false);
        };
      container.appendChild(span);
      if (i < pathHistory.length - 1) {
        const sep = document.createElement("span");
        sep.className = "breadcrumb-separator shrink-0";
        sep.innerText = "/";
        container.appendChild(sep);
      }
    });
  } else {
    const rootSpan = document.createElement("button");
    rootSpan.innerText = pathHistory[0].name;
    rootSpan.className =
      "breadcrumb-item transition md:hover:text-accent truncate max-w-[80px] shrink-0";
    rootSpan.onclick = () => {
      pathHistory = pathHistory.slice(0, 1);
      fetchMyFolders(pathHistory[0].id, pathHistory[0].name, false);
    };
    container.appendChild(rootSpan);
    const sep1 = document.createElement("span");
    sep1.className = "breadcrumb-separator shrink-0";
    sep1.innerText = "/";
    container.appendChild(sep1);
    const wrap = document.createElement("div");
    wrap.className = "relative flex items-center shrink-0";
    const ell = document.createElement("button");
    ell.setAttribute("aria-label", "Parent folders");
    ell.innerText = "...";
    ell.className =
      "breadcrumb-item font-bold px-1 tracking-widest cursor-pointer md:hover:text-accent menu-btn";
    ell.onclick = (e) => toggleMenu(e, "drop-breadcrumb");
    wrap.appendChild(ell);
    const drop = document.createElement("div");
    drop.id = "drop-breadcrumb";
    drop.className =
      "dropdown-content absolute left-0 top-8 w-48 bg-white dark:bg-gray-800 rounded-xl shadow-xl border border-gray-100 dark:border-gray-700 py-1.5 z-50 hidden max-h-60 overflow-y-auto";
    for (let i = 1; i < pathHistory.length - 1; i++) {
      const a = document.createElement("button");
      a.className =
        "block px-4 py-2 text-sm text-gray-700 dark:text-gray-300 md:hover:bg-gray-100 md:dark:hover:bg-gray-700 cursor-pointer truncate";
      a.innerText = pathHistory[i].name;
      a.onclick = (e) => {
        e.stopPropagation();
        pathHistory = pathHistory.slice(0, i + 1);
        fetchMyFolders(pathHistory[i].id, pathHistory[i].name, false);
        drop.classList.add("hidden");
      };
      drop.appendChild(a);
    }
    wrap.appendChild(drop);
    container.appendChild(wrap);
    const sep2 = document.createElement("span");
    sep2.className = "breadcrumb-separator shrink-0";
    sep2.innerText = "/";
    container.appendChild(sep2);
    const currSpan = document.createElement("span");
    currSpan.innerText = pathHistory[pathHistory.length - 1].name;
    currSpan.className =
      "font-semibold text-gray-900 dark:text-white truncate max-w-[120px] sm:max-w-[200px] shrink-0";
    container.appendChild(currSpan);
  }
}
function saveRecentTrack(t) {
  let recent = readRecentTracks();
  recent = recent.filter((x) => x.id !== t.id);

  const fId = t.folderId || currentFolderId;
  const fName = t.folderName || pathHistory[pathHistory.length - 1].name;

  recent.unshift({
    id: t.id,
    name: t.name,
    folderId: fId,
    folderName: fName,
    url: t.url,
  });

  if (recent.length > 8) recent = recent.slice(0, 8);

  storage.setItem("recent_tracks", JSON.stringify(recent));
  renderRecentUI();
}

function renderRecentUI() {
  const container = document.getElementById("recent_history_content");
  const chipsContainer = document.getElementById("recent_tracks_list");
  const wrapper = document.getElementById("recent_tracks_wrapper");
  const recent = readRecentTracks();

  container.innerHTML = "";
  if (chipsContainer) chipsContainer.innerHTML = "";

  if (!recent.length) {
    container.innerHTML =
      '<div class="text-center py-10 text-gray-400 text-sm">No recent history</div>';
    if (wrapper) wrapper.classList.add("hidden");
    return;
  }

  if (wrapper) wrapper.classList.remove("hidden");

  recent.forEach((t) => {
    const { base } = splitName(t.name);

    const item = document.createElement("button");
    item.type = "button";
    item.className =
      "w-full text-left p-3 rounded-xl bg-gray-50/50 dark:bg-gray-800/30 border border-gray-100 dark:border-gray-700/30 cursor-pointer transition md:hover:bg-accent-soft ";

    item.onclick = () => playRecentTrack(t);

    item.innerHTML = `<p class="text-sm font-medium text-gray-800 dark:text-gray-200 truncate">${escapeHTML(base)}</p><p class="text-[10px] text-gray-400 mt-0.5">${escapeHTML(t.folderName || "Unknown Folder")}</p>`;
    container.appendChild(item);

    if (chipsContainer && chipsContainer.children.length < 3) {
      const chip = document.createElement("button");
      chip.className =
        "px-4 py-2 rounded-full bg-white dark:bg-gray-800 border border-gray-200 dark:border-gray-700 text-xs font-medium text-gray-600 dark:text-gray-300 whitespace-nowrap transition md:hover:border-accent truncate";

      chip.onclick = () => playRecentTrack(t);

      chip.innerText = base;
      chipsContainer.appendChild(chip);
    }
  });
}

async function playRecentTrack(track) {
  closeModal("recent_history_modal");
  if (sampleTracks.some((sample) => sample.id === track.id)) {
    if (!accountId) loadSampleTracks();
    playingQueue = [...sampleTracks];
    return playTrack(
      playingQueue.findIndex((sample) => sample.id === track.id),
    );
  }
  const folderId = track.folderId || "root";
  if (currentFolderId !== folderId) {
    pathHistory = [{ id: "root", name: "My Drive" }];
    if (folderId !== "root")
      pathHistory.push({ id: folderId, name: track.folderName || "Folder" });
    fetchMyFolders(folderId, track.folderName, false);
  }
  const restoreQueue = (data) => {
    const queue = data.value
      .filter((item) => item.file?.mimeType?.includes("audio"))
      .sort((a, b) => a.name.localeCompare(b.name))
      .map((item) => ({ ...item, folderId, folderName: track.folderName }));
    const index = queue.findIndex((item) => item.id === track.id);
    if (index < 0) return false;
    playingQueue = queue;
    currentIndex = index;
    return true;
  };
  if (folderCache[folderId] && restoreQueue(folderCache[folderId]))
    return playTrack(currentIndex);
  playingQueue = [track];
  currentIndex = 0;
  playTrack(0);
  const generation = playbackGeneration;
  try {
    const data = await fetchAllChildren(folderId);
    folderCache[folderId] = data;
    storage.setItem("folder_cache", JSON.stringify(folderCache));
    if (
      generation === playbackGeneration &&
      playingQueue[currentIndex]?.id === track.id
    )
      restoreQueue(data);
  } catch {
    /* The selected track can still play when its folder cannot be refreshed. */
  }
}

function openRecentHistoryModal() {
  const content = document.getElementById("recent_history_content");
  const recent = readRecentTracks();
  content.innerHTML = "";

  if (!recent.length) {
    content.innerHTML =
      '<div class="text-sm text-gray-400 text-center py-4">No recent tracks.</div>';
  } else {
    recent.forEach((t) => {
      const { base } = splitName(t.name);
      const item = document.createElement("button");
      item.type = "button";
      item.className =
        "w-full text-left flex items-center justify-between p-3 rounded-xl bg-gray-50/50 dark:bg-gray-800/30 border border-gray-100 dark:border-gray-700/50 cursor-pointer active:scale-95 transition md:hover:bg-accent-soft  md:hover:border-accent-border  group";
      item.innerHTML = `<span class="text-sm font-medium text-gray-800 dark:text-gray-200 truncate group-hover:text-accent  transition">${escapeHTML(base)}</span>
                                      <span class="text-[10px] text-gray-400 shrink-0 ml-2 border border-gray-150 dark:border-gray-700 px-2 py-0.5 rounded-md truncate max-w-[80px]">${escapeHTML(t.folderName)}</span>`;

      item.onclick = () => playRecentTrack(t);

      content.appendChild(item);
    });
  }
  openModal("recent_history_modal");
}

function updateMarqueeText(text) {
  fileInfoEl.innerText = text;
  setTimeout(() => {
    const container = fileInfoEl.parentElement;
    if (fileInfoEl.scrollWidth > container.clientWidth) {
      fileInfoEl.classList.remove("short");
      fileInfoEl.style.animationDuration = `${fileInfoEl.scrollWidth / 20}s`;
    } else {
      fileInfoEl.classList.add("short");
      fileInfoEl.style.animationDuration = "";
    }
  }, 50);
}
function setStatus(t, type) {
  const el = document.getElementById("status_info"),
    dot = document.getElementById("status_dot");
  el.innerText = t;
  el.className = `text-[10px] sm:text-xs font-bold tracking-[0.2em] uppercase ${type === "play" ? "text-green-500" : type === "load" ? "text-accent " : "text-gray-500 dark:text-gray-400"}`;
  dot.className = `w-2 h-2 sm:w-2.5 sm:h-2.5 rounded-full ${type === "play" ? "bg-green-500" : type === "load" ? "bg-accent  animate-pulse" : "bg-gray-300 dark:bg-gray-600"}`;
}
async function togglePlayPause() {
  if (!playingQueue.length) {
    if (currentFolderAudios.length) {
      playingQueue = [...currentFolderAudios];
      return playTrack(0);
    }
    return;
  }
  if (!audio.paused) {
    audio.pause();
    return;
  }
  if (!audio.getAttribute("src") || audio.error) return playTrack(currentIndex);
  try {
    await audio.play();
  } catch (error) {
    if (error.name !== "AbortError") playbackFailed();
  }
}

function toggleRepeat() {
  repeatMode = (repeatMode + 1) % 3;
  audio.loop = repeatMode === 2;
  storage.setItem("repeatMode", repeatMode);
  updateRepeatBtnUI();
}

function toggleShuffle() {
  isShuffle = !isShuffle;
  storage.setItem("isShuffle", isShuffle);
  updateShuffleBtnUI();
}
function toggleMenu(event, id) {
  event.stopPropagation();
  const target = document.getElementById(id);
  const open = target.classList.contains("hidden");
  document
    .querySelectorAll(".dropdown-content")
    .forEach((menu) => menu.classList.add("hidden"));
  document
    .querySelectorAll("[aria-expanded]")
    .forEach((button) => button.setAttribute("aria-expanded", "false"));
  target.classList.toggle("hidden", !open);
  event.currentTarget.setAttribute("aria-expanded", String(open));
}

function updateActiveTrackUI() {
  const currentId = playingQueue[currentIndex]?.id;
  currentFolderAudios.forEach((a, i) => {
    const row = document.getElementById(`track_row_${a.id}`),
      icon = document.getElementById(`track_icon_${a.id}`),
      text = document.getElementById(`track_text_${a.id}`);
    if (!row) return;
    row.dataset.playing = String(
      a.id === currentId && Boolean(audio.getAttribute("src")),
    );
    row
      .querySelector(".track-button")
      ?.setAttribute(
        "aria-label",
        `${a.id === currentId && !audio.paused ? "Pause" : "Play"} ${a.name}`,
      );
    if (a.id === currentId && audio.src) {
      row.className =
        "flex items-center justify-between p-3 sm:p-3.5 mb-1.5 cursor-pointer rounded-2xl bg-accent-soft  border border-accent-border  shadow-sm transition group w-full";
      icon.className =
        "w-10 h-10 sm:w-11 sm:h-11 shrink-0 rounded-xl flex items-center justify-center transition text-accent ";
      text.className =
        "text-sm sm:text-[15px] leading-tight truncate text-accent  font-bold transition";
      icon.innerHTML = audio.paused
        ? '<div class="playing-eq eq-paused"><span class="eq-bar eq-bar-1"></span><span class="eq-bar eq-bar-2"></span><span class="eq-bar eq-bar-3"></span></div>'
        : '<div class="playing-eq"><span class="eq-bar eq-bar-1"></span><span class="eq-bar eq-bar-2"></span><span class="eq-bar eq-bar-3"></span></div>';
    } else {
      row.className =
        "flex items-center justify-between p-3 sm:p-3.5 mb-1.5 cursor-pointer rounded-2xl border border-transparent md:hover:bg-gray-50 md:dark:hover:bg-gray-800/50 transition group w-full";
      icon.className =
        "w-10 h-10 sm:w-11 sm:h-11 shrink-0 rounded-xl flex items-center justify-center bg-gray-100 dark:bg-gray-900 text-gray-400 font-bold text-sm transition";
      text.className =
        "text-sm sm:text-[15px] font-medium truncate text-gray-800 dark:text-gray-200 transition";
      icon.innerHTML = i + 1;
    }
  });
}
function updateRepeatBtnUI() {
  const btn = document.getElementById("repeat_btn");
  btn.dataset.active = String(repeatMode > 0);
  btn.setAttribute(
    "aria-label",
    `Repeat: ${["off", "all", "one"][repeatMode]}`,
  );
  btn.title = btn.getAttribute("aria-label");
  if (repeatMode > 0) {
    btn.classList.add("text-accent");
    btn.classList.remove("text-gray-400", "dark:text-gray-500");
  } else {
    btn.classList.add("text-gray-400", "dark:text-gray-500");
    btn.classList.remove("text-accent");
  }
  const baseClass = "w-5 h-5 sm:w-6 sm:h-6";
  if (repeatMode === 2)
    btn.innerHTML = `<svg class="${baseClass} relative" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15"></path><text x="12" y="15" font-family="sans-serif" font-size="9" font-weight="bold" fill="currentColor" text-anchor="middle" stroke="none">1</text></svg>`;
  else
    btn.innerHTML = `<svg class="${baseClass}" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15"></path></svg>`;
}
function updateShuffleBtnUI() {
  const btn = document.getElementById("shuffle_btn");
  btn.setAttribute("aria-pressed", String(isShuffle));
  if (isShuffle) {
    btn.classList.add("text-accent");
    btn.classList.remove("text-gray-400", "dark:text-gray-500");
  } else {
    btn.classList.add("text-gray-400", "dark:text-gray-500");
    btn.classList.remove("text-accent");
  }
}
function playRandomFromFolder() {
  if (!currentFolderAudios.length) return;
  isShuffle = true;
  storage.setItem("isShuffle", "true");
  updateShuffleBtnUI();
  playingQueue = [...currentFolderAudios];
  currentIndex = Math.floor(Math.random() * playingQueue.length);
  playTrack(currentIndex);
}
function playNextTrack(isManual = false) {
  if (!playingQueue.length) return;
  if (repeatMode === 2 && !isManual) playTrack(currentIndex);
  else if (isShuffle) {
    let n = currentIndex;
    if (playingQueue.length > 1)
      while (n === currentIndex)
        n = Math.floor(Math.random() * playingQueue.length);
    playTrack(n);
  } else if (currentIndex + 1 < playingQueue.length)
    playTrack(currentIndex + 1);
  else if (repeatMode === 1) playTrack(0);
  else {
    playbackFinished = true;
    ++playbackGeneration;
    playbackController?.abort();
    isLoadingTrack = false;
    audio.pause();
    audio.removeAttribute("src");
    audio.load();
    document.getElementById("player_card").classList.remove("shimmer-active");
    progressBar.value = 0;
    progressBar.disabled = true;
    currentTimeEl.textContent = "0:00";
    document.getElementById("play_pause_btn").setAttribute("aria-label", "Play");
    document.getElementById("play_pause_btn").title = "Play";
    setStatus("FINISHED", "standby");
    updateActiveTrackUI();
    document.getElementById("icon_pause").classList.add("hidden");
    document.getElementById("icon_play").classList.remove("hidden");
  }
}
function playPrevTrack() {
  if (!playingQueue.length) return;
  if (audio.currentTime > 3) audio.currentTime = 0;
  else if (currentIndex > 0) playTrack(currentIndex - 1);
}
async function deleteItem(id, name, isFolder) {
  if (
    !confirm(`Delete ${isFolder ? "folder" : "track"} '${name}' from OneDrive?`)
  )
    return;
  const folderId = currentFolderId;
  try {
    await graphRequest(`me/drive/items/${encodeURIComponent(id)}`, {
      method: "DELETE",
    });
    delete folderCache[folderId];
    if (currentFolderId === folderId)
      await fetchMyFolders(folderId, pathHistory.at(-1).name, false);
  } catch (error) {
    notify(friendlyError(error));
  }
}
let uploadController = null;
async function handleFileUpload(event) {
  const input = event.target,
    files = [...input.files];
  if (!files.length || uploadController) return;
  const folderId = currentFolderId,
    folderName = pathHistory.at(-1).name;
  uploadController = new AbortController();
  const signal = uploadController.signal;
  let sessionUrl = null,
    completed = 0;
  const container = document.getElementById("file_list_container");
  const state = document.createElement("div");
  state.className = "empty-state";
  const label = document.createElement("p");
  label.setAttribute("role", "status");
  const progress = document.createElement("progress");
  progress.id = "upload_progress";
  progress.max = 100;
  progress.value = 0;
  progress.setAttribute("aria-label", "Upload progress");
  const cancel = document.createElement("button");
  cancel.textContent = "Cancel upload";
  cancel.onclick = () => uploadController?.abort();
  state.append(label, progress, cancel);
  container.replaceChildren(state);
  document.getElementById("action_buttons").inert = true;
  try {
    for (const file of files) {
      if (signal.aborted) throw new DOMException("Cancelled", "AbortError");
      if (!file.size) throw new Error("Empty audio files cannot be uploaded.");
      label.textContent = `Uploading ${file.name} (${completed + 1} of ${files.length})`;
      const session = await graphRequest(
        `me/drive/items/${encodeURIComponent(folderId)}:/${encodeURIComponent(file.name)}:/createUploadSession`,
        { method: "POST", signal, body: "{}" },
      );
      sessionUrl = session.uploadUrl;
      if (!sessionUrl || new URL(sessionUrl).protocol !== "https:")
        throw new Error("Could not start the upload.");
      // OneDrive requires 320 KiB multiples for non-final fragments.
      const chunkSize = 10 * 320 * 1024;
      for (let start = 0; start < file.size; start += chunkSize) {
        const end = Math.min(start + chunkSize, file.size);
        const response = await fetch(sessionUrl, {
          method: "PUT",
          signal,
          headers: {
            "Content-Range": `bytes ${start}-${end - 1}/${file.size}`,
          },
          body: file.slice(start, end),
        });
        if (!response.ok) {
          const error = new Error("Upload failed");
          error.status = response.status;
          throw error;
        }
        if (end === file.size && ![200, 201].includes(response.status))
          throw new Error("The upload could not be completed.");
        progress.value = ((completed + end / file.size) / files.length) * 100;
      }
      sessionUrl = null;
      completed++;
    }
    notify(
      `${completed} ${completed === 1 ? "file uploaded" : "files uploaded"}.`,
    );
  } catch (error) {
    if (sessionUrl) {
      try {
        await fetch(sessionUrl, {
          method: "DELETE",
          signal: AbortSignal.timeout(5000),
        });
      } catch {
        /* OneDrive expires unfinished sessions. */
      }
    }
    notify(
      error.name === "AbortError"
        ? `Upload cancelled. ${completed} completed.`
        : `Upload failed. ${completed} completed. ${friendlyError(error)}`,
    );
  } finally {
    input.value = "";
    uploadController = null;
    document.getElementById("action_buttons").inert = false;
    delete folderCache[folderId];
    if (currentFolderId === folderId)
      await fetchMyFolders(folderId, folderName, false);
  }
}

async function requestWakeLock() {
  try {
    if ("wakeLock" in navigator) {
      wakeLock = await navigator.wakeLock.request("screen");
      const ind = document.getElementById("wake_lock_indicator");
      if (ind) {
        ind.classList.remove("opacity-0", "scale-95");
        ind.classList.add("opacity-100", "scale-100");
      }
    }
  } catch (e) {
    console.log("Wake lock unavailable:", e);
  }
}

document.addEventListener("visibilitychange", () => {
  if (document.visibilityState === "visible" && !audio.paused) {
    requestWakeLock();
  }
});
audio.addEventListener("timeupdate", () => {
  if (!Number.isFinite(audio.duration) || audio.duration <= 0) return;
  progressBar.value = (audio.currentTime / audio.duration) * 100;
  currentTimeEl.innerText = formatTime(audio.currentTime);

  if ("mediaSession" in navigator && !isNaN(audio.duration)) {
    try {
      navigator.mediaSession.setPositionState({
        duration: audio.duration,
        playbackRate: audio.playbackRate,
        position: audio.currentTime,
      });
    } catch (e) {}
  }
});
audio.addEventListener("loadedmetadata", () => {
  durationTimeEl.innerText = formatTime(audio.duration);
  progressBar.disabled = false;
});
audio.addEventListener("play", () => {
  document.getElementById("play_pause_btn").setAttribute("aria-label", "Pause");
  document.getElementById("play_pause_btn").title = "Pause";
  document.getElementById("icon_play").classList.add("hidden");
  document.getElementById("icon_pause").classList.remove("hidden");
  setStatus("NOW PLAYING", "play");
  updateActiveTrackUI();
  requestWakeLock();
});
audio.addEventListener("pause", () => {
  if (!audio.paused) return;
  document.getElementById("play_pause_btn").setAttribute("aria-label", "Play");
  document.getElementById("play_pause_btn").title = "Play";
  document.getElementById("icon_pause").classList.add("hidden");
  document.getElementById("icon_play").classList.remove("hidden");
  if (!isLoadingTrack && !playbackFinished) {
    setStatus("PAUSED", "standby");
  }
  updateActiveTrackUI();

  if (wakeLock) {
    wakeLock.release();
    wakeLock = null;
  }
  const ind = document.getElementById("wake_lock_indicator");
  if (ind) {
    ind.classList.add("opacity-0", "scale-95");
    ind.classList.remove("opacity-100", "scale-100");
  }
});
audio.addEventListener("ended", () => playNextTrack(false));
progressBar.addEventListener("input", (event) => {
  if (Number.isFinite(audio.duration) && audio.duration > 0)
    audio.currentTime = (event.target.value / 100) * audio.duration;
});
function formatTime(s) {
  if (!Number.isFinite(s) || s < 0) return "0:00";
  const m = Math.floor(s / 60),
    r = Math.floor(s % 60);
  return `${m}:${r < 10 ? "0" : ""}${r}`;
}

document.addEventListener("DOMContentLoaded", () => {
  const isMobile = /Mobi|Android|iPhone|iPad|iPod/i.test(navigator.userAgent);
  const indicator = document.getElementById("wake_lock_indicator");
  if (!isMobile) indicator.remove();
  else {
    indicator.classList.remove("hidden");
    indicator.classList.add("flex");
  }
  document.querySelectorAll('[id$="_modal"]').forEach((modal) => {
    modal.setAttribute("role", "dialog");
    modal.setAttribute("aria-modal", "true");
    modal.tabIndex = -1;
    const heading = modal.querySelector("h2, h3");
    if (heading) {
      heading.id ||= `${modal.id}_title`;
      modal.setAttribute("aria-labelledby", heading.id);
    }
  });
  updateRepeatBtnUI();
  updateShuffleBtnUI();
  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape") {
      if (activeModal) closeModal(activeModal.id);
      else
        document
          .querySelectorAll(".dropdown-content:not(.hidden)")
          .forEach((menu) => {
            menu.classList.add("hidden");
            const trigger = [
              ...document.querySelectorAll("[aria-controls]"),
            ].find(
              (button) => button.getAttribute("aria-controls") === menu.id,
            );
            trigger?.setAttribute("aria-expanded", "false");
            trigger?.focus();
          });
      return;
    }
    if (activeModal && event.key === "Tab") {
      const elements = [
        ...activeModal.querySelectorAll(
          'button, input, textarea, a[href], [tabindex="0"]',
        ),
      ].filter(
        (element) => !element.disabled && element.getClientRects().length,
      );
      const first = elements[0],
        last = elements.at(-1);
      if (
        event.shiftKey &&
        (document.activeElement === first ||
          document.activeElement === activeModal)
      ) {
        event.preventDefault();
        last?.focus();
      } else if (!event.shiftKey && document.activeElement === last) {
        event.preventDefault();
        first?.focus();
      }
      return;
    }
    if (
      event.target.matches('[role="button"]') &&
      (event.key === "Enter" || event.code === "Space")
    ) {
      event.preventDefault();
      event.target.click();
      return;
    }
    if (
      activeModal ||
      event.target.closest(
        'button, input, textarea, a, [contenteditable="true"], [role="button"]',
      )
    )
      return;
    if (event.code === "Space") {
      event.preventDefault();
      togglePlayPause();
    }
  });
  initializeAccount();
});
