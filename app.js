const CLIENT_ID = "9a019c14-5029-4055-b5c8-0e42509f6999";
const REDIRECT_URI = "https://softsound557.github.io/stunning-doodle/";
const SCOPES = ["Files.Read"];
const FOLDER_NAME = "荷物管理エクスポート";
const GRAPH_ROOT = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(FOLDER_NAME)}`;

const msalInstance = new msal.PublicClientApplication({
  auth: {
    clientId: CLIENT_ID,
    authority: "https://login.microsoftonline.com/consumers",
    redirectUri: REDIRECT_URI,
  },
  cache: {
    cacheLocation: "localStorage",
  },
});

let items = [];
const photoUrlCache = new Map();

const setupView = document.getElementById("setup-view");
const listView = document.getElementById("list-view");
const detailView = document.getElementById("detail-view");
const itemListEl = document.getElementById("item-list");
const countEl = document.getElementById("count");
const searchInput = document.getElementById("search-input");
const statusEl = document.getElementById("status");

function showStatus(message) {
  statusEl.textContent = message;
  statusEl.hidden = !message;
}

function showSignIn() {
  setupView.hidden = false;
  listView.hidden = true;
  detailView.hidden = true;
}

async function getToken() {
  const account = msalInstance.getActiveAccount();
  try {
    const result = await msalInstance.acquireTokenSilent({ scopes: SCOPES, account });
    return result.accessToken;
  } catch (e) {
    await msalInstance.acquireTokenRedirect({ scopes: SCOPES });
    throw e;
  }
}

async function loadPhotoUrl(filename) {
  if (photoUrlCache.has(filename)) return photoUrlCache.get(filename);
  const token = await getToken();
  const res = await fetch(`${GRAPH_ROOT}/photos/${encodeURIComponent(filename)}:/content`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) throw new Error("photo fetch failed: " + res.status);
  const blob = await res.blob();
  const url = URL.createObjectURL(blob);
  photoUrlCache.set(filename, url);
  return url;
}

const thumbObserver = new IntersectionObserver((entries, obs) => {
  for (const entry of entries) {
    if (!entry.isIntersecting) continue;
    const el = entry.target;
    obs.unobserve(el);
    const filename = el.dataset.filename;
    loadPhotoUrl(filename)
      .then((url) => { el.src = url; })
      .catch(() => {});
  }
});

async function loadData() {
  setupView.hidden = true;
  listView.hidden = false;

  try {
    const token = await getToken();
    const res = await fetch(`${GRAPH_ROOT}/data.json:/content`, {
      headers: { Authorization: `Bearer ${token}` },
      cache: "no-store",
    });
    if (!res.ok) throw new Error("status " + res.status);
    items = await res.json();
    showStatus("");
  } catch (e) {
    showStatus("最新データを取得できませんでした(前回表示分を表示しています)");
  }
  renderList();
}

function renderList() {
  const q = searchInput.value.trim();
  const filtered = q
    ? items.filter((it) => [it.name, it.note, it.category].some((v) => (v || "").includes(q)))
    : items;

  countEl.textContent = `${filtered.length}件`;
  itemListEl.innerHTML = "";

  if (filtered.length === 0) {
    itemListEl.innerHTML = '<li class="empty">見つかりませんでした</li>';
    return;
  }

  for (const it of filtered) {
    const li = document.createElement("li");
    const a = document.createElement("a");
    a.href = "#";
    a.className = "item-row";
    a.addEventListener("click", (e) => {
      e.preventDefault();
      showDetail(it.id);
    });

    const hasPhoto = it.photos.length > 0;
    const thumb = document.createElement(hasPhoto ? "img" : "div");
    thumb.className = "thumb" + (hasPhoto ? "" : " no-photo");
    if (hasPhoto) {
      thumb.dataset.filename = it.photos[0];
      thumb.loading = "lazy";
      thumb.alt = "";
      thumbObserver.observe(thumb);
    } else {
      thumb.textContent = "写真なし";
    }

    const info = document.createElement("div");
    info.className = "item-info";
    const nameEl = document.createElement("div");
    nameEl.className = "item-name";
    nameEl.textContent = it.name;
    const locEl = document.createElement("div");
    locEl.className = "item-location";
    locEl.textContent = it.location;
    info.appendChild(nameEl);
    info.appendChild(locEl);

    a.appendChild(thumb);
    a.appendChild(info);

    if (it.discarded_date) {
      const badge = document.createElement("div");
      badge.className = "badge";
      badge.textContent = "処分済み";
      a.appendChild(badge);
    }

    li.appendChild(a);
    itemListEl.appendChild(li);
  }
}

async function showDetail(id) {
  const it = items.find((x) => x.id === id);
  if (!it) return;

  document.getElementById("detail-name").textContent = it.name;
  document.getElementById("detail-location").textContent = it.location || "-";
  document.getElementById("detail-category").textContent = it.category || "-";
  document.getElementById("detail-note").textContent = it.note || "-";
  document.getElementById("detail-added").textContent = it.added_date || "-";

  const badge = document.getElementById("detail-badge");
  if (it.discarded_date) {
    badge.hidden = false;
    badge.textContent = `処分済み（${it.discarded_date}）`;
  } else {
    badge.hidden = true;
  }

  const photosEl = document.getElementById("detail-photos");
  photosEl.innerHTML = "";
  if (it.photos.length === 0) {
    photosEl.innerHTML = '<div class="no-photo-large">写真なし</div>';
  } else {
    for (const filename of it.photos) {
      const img = document.createElement("img");
      img.className = "photo";
      photosEl.appendChild(img);
      loadPhotoUrl(filename)
        .then((url) => { img.src = url; })
        .catch(() => {});
    }
  }

  listView.hidden = true;
  detailView.hidden = false;
}

document.getElementById("back-link").addEventListener("click", (e) => {
  e.preventDefault();
  detailView.hidden = true;
  listView.hidden = false;
});

document.getElementById("settings-link").addEventListener("click", (e) => {
  e.preventDefault();
  msalInstance.logoutRedirect();
});

document.getElementById("signin-button").addEventListener("click", () => {
  msalInstance.loginRedirect({ scopes: SCOPES });
});

searchInput.addEventListener("input", renderList);

if ("serviceWorker" in navigator) {
  navigator.serviceWorker.register("sw.js");
}

async function init() {
  const result = await msalInstance.handleRedirectPromise();
  let account = msalInstance.getAllAccounts()[0];
  if (result && result.account) {
    account = result.account;
  }
  if (!account) {
    showSignIn();
    return;
  }
  msalInstance.setActiveAccount(account);
  await loadData();
}

init();
