const STORAGE_KEY = "share_url";

function encodeShareUrl(url) {
  const base64 = btoa(url)
    .replace(/=+$/, "")
    .replace(/\//g, "_")
    .replace(/\+/g, "-");
  return "u!" + base64;
}

function getShareUrl() {
  return localStorage.getItem(STORAGE_KEY) || "";
}

function setShareUrl(url) {
  localStorage.setItem(STORAGE_KEY, url);
}

let SHARE_URL = getShareUrl();
let GRAPH_ROOT = SHARE_URL ? `https://graph.microsoft.com/v1.0/shares/${encodeShareUrl(SHARE_URL)}/driveItem` : null;
let items = [];

const setupView = document.getElementById("setup-view");
const listView = document.getElementById("list-view");
const detailView = document.getElementById("detail-view");
const itemListEl = document.getElementById("item-list");
const countEl = document.getElementById("count");
const searchInput = document.getElementById("search-input");
const statusEl = document.getElementById("status");

function photoUrl(filename) {
  return `${GRAPH_ROOT}:/photos/${filename}:/content`;
}

function showStatus(message) {
  statusEl.textContent = message;
  statusEl.hidden = !message;
}

function showSetup() {
  setupView.hidden = false;
  listView.hidden = true;
  detailView.hidden = true;
}

async function loadData() {
  if (!GRAPH_ROOT) {
    showSetup();
    return;
  }
  setupView.hidden = true;
  listView.hidden = false;

  try {
    const res = await fetch(`${GRAPH_ROOT}:/data.json:/content`, { cache: "no-store" });
    if (!res.ok) throw new Error("status " + res.status);
    items = await res.json();
    showStatus("");
  } catch (e) {
    showStatus("最新データを取得できませんでした（前回表示分を表示しています）");
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
      thumb.src = photoUrl(it.photos[0]);
      thumb.loading = "lazy";
      thumb.alt = "";
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

function showDetail(id) {
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
      img.loading = "lazy";
      img.src = photoUrl(filename);
      photosEl.appendChild(img);
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
  showSetup();
});

document.getElementById("setup-form").addEventListener("submit", (e) => {
  e.preventDefault();
  const url = document.getElementById("setup-input").value.trim();
  if (!url) return;
  setShareUrl(url);
  SHARE_URL = url;
  GRAPH_ROOT = `https://graph.microsoft.com/v1.0/shares/${encodeShareUrl(SHARE_URL)}/driveItem`;
  loadData();
});

searchInput.addEventListener("input", renderList);

if ("serviceWorker" in navigator) {
  navigator.serviceWorker.register("sw.js");
}

loadData();
