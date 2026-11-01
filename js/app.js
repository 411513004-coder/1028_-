/**
 * 花蓮生活地圖 - 核心互動邏輯
 * 支援三大主題切換：蔬食美食、精選景點、質感咖啡廳
 */

let map;
let markers = [];
let currentType = "全部";      // 大主題：全部 / 素食 / 景點 / 咖啡廳
let currentCategory = "全部";  // 素食子分類：全部 / 全素 / 蛋奶素 / 蔬食友善
let currentTown = "全部";      // 行政區
let currentSearch = "";        // 搜尋關鍵字
let activeCardId = null;

// 初始化應用
document.addEventListener("DOMContentLoaded", () => {
  initMap();
  initTypeTabs();
  initFilters();
  renderPlaces();
  setupMobileToggle();
});

// 初始化 Leaflet 地圖
function initMap() {
  const hualienCenter = [23.980, 121.605];
  
  map = L.map("map", {
    center: hualienCenter,
    zoom: 12,
    zoomControl: true
  });

  // CARTO Voyager 免費開源圖資 (無金鑰、不擋本機檔案)
  const cartoVoyager = L.tileLayer("https://{s}.basemaps.cartocdn.com/rastertiles/voyager/{z}/{x}/{y}{r}.png", {
    attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> &copy; <a href="https://carto.com/attributions">CARTO</a> | 花蓮探索地圖',
    subdomains: 'abcd',
    maxZoom: 20
  });

  // 備用圖層：Esri 世界街道圖
  const esriStreets = L.tileLayer("https://server.arcgisonline.com/ArcGIS/rest/services/World_Street_Map/MapServer/tile/{z}/{y}/{x}", {
    attribution: 'Tiles &copy; Esri &mdash; Source: Esri, DeLorme, NAVTEQ',
    maxZoom: 19
  });

  cartoVoyager.addTo(map);

  const baseMaps = {
    "🌿 清新探訪風 (推薦)": cartoVoyager,
    "🗺️ 經典街道圖": esriStreets
  };
  L.control.layers(baseMaps, null, { position: "topright" }).addTo(map);
}

// 建立自訂彩色圖標 Pin
function createCustomPin(item) {
  let iconEmoji = "📍";
  let pinClass = "pin-全素";

  if (item.type === "景點") {
    iconEmoji = "🏛️";
    pinClass = "pin-景點";
  } else if (item.type === "咖啡廳") {
    iconEmoji = "☕";
    pinClass = "pin-咖啡廳";
  } else {
    // 蔬食分類
    if (item.category === "全素") {
      iconEmoji = "🌱";
      pinClass = "pin-全素";
    } else if (item.category === "蛋奶素") {
      iconEmoji = "🥚";
      pinClass = "pin-蛋奶素";
    } else if (item.category === "蔬食友善") {
      iconEmoji = "🥗";
      pinClass = "pin-蔬食友善";
    }
  }

  return L.divIcon({
    className: "custom-div-icon",
    html: `<div class="custom-marker-pin ${pinClass}"><div class="custom-marker-inner">${iconEmoji}</div></div>`,
    iconSize: [34, 34],
    iconAnchor: [17, 34],
    popupAnchor: [0, -34]
  });
}

// 建立 Popup 彈出卡片內容
function createPopupContent(item) {
  const specialtyTags = item.specialty
    .map(s => `<span class="card-tag">✨ ${s}</span>`)
    .join("");

  const isPhoneAvailable = item.phone && !item.phone.includes("無電話") && !item.phone.includes("粉專");
  const phoneBtn = isPhoneAvailable 
    ? `<a href="tel:${item.phone}" class="action-btn call-btn">📞 一鍵撥號</a>`
    : `<button class="action-btn call-btn" style="opacity:0.75; cursor:default;" title="${item.phone}">💬 社群/預約制</button>`;

  return `
    <div class="popup-card">
      <div class="popup-header">
        <h3 class="popup-title">${item.name}</h3>
        <div class="popup-badges">
          <span class="card-category category-${item.category}">${item.category}</span>
          <span class="badge-tag">📍 ${item.town}</span>
          <span class="badge-tag" style="background:#fef3c7; color:#b45309; border-color:#fde68a;">💰 ${item.price}</span>
        </div>
      </div>
      
      <p class="popup-desc">${item.description}</p>
      
      <div class="popup-meta">
        <div>🕒 <strong>營業/開放：</strong>${item.business_hours}</div>
        <div>📍 <strong>地址：</strong>${item.address}</div>
        <div>📞 <strong>聯絡/電話：</strong>${item.phone}</div>
      </div>

      <div>
        <div style="font-size:0.75rem; color:#64748b; margin-bottom:4px; font-weight:600;">特色亮點：</div>
        <div class="card-tags">${specialtyTags}</div>
      </div>

      <div class="popup-actions">
        ${phoneBtn}
        <a href="https://www.google.com/maps/dir/?api=1&destination=${item.lat},${item.lng}" target="_blank" class="action-btn nav-btn">
          🧭 導航前往
        </a>
      </div>
    </div>
  `;
}

// 渲染列表與地圖標記
function renderPlaces() {
  const listContainer = document.getElementById("restaurantList");
  const resultCount = document.getElementById("resultCount");

  // 清空地圖現有標記
  markers.forEach(m => map.removeLayer(m.marker));
  markers = [];
  listContainer.innerHTML = "";

  // 篩選邏輯
  const filtered = PLACES_DATA.filter(item => {
    // 1. 大分類過濾 (全部 / 素食 / 景點 / 咖啡廳)
    if (currentType !== "全部" && item.type !== currentType) {
      return false;
    }
    // 2. 素食子類別過濾 (全素 / 蛋奶素 / 蔬食友善)
    if (currentCategory !== "全部") {
      if (item.category !== currentCategory) {
        return false;
      }
    }
    // 3. 行政區過濾
    if (currentTown !== "全部" && item.town !== currentTown) {
      return false;
    }
    // 4. 關鍵字搜尋過濾
    if (currentSearch.trim() !== "") {
      const q = currentSearch.toLowerCase();
      const matchName = item.name.toLowerCase().includes(q);
      const matchAddress = item.address.toLowerCase().includes(q);
      const matchTown = item.town.toLowerCase().includes(q);
      const matchDesc = item.description.toLowerCase().includes(q);
      const matchSpecialty = item.specialty.some(s => s.toLowerCase().includes(q));
      if (!matchName && !matchAddress && !matchTown && !matchDesc && !matchSpecialty) {
        return false;
      }
    }
    return true;
  });

  resultCount.textContent = `共找到 ${filtered.length} 個精選地點`;

  if (filtered.length === 0) {
    listContainer.innerHTML = `
      <div style="text-align:center; padding:40px 20px; color:#94a3b8;">
        <div style="font-size:40px; margin-bottom:10px;">🔍</div>
        <p>查無符合條件的地點</p>
        <p style="font-size:0.8rem; margin-top:4px;">請切換分類標籤或清除搜尋條件重試</p>
      </div>
    `;
    return;
  }

  const latlngBounds = [];

  filtered.forEach(item => {
    // 1. 地圖標記
    const pinIcon = createCustomPin(item);
    const marker = L.marker([item.lat, item.lng], { icon: pinIcon }).addTo(map);
    const popupContent = createPopupContent(item);
    
    marker.bindPopup(popupContent, { maxWidth: 320 });
    
    marker.on("click", () => {
      highlightCard(item.id);
    });

    markers.push({ id: item.id, marker });
    latlngBounds.push([item.lat, item.lng]);

    // 2. 左側卡片元素
    const card = document.createElement("div");
    card.className = `restaurant-card ${activeCardId === item.id ? "active" : ""}`;
    card.id = `card-${item.id}`;

    const specialtyBadges = item.specialty
      .slice(0, 3)
      .map(s => `<span class="card-tag">✨ ${s}</span>`)
      .join("");

    card.innerHTML = `
      <div class="card-top">
        <div class="card-title">${item.name}</div>
        <span class="card-category category-${item.category}">${item.category}</span>
      </div>
      <div class="card-info-item">
        <span>📍 ${item.town} · ${item.address}</span>
      </div>
      <div class="card-info-item">
        <span>🕒 ${item.business_hours}</span>
      </div>
      <div class="card-tags">
        ${specialtyBadges}
      </div>
    `;

    card.addEventListener("click", () => {
      focusPlace(item);
    });

    listContainer.appendChild(card);
  });

  // 自動調整視角以包含所有結果
  if (latlngBounds.length > 0) {
    map.fitBounds(latlngBounds, { padding: [40, 40], maxZoom: 15 });
  }
}

// 聚焦特定卡片地點
function focusPlace(item) {
  activeCardId = item.id;
  highlightCard(item.id);

  map.setView([item.lat, item.lng], 15, { animate: true });
  
  const target = markers.find(m => m.id === item.id);
  if (target) {
    target.marker.openPopup();
  }

  // 手機版自動切換回地圖
  if (window.innerWidth <= 768) {
    const sidebar = document.getElementById("sidebar");
    sidebar.classList.add("hidden");
    const toggleBtn = document.getElementById("mobileToggleBtn");
    toggleBtn.textContent = "📋 查看地點清單";
  }
}

// 高亮卡片並自動滾動
function highlightCard(id) {
  activeCardId = id;
  document.querySelectorAll(".restaurant-card").forEach(c => {
    c.classList.remove("active");
  });
  const targetCard = document.getElementById(`card-${id}`);
  if (targetCard) {
    targetCard.classList.add("active");
    targetCard.scrollIntoView({ behavior: "smooth", block: "nearest" });
  }
}

// 大分類 Tab 標籤切換
function initTypeTabs() {
  const tabBtns = document.querySelectorAll(".tab-btn");
  const subCategoryFilter = document.getElementById("subCategoryFilter");

  tabBtns.forEach(btn => {
    btn.addEventListener("click", () => {
      tabBtns.forEach(b => b.classList.remove("active"));
      btn.classList.add("active");
      currentType = btn.dataset.type;

      // 若切換到非蔬食分類，隱藏蔬食子分類按鈕列；切回蔬食時顯示
      if (currentType === "素食" || currentType === "全部") {
        subCategoryFilter.style.display = "flex";
      } else {
        subCategoryFilter.style.display = "none";
        currentCategory = "全部";
      }

      renderPlaces();
    });
  });
}

// 初始化篩選器事件
function initFilters() {
  // 蔬食子分類標籤點擊
  const categoryBtns = document.querySelectorAll(".filter-btn");
  categoryBtns.forEach(btn => {
    btn.addEventListener("click", () => {
      categoryBtns.forEach(b => b.classList.remove("active"));
      btn.classList.add("active");
      currentCategory = btn.dataset.category;
      renderPlaces();
    });
  });

  // 鄉鎮選單變更
  const townSelect = document.getElementById("townSelect");
  townSelect.addEventListener("change", (e) => {
    currentTown = e.target.value;
    renderPlaces();
  });

  // 搜尋框輸入
  const searchInput = document.getElementById("searchInput");
  searchInput.addEventListener("input", (e) => {
    currentSearch = e.target.value;
    renderPlaces();
  });

  // 重置按鈕
  const resetBtn = document.getElementById("resetBtn");
  resetBtn.addEventListener("click", () => {
    currentType = "全部";
    currentCategory = "全部";
    currentTown = "全部";
    currentSearch = "";
    searchInput.value = "";
    townSelect.value = "全部";
    
    document.querySelectorAll(".tab-btn").forEach(b => b.classList.toggle("active", b.dataset.type === "全部"));
    document.getElementById("subCategoryFilter").style.display = "flex";
    categoryBtns.forEach(b => b.classList.toggle("active", b.dataset.category === "全部"));
    
    renderPlaces();
  });
}

// 手機版切換按鈕
function setupMobileToggle() {
  const toggleBtn = document.getElementById("mobileToggleBtn");
  const sidebar = document.getElementById("sidebar");

  toggleBtn.addEventListener("click", () => {
    const isHidden = sidebar.classList.contains("hidden");
    if (isHidden) {
      sidebar.classList.remove("hidden");
      toggleBtn.textContent = "🗺️ 查看地圖";
    } else {
      sidebar.classList.add("hidden");
      toggleBtn.textContent = "📋 查看地點清單";
    }
  });
}
