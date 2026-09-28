const tg = window.Telegram ? window.Telegram.WebApp : null;

// Telegram oynasini to'liq ekranga yoyish
if (tg) {
  tg.ready();
  tg.expand();
}

// --- TELEGRAM BOT SOZLAMALARI ---
// O'zingizning Bot Token va Telegram ID ingizni kiriting
const TELEGRAM_BOT_TOKEN = "8825305562:AAFtgCijJ6puOljs12GOWaMtFwIAd7xOagM";
const MY_TELEGRAM_ID = "1347548152";

// Boshlang'ich topshiriqlar ro'yxati (String list)
const DEFAULT_TASKS_LIST = [
  "Arenda",
  "Reklama",
  "Chek qog'oz",
  "Ekran",
  "Sistemniy Blok",
  "Paket",
  "Boshqa",
];

// LocalStorage ni xavfsiz o'qish funksiyasi
function safeGetStorage(key, defaultValue) {
  try {
    const item = localStorage.getItem(key);
    return item ? JSON.parse(item) : defaultValue;
  } catch (e) {
    console.error(`LocalStorage xatosi (${key}):`, e);
    return defaultValue;
  }
}

let db = {
  baza: safeGetStorage("inkassa_baza", []),
  marshrutIds: safeGetStorage("inkassa_marshrut", []),
  tarix: safeGetStorage("inkassa_tarix" || []),
  topshiriqAtms: safeGetStorage("topshiriq_atms", []),
  selectedTasks: safeGetStorage("topshiriq_selected_tasks", DEFAULT_TASKS_LIST),
  topshiriqData: safeGetStorage("topshiriq_data", {}),
};

let myMap = null;
let locationControl = null;
let pendingConfirmAction = null;
let currentEditingAtmId = null;

window.onload = function () {
  setupNavigation();
  setupExcelImport();
  renderAllViews();
  initYandexMap();
};

function saveData() {
  localStorage.setItem("inkassa_baza", JSON.stringify(db.baza));
  localStorage.setItem("inkassa_marshrut", JSON.stringify(db.marshrutIds));
  localStorage.setItem("inkassa_tarix", JSON.stringify(db.tarix));
  localStorage.setItem("topshiriq_atms", JSON.stringify(db.topshiriqAtms));
  localStorage.setItem(
    "topshiriq_selected_tasks",
    JSON.stringify(db.selectedTasks),
  );
  localStorage.setItem("topshiriq_data", JSON.stringify(db.topshiriqData));
  renderAllViews();
}

function setupNavigation() {
  const navItems = document.querySelectorAll(".nav-item");
  navItems.forEach((item) => {
    item.addEventListener("click", () => {
      switchTab(item.getAttribute("data-target"));
    });
  });
}

function switchTab(targetId) {
  document
    .querySelectorAll(".nav-item")
    .forEach((i) => i.classList.remove("active"));
  document
    .querySelectorAll(".view")
    .forEach((v) => v.classList.remove("active"));

  const activeNav = document.querySelector(
    `.nav-item[data-target="${targetId}"]`,
  );
  if (activeNav) activeNav.classList.add("active");

  const activeView = document.getElementById(targetId);
  if (activeView) activeView.classList.add("active");

  if (targetId === "view-xarita" && myMap) {
    myMap.container.fitToViewport();
    updateMapMarkers();
  }
}

function renderAllViews() {
  renderBazaView();
  renderMarshrutView();
  renderTarixView();
  renderTopshiriqView();
  if (myMap) updateMapMarkers();
}

function showLoader() {
  const overlay = document.getElementById("loading-overlay");
  if (overlay) overlay.classList.remove("hidden");
}

function hideLoader() {
  const overlay = document.getElementById("loading-overlay");
  if (overlay) overlay.classList.add("hidden");
}

function showToast(message, type = "success") {
  const toast = document.getElementById("toast-notification");
  if (!toast) return;

  toast.innerText = message;
  toast.className = `toast ${type}`;

  setTimeout(() => {
    toast.classList.add("hidden");
  }, 3500);
}

// 1-OYNA: BAZA RENDER
function renderBazaView() {
  const actionsEl = document.getElementById("baza-actions");
  const countEl = document.getElementById("baza-total-count");
  const listEl = document.getElementById("baza-list");

  countEl.innerText = db.baza.length;

  if (db.baza.length === 0) {
    actionsEl.innerHTML = `<button class="btn btn-primary" onclick="triggerFileInput()">Import</button>`;
  } else {
    actionsEl.innerHTML = `
      <button class="btn btn-success" onclick="exportExcel()">Export</button>
      <button class="btn btn-danger" onclick="clearFullSystem()">Tozalash</button>
    `;
  }

  listEl.innerHTML = "";
  db.baza.forEach((atm) => {
    const card = document.createElement("div");
    card.className = "atm-card";
    card.style.borderLeftColor = atm.color || "#2563eb";
    card.innerHTML = `
      <div class="atm-info">
        <span class="color-dot" style="background:${atm.color || "#2563eb"}"></span>
        <span class="atm-name">#${atm.id}. ${atm.name}</span>
      </div>
    `;
    listEl.appendChild(card);
  });
}

function triggerFileInput() {
  document.getElementById("excel-file-input").click();
}

function setupExcelImport() {
  document
    .getElementById("excel-file-input")
    .addEventListener("change", function (e) {
      const file = e.target.files[0];
      if (!file) return;

      const reader = new FileReader();
      reader.onload = function (e) {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: "array" });
          const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
          const jsonData = XLSX.utils.sheet_to_json(firstSheet);

          db.baza = jsonData.map((item) => ({
            id: item.id,
            name: item.name,
            group: item.group || "Umumiy",
            color: item.color || "#FF0000",
            lat: parseFloat(item.lat),
            lng: parseFloat(item.lng),
          }));

          saveData();
          alert("Excel muvaffaqiyatli import qilindi!");
        } catch (err) {
          alert(
            "Excel faylini o'qishda xatolik! Ustunlar: id, name, group, color, lat, lng bo'lishi kerak.",
          );
        }
      };
      reader.readAsArrayBuffer(file);
      this.value = "";
    });
}

function exportExcel() {
  if (db.baza.length === 0) return;
  const worksheet = XLSX.utils.json_to_sheet(db.baza);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Bankomatlar");
  XLSX.writeFile(workbook, "Bankomatlar_Baza.xlsx");
}

function clearFullSystem() {
  showConfirm(
    "Haqiqatan ham barcha bazani va ma'lumotlarni o'chirib, dasturni 0 holatiga qaytarmoqchimisiz?",
    function () {
      db.baza = [];
      db.marshrutIds = [];
      db.tarix = [];
      db.topshiriqAtms = [];
      db.topshiriqData = {};
      saveData();
    },
  );
}

// 2-OYNA: MARSHRUT RENDER
function renderMarshrutView() {
  const listEl = document.getElementById("marshrut-list");
  const activeMarshrut = db.baza.filter(
    (a) =>
      db.marshrutIds.includes(a.id) && !db.tarix.some((t) => t.id === a.id),
  );
  const doneCount = db.marshrutIds.filter((id) =>
    db.tarix.some((t) => t.id === id),
  ).length;

  document.getElementById("m-total").innerText = db.marshrutIds.length;
  document.getElementById("m-done").innerText = doneCount;
  document.getElementById("m-left").innerText = activeMarshrut.length;

  listEl.innerHTML = "";
  activeMarshrut.forEach((atm) => {
    const card = document.createElement("div");
    card.className = "atm-card";
    card.style.borderLeftColor = atm.color;
    card.innerHTML = `
      <div class="atm-info">
        <span class="color-dot" style="background:${atm.color}"></span>
        <span class="atm-name">#${atm.id}. ${atm.name}</span>
      </div>
      <div class="atm-actions">
        <button class="btn-icon-map" title="Xaritada ko'rish" onclick="showOnMap(${atm.id})">🗺️</button>
        <button class="btn-icon-map" title="Nosoz deb belgilash" onclick="confirmBroken(${atm.id})">🚫</button>
        <button class="btn-icon-map" title="Inkassa qilish" onclick="confirmInkassa(${atm.id})">✅</button>
      </div>
    `;
    listEl.appendChild(card);
  });
}

function openSelectModal() {
  const listEl = document.getElementById("select-checkbox-list");
  listEl.innerHTML = "";

  if (db.baza.length === 0) {
    alert("Avval Baza bo'limida Excel fayl import qiling!");
    return;
  }

  const groups = {};
  db.baza.forEach((atm) => {
    const groupName = atm.group || "Umumiy";
    if (!groups[groupName]) groups[groupName] = [];
    groups[groupName].push(atm);
  });

  Object.keys(groups).forEach((groupName, index) => {
    const groupAtms = groups[groupName];
    const allChecked = groupAtms.every((atm) =>
      db.marshrutIds.includes(atm.id),
    );

    const groupWrapper = document.createElement("div");
    groupWrapper.className = "group-wrapper";

    const groupHeader = document.createElement("div");
    groupHeader.className = "group-header";
    groupHeader.innerHTML = `
      <div class="group-title-area">
        <input type="checkbox" class="group-checkbox" data-group="${groupName}" ${allChecked ? "checked" : ""} onchange="toggleGroupCheck(this, '${groupName}')">
        <strong onclick="toggleGroupAccordion('group-items-${index}', this.closest('.group-wrapper'))" style="cursor:pointer;">📂 ${groupName} (${groupAtms.length} ta)</strong>
      </div>
      <span class="accordion-icon" onclick="toggleGroupAccordion('group-items-${index}', this.closest('.group-wrapper'))">▼</span>
    `;

    const itemsContainer = document.createElement("div");
    itemsContainer.id = `group-items-${index}`;
    itemsContainer.className = "group-items hidden";

    groupAtms.forEach((atm) => {
      const isChecked = db.marshrutIds.includes(atm.id) ? "checked" : "";
      const item = document.createElement("label");
      item.className = "checkbox-item";
      item.innerHTML = `
        <input type="checkbox" class="atm-checkbox" data-group="${groupName}" value="${atm.id}" ${isChecked} onchange="updateGroupCheckboxState('${groupName}')">
        <span class="color-dot" style="background:${atm.color}"></span>
        <span>#${atm.id}. ${atm.name}</span>
      `;
      itemsContainer.appendChild(item);
    });

    groupWrapper.appendChild(groupHeader);
    groupWrapper.appendChild(itemsContainer);
    listEl.appendChild(groupWrapper);
  });

  document.getElementById("select-modal").classList.remove("hidden");
}

function toggleGroupCheck(groupMasterCb, groupName) {
  const isChecked = groupMasterCb.checked;
  const atmCheckboxes = document.querySelectorAll(
    `.atm-checkbox[data-group="${groupName}"]`,
  );
  atmCheckboxes.forEach((cb) => {
    cb.checked = isChecked;
  });
}

function updateGroupCheckboxState(groupName) {
  const masterCb = document.querySelector(
    `.group-checkbox[data-group="${groupName}"]`,
  );
  const atmCheckboxes = document.querySelectorAll(
    `.atm-checkbox[data-group="${groupName}"]`,
  );
  const allChecked = Array.from(atmCheckboxes).every((cb) => cb.checked);
  if (masterCb) masterCb.checked = allChecked;
}

function toggleGroupAccordion(containerId, targetWrapper) {
  const targetContainer = document.getElementById(containerId);
  const modalContent = document.querySelector("#select-modal .modal-content");
  const isCurrentlyHidden = targetContainer.classList.contains("hidden");

  document
    .querySelectorAll(".group-items")
    .forEach((el) => el.classList.add("hidden"));
  document
    .querySelectorAll(".group-wrapper")
    .forEach((el) => el.classList.remove("open"));

  if (isCurrentlyHidden) {
    targetContainer.classList.remove("hidden");
    targetWrapper.classList.add("open");
    if (modalContent) modalContent.classList.add("expanded");
  } else {
    if (modalContent) modalContent.classList.remove("expanded");
  }
}

function closeSelectModal() {
  const modalContent = document.querySelector("#select-modal .modal-content");
  if (modalContent) modalContent.classList.remove("expanded");
  document.getElementById("select-modal").classList.add("hidden");
}

function saveSelectedMarshrut() {
  const checkboxes = document.querySelectorAll(".atm-checkbox");
  const selected = [];
  checkboxes.forEach((cb) => {
    if (cb.checked) selected.push(parseInt(cb.value));
  });
  db.marshrutIds = selected;
  saveData();
  closeSelectModal();
}

function clearMarshrutData() {
  showConfirm(
    "Haqiqatan ham bugungi marshrutni tozalamoqchimisiz?",
    function () {
      db.marshrutIds = [];
      saveData();
    },
  );
}

// Kunni yakunlash va Telegramga hisobot yuborish
function finishDay() {
  // 1. Ma'lumotlar bor-yo'qligini tekshirish
  if (db.marshrutIds.length === 0 && db.tarix.length === 0) {
    alert("Yakunlash uchun faol marshrut yoki tarix mavjud emas!");
    return;
  }

  // 2. Maxsus loyiha confirm modalidan foydalanish
  showConfirm(
    "Rostdan ham kunni yakunlab, ma'lumotlarni tozalamoqchimisiz?",
    function () {
      executeFinishDay();
    },
  );
}

// Kunni yakunlash mantig'i va Telegramga yuborish
async function executeFinishDay() {
  // Bugungi sana (DD.MM.YYYY)
  const today = new Date();
  const dateStr =
    today.getDate().toString().padStart(2, "0") +
    "." +
    (today.getMonth() + 1).toString().padStart(2, "0") +
    "." +
    today.getFullYear();

  // 1. Success va Broken hisoblash
  const successCount = db.tarix.filter((t) => t.status === "success").length;
  const brokenCount = db.tarix.length - successCount;

  // 2. Unhandled (Left) bankomatlarni aniqlash
  const unhandledAtms = [];
  db.marshrutIds.forEach((id) => {
    const inHistory = db.tarix.some((t) => t.id === id);
    if (!inHistory) {
      const atmData = db.baza.find((a) => a.id === id);
      if (atmData) {
        unhandledAtms.push(atmData);
      }
    }
  });

  const leftCount = unhandledAtms.length;

  // 3. Jami (All) bankomatlar soni
  const total = successCount + brokenCount + leftCount;

  // Tarix va harakatlar ro'yxati
  let historyText = "";

  if (db.tarix.length === 0) {
    historyText = "• Bugun hech qanday harakat bajarilmadi.\n";
  } else {
    db.tarix.forEach((item) => {
      const statusLabel = item.status === "broken" ? "Nosoz 🚫" : "Inkassa ✅";
      historyText += `• ${item.time} — #${item.id}. ${item.name} — ${statusLabel}\n`;
    });
  }

  // Marshrutda qolib ketgan (bajarilmagan) bankomatlarni kiritish

  if (unhandledAtms.length > 0) {
    historyText += "\n⚠️ BAJARILMAGANLAR:\n";
    unhandledAtms.forEach((atm) => {
      historyText += `• #${atm.id}. ${atm.name}\n`;
    });
  }

  // Mas'ul xodim ma'lumoti
  let workerInfo = "Noma'lum xodim";
  if (tg && tg.initDataUnsafe && tg.initDataUnsafe.user) {
    const u = tg.initDataUnsafe.user;
    const name = `${u.first_name || ""} ${u.last_name || ""}`.trim();
    const username = u.username ? ` (@${u.username})` : "";
    workerInfo = `${name}${username}`.trim();
  }

  // Telegramga yuboriladigan yakuniy hisobot matni
  const reportMessage = `📅 Sana: ${dateStr}

📊 STATISTIKA:
• Jami marshrut: ${total} ta
• Inkassa qilindi: ${successCount} ta ✅
• Nosoz deb belgilandi: ${brokenCount} ta 🚫
• Bajarilmadi: ${leftCount} ta ⏳

📜 KUN DAVOMIDAGI HARAKATLAR:
${historyText.trim()}

👤 Mas'ul xodim: ${workerInfo}`;

  // Telegram botga ma'lumotni uzatish
  if (tg) {
    showLoader(); // Ekranni bloklab loaderni yoqish
    try {
      const response = await fetch(
        `https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/sendMessage`,
        {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            chat_id: MY_TELEGRAM_ID,
            text: reportMessage,
            parse_mode: "HTML",
          }),
        },
      );

      const result = await response.json();

      // FAQAT XABAR MUVAFFAQIYATLI KETGANDA TOZALAYMIZ
      if (result.ok) {
        // Marshrut va Tarix ma'lumotlarini tozalash
        db.marshrutIds = [];
        db.tarix = [];

        // Lokal xotiraga saqlash va ekran hamda xaritani qayta render qilish
        saveData();

        hideLoader(); // Loaderni o'chirish
        showToast("Kun muvaffaqiyatli yakunlandi!", "success");
      } else {
        hideLoader();
        showToast("Xatolik: Kun yakunlanmadi!", "error");
        console.error("Telegram API Error Response:", result);
      }
    } catch (error) {
      // Internet uzilganda yoki tarmoq xatosida ma'lumotlar toza emas, xavfsiz saqlanadi
      hideLoader();
      showToast("Internet bilan aloqa yo'q!", "error");
      console.error("Fetch Network Error:", error);
    }
  } else {
    console.log(
      "Telegram WebApp topilmadi. Yuborilgan xabar:\n",
      reportMessage,
    );
    showToast("Test rejimida bajarildi", "success");
  }
}

// 3-OYNA: TARIX RENDER
function renderTarixView() {
  const listEl = document.getElementById("tarix-list");
  const countEl = document.getElementById("tarix-total-count");
  const successEl = document.getElementById("tarix-success-count");
  const brokenEl = document.getElementById("tarix-broken-count");

  const successCount = db.tarix.filter(
    (t) => t.status === "success" || !t.status,
  ).length;
  const brokenCount = db.tarix.filter((t) => t.status === "broken").length;

  if (countEl) countEl.innerText = db.tarix.length;
  if (successEl) successEl.innerText = successCount;
  if (brokenEl) brokenEl.innerText = brokenCount;

  listEl.innerHTML = "";
  const reversedTarix = [...db.tarix].reverse();

  reversedTarix.forEach((item) => {
    const el = document.createElement("div");
    const isBroken = item.status === "broken";
    el.className = `tarix-item ${isBroken ? "tarix-item-broken" : "tarix-item-success"}`;
    el.innerHTML = `
      <span class="tarix-info">${item.time} — #${item.id}. ${item.name}</span>
      <div class="tarix-actions">
        ${isBroken ? `<span class="tarix-warning-icon">⚠️</span>` : ""}
        <button class="btn-remove-tarix" title="Tarixdan o'chirish" onclick="removeFromTarix(${item.id})">❌</button>
      </div>
    `;
    listEl.appendChild(el);
  });
}

function removeFromTarix(id) {
  showConfirm(
    "Haqiqatan ham ushbu bankomatni tarixdan o'chirib, marshrutga qaytarmoqchimisiz?",
    function () {
      db.tarix = db.tarix.filter((t) => t.id !== id);
      if (!db.marshrutIds.includes(id)) {
        db.marshrutIds.push(id);
      }
      saveData();
    },
  );
}

function clearTarixData() {
  showConfirm("Haqiqatan ham inkassa tarixini tozalamoqchimisiz?", function () {
    db.tarix = [];
    saveData();
  });
}

function confirmInkassa(id) {
  showConfirm("Ushbu bankomatni inkassa qildingizmi?", function () {
    const atm = db.baza.find((a) => a.id === id);
    if (!atm) return;

    const now = new Date();
    const timeStr =
      now.getHours().toString().padStart(2, "0") +
      ":" +
      now.getMinutes().toString().padStart(2, "0");

    db.tarix.push({
      id: atm.id,
      name: atm.name,
      time: timeStr,
      status: "success",
    });

    saveData();
  });
}

function confirmBroken(id) {
  showConfirm("Ushbu bankomatni nosoz deb belgilamoqchimisiz?", function () {
    const atm = db.baza.find((a) => a.id === id);
    if (!atm) return;

    const now = new Date();
    const timeStr =
      now.getHours().toString().padStart(2, "0") +
      ":" +
      now.getMinutes().toString().padStart(2, "0");

    db.tarix.push({
      id: atm.id,
      name: atm.name,
      time: timeStr,
      status: "broken",
    });

    saveData();
  });
}

function showConfirm(msg, yesCallback) {
  document.getElementById("confirm-message").innerText = msg;
  pendingConfirmAction = yesCallback;
  document.getElementById("confirm-modal").classList.remove("hidden");
}

function closeConfirmModal() {
  document.getElementById("confirm-modal").classList.add("hidden");
  pendingConfirmAction = null;
}

document
  .getElementById("confirm-yes-btn")
  .addEventListener("click", function () {
    if (pendingConfirmAction) pendingConfirmAction();
    closeConfirmModal();
  });

// 4-OYNA: YANDEX MAPS
function initYandexMap() {
  ymaps.ready(() => {
    myMap = new ymaps.Map("yandex-map", {
      center: [37.2242, 67.2783],
      zoom: 13,
      controls: [],
    });

    locationControl = new ymaps.control.GeolocationControl({
      options: { noPlacemark: false, visible: false },
    });
    myMap.controls.add(locationControl);

    updateMapMarkers();
  });
}

function updateMapMarkers() {
  if (!myMap) return;
  myMap.geoObjects.removeAll();

  const activeAtms = db.baza.filter(
    (a) =>
      db.marshrutIds.includes(a.id) && !db.tarix.some((t) => t.id === a.id),
  );

  activeAtms.forEach((atm) => {
    const placemark = new ymaps.Placemark(
      [atm.lat, atm.lng],
      {
        balloonContentHeader: `<b>#${atm.id}. ${atm.name}</b>`,
        balloonContentBody: `
        <div class="map-balloon-grid">
          <button class="map-btn map-btn-success" onclick="confirmInkassa(${atm.id})">Inkassa</button>
          <button class="map-btn map-btn-danger" onclick="confirmBroken(${atm.id})">Nosoz</button>
          <button class="map-btn map-btn-primary" onclick="openYandexNavi(${atm.lat}, ${atm.lng})">Marshrut</button>
        </div>
      `,
      },
      {
        preset: "islands#icon",
        iconColor: atm.color || "#FF0000",
      },
    );
    myMap.geoObjects.add(placemark);
  });
}

function showOnMap(id) {
  const atm = db.baza.find((a) => a.id === id);
  if (!atm) return;

  switchTab("view-xarita");
  if (myMap) {
    myMap.setCenter([atm.lat, atm.lng], 16, { duration: 400 });
    myMap.geoObjects.each((geoObj) => {
      const coords = geoObj.geometry.getCoordinates();
      if (coords && coords[0] === atm.lat && coords[1] === atm.lng) {
        geoObj.balloon.open();
      }
    });
  }
}

function openYandexNavi(lat, lng) {
  window.open(
    `https://yandex.ru/maps/?rtext=~${lat},${lng}&rtt=auto`,
    "_blank",
  );
}

let userPlacemark = null;

function locateUser() {
  const btnGps = document.getElementById("btn-gps");
  if (btnGps) btnGps.disabled = true;

  if (!navigator.geolocation) {
    alert("Brauzeringizda Geolocation qo'llab-quvvatlanmaydi.");
    if (btnGps) btnGps.disabled = false;
    return;
  }

  navigator.geolocation.getCurrentPosition(
    (position) => {
      const lat = position.coords.latitude;
      const lng = position.coords.longitude;
      const currentLocation = [lat, lng];

      if (myMap) {
        myMap.setCenter(currentLocation, 16, {
          checkZoomRange: true,
          duration: 300,
        });
      }
      showUserMarker(currentLocation);
      if (btnGps) btnGps.disabled = false;
    },
    (error) => {
      alert("Geolokatsiyani aniqlashda xatolik yuz berdi.");
      if (btnGps) btnGps.disabled = false;
    },
    { enableHighAccuracy: true, timeout: 10000, maximumAge: 0 },
  );
}

function showUserMarker(coords) {
  if (!myMap) return;
  if (userPlacemark) {
    userPlacemark.geometry.setCoordinates(coords);
  } else {
    userPlacemark = new ymaps.Placemark(
      coords,
      { hintContent: "Sizning joylashuvinigiz" },
      { preset: "islands#circleDotIcon", iconColor: "#1E88E5" },
    );
    myMap.geoObjects.add(userPlacemark);
  }
}

// ==========================================
// 5-OYNA: TOPSHIRIKLAR MANIQLARI
// ==========================================

// Dropdown menyuni ochish / yopish
function toggleTaskDropdown(event) {
  event.stopPropagation();
  const dropdown = document.getElementById("task-dropdown-menu");
  dropdown.classList.toggle("show");
}

// Ekran bo'sh joyi bosilganda dropdown menyuni avtomatik yopish
window.addEventListener("click", () => {
  const dropdown = document.getElementById("task-dropdown-menu");
  if (dropdown && dropdown.classList.contains("show")) {
    dropdown.classList.remove("show");
  }
});

// Barcha topshiriq statuslarini nolga tushirish (Tozalash)
function resetAllTaskStatuses() {
  // Dropdown menyuni yopish
  const dropdown = document.getElementById("task-dropdown-menu");
  if (dropdown) dropdown.classList.remove("show");

  showConfirm(
    "Haqiqatan ham barcha bankomatlarning topshiriq bajarilish holatlarini tozalamoqchimisiz?",
    function () {
      db.topshiriqData = {}; // Barcha belgilangan statuslarni tozalaymiz
      saveData(); // Baza holatini saqlaymiz va ko'rinishni yangilaymiz (saveDb o'rniga saveData)
    },
  );
}

function renderTopshiriqView() {
  const uncompletedWrapper = document.getElementById(
    "uncompleted-table-wrapper",
  );
  const completedWrapper = document.getElementById("completed-table-wrapper");

  if (!uncompletedWrapper || !completedWrapper) return;

  uncompletedWrapper.innerHTML = "";
  completedWrapper.innerHTML = "";

  // Baza ichidan tanlangan bankomatlarni ajratamiz va ID bo'yicha saralaymiz
  const selectedAtms = db.baza
    .filter((a) => db.topshiriqAtms.includes(a.id))
    .sort((a, b) => a.id - b.id);

  const uncompletedAtms = [];
  const completedAtms = [];

  selectedAtms.forEach((atm) => {
    const atmTasks = db.topshiriqData[atm.id] || {};
    // Kamida 1 ta task '✓' yoki '✕' bo'lsa -> Completed
    const isDone = db.selectedTasks.some(
      (task) => atmTasks[task] === "✓" || atmTasks[task] === "✕",
    );
    if (isDone) {
      completedAtms.push(atm);
    } else {
      uncompletedAtms.push(atm);
    }
  });

  // Statistikalarni yangilash
  const totalEl = document.getElementById("t-total");
  const doneEl = document.getElementById("t-done");
  const leftEl = document.getElementById("t-left");

  if (totalEl) totalEl.innerText = selectedAtms.length;
  if (doneEl) doneEl.innerText = completedAtms.length;
  if (leftEl) leftEl.innerText = uncompletedAtms.length;

  // Ikkinchi parametr isCompleted: bajarilmaganlar uchun false, bajarilganlar uchun true
  uncompletedWrapper.appendChild(buildTaskTable(uncompletedAtms, false));
  completedWrapper.appendChild(buildTaskTable(completedAtms, true));
}

function buildTaskTable(atms, isCompleted = false) {
  const table = document.createElement("table");
  table.className = "task-table";

  // Header
  const thead = document.createElement("thead");
  const headerRow = document.createElement("tr");

  const nameTh = document.createElement("th");
  nameTh.innerText = "Nomi";
  headerRow.appendChild(nameTh);

  db.selectedTasks.forEach((taskName) => {
    const th = document.createElement("th");
    th.innerText = taskName;
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);
  table.appendChild(thead);

  // Body
  const tbody = document.createElement("tbody");
  if (atms.length === 0) {
    const emptyRow = document.createElement("tr");
    const emptyTd = document.createElement("td");
    emptyTd.colSpan = db.selectedTasks.length + 1;
    emptyTd.innerText = "Ma'lumot mavjud emas";
    emptyTd.className = "empty-td";
    emptyRow.appendChild(emptyTd);
    tbody.appendChild(emptyRow);
  } else {
    atms.forEach((atm) => {
      const row = document.createElement("tr");

      // Status xolatini aniqlash:
      const atmTasks = db.topshiriqData[atm.id] || {};
      const taskValues = db.selectedTasks.map((t) => atmTasks[t] || "▢");

      let statusClass = "atm-btn-pending"; // Bajarilmoqda/Neytral
      if (taskValues.some((val) => val === "✕")) {
        statusClass = "atm-btn-failed"; // Qoniqarsiz (kamida 1 ta ✕ bo'lsa)
      } else if (
        db.selectedTasks.length > 0 &&
        taskValues.every((val) => val === "✓")
      ) {
        statusClass = "atm-btn-done"; // Barchasi bajarilgan (barchasi ✓)
      }

      // Bankomat nomi (Text Button)
      const nameTd = document.createElement("td");
      const btn = document.createElement("button");
      btn.className = `btn-atm-name ${statusClass}`;
      btn.innerText = `#${atm.id}. ${atm.name}`;
      btn.onclick = () => openAtmStatusEditModal(atm.id);
      nameTd.appendChild(btn);
      row.appendChild(nameTd);

      // Status ustunlari
      db.selectedTasks.forEach((taskName) => {
        const td = document.createElement("td");
        td.className = "status-cell";
        const val = atmTasks[taskName] || "▢";
        td.innerText = val;

        if (val === "✓") td.classList.add("status-success");
        else if (val === "✕") td.classList.add("status-danger");
        else td.classList.add("status-neutral");

        row.appendChild(td);
      });

      tbody.appendChild(row);
    });
  }
  table.appendChild(tbody);

  // Footer (Faqat bajarilganlar jadvali uchun hamda ma'lumot bo'lganda hisoblanadi)
  if (isCompleted && atms.length > 0) {
    const tfoot = document.createElement("tfoot");
    const footerRow = document.createElement("tr");

    const labelTd = document.createElement("td");
    labelTd.className = "summary-label";
    labelTd.innerText = "Natija:";
    footerRow.appendChild(labelTd);

    db.selectedTasks.forEach((taskName) => {
      let successCount = 0;
      atms.forEach((atm) => {
        const atmTasks = db.topshiriqData[atm.id] || {};
        if (atmTasks[taskName] === "✓") {
          successCount++;
        }
      });

      const countTd = document.createElement("td");
      countTd.className = "summary-count-cell";
      countTd.innerHTML = `<span class="summary-count">${successCount}</span> <span class="summary-check">✓</span>`;
      footerRow.appendChild(countTd);
    });

    tfoot.appendChild(footerRow);
    table.appendChild(tfoot);
  }

  return table;
}

// 1. Bankomatlar tanlash Modali (Topshiriq)
function openTopshiriqAtmModal() {
  const listEl = document.getElementById("topshiriq-atm-checkbox-list");
  listEl.innerHTML = "";

  if (db.baza.length === 0) {
    alert("Avval Baza bo'limida Excel fayl import qiling!");
    return;
  }

  const groups = {};
  db.baza.forEach((atm) => {
    const groupName = atm.group || "Umumiy";
    if (!groups[groupName]) groups[groupName] = [];
    groups[groupName].push(atm);
  });

  Object.keys(groups).forEach((groupName, index) => {
    const groupAtms = groups[groupName];
    const allChecked = groupAtms.every((atm) =>
      db.topshiriqAtms.includes(atm.id),
    );

    const groupWrapper = document.createElement("div");
    groupWrapper.className = "group-wrapper";

    const groupHeader = document.createElement("div");
    groupHeader.className = "group-header";
    groupHeader.innerHTML = `
      <div class="group-title-area">
        <input type="checkbox" class="top-group-checkbox" data-group="${groupName}" ${allChecked ? "checked" : ""} onchange="toggleTopGroupCheck(this, '${groupName}')">
        <strong onclick="toggleGroupAccordion('top-group-items-${index}', this.closest('.group-wrapper'))" style="cursor:pointer;">📂 ${groupName} (${groupAtms.length} ta)</strong>
      </div>
      <span class="accordion-icon" onclick="toggleGroupAccordion('top-group-items-${index}', this.closest('.group-wrapper'))">▼</span>
    `;

    const itemsContainer = document.createElement("div");
    itemsContainer.id = `top-group-items-${index}`;
    itemsContainer.className = "group-items hidden";

    groupAtms.forEach((atm) => {
      const isChecked = db.topshiriqAtms.includes(atm.id) ? "checked" : "";
      const item = document.createElement("label");
      item.className = "checkbox-item";
      item.innerHTML = `
        <input type="checkbox" class="top-atm-checkbox" data-group="${groupName}" value="${atm.id}" ${isChecked} onchange="updateTopGroupCheckboxState('${groupName}')">
        <span class="color-dot" style="background:${atm.color}"></span>
        <span>#${atm.id}. ${atm.name}</span>
      `;
      itemsContainer.appendChild(item);
    });

    groupWrapper.appendChild(groupHeader);
    groupWrapper.appendChild(itemsContainer);
    listEl.appendChild(groupWrapper);
  });

  document.getElementById("topshiriq-atm-modal").classList.remove("hidden");
}

function toggleTopGroupCheck(groupMasterCb, groupName) {
  const isChecked = groupMasterCb.checked;
  const atmCheckboxes = document.querySelectorAll(
    `.top-atm-checkbox[data-group="${groupName}"]`,
  );
  atmCheckboxes.forEach((cb) => {
    cb.checked = isChecked;
  });
}

function updateTopGroupCheckboxState(groupName) {
  const masterCb = document.querySelector(
    `.top-group-checkbox[data-group="${groupName}"]`,
  );
  const atmCheckboxes = document.querySelectorAll(
    `.top-atm-checkbox[data-group="${groupName}"]`,
  );
  const allChecked = Array.from(atmCheckboxes).every((cb) => cb.checked);
  if (masterCb) masterCb.checked = allChecked;
}

function closeTopshiriqAtmModal() {
  document.getElementById("topshiriq-atm-modal").classList.add("hidden");
}

function saveTopshiriqAtms() {
  const checkboxes = document.querySelectorAll(".top-atm-checkbox");
  const selected = [];
  checkboxes.forEach((cb) => {
    if (cb.checked) selected.push(parseInt(cb.value));
  });
  db.topshiriqAtms = selected;
  saveData();
  closeTopshiriqAtmModal();
}

// 2. Topshiriqlar listini tanlash Modali
function openTopshiriqTaskModal() {
  const listEl = document.getElementById("topshiriq-task-checkbox-list");
  listEl.innerHTML = "";

  DEFAULT_TASKS_LIST.forEach((taskName, idx) => {
    const isChecked = db.selectedTasks.includes(taskName) ? "checked" : "";
    const item = document.createElement("label");
    item.className = "checkbox-item";
    item.innerHTML = `
      <input type="checkbox" class="task-checkbox" value="${taskName}" ${isChecked}>
      <span>${taskName}</span>
    `;
    listEl.appendChild(item);
  });

  document.getElementById("topshiriq-task-modal").classList.remove("hidden");
}

function closeTopshiriqTaskModal() {
  document.getElementById("topshiriq-task-modal").classList.add("hidden");
}

function saveTopshiriqTasks() {
  const checkboxes = document.querySelectorAll(".task-checkbox");
  const selected = [];
  checkboxes.forEach((cb) => {
    if (cb.checked) selected.push(cb.value);
  });
  db.selectedTasks = selected;
  saveData();
  closeTopshiriqTaskModal();
}

// 3. Bankomat topshiriqlari holatini o'zgartirish (Radio Buttons)
function openAtmStatusEditModal(atmId) {
  currentEditingAtmId = atmId;
  const atm = db.baza.find((a) => a.id === atmId);
  if (!atm) return;

  document.getElementById("modal-atm-title").innerText =
    `#${atm.id}. ${atm.name}`;
  const listEl = document.getElementById("atm-tasks-status-list");
  listEl.innerHTML = "";

  const atmTasks = db.topshiriqData[atmId] || {};

  db.selectedTasks.forEach((taskName, idx) => {
    const currentVal = atmTasks[taskName] || "▢";

    const item = document.createElement("div");
    item.className = "task-radio-item";
    item.innerHTML = `
      <span class="task-name-label">${taskName}</span>
      <div class="radio-options">
        <label class="radio-label">
          <input type="radio" name="task_radio_${idx}" value="✓" ${currentVal === "✓" ? "checked" : ""}>
          <span class="radio-custom status-success">✓</span>
        </label>
        <label class="radio-label">
          <input type="radio" name="task_radio_${idx}" value="✕" ${currentVal === "✕" ? "checked" : ""}>
          <span class="radio-custom status-danger">✕</span>
        </label>
        <label class="radio-label">
          <input type="radio" name="task_radio_${idx}" value="▢" ${currentVal === "▢" ? "checked" : ""}>
          <span class="radio-custom status-neutral">▢</span>
        </label>
      </div>
    `;
    listEl.appendChild(item);
  });

  document.getElementById("atm-status-edit-modal").classList.remove("hidden");
}

function closeAtmStatusEditModal() {
  document.getElementById("atm-status-edit-modal").classList.add("hidden");
  currentEditingAtmId = null;
}

function saveAtmTaskStatus() {
  if (!currentEditingAtmId) return;

  if (!db.topshiriqData[currentEditingAtmId]) {
    db.topshiriqData[currentEditingAtmId] = {};
  }

  db.selectedTasks.forEach((taskName, idx) => {
    const radios = document.getElementsByName(`task_radio_${idx}`);
    let selectedVal = "▢";
    radios.forEach((r) => {
      if (r.checked) selectedVal = r.value;
    });
    db.topshiriqData[currentEditingAtmId][taskName] = selectedVal;
  });

  saveData();
  closeAtmStatusEditModal();
}
