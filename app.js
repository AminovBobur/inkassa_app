// 1. Ruxsat berilgan Telegram foydalanuvchilarining ID lari
const ALLOWED_TELEGRAM_IDS = [
  1347548152, // Bobur
  5013974621, // Jurabek Akam
];

// 2. Telegram WebApp obyektini tekshirish
const tg = window.Telegram ? window.Telegram.WebApp : null;

function checkAccess() {
  // // Agar foydalanuvchi Telegram ichidan kirmagan bo'lsa
  // if (!tg || !tg.initDataUnsafe || !tg.initDataUnsafe.user) {
  //   document.body.innerHTML = `
  //           <div style="display:flex; height:100vh; align-items:center; justify-content:center; text-align:center; padding:20px; font-family:sans-serif;">
  //               <h2 style="color:#dc2626;">❌ Ruxsat berilmagan! <br><br> Bu ilovadan faqat Telegram Bot orqali foydalanish mumkin.</h2>
  //           </div>
  //       `;
  //   return false;
  // }

  // const currentUserId = tg.initDataUnsafe.user.id;

  // // Agar foydalanuvchi ID si ruxsat berilganlar ro'yxatida bo'lmasa
  // if (!ALLOWED_TELEGRAM_IDS.includes(currentUserId)) {
  //   document.body.innerHTML = `
  //           <div style="display:flex; height:100vh; align-items:center; justify-content:center; text-align:center; padding:20px; font-family:sans-serif;">
  //               <h2 style="color:#dc2626;">🚫 Kirish taqiqlangan! <br><br> Sizning Telegram ID (${currentUserId}) ushbu tizimga ulangan emas.</h2>
  //           </div>
  //       `;
  //   return false;
  // }

  return true;
}

// Sahifa yuklanishidan oldin tekshiramiz
if (!checkAccess()) {
  throw new Error("Ruxsatsiz kirishga urunish to'xtatildi.");
}

let db = {
  baza: JSON.parse(localStorage.getItem("inkassa_baza")) || [],
  marshrutIds: JSON.parse(localStorage.getItem("inkassa_marshrut")) || [],
  tarix: JSON.parse(localStorage.getItem("inkassa_tarix")) || [],
};

let myMap = null;
let locationControl = null;
let pendingConfirmAction = null;

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
  if (myMap) updateMapMarkers();
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
            group: item.group || "Umumiy", // Agar guruh ko'rsatilmagan bo'lsa 'Umumiy' bo'ladi
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
        <button class="btn btn-success" onclick="confirmInkassa(${atm.id})">Inkassa</button>
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

  // Bankomatlarni guruhlar bo'yicha ajratamiz
  const groups = {};
  db.baza.forEach((atm) => {
    const groupName = atm.group || "Umumiy";
    if (!groups[groupName]) groups[groupName] = [];
    groups[groupName].push(atm);
  });

  // Har bir guruh uchun blok yaratamiz
  Object.keys(groups).forEach((groupName, index) => {
    const groupAtms = groups[groupName];

    // Guruhdagi barcha bankomatlar belgilanganmi?
    const allChecked = groupAtms.every((atm) =>
      db.marshrutIds.includes(atm.id),
    );

    // openSelectModal funksiyasi ichida:
    const groupWrapper = document.createElement("div");
    groupWrapper.className = "group-wrapper";

    // Guruh sarlavhasi va Bosh Checkbox
    const groupHeader = document.createElement("div");
    groupHeader.className = "group-header";

    groupHeader.innerHTML = `
            <div class="group-title-area">
              <input type="checkbox" class="group-checkbox" data-group="${groupName}" ${allChecked ? "checked" : ""} onchange="toggleGroupCheck(this, '${groupName}')">
              <strong onclick="toggleGroupAccordion('group-items-${index}', this.closest('.group-wrapper'))" style="cursor:pointer;">📂 ${groupName} (${groupAtms.length} ta)</strong>
            </div>
            <span class="accordion-icon" onclick="toggleGroupAccordion('group-items-${index}', this.closest('.group-wrapper'))">▼</span>
        `;

    // Guruh ichidagi bankomatlar ro'yxati
    const itemsContainer = document.createElement("div");
    itemsContainer.id = `group-items-${index}`;
    itemsContainer.className = "group-items hidden"; // Boshida yig'ilgan turadi

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

// GURUHNIBOSH CHECKBOX'I BOSILGANDA ICHIDAGI BARCHASINI BELGILASH / OCHIRISH
function toggleGroupCheck(groupMasterCb, groupName) {
  const isChecked = groupMasterCb.checked;
  const atmCheckboxes = document.querySelectorAll(
    `.atm-checkbox[data-group="${groupName}"]`,
  );

  atmCheckboxes.forEach((cb) => {
    cb.checked = isChecked;
  });
}

// ICHKI CHECKBOX O'ZGARSA, GURUHNINKINI TEKSHIRISH
function updateGroupCheckboxState(groupName) {
  const masterCb = document.querySelector(
    `.group-checkbox[data-group="${groupName}"]`,
  );
  const atmCheckboxes = document.querySelectorAll(
    `.atm-checkbox[data-group="${groupName}"]`,
  );

  const allChecked = Array.from(atmCheckboxes).every((cb) => cb.checked);
  masterCb.checked = allChecked;
}

// AKKORDEON: YAGONA OCHILISH VA DYNAMIK BALANDLIK MANTIQI
function toggleGroupAccordion(containerId, targetWrapper) {
  const targetContainer = document.getElementById(containerId);
  const modalContent = document.querySelector("#select-modal .modal-content");
  const isCurrentlyHidden = targetContainer.classList.contains("hidden");

  // 1. Barcha guruhlarni va ularning ochiq holatlarini yopamiz
  document
    .querySelectorAll(".group-items")
    .forEach((el) => el.classList.add("hidden"));
  document
    .querySelectorAll(".group-wrapper")
    .forEach((el) => el.classList.remove("open"));

  // 2. Agar bosilgan guruh yopiq bo'lgan bo'lsa, uni ochamiz
  if (isCurrentlyHidden) {
    targetContainer.classList.remove("hidden");
    targetWrapper.classList.add("open");
    modalContent.classList.add("expanded"); // Modal balandligini 80vh qiladi
  } else {
    // Agar barcha guruhlar yopilsa, modalni compact (ixcham) holatga qaytaramiz
    modalContent.classList.remove("expanded");
  }
}

// MODAL YOPILGANDA BALANDLIKNI DASTLABKI HOLATGA QAYTARISH
function closeSelectModal() {
  const modalContent = document.querySelector("#select-modal .modal-content");
  if (modalContent) modalContent.classList.remove("expanded");
  document.getElementById("select-modal").classList.add("hidden");
}

function saveSelectedMarshrut() {
  // const checkboxes = document.querySelectorAll('#select-checkbox-list input[type="checkbox"]');

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

// 3-OYNA: TARIX RENDER
function renderTarixView() {
  const listEl = document.getElementById("tarix-list");
  const countEl = document.getElementById("tarix-total-count");

  countEl.innerText = db.tarix.length;
  listEl.innerHTML = "";

  const reversedTarix = [...db.tarix].reverse();

  reversedTarix.forEach((item) => {
    const el = document.createElement("div");
    el.className = "tarix-item";
    el.innerText = `${item.time} — #${item.id}. ${item.name}`;
    listEl.appendChild(el);
  });
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

// 4-OYNA: YANDEX MAPS VA GEOLOKATSIYA
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
        <div style="display:flex; gap:8px; margin-top:8px;">
          <button style="padding:6px 10px; background:#16a34a; color:white; border:none; border-radius:4px; font-weight:bold; cursor:pointer;" onclick="confirmInkassa(${atm.id})">Inkassa qilish</button>
          <button style="padding:6px 10px; background:#2563eb; color:white; border:none; border-radius:4px; font-weight:bold; cursor:pointer;" onclick="openYandexNavi(${atm.lat}, ${atm.lng})">Marshrut</button>
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

// Foydalanuvchi joylashuvini xotirada saqlash va icon obyektini ushlab turish uchun o'zgaruvchilar
let userLocation = null;
let userPlacemark = null;

function locateUser() {
  // 1. Agar joylashuv avval aniqlangan bo'lsa, qayta ruxsat so'ramaymiz
  if (userLocation) {
    myMap.setCenter(userLocation, 16, { checkZoomRange: true, duration: 300 });
    showUserMarker(userLocation);
    return;
  }

  // 2. Birinchi marta bosilganda brauzerdan ruxsat so'raymiz
  if (navigator.geolocation) {
    navigator.geolocation.getCurrentPosition(
      (position) => {
        const lat = position.coords.latitude;
        const lng = position.coords.longitude;
        userLocation = [lat, lng]; // Xotiraga saqlaymiz

        // Xaritaning markazini suramiz
        myMap.setCenter(userLocation, 16, {
          checkZoomRange: true,
          duration: 300,
        });

        // Icon (Marker) qo'shamiz
        showUserMarker(userLocation);
      },
      (error) => {
        alert(
          "GPS geolokatsiyani aniqlab bo'lmadi. Telefon geolokatsiyasi yoqilganini tekshiring.",
        );
        console.error(error);
      },
      { enableHighAccuracy: true },
    );
  } else {
    alert("Brauzeringizda Geolocation qo'llab-quvvatlanmaydi.");
  }
}

// Xaritaga foydalanuvchi iconini chiqaruvchi funksiya
function showUserMarker(coords) {
  // Agar icon avval yaratilgan bo'lsa, shunchaki o'rnini yangilaymiz
  if (userPlacemark) {
    userPlacemark.geometry.setCoordinates(coords);
  } else {
    // Yangi ajralib turuvchi ko'k icon yaratamiz
    userPlacemark = new ymaps.Placemark(
      coords,
      { hintContent: "Sizning joylashuvinigiz" },
      {
        preset: "islands#circleDotIcon", // Nuqtali doira shaklidagi standart icon
        iconColor: "#1E88E5", // Yorqin ko'k rang
      },
    );
    myMap.geoObjects.add(userPlacemark);
  }
}

// Telegram oynasini to'liq ekranga yoyish va tayyor holatga keltirish
if (window.Telegram && window.Telegram.WebApp) {
  window.Telegram.WebApp.ready();
  window.Telegram.WebApp.expand(); // Ilovani to'liq ekranda ochish
}
