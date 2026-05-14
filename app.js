let data = [];
let imagesData = [];
let filteredData = [];
let imageFiles = {};
let imageMap = {};

let map;
let markerClusterGroup;

const excelFile = document.getElementById("excelFile");
const imageFolder = document.getElementById("imageFolder");

const searchInput = document.getElementById("searchInput");

const prefixFilter = document.getElementById("prefixFilter");
const cropFilter = document.getElementById("cropFilter");
const resultFilter = document.getElementById("resultFilter");
const locationFilter = document.getElementById("locationFilter");

excelFile.addEventListener("change", loadExcel);
imageFolder.addEventListener("change", loadImages);

searchInput.addEventListener("input", applyFilters);

document.querySelectorAll(".dropdown-button").forEach(button => {
  button.addEventListener("click", event => {
    event.stopPropagation();

    const targetId = button.dataset.target;
    const menu = document.getElementById(targetId);

    document.querySelectorAll(".dropdown-menu").forEach(otherMenu => {
      if (otherMenu !== menu) {
        otherMenu.classList.remove("open");
      }
    });

    menu.classList.toggle("open");
  });
});

document.addEventListener("click", event => {
  if (!event.target.closest(".dropdown-filter")) {
    document.querySelectorAll(".dropdown-menu").forEach(menu => {
      menu.classList.remove("open");
    });
  }
});

function normalizePath(path) {
  return String(path || "").replaceAll("\\", "/");
}

function getFileName(path) {
  const normalized = normalizePath(path);
  return normalized.split("/").pop();
}

function getNamePrefix(row) {
  return String(row.name || "").substring(0, 6);
}

function getCheckedValues(container) {
  return Array.from(
    container.querySelectorAll("input[type='checkbox']:checked")
  ).map(cb => cb.value);
}

function updateDropdownButton(container, label) {
  const selectedValues = getCheckedValues(container);
  const button = document.querySelector(`[data-target="${container.id}"]`);

  if (selectedValues.length === 0) {
    button.textContent = `選擇 ${label}`;
  } else if (selectedValues.length === 1) {
    button.textContent = selectedValues[0];
  } else {
    button.textContent = `${label}：已選 ${selectedValues.length} 項`;
  }
}

function createCheckboxes(container, values, name, label) {
  container.innerHTML = "";

  values.forEach(value => {
    const wrapper = document.createElement("div");
    wrapper.className = "checkbox-item";

    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.value = value;
    checkbox.name = name;

    checkbox.addEventListener("click", event => {
      event.stopPropagation();
    });

    checkbox.addEventListener("change", () => {
      updateDropdownButton(container, label);
      applyFilters();
    });

    const checkboxLabel = document.createElement("label");
    checkboxLabel.textContent = value;

    wrapper.appendChild(checkbox);
    wrapper.appendChild(checkboxLabel);

    container.appendChild(wrapper);
  });

  updateDropdownButton(container, label);
}

function loadImages(event) {
  imageFiles = {};

  for (const file of event.target.files) {
    imageFiles[file.name] = URL.createObjectURL(file);

    if (file.webkitRelativePath) {
      imageFiles[
        normalizePath(file.webkitRelativePath)
      ] = URL.createObjectURL(file);
    }
  }
}

function loadExcel(event) {
  const file = event.target.files[0];

  const reader = new FileReader();

  reader.onload = function(e) {
    const rawData = new Uint8Array(e.target.result);

    const workbook = XLSX.read(rawData, { type: "array" });

    const samplesSheet = workbook.Sheets["Samples"];
    const imagesSheet = workbook.Sheets["Images"];

    data = XLSX.utils.sheet_to_json(samplesSheet);

    imagesData = imagesSheet
      ? XLSX.utils.sheet_to_json(imagesSheet)
      : [];

    buildImageMap();

    setupFilters();

    applyFilters();
  };

  reader.readAsArrayBuffer(file);
}

function buildImageMap() {
  imageMap = {};

  imagesData.forEach(row => {
    if (row.sample_id && row.image_path) {
      imageMap[row.sample_id] = row.image_path;
    }
  });
}

function setupFilters() {
  const prefixes =
    [...new Set(data.map(row => getNamePrefix(row)).filter(Boolean))];

  const crops =
    [...new Set(data.map(row => row.crop).filter(Boolean))];

  const results =
    [...new Set(data.map(row => row.result).filter(Boolean))];

  const locations =
    [...new Set(data.map(row => row.location).filter(Boolean))];

  createCheckboxes(prefixFilter, prefixes, "prefix", "Prefix");
  createCheckboxes(cropFilter, crops, "crop", "Crop");
  createCheckboxes(resultFilter, results, "result", "Result");
  createCheckboxes(locationFilter, locations, "location", "Location");
}

function applyFilters() {
  const keyword = searchInput.value.toLowerCase();

  const selectedPrefixes = getCheckedValues(prefixFilter);

  const selectedCrops = getCheckedValues(cropFilter);

  const selectedResults = getCheckedValues(resultFilter);

  const selectedLocations = getCheckedValues(locationFilter);

  filteredData = data.filter(row => {
    const matchKeyword =
      Object.values(row)
        .join(" ")
        .toLowerCase()
        .includes(keyword);

    const matchPrefix =
      selectedPrefixes.length === 0 ||
      selectedPrefixes.includes(getNamePrefix(row));

    const matchCrop =
      selectedCrops.length === 0 ||
      selectedCrops.includes(row.crop);

    const matchResult =
      selectedResults.length === 0 ||
      selectedResults.includes(row.result);

    const matchLocation =
      selectedLocations.length === 0 ||
      selectedLocations.includes(row.location);

    return (
      matchKeyword &&
      matchPrefix &&
      matchCrop &&
      matchResult &&
      matchLocation
    );
  });

  renderTable();
  updateMap();
}

function renderTable() {
  const tbody = document.querySelector("#dataTable tbody");

  tbody.innerHTML = "";

  filteredData.forEach(row => {
    const tr = document.createElement("tr");

    tr.innerHTML = `
      <td>${row.name || ""}</td>
      <td>${row.crop || ""}</td>
      <td>${row.location || ""}</td>
      <td>${row.date || ""}</td>
      <td>${row.result || ""}</td>
      <td>${row.latitude || ""}</td>
      <td>${row.longitude || ""}</td>
    `;

    tr.addEventListener("click", () => showDetail(row));

    tbody.appendChild(tr);
  });
}

function showDetail(row) {
  const detailInfo = document.getElementById("detailInfo");

  const imagePath = imageMap[row.id];

  const imageFileName = getFileName(imagePath);

  const normalizedImagePath = normalizePath(imagePath);

  detailInfo.innerHTML = `
    <p><b>ID:</b> ${row.id || ""}</p>
    <p><b>Name:</b> ${row.name || ""}</p>
    <p><b>Crop:</b> ${row.crop || ""}</p>
    <p><b>Location:</b> ${row.location || ""}</p>
    <p><b>Date:</b> ${row.date || ""}</p>
    <p><b>Result:</b> ${row.result || ""}</p>
    <p><b>Latitude:</b> ${row.latitude || ""}</p>
    <p><b>Longitude:</b> ${row.longitude || ""}</p>
    <p><b>Note:</b> ${row.note || ""}</p>
    <p><b>Image:</b> ${imagePath || "沒有圖片路徑"}</p>
  `;

  const image = document.getElementById("dataImage");

  if (imagePath && imageFiles[imageFileName]) {
    image.src = imageFiles[imageFileName];
  } else if (imagePath && imageFiles[normalizedImagePath]) {
    image.src = imageFiles[normalizedImagePath];
  } else {
    image.src = "";
  }
}

function initMap() {
  if (map) {
    return;
  }

  map = L.map("map").setView([23.7, 120.9], 7);

  L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
    maxZoom: 19
  }).addTo(map);

  markerClusterGroup = L.markerClusterGroup({
    showCoverageOnHover: false,
    spiderfyOnMaxZoom: true,
    disableClusteringAtZoom: 18
  });

  map.addLayer(markerClusterGroup);
}

function updateMap() {
  initMap();

  markerClusterGroup.clearLayers();

  const validRows = filteredData.filter(row => {
    return row.latitude && row.longitude;
  });

  const markers = [];

  validRows.forEach(row => {
    const lat = parseFloat(row.latitude);
    const lng = parseFloat(row.longitude);

    if (isNaN(lat) || isNaN(lng)) {
      return;
    }

    const marker = L.marker([lat, lng]);

    marker.bindPopup(`
      <b>${row.name || ""}</b><br>
      Crop: ${row.crop || ""}<br>
      Location: ${row.location || ""}<br>
      Date: ${row.date || ""}<br>
      Result: ${row.result || ""}
    `);

    marker.on("click", () => showDetail(row));

    markerClusterGroup.addLayer(marker);
    markers.push(marker);
  });

  if (markers.length > 0) {
    const group = L.featureGroup(markers);
    map.fitBounds(group.getBounds(), {
      padding: [30, 30]
    });
  }
}