let data = [];
let imagesData = [];
let filteredData = [];
let imageFiles = {};
let imageMap = {};

const excelFile = document.getElementById("excelFile");
const imageFolder = document.getElementById("imageFolder");

const searchInput = document.getElementById("searchInput");

const cropFilter = document.getElementById("cropFilter");
const resultFilter = document.getElementById("resultFilter");
const locationFilter = document.getElementById("locationFilter");

excelFile.addEventListener("change", loadExcel);
imageFolder.addEventListener("change", loadImages);

searchInput.addEventListener("input", applyFilters);

cropFilter.addEventListener("change", applyFilters);
resultFilter.addEventListener("change", applyFilters);
locationFilter.addEventListener("change", applyFilters);

function normalizePath(path) {
  return String(path || "").replaceAll("\\", "/");
}

function getFileName(path) {
  const normalized = normalizePath(path);
  return normalized.split("/").pop();
}

function getSelectedValues(select) {
  return Array.from(select.selectedOptions).map(option => option.value);
}

function loadImages(event) {
  imageFiles = {};

  for (const file of event.target.files) {

    imageFiles[file.name] = URL.createObjectURL(file);

    if (file.webkitRelativePath) {
      imageFiles[normalizePath(file.webkitRelativePath)] =
        URL.createObjectURL(file);
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

  const crops =
    [...new Set(data.map(row => row.crop).filter(Boolean))];

  const results =
    [...new Set(data.map(row => row.result).filter(Boolean))];

  const locations =
    [...new Set(data.map(row => row.location).filter(Boolean))];

  cropFilter.innerHTML = "";
  resultFilter.innerHTML = "";
  locationFilter.innerHTML = "";

  crops.forEach(crop => {

    const option = document.createElement("option");

    option.value = crop;
    option.textContent = crop;

    cropFilter.appendChild(option);

  });

  results.forEach(result => {

    const option = document.createElement("option");

    option.value = result;
    option.textContent = result;

    resultFilter.appendChild(option);

  });

  locations.forEach(location => {

    const option = document.createElement("option");

    option.value = location;
    option.textContent = location;

    locationFilter.appendChild(option);

  });

}

function applyFilters() {

  const keyword = searchInput.value.toLowerCase();

  const selectedCrops = getSelectedValues(cropFilter);

  const selectedResults = getSelectedValues(resultFilter);

  const selectedLocations = getSelectedValues(locationFilter);

  filteredData = data.filter(row => {

    const matchKeyword =
      Object.values(row)
        .join(" ")
        .toLowerCase()
        .includes(keyword);

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
      matchCrop &&
      matchResult &&
      matchLocation
    );

  });

  renderTable();

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