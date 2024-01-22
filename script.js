let map;
let marker;
let coordinates = document.getElementById('coordinates');
let addingLocationEnabled = false;
let markers = [];

map = L.map("map").setView([39.9334, 32.8597], 6);

L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
  attribution: '© OpenStreetMap contributors'
}).addTo(map);

L.Control.geocoder().addTo(map);

let map_prop = document.getElementById('map-prop');
let right_side = document.getElementById('right-side');
let buttons_container = document.getElementById('button-properties');

document.getElementById("add-location-btn").addEventListener("click", function () {
  if (!addingLocationEnabled) {
    addingLocationEnabled = true;
    map.on("click", function (e) {
      if (addingLocationEnabled) {
        map_prop.style.width = '80vw';
        right_side.style.width = '20vw';
        buttons_container.style.width = '18vw';
        let lat = e.latlng.lat.toFixed(4);
        let lng = e.latlng.lng.toFixed(4);
        marker = L.marker(e.latlng).addTo(map);
        markers.push(marker);
        showInfoForm(e.latlng);
        coordinates.value = `Enlem: ${lat}, Boylam: ${lng}`;
      }
    });
  }
});

function showInfoForm(latlng) {
  document.getElementById("info-form").classList.remove("hidden");
  document.getElementById("location-form").reset();
}

function saveLocation() {
  if (addingLocationEnabled && marker) {
    addingLocationEnabled = false;
    const name = document.getElementById("name").value;
    const surname = document.getElementById("surname").value;
    const company = document.getElementById("company").value;
    const docType = document.getElementById("document-type").value;
    const stock = document.getElementById("stock").value;
    const exit = document.getElementById("exit").value;
    const unit = document.getElementById("unit").value;
    const date = document.getElementById("date").value;

    marker.bindPopup(
      `<b>Ad:</b> ${name} ${surname}<br>` +
      `<b>Firma:</b> ${company}<br>` +
      `<b>Evrak Tipi:</b> ${docType}<br>` +
      `<b>Stok:</b> ${stock}<br>` +
      `<b>Çıkış:</b> ${exit}<br>` +
      `<b>Birim:</b> ${unit}<br>` +
      `<b>Tarih:</b> ${date}`
    ).openPopup();

    document.getElementById("info-form").classList.add("hidden");
  }
}

document.getElementById("save-btn").addEventListener("click", saveLocation);
document.getElementById("export-btn").addEventListener("click", exportToExcel);

function exportToExcel() {
  const data = markers.map(marker => {
    // Check if the marker has a popup
    if (marker.getPopup()) {
      const popupContent = marker.getPopup().getContent();

      if (!exportToExcel.columnNames) {
        exportToExcel.columnNames = popupContent
          .split('<br>')
          .map(line => line.replace(/<b>|<\/b>/g, '').trim().split(': ')[0]);
        exportToExcel.columnNames.push('Latitude', 'Longitude');
      }

      const values = popupContent
        .split('<br>')
        .map(line => line.replace(/<b>|<\/b>/g, '').trim().split(': ')[1]);

      const [name, surname] = values[0].split(' ');

      const locationData = {
        ...Object.fromEntries(exportToExcel.columnNames.slice(0, -2).map((name, index) => [name, values[index]])),
        'Latitude': marker.getLatLng().lat,
        'Longitude': marker.getLatLng().lng,
        'İsim': name,
        'Soyisim': surname,
      };

      return locationData;
    } else {
      // Handle the case where the marker doesn't have a popup
      console.warn('Marker does not have a popup:', marker);
      return null;
    }
  });

  // Filter out markers without a popup
  const filteredData = data.filter(entry => entry !== null);

  const ws = XLSX.utils.json_to_sheet(filteredData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
  XLSX.writeFile(wb, 'yardim_bilgileri.xlsx');
}
