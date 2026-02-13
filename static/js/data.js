$(document).ready(function () {
  setInterval(updateClock, 1000);
  setInterval(getData, 10000);
  updateClock();
  getData();
});

function getData() {
  const zone = document.body.dataset.zone;

  if (zone === "all") {
    $.get("/api/all", function (response) {
      if (!response || response.hijau?.offline || response.merah?.offline) {
        showOfflineAlert();
        return;
      }

      $("#offline-alert").hide();
      renderAllData(response.hijau, response.merah);
    }).fail(showOfflineAlert);
  } else {
    const endpoint = zone === "merah" ? "/api/merah" : "/api/data";
    $.get(endpoint, function (response) {
      if (response.offline) {
        showOfflineAlert();
        return;
      }

      $("#offline-alert").hide();
      $("#totalin").text(response.totalin);
      $("#totalout").text(response.totalout);
      $("#totalcur").text(response.totalcur);

      let html = '';
      response.data.forEach(dept => {
        html += `
          <tr>
            <td class="text-left"><strong>${dept.dept}</strong></td>
            <td><strong>${dept.in}</strong></td>
            <td><strong>${dept.out}</strong></td>
            <td><strong>${dept.cur}</strong></td>
          </tr>
        `;
      });
      $("#dept-table").html(html);
    }).fail(showOfflineAlert);
  }
}

function renderAllData(hijau, merah) {
  $("#totalin").text(hijau.totalin ?? 0);
  $("#totalout").text(hijau.totalout ?? 0);
  $("#totalcur").text(hijau.totalcur ?? 0);

  const dataHijau = hijau.data || [];
  const dataMerah = merah.data || [];

  const deptMap = {};

  dataHijau.forEach(d => {
    deptMap[d.dept] = {
      dept: d.dept,
      hijau_in: d.in || 0,
      hijau_out: d.out || 0,
      hijau_cur: d.cur || 0,
      merah_in: 0,
      merah_out: 0,
      merah_cur: 0
    };
  });

  dataMerah.forEach(d => {
    if (!deptMap[d.dept]) {
      deptMap[d.dept] = {
        dept: d.dept,
        hijau_in: 0,
        hijau_out: 0,
        hijau_cur: 0,
        merah_in: d.in || 0,
        merah_out: d.out || 0,
        merah_cur: d.cur || 0
      };
    } else {
      deptMap[d.dept].merah_in = d.in || 0;
      deptMap[d.dept].merah_out = d.out || 0;
      deptMap[d.dept].merah_cur = d.cur || 0;
    }
  });

  let html = '';
  Object.values(deptMap).forEach(dept => {
    html += `
      <tr>
        <td class="text-left"><strong>${dept.dept}</strong></td>
        <td><strong>${dept.hijau_in}</strong></td>
        <td><strong>${dept.hijau_out}</strong></td>
        <td><strong>${dept.merah_in}</strong></td>
        <td><strong>${dept.merah_out}</strong></td>
        <td><strong>${dept.hijau_cur}</strong></td>
        <td><strong>${dept.merah_cur}</strong></td>
      </tr>
    `;
  });

  $("#dept-table").html(html);

  $("#dept-table-header").html(`
    <th class="text-left"><strong>DEPARTEMEN</strong></th>
    <th><strong>TERBATAS<br>IN (${hijau.totalin ?? 0})</strong></th>
    <th><strong>TERBATAS<br>OUT (${hijau.totalout ?? 0})</strong></th>
    <th><strong>TERLARANG<br>IN (${merah.totalin ?? 0})</strong></th>
    <th><strong>TERLARANG<br>OUT (${merah.totalout ?? 0})</strong></th>
    <th><strong>TERBATAS<br>CURRENT (${hijau.totalcur ?? 0})</strong></th>
    <th><strong>TERLARANG<br>CURRENT (${merah.totalcur ?? 0})</strong></th>
  `);
}

function showOfflineAlert() {
  $("#offline-alert").show();
  $("#totalin").text("-");
  $("#totalout").text("-");
  $("#totalcur").text("-");
  $("#dept-table").html(`<tr><td colspan="7" class="text-center text-danger">Data tidak tersedia</td></tr>`);
}

function updateClock() {
  const now = new Date();
  const offset = (now.getTimezoneOffset() === 0) ? 7 * 3600000 : 0;
  now.setTime(now.getTime() + offset);

  const jam = now.getHours().toString().padStart(2, '0');
  const menit = now.getMinutes().toString().padStart(2, '0');
  const detik = now.getSeconds().toString().padStart(2, '0');

  const hariArray = ["Minggu,", "Senin,", "Selasa,", "Rabu,", "Kamis,", "Jum'at,", "Sabtu,"];
  const bulanArray = ["Januari", "Februari", "Maret", "April", "Mei", "Juni",
                      "Juli", "Agustus", "September", "Oktober", "Nopember", "Desember"];

  $("#jam").text(jam);
  $("#menit").text(menit);
  $("#detik").text(detik);
  $("#tanggalwaktu").text(`${hariArray[now.getDay()]} ${now.getDate()} ${bulanArray[now.getMonth()]} ${now.getFullYear()}`);
}
