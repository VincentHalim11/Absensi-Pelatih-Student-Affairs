  // const FOLDER_SELFIE_ID = "1bxo4SfVu_0Pa2tP2M2g98ZSLgegmRgk9";
  const FOLDER_SELFIE_ID = "1HgA0TfJHqlsMXKQdCuHWKloN4-Og0SXB";
  const AUTO_CLOCK_OUT_TIME = "21:00";
  const AUTO_CLOCK_OUT_NOTE = "Clock Out otomatis oleh sistem karena pelatih tidak melakukan absen pulang sebelum pergantian hari.";

  const COL = {
  TANGGAL: 1,
  ID: 2,
  NAMA: 3,
  UKM: 4,
  JAM_MASUK: 5,
  LAT_MASUK: 6,
  LONG_MASUK: 7,
  FOTO_MASUK: 8,
  JAM_PULANG: 9,
  FOTO_PULANG: 10,
  LAT_PULANG: 11,
  LONG_PULANG: 12,
  LOGBOOK: 13,
  STATUS: 14,
  CATATAN: 15 
};

const IDX = {
  TANGGAL: 0,
  ID: 1,
  NAMA: 2,
  UKM: 3,
  JAM_MASUK: 4,
  LAT_MASUK: 5,
  LONG_MASUK: 6,
  FOTO_MASUK: 7,
  JAM_PULANG: 8,
  FOTO_PULANG: 9,
  LAT_PULANG: 10,
  LONG_PULANG: 11,
  LOGBOOK: 12,
  STATUS: 13
};

function autoCheckGoogleAccount(email) {

  if (!email || email.trim() === "") {
    return { status: "error" };
  }

  const ss = SpreadsheetApp.openById("1Y1ZnugwO8YE5uQaz8ymBSWUW-h3e5V-ABG1N7JdfSvw");
  const sheet = ss.getSheetByName("AllowedUsers");

  if (!sheet) {
    Logger.log("Sheet AllowedUsers tidak ditemukan!");
    return { status: "error" };
  }

  const data = sheet.getDataRange().getValues();

  Logger.log("Email Google : " + email);

  for (let i = 1; i < data.length; i++) {
    const emailDB = String(data[i][1]).trim().toLowerCase();
    Logger.log("Email di Sheet : " + emailDB);
    if (emailDB === email.trim().toLowerCase()) {
      const statusDB = String(data[i][5]).trim().toLowerCase();
          // Cek apakah akun masih aktif
          if (statusDB !== "aktif") {
            return {
              status: "error",
              message: "Akun Anda sudah tidak aktif."
            };
          }

      Logger.log("EMAIL DITEMUKAN!");

      return {
        status: "success",
        user: {
          id: data[i][0],
          email: data[i][1],
          nama: data[i][2],
          ukm: data[i][3],
          telp: data[i][4],
          status: data[i][5],
          role: data[i][6]
        }
      };
    }
  }

  

  return { status: "error" };
}


function doGet(e) {
  

  const email = Session.getActiveUser().getEmail(); 
  Logger.log("EMAIL GOOGLE : " + email); 

  const template = HtmlService.createTemplateFromFile("index"); 
  const inviteCode = (e && e.parameter && e.parameter.invite) ? e.parameter.invite.trim() : ""; 
  const check = autoCheckGoogleAccount(email); 
  const invite = checkInvitation(inviteCode); 
  
  Logger.log(JSON.stringify(check)); 
  Logger.log(JSON.stringify(invite)); 

  template.verifiedUser = (check.status === "success") ? check.user : null; 
  template.inviteCode = inviteCode; 
  template.inviteStatus = invite.status; 
  template.detectedEmail = email; 
  template.accessDeniedMessage = check.message || "";
  

  return template
    .evaluate()
    .setTitle("Portal Absensi")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService
      .createHtmlOutputFromFile(filename)
      .getContent();
}

function checkInvitation(inviteCode) {

  const ss = SpreadsheetApp.openById("1Y1ZnugwO8YE5uQaz8ymBSWUW-h3e5V-ABG1N7JdfSvw");
  const sheet = ss.getSheetByName("Invitations");

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {

    const code = String(data[i][0]).trim();
    const status = String(data[i][1]).trim();
    if (code === inviteCode) {

      if (status === "Available") {
        return {
          status: "success",
          row: i + 1
        };
      }
      return {
        status: "used"
      };
    }
  }
  return {
    status: "invalid"
  };
}

function registerByInvitation(formData, inviteCode) {

  const email = Session.getActiveUser().getEmail();

  if (!email) {
    return {
      status: "error",
      message: "Email Google tidak dapat dideteksi."
    };
  }

  // Cek invitation
  const invite = checkInvitation(inviteCode);

  if (invite.status !== "success") {
    return {
      status: "error",
      message: "Link undangan sudah tidak berlaku."
    };
  }

  const ss = SpreadsheetApp.openById("1Y1ZnugwO8YE5uQaz8ymBSWUW-h3e5V-ABG1N7JdfSvw");

  const sheetUser = ss.getSheetByName("AllowedUsers");
  const sheetInvite = ss.getSheetByName("Invitations");

  const data = sheetUser.getDataRange().getValues();

  // Cek email sudah terdaftar
  for (let i = 1; i < data.length; i++) {

    if (String(data[i][1]).trim().toLowerCase() === email.toLowerCase()) {

      return {
        status: "error",
        message: "Email ini sudah terdaftar."
      };

    }

  }

  // Generate ID baru
  const newID = generateNewCoachID(sheetUser);

sheetUser.appendRow([
    newID,
    email,
    formData.nama,
    formData.kegiatan,
    formData.telp,
    "aktif",
    "Pelatih",
    formData.jenis,
    JSON.stringify(formData.signature)
]);

  // Ubah invitation menjadi Used
  // Tandai invitation sudah digunakan
  sheetInvite.getRange(invite.row, 2).setValue("Used");

  // Simpan email yang menggunakan invitation
  sheetInvite.getRange(invite.row, 4).setValue(email);

  // Simpan waktu penggunaan
  const waktuSekarang = Utilities.formatDate(
  new Date(),
  Session.getScriptTimeZone(),
  "dd/MM/yyyy HH:mm:ss"
);

sheetInvite.getRange(invite.row, 5).setValue(waktuSekarang);

  try {

    kirimEmailRegistrasi({
        id: newID,
        nama: formData.nama,
        email: email,
        kegiatan: formData.kegiatan
    });

  } catch(err) {

    Logger.log("Gagal mengirim email: " + err);

  }

  SpreadsheetApp.flush();

  return {
    status: "success",
    message: "Registrasi berhasil.",
    user:{
      id:newID,
      email:email,
      nama:formData.nama,
      ukm:formData.kegiatan,
      jenis:formData.jenis
    }
  };
}

function generateNewCoachID(sheet) {

  const data = sheet.getDataRange().getValues();

  let lastNumber = 0;

  for (let i = 1; i < data.length; i++) {

    const id = String(data[i][0]).trim();

    if (id.startsWith("PL")) {

      const number = parseInt(id.substring(2), 10);

      if (!isNaN(number) && number > lastNumber) {
        lastNumber = number;
      }

    }

  }

  const nextNumber = lastNumber + 1;

  return "PL" + String(nextNumber).padStart(3, "0");

}

function kirimEmailRegistrasi(user) {

    const subject = "Registrasi Pelatih UKM Berhasil";

    const body =
    `Halo ${user.nama},

    Selamat!

    Akun Anda telah berhasil didaftarkan sebagai Pelatih UKM Universitas Ciputra.

    Berikut informasi akun Anda:

    🆔 ID Pelatih  : ${user.id}
    👤 Nama        : ${user.nama}
    🎯 Email       : ${user.email}
    🏷️ Kategori    : ${user.jenis}
    🎯 Kegiatan    : ${user.kegiatan}

    Silakan gunakan akun Google yang sama setiap kali mengakses sistem absensi.

    Terima kasih.`;

      MailApp.sendEmail(
        user.email,
        subject,
        body
      );

}

function clockInServer(data) {
  const ss = SpreadsheetApp.openById("1Y1ZnugwO8YE5uQaz8ymBSWUW-h3e5V-ABG1N7JdfSvw");
  const sheet = ss.getSheetByName("Log Absensi");
  let urlFoto = "Latihan di Kampus";

  if (data.fotoMasuk) {

    try {

      const namaFile =
        data.idPelatih + "_" +
        Utilities.formatDate(
          new Date(),
          Session.getScriptTimeZone(),
          "yyyyMMdd_HHmmss"
        ) + ".jpg";

      urlFoto = uploadSelfieToDrive(
        data.fotoMasuk,
        namaFile
      );

    } catch(err) {

      return {
        status: "error",
        message: err.toString()
      };

    }

  }

  sheet.appendRow([
    new Date(),
    data.idPelatih,
    data.nama,
    data.ukm,
    data.jamMasuk,
    formatKoordinat(data.latMasuk),
    formatKoordinat(data.longMasuk),
    urlFoto,
    "",
    "",
    "",
    "",
    "",
    "Sedang Latihan"
  ]);

  SpreadsheetApp.flush();

  return {
    status: "success",
    message: "Absen masuk berhasil disimpan."
  };
}

 function clockOutServer(data) {

  const ss = SpreadsheetApp.openById("1Y1ZnugwO8YE5uQaz8ymBSWUW-h3e5V-ABG1N7JdfSvw");
  const sheet = ss.getSheetByName("Log Absensi");

  const today = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "dd/MM/yyyy"
  );

  const values = sheet.getDataRange().getValues();

  for (let i = values.length - 1; i >= 1; i--) {

    const tanggal = Utilities.formatDate(
      values[i][IDX.TANGGAL],
      Session.getScriptTimeZone(),
      "dd/MM/yyyy"
    );

    const id = String(values[i][IDX.ID]).trim();

    if (tanggal === today && id === data.idPelatih.trim()) {

      if (String(values[i][IDX.JAM_PULANG]).trim() !== "") {
        return {
          status: "error",
          message: "Anda sudah melakukan Clock Out hari ini."
        };
      }

      let urlFoto = "";

      try {

        const namaFile =
          data.idPelatih + "_" +
          Utilities.formatDate(
            new Date(),
            Session.getScriptTimeZone(),
            "yyyyMMdd_HHmmss"
          ) + ".jpg";

        urlFoto = uploadSelfieToDrive(
          data.fotoPulang,
          namaFile
        );

      } catch (err) {

        return {
          status: "error",
          message: err.toString()
        };

      }

      sheet.getRange(i + 1, COL.JAM_PULANG).setValue(data.jamPulang);
      sheet.getRange(i + 1, COL.FOTO_PULANG).setValue(urlFoto);
      sheet.getRange(i + 1, COL.LAT_PULANG).setValue(formatKoordinat(data.latPulang));
      sheet.getRange(i + 1, COL.LONG_PULANG).setValue(formatKoordinat(data.longPulang));
      sheet.getRange(i + 1, COL.LOGBOOK).setValue(data.logbook);
      sheet.getRange(i + 1, COL.STATUS).setValue("Selesai");

      SpreadsheetApp.flush();

      return {
        status: "success",
        message: "Absen pulang berhasil disimpan."
      };

    }

  }

  return {
    status: "error",
    message: "Data Clock In hari ini tidak ditemukan."
  };

}

function testFolderAccess() {
  const folder = DriveApp.getFolderById(FOLDER_SELFIE_ID);
  Logger.log(folder.getName());
}

function uploadSelfieToDrive(base64Data, fileName) {

  Logger.log("=== MULAI UPLOAD ===");
  const folder = DriveApp.getFolderById(FOLDER_SELFIE_ID);
  Logger.log("1. Folder OK");
  const base64 = base64Data.replace(/^data:image\/\w+;base64,/, "");
  Logger.log("2. Prefix dihapus");
  const bytes = Utilities.base64Decode(base64);
  Logger.log("3. Decode berhasil");
  const blob = Utilities.newBlob(bytes, MimeType.JPEG, fileName);
  Logger.log("4. Blob berhasil");
  const file = folder.createFile(blob);
  Logger.log("5. File berhasil dibuat");
  // file.setSharing(
  //   DriveApp.Access.ANYONE_WITH_LINK,
  //   DriveApp.Permission.VIEW
  // );
  Logger.log("6. Sharing berhasil");
  Logger.log("=== SELESAI ===");
  return file.getUrl();
}

function formatKoordinat(value) {
  if (value === null || value === undefined || value === "") {
    return "";
  }

  const angka = Number(value);

  if (isNaN(angka)) {
    return String(value);
  }

  return "'" + angka.toFixed(9);
}

function getAttendanceStatus() {

  const ss = SpreadsheetApp.openById("1Y1ZnugwO8YE5uQaz8ymBSWUW-h3e5V-ABG1N7JdfSvw");
  const sheet = ss.getSheetByName("Log Absensi");
  const email = Session.getActiveUser().getEmail();
  const user = autoCheckGoogleAccount(email);

  if (user.status !== "success") {
    return {
      status: "unauthorized"
    };
  }

  const idPelatih = user.user.id;
  const data = sheet.getDataRange().getValues();
  const today = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "dd/MM/yyyy"
  );

  //---------------------------------------------------
  // 1. CEK APAKAH MASIH ADA ABSENSI BELUM SELESAI
  //---------------------------------------------------

  for (let i = data.length - 1; i >= 1; i--) {
    const id = String(data[i][IDX.ID]).trim();
    if (id !== idPelatih) continue;
    const jamMasuk = data[i][IDX.JAM_MASUK];
    const jamPulang = data[i][IDX.JAM_PULANG];
    const tanggal = Utilities.formatDate(
      data[i][IDX.TANGGAL],
      Session.getScriptTimeZone(),
      "dd/MM/yyyy"
    );

    // Hari sebelumnya
    if (tanggal !== today && jamMasuk && !jamPulang) {

      return {
        status: "pending",
        row: i + 1,
        tanggal: tanggal,
        jamMasuk: Utilities.formatDate(
          new Date(jamMasuk),
          Session.getScriptTimeZone(),
          "HH:mm"
        )
      };
    }
  }

  //---------------------------------------------------
  // 2. CEK ABSENSI HARI INI
  //---------------------------------------------------

  for (let i = data.length - 1; i >= 1; i--) {
    const id = String(data[i][IDX.ID]).trim();
    if (id !== idPelatih) continue;
    const tanggal = Utilities.formatDate(
      data[i][IDX.TANGGAL],
      Session.getScriptTimeZone(),
      "dd/MM/yyyy"
    );

    if (tanggal !== today) continue;

    return {
      status: "today",
      clockIn: data[i][IDX.JAM_MASUK] != "",
      clockOut: data[i][IDX.JAM_PULANG] != ""
    };

  }

  //---------------------------------------------------
  // 3. BELUM ADA ABSENSI HARI INI
  //---------------------------------------------------

  return {
    status: "not_found",
    clockIn: false,
    clockOut: false
  };

}

function testCreateFile() {
  const folder = DriveApp.getFolderById(FOLDER_SELFIE_ID);

  const file = folder.createFile(
    "test.txt",
    "Hello World"
  );

  return file.getUrl();
}

function debugTanggal() {

  const ss = SpreadsheetApp.openById("1Y1ZnugwO8YE5uQaz8ymBSWUW-h3e5V-ABG1N7JdfSvw");
  const sheet = ss.getSheetByName("Log Absensi");

  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {

    Logger.log(values[i][0]);
    Logger.log(typeof values[i][0]);

  }

}

function getHistoryPelatih() {
  

  Logger.log("=== GET HISTORY DIPANGGIL ===");
  const ss = SpreadsheetApp.openById("1Y1ZnugwO8YE5uQaz8ymBSWUW-h3e5V-ABG1N7JdfSvw");
  const sheet = ss.getSheetByName("Log Absensi");

  const email = Session.getActiveUser().getEmail();

  const user = autoCheckGoogleAccount(email);

  if (user.status !== "success") {
    return [];
  }

  const idPelatih = user.user.id;

  Logger.log("ID Login : " + idPelatih);

  const values = sheet.getDataRange().getValues();
  const history = [];

  for (let i = values.length - 1; i >= 1; i--) {

    Logger.log("Row " + i);
    Logger.log(values[i]);
    Logger.log("ID Sheet : " + values[i][IDX.ID]);
    Logger.log("ID Login : " + idPelatih);

    if (String(values[i][IDX.ID]).trim() !== idPelatih) {
        Logger.log("Tidak sama");
        continue;
    }

    Logger.log("MATCH!");

    history.push({

      tanggal: Utilities.formatDate(
          new Date(values[i][IDX.TANGGAL]),
          Session.getScriptTimeZone(),
          "dd/MM/yyyy"
      ),

      ukm: String(values[i][IDX.UKM]),

      jamMasuk: Utilities.formatDate(
          new Date(values[i][IDX.JAM_MASUK]),
          Session.getScriptTimeZone(),
          "HH:mm"
      ),

      jamPulang: Utilities.formatDate(
          new Date(values[i][IDX.JAM_PULANG]),
          Session.getScriptTimeZone(),
          "HH:mm"
      ),

      status: String(values[i][IDX.STATUS]),

      logbook: String(values[i][IDX.LOGBOOK]),

      fotoMasuk: String(values[i][IDX.FOTO_MASUK]),

      fotoPulang: String(values[i][IDX.FOTO_PULANG])

    });
  }
  Logger.log(history);
  Logger.log(JSON.stringify(history));
  return history;
}

function getListUKM() {

  const ss = SpreadsheetApp.openById("1Y1ZnugwO8YE5uQaz8ymBSWUW-h3e5V-ABG1N7JdfSvw");
  const sheet = ss.getSheetByName("UKM");

  const data = sheet
      .getRange(2, 1, sheet.getLastRow() - 1, 1)
      .getValues();

  return data.flat().filter(String);

}

function getListClub(){

    const ss = SpreadsheetApp.openById("1Y1ZnugwO8YE5uQaz8ymBSWUW-h3e5V-ABG1N7JdfSvw");
    const sheet = ss.getSheetByName("CLUB");
    const values = sheet.getDataRange().getValues();

    values.shift();

    return values.map(function(r){

        return r[0];

    });

}

function getDashboardAdmin(){

  const ssUser = SpreadsheetApp.openById("1Y1ZnugwO8YE5uQaz8ymBSWUW-h3e5V-ABG1N7JdfSvw");
  const userSheet = ssUser.getSheetByName("AllowedUsers");

  const ssAbsen = SpreadsheetApp.openById("1Y1ZnugwO8YE5uQaz8ymBSWUW-h3e5V-ABG1N7JdfSvw");
  const logSheet = ssAbsen.getSheetByName("Log Absensi");
  const ukmSheet = ssAbsen.getSheetByName("UKM");

  const userData = userSheet.getDataRange().getValues();
  const logData = logSheet.getDataRange().getValues();
  const ukmData = ukmSheet.getDataRange().getValues();

  let totalPelatih = 0;
  let sedangLatihan = 0;
  let absensiHariIni = 0;
  let aktivitasTerbaru = [];

  const today = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "yyyyMMdd"
  );

  // hitung pelatih
  for(let i=1;i<userData.length;i++){

      if(userData[i][6]=="Pelatih"){
          totalPelatih++;
      }

  }

  // hitung log absensi
  for(let i=1;i<logData.length;i++){

      if(logData[i][13]=="Sedang Latihan"){
          sedangLatihan++;
      }

      const tgl = Utilities.formatDate(
          new Date(logData[i][0]),
          Session.getScriptTimeZone(),
          "yyyyMMdd"
      );

      if(tgl==today){
          absensiHariIni++;
      }

  }

    const recent = logData
      .slice(1)
      .reverse()
      .slice(0,5);

  recent.forEach(function(r){

      aktivitasTerbaru.push({

          nama:r[2],

          ukm:r[3],

          jam:r[13]=="Selesai"
              ? Utilities.formatDate(
                  new Date(r[8]),
                  Session.getScriptTimeZone(),
                  "HH:mm"
                )
              : Utilities.formatDate(
                  new Date(r[4]),
                  Session.getScriptTimeZone(),
                  "HH:mm"
                ),

          status:r[13]

      });

  });

return{
    totalPelatih,
    totalUKM:ukmData.length-1,
    sedangLatihan,
    absensiHariIni,
    aktivitasTerbaru
};

}

function getListPelatih(){
    const ss = SpreadsheetApp.openById("1Y1ZnugwO8YE5uQaz8ymBSWUW-h3e5V-ABG1N7JdfSvw");
    const sheet = ss.getSheetByName("AllowedUsers");
    const data = sheet.getDataRange().getValues();
    const hasil = [];
    for(let i=1;i<data.length;i++){
        if(data[i][6] != "Pelatih") continue;
        hasil.push({
          id: data[i][0],
          email: data[i][1],
          nama: data[i][2],
          ukm: data[i][3],      // kegiatan
          telp: data[i][4],
          status: data[i][5],
          jenis: data[i][7]     // <-- TAMBAHKAN
        });
    }
    return hasil;
}

function loadPelatih(){
    google.script.run
    .withSuccessHandler(function(data){
    allPelatih = data;
    filteredPelatih = data;
    renderPelatih(filteredPelatih);
  })
  .getListPelatih();
}

function updatePelatih(data){
    const ss = SpreadsheetApp.openById("1Y1ZnugwO8YE5uQaz8ymBSWUW-h3e5V-ABG1N7JdfSvw");
    const sheet = ss.getSheetByName("AllowedUsers");
    const values = sheet.getDataRange().getValues();

    for(let i=1;i<values.length;i++)
    {
      if(values[i][0] == data.id)
      {
        sheet.getRange(i+1,3).setValue(data.nama);    // Nama
        sheet.getRange(i+1,4).setValue(data.ukm);     // UKM
        sheet.getRange(i+1,5).setValue(data.telp);    // Telp
        sheet.getRange(i+1,6).setValue(data.status);  // Status

        SpreadsheetApp.flush();

        return{
            status:"success",
            message:"Data berhasil diperbarui."
        };
      }
    }
    return{
        status:"error",
        message:"Data tidak ditemukan."
    };
}

function getListAbsensi(){

  const ss = SpreadsheetApp.openById("1Y1ZnugwO8YE5uQaz8ymBSWUW-h3e5V-ABG1N7JdfSvw");
  const sheet = ss.getSheetByName("Log Absensi");

  const values = sheet.getDataRange().getValues();
  values.shift();

  return values.map(function(r,index){
    return{
      row: index + 2, 
      tanggal:r[0]? Utilities.formatDate(new Date(r[0]),Session.getScriptTimeZone(),"yyyy-MM-dd"): "",
      tanggalDisplay:r[0]? Utilities.formatDate(new Date(r[0]),Session.getScriptTimeZone(),"dd/MM/yyyy" ): "",
      id: String(r[1] || ""),
      nama: String(r[2] || ""),
      ukm: String(r[3] || ""),
      jamMasuk: r[4]? Utilities.formatDate(new Date(r[4]),Session.getScriptTimeZone(),"HH:mm"): "",
      latMasuk: String(r[5] || ""),
      longMasuk: String(r[6] || ""),
      fotoMasuk: String(r[7] || ""),
      jamPulang: r[8]? Utilities.formatDate( new Date(r[8]),Session.getScriptTimeZone(),"HH:mm"): "",
      fotoPulang: String(r[9] || ""),
      latPulang: String(r[10] || ""),
      longPulang: String(r[11] || ""),
      logbook: String(r[12] || ""),
      status: String(r[13] || ""),
      catatan:String(r[14]||"")   // << kolom baru
    };
  });

}

function testGetListAbsensi(){
  const data = getListAbsensi();
  Logger.log(data);
}

function getListUKM() {

  const ss = SpreadsheetApp.openById("1Y1ZnugwO8YE5uQaz8ymBSWUW-h3e5V-ABG1N7JdfSvw");
  const sheet = ss.getSheetByName("UKM");

  const data = sheet
      .getRange(2,1,sheet.getLastRow()-1,1)
      .getValues();

  return data.flat().filter(String);
}

function addUKM(nama){

  const sheet = SpreadsheetApp
      .openById("1Y1ZnugwO8YE5uQaz8ymBSWUW-h3e5V-ABG1N7JdfSvw")
      .getSheetByName("UKM");

  sheet.appendRow([nama]);

  return {
    success:true
  };
}

function updateUKM(oldNama,newNama){

  const sheet = SpreadsheetApp
      .openById("1Y1ZnugwO8YE5uQaz8ymBSWUW-h3e5V-ABG1N7JdfSvw")
      .getSheetByName("UKM");

  const data = sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues();

  for(let i=0;i<data.length;i++){

    if(data[i][0]==oldNama){

      sheet.getRange(i+2,1).setValue(newNama);
      break;

    }

  }
}

function deleteUKM(nama){

  const sheet = SpreadsheetApp
      .openById("1Y1ZnugwO8YE5uQaz8ymBSWUW-h3e5V-ABG1N7JdfSvw")
      .getSheetByName("UKM");

  const data = sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues();

  for(let i=0;i<data.length;i++){

    if(data[i][0]==nama){

      sheet.deleteRow(i+2);
      break;

    }

  }

}

function getListClub(){

  const sheet = SpreadsheetApp
      .openById("1Y1ZnugwO8YE5uQaz8ymBSWUW-h3e5V-ABG1N7JdfSvw")
      .getSheetByName("CLUB");

  const data = sheet
      .getRange(2,1,sheet.getLastRow()-1,1)
      .getValues();

  return data.flat().filter(String);

}

function addClub(nama){
  const sheet = SpreadsheetApp
      .openById("1Y1ZnugwO8YE5uQaz8ymBSWUW-h3e5V-ABG1N7JdfSvw")
      .getSheetByName("CLUB");
  sheet.appendRow([nama]);
}

function updateClub(oldNama,newNama){
  const sheet = SpreadsheetApp
      .openById("1Y1ZnugwO8YE5uQaz8ymBSWUW-h3e5V-ABG1N7JdfSvw")
      .getSheetByName("CLUB");
  const data = sheet
      .getRange(2,1,sheet.getLastRow()-1,1)
      .getValues();
  for(let i=0;i<data.length;i++){
      if(data[i][0]==oldNama){
          sheet.getRange(i+2,1).setValue(newNama);
          return;
      }
  }
}

function updateAbsensi(data){
  Logger.log(JSON.stringify(data));
  const ss = SpreadsheetApp.openById("1Y1ZnugwO8YE5uQaz8ymBSWUW-h3e5V-ABG1N7JdfSvw");
  const sheet = ss.getSheetByName("Log Absensi");
  const row = Number(data.row);
  Logger.log("ROW = " + row);

  if (!row || row < 2) {
    throw new Error("Row tidak valid: " + row);
  }

  const values = sheet.getRange(row, 1, 1, 15).getValues()[0];

  values[4]  = data.jamMasuk;
  values[8]  = data.jamPulang;
  values[12] = data.logbook;
  values[13] = data.status;
  values[14] = data.catatan;

  sheet.getRange(row, 1, 1, 15).setValues([values]);

  return true;
}

function deleteClub(nama){
  const sheet = SpreadsheetApp
      .openById("1Y1ZnugwO8YE5uQaz8ymBSWUW-h3e5V-ABG1N7JdfSvw")
      .getSheetByName("CLUB");

  const data = sheet
      .getRange(2,1,sheet.getLastRow()-1,1)
      .getValues();

  for(let i=0;i<data.length;i++){
      if(data[i][0]==nama){
          sheet.deleteRow(i+2);
          return;
      }
  }
}

function getListPelatihLogbook(){
  const ss = SpreadsheetApp.openById("1Y1ZnugwO8YE5uQaz8ymBSWUW-h3e5V-ABG1N7JdfSvw");
  const sheet = ss.getSheetByName("AllowedUsers");
  const data = sheet.getDataRange().getValues();
  const hasil = [];
  for(let i=1;i<data.length;i++){

    if(data[i][6] != "Pelatih") continue;

    hasil.push({
      id:data[i][0],
      nama:data[i][2]
    });

  }
  return hasil;
}

function loadTahunLogbook(){
    const select =
        document.getElementById("logbookTahun");

    const tahunSekarang =
        new Date().getFullYear();

    select.innerHTML = "";
    for(let tahun=tahunSekarang; tahun>=2026; tahun--){

        select.innerHTML += `
            <option value="${tahun}">
                ${tahun}
            </option>
        `;
    }
}

function getPreviewLogbook(idPelatih, tanggalMulai, tanggalSelesai){

  const ss = SpreadsheetApp.openById("1Y1ZnugwO8YE5uQaz8ymBSWUW-h3e5V-ABG1N7JdfSvw");
  const sheetUser = ss.getSheetByName("AllowedUsers");
  const sheetAbsen = ss.getSheetByName("Log Absensi");
  const userData = sheetUser.getDataRange().getValues();
  const absensi = sheetAbsen.getDataRange().getValues();
  let pelatih = null;

  // ==========================
  // Cari data pelatih
  // ==========================
  for(let i=1;i<userData.length;i++){
    if(userData[i][0] == idPelatih){
      pelatih = {
        id : userData[i][0],
        email : userData[i][1],
        nama : userData[i][2],
        kegiatan : userData[i][3],
        telp : userData[i][4],
        status : userData[i][5],
        role : userData[i][6],
        jenis : userData[i][7],
        tandaTangan : userData[i][8]
      };
      break;
    }
  }

  // Kalau ID tidak ditemukan
  if(!pelatih){
    return {
      status:"error",
      message:"Pelatih tidak ditemukan."
    };
  }

  // ==========================
  // Ambil data absensi
  // ==========================
  const hasil = [];
  for(let i=1;i<absensi.length;i++){

    // Kolom B = ID Pelatih
    Logger.log(absensi[i][0]);
    Logger.log(typeof absensi[i][0]);
    if(absensi[i][1] != idPelatih){
        Logger.log({
            idSheet: absensi[i][1],
            idDipilih: idPelatih,
            sama: absensi[i][1] == idPelatih
        });
      continue;
    }

    const tanggalData = Utilities.formatDate(
    new Date(absensi[i][0]),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd"
    );

    if (tanggalData < tanggalMulai || tanggalData > tanggalSelesai) {
      continue;
    }

    if(absensi[i][13] !== "Selesai"){
      continue;
    }

   hasil.push({
    tanggal: Utilities.formatDate(
        new Date(absensi[i][0]),
        Session.getScriptTimeZone(),
        "dd/MM/yyyy"
    ),

    jamMasuk: absensi[i][4]
        ? Utilities.formatDate(
            new Date(absensi[i][4]),
            Session.getScriptTimeZone(),
            "HH:mm"
        )
        : "",

    jamPulang: absensi[i][8]
        ? Utilities.formatDate(
            new Date(absensi[i][8]),
            Session.getScriptTimeZone(),
            "HH:mm"
        )
        : "",
    logbook: absensi[i][12] || "",
    status: absensi[i][13] || ""
});
  }
  // ==========================
  // Return hasil
  // ==========================
  return {
    status: "success",
    pelatih: pelatih,
    absensi: hasil,
    tanggalMulai: tanggalMulai,
    tanggalSelesai: tanggalSelesai
  };
}

function submitRecovery(data){
  const ss = SpreadsheetApp.openById("1Y1ZnugwO8YE5uQaz8ymBSWUW-h3e5V-ABG1N7JdfSvw");
  const sheet = ss.getSheetByName("Log Absensi");
  try{

    //----------------------------------------
    // Upload foto
    //----------------------------------------

    const namaFile =
      "RECOVERY_" +
      Utilities.formatDate(
        new Date(),
        Session.getScriptTimeZone(),
        "yyyyMMdd_HHmmss"
      ) +
      ".jpg";

    const urlFoto = uploadSelfieToDrive(
      data.foto,
      namaFile
    );

    //----------------------------------------
    // Update data
    //----------------------------------------
    sheet.getRange(data.row, COL.JAM_PULANG).setValue(data.jamPulang);
    sheet.getRange(data.row, COL.FOTO_PULANG).setValue(urlFoto);
    sheet.getRange(data.row, COL.LOGBOOK).setValue(data.logbook);
    sheet.getRange(data.row, COL.STATUS).setValue("Selesai");
    sheet.getRange(data.row, COL.CATATAN)
        .setValue(
          "Clock Out diselesaikan oleh pelatih pada " +
          Utilities.formatDate(
            new Date(),
            Session.getScriptTimeZone(),
            "dd/MM/yyyy HH:mm:ss"
          )
        );
    SpreadsheetApp.flush();

    return{
      status:"success",
      message:"Absensi sebelumnya berhasil diselesaikan."
    };
  }catch(err){
    return{
      status:"error",
      message:err.toString()
    };
  }
}

function getEmailByID(idPelatih){

  const ss = SpreadsheetApp.openById("1Y1ZnugwO8YE5uQaz8ymBSWUW-h3e5V-ABG1N7JdfSvw");
  const sheet = ss.getSheetByName("AllowedUsers");
  const data = sheet.getDataRange().getValues();
  for(let i=1;i<data.length;i++){
      if(data[i][0] == idPelatih){
          return {
              email: data[i][1],
              nama : data[i][2]
          };
      }
  }
  return null;
}

function reminderClockOut(){
  const ss = SpreadsheetApp.openById("1Y1ZnugwO8YE5uQaz8ymBSWUW-h3e5V-ABG1N7JdfSvw");
  const sheet = ss.getSheetByName("Log Absensi");
  const data = sheet.getDataRange().getValues();
  const today = Utilities.formatDate(
      new Date(),
      "Asia/Jakarta",
      "yyyy-MM-dd"
  );

  for(let i=1;i<data.length;i++){
    const row = data[i];
    const tanggal = Utilities.formatDate(
        new Date(row[0]),
        "Asia/Jakarta",
        "yyyy-MM-dd"
    );
    const id = row[1];
    const jamPulang = row[8];
    const status = row[13];
    Logger.log({
      tanggal: tanggal,
      today: today,
      status: status,
      jamPulang: jamPulang,
      id: id
    });
    if(
        tanggal != today &&
        status == "Sedang Latihan" &&
        !jamPulang
    ){
    const user = getEmailByID(id);
      if(user){
        MailApp.sendEmail({
          to: user.email,
          subject: "Pengingat Clock Out",
          htmlBody: `
            Halo <b>${user.nama}</b>,<br><br>

            Sistem mendeteksi bahwa Anda belum melakukan
            <b>Clock Out</b> pada latihan hari ini.

            <br><br>

            Mohon segera membuka Portal Absensi Pelatih
            dan menyelesaikan absensi.

            <br><br>

            Terima kasih.<br>
            Student Affairs
          `
        });
      }
    }
  }
}
//-------------------------------------CEK HISTORY LAPORAN--------------------------------------------------------//
function cekRiwayatLaporan(idPelatih, tanggalMulai, tanggalSelesai){
  const ss = SpreadsheetApp.openById("1Y1ZnugwO8YE5uQaz8ymBSWUW-h3e5V-ABG1N7JdfSvw");
  const sheet = ss.getSheetByName("HistoryLaporan");
  const data = sheet.getDataRange().getValues();

  for(let i = 1; i < data.length; i++){

    const id = data[i][0];

    if(id != idPelatih){
      continue;
    }

    const mulaiLama = Utilities.formatDate(
      new Date(data[i][2]),
      Session.getScriptTimeZone(),
      "yyyy-MM-dd"
    );

    const selesaiLama = Utilities.formatDate(
      new Date(data[i][3]),
      Session.getScriptTimeZone(),
      "yyyy-MM-dd"
    );

    // Cek apakah rentang tanggal bertabrakan
    if(
      tanggalMulai <= selesaiLama &&
      tanggalSelesai >= mulaiLama
    )
    {
      
      return {
        status: "error",
        message: "Periode tersebut sudah pernah diekspor.",
        tanggalMulai: mulaiLama,
        tanggalSelesai: selesaiLama
      };
    }
  }
  return {
    status: "success"
  };
}

function simpanRiwayatLaporan(idPelatih,namaPelatih,tanggalMulai,tanggalSelesai){
    const ss = SpreadsheetApp.openById("1Y1ZnugwO8YE5uQaz8ymBSWUW-h3e5V-ABG1N7JdfSvw");
    const sheet = ss.getSheetByName("HistoryLaporan");

    sheet.appendRow([
        idPelatih,
        namaPelatih,
        tanggalMulai,
        tanggalSelesai,
        new Date(),
        Session.getActiveUser().getEmail()
    ]);

    return {
        status: "success"
    };
}


// ---------------------------------------------------------------------------------------------------------------//
function generateAndBuildLinks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const WEB_APP_URL =
"https://script.google.com/macros/s/AKfycbwdYF_OZUDBDEFYiOOO1CWhJmcHnhC-LyEzr_VBIpLiFnq_dqh5MlwcVfw_ckXcjJ0/exec";
  const inviteSheet = ss.getSheetByName("Invitations");
  if (!inviteSheet) {
    SpreadsheetApp.getUi().alert("Sheet 'Invitations' tidak ditemukan!");
    return;
  }

  const webAppUrl = WEB_APP_URL;
  
  if (!webAppUrl || webAppUrl.indexOf("exec") === -1) {
    throw new Error("Deploy dulu script ini sebagai Web App untuk mendapatkan URL yang valid!");
  }

  const tokensToCreate = 10;
  for (let i = 0; i < tokensToCreate; i++) {
    const randomToken = "INV-" + Math.floor(100000 + Math.random() * 900000);
    const fullInviteLink = webAppUrl + "?invite=" + randomToken;
    
    inviteSheet.appendRow([
      randomToken,
      "Available",
      fullInviteLink,
      ""
    ]);
  }

  SpreadsheetApp.getUi().alert("🎉 10 link undangan baru berhasil dibuat di sheet Invitations.");
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🎫 Invitation Tool')
    .addItem('Generate 10 Link Baru', 'generateAndBuildLinks')
    .addToUi();
}

function getStatistikAbsensi7Hari(){
  const ss = SpreadsheetApp.openById("1Y1ZnugwO8YE5uQaz8ymBSWUW-h3e5V-ABG1N7JdfSvw");
  const sheet = ss.getSheetByName("Log Absensi");
  const data = sheet.getDataRange().getValues();
  data.shift();
  let hasil = {};
  data.forEach(function(r){
      if(!r[0]) return;
      const tanggal = Utilities.formatDate(
          new Date(r[0]),
          Session.getScriptTimeZone(),
          "dd/MM"
      );
      hasil[tanggal] = (hasil[tanggal] || 0) + 1;
  });

  const labels = Object.keys(hasil).slice(-7);
  const values = labels.map(t => hasil[t]);

  return {
      labels: labels,
      values: values
  };
}
