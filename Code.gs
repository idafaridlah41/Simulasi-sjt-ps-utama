const SHEET_SOAL = 'Bank_Soal';
const SHEET_HASIL = 'Hasil_Ujian';

/**
 * Mengirim data soal dalam format JSON saat diakses oleh GitHub Pages
 */
function doGet(e) {
  const hasil = getDaftarSoal();
  return ContentService
    .createTextOutput(JSON.stringify(hasil))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Menerima dan menyimpan hasil ujian yang dikirim dari GitHub Pages
 */
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const hasil = simpanHasilUjian(payload);
    return ContentService
      .createTextOutput(JSON.stringify(hasil))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'error',
        message: 'Format data tidak valid: ' + error.message
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Mengambil seluruh daftar soal dari Google Sheet 'Bank_Soal'
 */
function getDaftarSoal() {
  try {
    inisialisasiDatabase();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_SOAL);
    const data = sheet.getDataRange().getValues();

    const daftarSoal = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;

      daftarSoal.push({
        no: Number(row[0]),
        kategori: String(row[1]),
        pertanyaan: String(row[2]),
        opsi: {
          A: String(row[3]),
          B: String(row[4]),
          C: String(row[5]),
          D: String(row[6])
        },
        kunci: String(row[7]).trim().toUpperCase(),
        analisis: String(row[8])
      });
    }

    return {
      status: 'success',
      data: daftarSoal
    };
  } catch (error) {
    return {
      status: 'error',
      message: 'Gagal memuat soal: ' + error.message
    };
  }
}

/**
 * Menyimpan hasil pengerjaan peserta ke Google Sheet 'Hasil_Ujian'
 */
function simpanHasilUjian(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheetHasil = ss.getSheetByName(SHEET_HASIL);

    if (!sheetHasil) {
      sheetHasil = ss.insertSheet(SHEET_HASIL);
      sheetHasil.appendRow([
        'Waktu Selesai',
        'Nama Peserta',
        'NIP / Instansi',
        'Skor Total',
        'Benar',
        'Salah / Kosong',
        'Skor Manajerial',
        'Skor Sosial Kultural',
        'Skor Teknis',
        'Durasi Pengerjaan'
      ]);
      formatHeaderSheet(sheetHasil, 10);
    }

    const waktu = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');

    sheetHasil.appendRow([
      waktu,
      payload.nama || 'Anonim',
      payload.instansi || '-',
      payload.skorTotal,
      payload.jumlahBenar,
      payload.jumlahSalah,
      payload.skorManajerial,
      payload.skorSosiokultural,
      payload.skorTeknis,
      payload.durasi
    ]);

    return {
      status: 'success',
      message: 'Hasil ujian berhasil disimpan ke Google Sheets.'
    };
  } catch (error) {
    return {
      status: 'error',
      message: 'Gagal menyimpan hasil: ' + error.message
    };
  }
}

/**
 * Memformat baris pertama (header) sheet agar rapi
 */
function formatHeaderSheet(sheet, numCols) {
  const headerRange = sheet.getRange(1, 1, 1, numCols);
  headerRange.setBackground('#1E3A8A');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  sheet.setFrozenRows(1);
}

/**
 * Membuat sheet 'Bank_Soal' & 'Hasil_Ujian' secara otomatis jika belum tersedia
 */
function inisialisasiDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sheetHasil = ss.getSheetByName(SHEET_HASIL);
  if (!sheetHasil) {
    sheetHasil = ss.insertSheet(SHEET_HASIL);
    sheetHasil.appendRow([
      'Waktu Selesai',
      'Nama Peserta',
      'NIP / Instansi',
      'Skor Total',
      'Benar',
      'Salah / Kosong',
      'Skor Manajerial',
      'Skor Sosial Kultural',
      'Skor Teknis',
      'Durasi Pengerjaan'
    ]);
    formatHeaderSheet(sheetHasil, 10);
  }

  let sheetSoal = ss.getSheetByName(SHEET_SOAL);
  if (sheetSoal && sheetSoal.getLastRow() > 1) {
    return;
  }

  if (!sheetSoal) {
    sheetSoal = ss.insertSheet(SHEET_SOAL);
  }

  sheetSoal.clear();
  sheetSoal.appendRow([
    'No',
    'Kategori',
    'Pertanyaan',
    'Opsi_A',
    'Opsi_B',
    'Opsi_C',
    'Opsi_D',
    'Kunci_Benar',
    'Analisis_Pembahasan'
  ]);
  formatHeaderSheet(sheetSoal, 9);

  const bankData = getDataAwal45Soal();
  sheetSoal.getRange(2, 1, bankData.length, 9).setValues(bankData);
  sheetSoal.setColumnWidth(1, 50);
  sheetSoal.setColumnWidth(2, 140);
  sheetSoal.setColumnWidth(3, 400);
  sheetSoal.setColumnWidths(4, 4, 250);
  sheetSoal.setColumnWidth(8, 90);
  sheetSoal.setColumnWidth(9, 400);
  sheetSoal.getRange(2, 1, bankData.length, 9).setWrap(true);
}

/**
 * Dataset lengkap 45 Soal SJT Pengawas Sekolah (Ahli Madya ke Ahli Utama)
 */
function getDataAwal45Soal() {
  return [
    [1, 'Manajerial', 'Sebagai Pengawas Sekolah, Anda sering mendengar beberapa Kepala Sekolah mengeluh tentang kinerja pengawas sekolah lainnya. Beberapa Kepala Sekolah merasa kecewa dan ingin pengawas sekolah yang mendampingi mereka lebih memahami kondisi dan kebutuhan sekolah. Dalam situasi ini, upaya yang seharusnya dilakukan sebagai Pengawas Sekolah adalah ....',
      'Menyampaikan keresahannya dengan apa yang terjadi di lapangan dalam forum pengawas sekolah, agar pengawas sekolah lain sadar dan berubah untuk berperilaku sesuai kode etik dan norma pengawas sekolah.',
      'Melakukan refleksi diri setiap saat, agar tetap berpegang teguh pada kode etik dan norma sebagai pengawas sekolah, dan sesekali meminta umpan balik dari pengawas sekolah lainnya.',
      'Memberikan usul untuk membuat sesi rutin refleksi dan umpan balik dalam forum pengawas sekolah, sehingga hal-hal yang melanggar kode etik akan terus direfleksikan dan diperbaiki bersama.',
      'Mengingatkan pengawas sekolah lain untuk tetap berpedoman pada kode etik dan norma sebagai seorang pengawas sekolah.',
      'C', 'Jenjang Ahli Utama menuntut kemampuan membangun sistem dan budaya organisasi secara kolektif. Mengusulkan sesi rutin refleksi dan umpan balik dalam forum pengawas (Opsi C) menciptakan mekanisme perbaikan berkelanjutan yang menjaga kode etik profesi tanpa terkesan menggurui rekan sejawat secara personal.'],

    [2, 'Manajerial', 'Dalam masa transisi kebijakan Merdeka Belajar, sebagian kepala sekolah dan guru di sekolah binaan Anda masih merasa kebingungan serta terbebani dengan perubahan pola pendampingan baru yang diterapkan. Sebagai Pengawas Sekolah, tindakan yang paling tepat Anda lakukan untuk membangun pemahaman bersama dan mendorong komitmen perubahan di sekolah binaan adalah ....',
      'Menginstruksikan kepala sekolah dan guru untuk segera menyesuaikan diri dengan kebijakan baru karena aturan tersebut bersifat wajib dari pusat.',
      'Menyampaikan rincian tugas administratif yang harus dipenuhi sekolah secara tertulis agar tidak terjadi kesalahan dalam pelaporan.',
      'Membuka ruang dialog bersama kepala sekolah dan guru untuk mendengarkan kendala mereka, lalu menyepakati tahapan pendampingan yang sesuai dengan kesiapan masing-masing sekolah.',
      'Menunggu sekolah binaan siap secara mandiri sebelum memberikan pendampingan lebih lanjut agar tidak menambah beban kerja guru.',
      'C', 'Menerapkan paradigma pendampingan partisipatif dan berdiferensiasi. Pengawas mendengarkan kendala satuan pendidikan terlebih dahulu untuk membangun rasa aman psikologis, kemudian menyepakati tahapan pendampingan yang proporsional sesuai tingkat kesiapan sekolah.'],

    [3, 'Manajerial', 'Di era Kurikulum Merdeka banyak sekali informasi yang perlu disampaikan kepada sekolah binaan, dari dasar dan tujuan, pembelajaran dan asesmen serta administrasi. Tindakan yang seharusnya Anda lakukan sebagai pengawas sekolah, agar informasi tersebut dipahami dan dijalankan sesuai kesiapan sekolah di daerah Anda adalah ....',
      'Menggunakan gaya komunikasi informal agar sekolah binaan merasa nyaman saat kegiatan sosialisasi informasi tersebut.',
      'Membuat media penyampaian informasi yang memudahkan sekolah binaan memahami informasi Kurikulum Merdeka.',
      'Membuat sesi sosialisasi tentang Kurikulum Merdeka ke dalam beberapa tahapan sosialisasi, agar sekolah binaan bisa memahami sedikit demi sedikit.',
      'Mengintegrasikan informasi dengan membuat pemahaman konsep yang kontekstual dan mudah dipahami oleh sekolah binaan.',
      'D', 'Tantangan utama implementasi kebijakan baru adalah kesenjangan antara teori regulasi dengan realitas lapangan. Mengintegrasikan informasi menjadi pemahaman konsep yang kontekstual membuat sekolah mampu menerjemahkan kebijakan sesuai kesiapan dan karakteristik lokalnya.'],

    [4, 'Manajerial', 'Selama menjabat sebagai pengawas sekolah, Anda telah menemukan banyak praktik baik yang efektif dalam mendampingi sekolah binaan. Namun, Anda juga menyadari bahwa tidak semua pengawas sekolah melakukan hal yang sama. Oleh karena itu, sebagai upaya untuk menginspirasi pengawas sekolah lain, yang akan Anda lakukan untuk membagikan praktik baik yang telah Anda temukan dan mendorong mereka untuk mengikuti jejak Anda dalam mendampingi sekolah binaan adalah ....',
      'Mengapresiasi pengawas sekolah lain yang sudah melakukan praktik baik agar terus mempertahankannya.',
      'Dalam organisasi profesi pengawas sekolah yang Anda ikuti, mengusulkan beberapa praktik baik yang teruji menjadi kebijakan dalam organisasi untuk bisa diikuti oleh pengawas sekolah lainnya.',
      'Memberikan masukan kepada pengawas sekolah lain untuk terus memperbaiki peranannya.',
      'Menyampaikan praktik baik yang teruji dalam forum pengawas sekolah.',
      'B', 'Level strategis Ahli Utama tidak berhenti pada berbagi cerita secara lisan (Opsi D), melainkan melembagakan praktik baik yang sudah teruji menjadi standar atau kebijakan dalam organisasi profesi (Opsi B) sehingga dampaknya terstruktur dan berkelanjutan.'],

    [5, 'Manajerial', 'Sekolah binaan Anda memiliki beberapa kelas paralel di tingkatan yang sama yang diajari oleh guru yang berbeda di setiap kelasnya. Setelah Anda amati dengan seksama, terdapat perbedaan metode pengajaran yang signifikan antar kelas ditingkatan yang sama. Sebagai upaya menyamakan setiap kelas paralel, yang Anda lakukan adalah ....',
      'Melakukan pelatihan untuk menyamakan kompetensi dan metode pengajaran setiap guru yang ada serta melakukan pemantauan secara berkala.',
      'Perbedaan metode pengajaran tidak membuat prestasi siswa berbeda-beda sehingga Anda membiarkan saja perbedaan yang ada.',
      'Anda mengumpulkan guru-guru terkait dan menyampaikan apa yang harus mereka lakukan secara langsung.',
      'Melakukan koordinasi dengan Kepala Sekolah agar Kepala Sekolah melakukan penyetaraan di setiap kelas',
      'A', 'Menjamin pemerataan mutu layanan pendidikan bagi peserta didik di kelas paralel melalui intervensi peningkatan kapasitas guru (pelatihan) yang diikuti pemantauan berkala, bukan sekadar instruksi satu arah (C) atau membiarkan kesenjangan mutu (B).'],

    [6, 'Manajerial', 'Salah satu sekolah binaan Anda memiliki gedung serbaguna yang cukup luas dan sering digunakan oleh masyarakat sekitar untuk acara-acara non akademik. Setelah beberapa kali penggunaan, Anda menemukan beberapa kerusakan dan masalah kebersihan setelah gedung tersebut digunakan oleh maasyarakat. Agar hubungan antara pihak sekolah dan masyarakat tetap terjalin baik namun tidak ada pihak manapun yang dirugikan yang Anda lakukan adalah ....',
      'Mengadakan uang iuran gedung bagi masyarakat sekitar sehingga masyarakat memiliki tanggung jawab untuk merawat gedung. Iuran digunakan untuk pemeliharaan fasilitas gedung secara berkala.',
      'Menerapkan biaya sewa gedung per jam untuk menurunkan keinginan warga sekitar menggunakan gedung serbaguna miliki sekolah.',
      'Membuat kebijakan baru dengan mengijinkan penggunaan gedung serbaguna oleh semua pihak namun akan dikenakan denda jika ditemukan kerusakan setelah penggunaan gedung.',
      'Mengadakan peraturan baru yang disosialisasikan ke masyarakat setempat bahwa gedung serbaguna hanya dapat dipakai di jam sekolah dan hanya untuk kepentingan sekolah saja.',
      'A', 'Mencerminkan kompetensi Pelayanan Publik dan kemitraan sekolah–masyarakat yang saling menguntungkan (win-win solution). Fasilitas publik tetap dapat dimanfaatkan masyarakat sekaligus menumbuhkan rasa memiliki serta menjamin ketersediaan dana pemeliharaan berkala.'],

    [7, 'Manajerial', 'Menurut pengamatan Anda, sekolah binaan Anda cukup tertinggal dari segi proses pembelajaran dan metode pengajaran. Tindakan yang Anda lakukan adalah ....',
      'Berdiskusi dengan kepala sekolah kemungkinan-kemungkinan inovasi baru yang dapat dilakukan.',
      'Menyampaikan pandangan Anda kepada Kepala Sekolah untuk dipikirkan oleh Kepala Sekolah.',
      'Membiarkan sekolah binaan Anda berjalan seperti biasa, merubah tradisi terlalu beresiko.',
      'Perlahan-lahan mengubah metode yang lama, sambil mendampingi dan memberikan pelatihan yang sesuai kepada guru.',
      'D', 'Transformasi pembelajaran pada satuan pendidikan yang tertinggal membutuhkan manajemen perubahan bertahap (incremental change) yang disertai pendampingan teknis dan pelatihan kontekstual kepada guru agar perubahan dapat diadopsi tanpa menimbulkan resistensi.'],

    [8, 'Manajerial', 'Saat mendampingi sekolah binaan dalam pelaksanaan Proyek Penguatan Profil Pelajar Pancasila (P5), ada beberapa tantangan yang ditemukan. Salah satunya adalah menganggap, bahwa Produk menjadi tujuan P5, sehingga tujuan utama penguatan Profil Pelajar Pancasila tidak tercapai. Hal itu juga terjadi di banyak sekolah di daerah Anda. Langkah yang bisa dilakukan agar masalah tersebut bisa diatasi dengan efektif adalah ....',
      'Menyampaikan temuan di lapangan tentang pelaksanaan P5, dan data-data tidak tercapainya tujuan dimensi yang diharapkan. Mengajak pengawas lain melakukan refleksi, dan mendiskusikan strategi yang lebih baik.',
      'Mengusulkan cara baru dalam sosialisasi P5 agar tidak lagi fokus pada produk, melainkan pada dimensi P5.',
      'Mengajak sekolah binaan untuk merefleksikan pelaksanaan P5, dan memperbaiki prosesnya ke depan agar tidak fokus pada produk lagi.',
      'Mengajak pengawas lain untuk merefleksikan pelaksanaan P5 di sekolah-sekolah binaan, apakah sudah sesuai dengan tujuan esensi yang diharapkan atau masih fokus dengan produk.',
      'A', 'Karena permasalahan terjadi secara masif di banyak sekolah di daerah tersebut, pengawas Ahli Utama perlu melakukan intervensi berbasis bukti (membawa data temuan lapangan), memfasilitasi refleksi kolektif antarpengawas, sekaligus merumuskan strategi perbaikan bersama (Opsi A).'],

    [9, 'Manajerial', 'Meskipun Pengawas Sekolah seharusnya berperan sebagai pendamping, seringkali terlihat banyak pengawas sekolah yang masih terlalu mengendalikan situasi di sekolah binaannya. Upaya yang bisa dilakukan agar lebih banyak pengawas sekolah yang mulai memahami peran dan fungsinya sebagai pendamping bagi sekolah binaannya adalah ....',
      'Memberikan masukan pada forum pengawas sekolah di daerah untuk membuat kebijakan peran pengawas sekolah, agar tidak lagi melakukan miskonsepsi.',
      'Dalam forum pengawas sekolah meminta pengawas sekolah lain untuk kembali merefleksikan peran mereka di sekolah.',
      'Memberitahu pengawas sekolah lain untuk menjadi pengawas seolah yang bukan mengendalikan sekolah binaan, tetapi mendampingi, menemani.',
      'Dalam organisasi profesi pengawas sekolah yang Anda ikuti, terus menunjukkan peran-peran pengawas sekolah sebagai pendamping dalam setiap interaksi dengan pengawas sekolah lain.',
      'D', 'Keteladanan perilaku nyata (role modeling) secara konsisten dalam setiap interaksi profesional jauh lebih efektif menggeser paradigma pengawas dari pengendali menjadi pendamping dibandingkan sekadar imbauan lisan (C) atau teguran langsung (B).'],

    [10, 'Manajerial', 'Sebagai Pengawas Sekolah, terkadang Anda harus mendampingi lebih dari satu sekolah. Isu kesenjangan antara sekolah kadang muncul, di mana beberapa sekolah merasa bahwa Anda lebih memperhatikan salah satu sekolah daripada yang lain. Dalam situasi ini, upaya yang seharusnya dilakukan agar Anda dapat bersikap adil dan setara terhadap semua sekolah binaan Anda adalah ....',
      'Membuat forum sekolah binaan untuk mendengarkan kebutuhan dan tantangan yang dihadapi sekolah, serta membuat kesepakatan pendampingan yang diketahui bersama.',
      'Memperlakukan cara yang sama pada setiap sekolah binaan.',
      'Membagi waktu pendampingan dengan adil, disamaratakan antar sekolah agar tidak terjadi kecemburuan.',
      'Tidak membanding-bandingkan sekolah binaan satu dan lainnya, karena setiap sekolah memiliki kelebihan dan kebutuhan masing-masing.',
      'A', 'Prinsip keadilan dalam kepengawasan bukanlah penyamarataan waktu secara kaku (B dan C), melainkan keadilan proporsional yang didasarkan pada pemetaan kebutuhan dan tantangan tiap sekolah melalui kesepakatan terbuka dalam forum bersama.'],

    [11, 'Manajerial', 'Kepala Sekolah di sekolah binaan Anda membuat masalah yang cukup berat, yaitu melanggar Hak Asasi Manusia dan membuat suasana di sekolah tersebut tidak kondusif. Semua guru dengan jelas menunjukkan ketidaksukaan mereka terhadap kepala sekolah. Untuk mengembalikan suasana kondusif di sekolah, yang Anda lakukan adalah ....',
      'Anda tidak melakukan apa-apa, kegiatan belajar mengajar berlangsung seperti biasa tetapi Anda membawa masalah ini ke Kepala Dinas dan menunggu keputusan dari Kepala Dinas untuk langkah selanjutnya.',
      'Membawa masalah kepada Kepala Dinas untuk ditindaklanjuti dan menghimbau semua warga sekolah untuk mengawasi tingkah laku Kepala Sekolah jika ada gerak gerik yang mencurigakan harap dilaporkan kepada Anda.',
      'Memberikan peringatan keras bagi Kepala Sekolah dan mengawasi Kepala Sekolah tersebut dengan ketat selama beberapa bulan.',
      'Memberikan peringatan kepada Kepala Sekolahdan membawa masalah ini ke Kepala Dinas untuk ditindaklanjuti serta memberikan pembinaan kepada Kepala Sekolah.',
      'D', 'Pelanggaran berat terkait HAM membutuhkan ketegasan yang sesuai batas kewenangan: pengawas memberikan peringatan dan pembinaan langsung untuk memulihkan kondusivitas sekolah, sekaligus melaporkan kasus tersebut kepada Kepala Dinas sebagai pejabat pembina kepegawaian.'],

    [12, 'Manajerial', 'Sekolah binaan Anda terletak di wilayah rawan banjir. Setiap tahunnya banjir selalu terjadi dan memberikan dampak yang besar bagi wilayah sekitar sekolah Anda. Anda mengamati penyebab dari terjadinya banjir ialah kebiasaan membuang sampah sembarangan. Sebagai salah satu petinggi di lembaga pendidikan, tindakan yang Anda lakukan sebagai Pengawas Sekolah adalah ....',
      'Sebulan sekali melakukan kegiatan gotong royong bersih-bersih yang dilakukan oleh seluruh warga sekolah Anda.',
      'Menghimbau para guru untuk menerapkan pentingnya menjaga kebersihan lingkungan kepada siswa sehingga mengurangi resiko banjir di masa depan.',
      'Bersama dengan warga sekolah dan masyarakat setempat melakukan kegiatan gotong royong bersih-bersih sebulan sekali.',
      'Menggerakan setiap bagian dari sekolah binaan Anda untuk menerapkan kebiasaan membuang sampah yang benar.',
      'D', 'Menggerakkan seluruh ekosistem sekolah untuk membangun pembiasaan perilaku bersih sehari-hari jauh lebih sistemik dan berkelanjutan untuk mencegah banjir daripada kegiatan gotong royong insidental sebulan sekali (A dan C) atau imbauan pasif (B).'],

    [13, 'Manajerial', 'Ketika membaca Rapor Pendidikan sekolah binaan, Anda menemukan bahwa setiap sekolah memiliki kompetensi yang rendah. Anda merasa perlu untuk mengajak setiap sekolah melakukan refleksi dan mengambil tindakan dari hasil Rapor Pendidikan tersebut. Sebagai Pengawas Sekolah, tindakan yang seharusnya Anda lakukan dalam situasi ini untuk membantu sekolah binaan meningkatkan kompetensinya adalah ....',
      'Memberikan solusi-solusi sesuai kebutuhan sekolah, agar bisa langsung diimplementasikan pada masing-masing sekolah.',
      'Melakukan refleksi masing-masing sekolah, memetakan kebutuhan belajar bersama, memberikan umpan balik dan mengapresiasi dari ide-ide yang disampaikan.',
      'Mendengarkan apa yang ingin dilakukan masing-masing sekolah dari hasil rapor pendidikan tersebut, dan memberikan beberapa solusi yang perlu dilakukan sekolah.',
      'Mengajak refleksi kepala sekolah, dan mendiskusikan masalah serta solusi yang perlu dilakukan untuk meningkatkan poin-poin yang rendah dalam Rapor Pendidikan.',
      'B', 'Selaras dengan siklus Perencanaan Berbasis Data (Identifikasi–Refleksi–Benahi) dan pendekatan coaching. Pengawas tidak menyuapi solusi siap pakai (A), melainkan memfasilitasi refleksi, memetakan kebutuhan belajar bersama, memberi umpan balik konstruktif, dan mengapresiasi inisiatif sekolah.'],

    [14, 'Manajerial', 'Keluarga dekat Anda mengetahui Anda adalah Pengawas Sekolah dan meminta Anda untuk mengosongkan satu tempat pada penerimaan siswa baru tahun ajaran berikutnya untuk anaknya. Kebetulan sekolah binaan Anda adalah salah satu sekolah favorit di daerah tersebut dan banyak orang yang ingin memasukkan anaknya ke sekolah tersebut. Tindakan yang Anda lakukan adalah ....',
      'Menyarankan mereka mengikuti tata cara yang seharusnya, tapi tanpa sepengetahuan keluarga Anda, Anda sudah menyiapkan satu tempat untuk anak tersebut agar tetap terlihat taat dengan peraturan.',
      'Menolak memasukkan ke sekolah terfavorit, tetapi memasukkan ke sekolah terfavorit kedua di daerah tersebut agar tidak terlihat mencurigakan oleh masyarakat sekitar.',
      'Langsung menyetujui permintaan keluarga Anda tersebut dan memberikan jaminan 100% anak tersebut otomatis menjadi siswa di tahun ajaran berikutnya.',
      'Menolak permintaan tersebut karena tidak sesuai dengan prosedur serta merugikan banyak pihak dan memintanya untuk mengikuti alur pendaftaran yang seharusnya.',
      'D', 'Menguji kompetensi Integritas tingkat tinggi. Pengawas wajib menolak segala bentuk penyalahgunaan wewenang dan nepotisme secara tegas, transparan, dan mengarahkan pihak keluarga mengikuti prosedur penerimaan peserta didik yang berlaku.'],

    [15, 'Manajerial', 'Salah satu guru di sekolah binaan Anda memiliki masalah dengan masyarakat setempat sehingga terjadi permusuhan antara guru dan masyarakat sekitar. Suasana di sekitar sekolah mulai tidak kondusif. Agar keadaan kembali kondusif yang Anda lakukan adalah ....',
      'Mempertemukan kedua pihak yang bermasalah dan merundingkan apa yang diinginkan oleh kedua pihak jika ingin berdamai.',
      'Memberikan peringatan tegas kepada guru yang bermasalah serta meminta maaf kepada masyarakat sekitar.',
      'Menegur baik guru dan kepala sekolah bersangkutan atas masalah yang ditimbulkan.',
      'Memindahkan guru yang bermasalah ke sekolah binaan lain yang terletak tidak jauh dari sekolah awal tetapi masyarakat sekitar tidak memiliki pengetahuan apa pun mengenai guru tersebut.',
      'A', 'Resolusi konflik sosial antara unsur sekolah dan masyarakat dilakukan melalui mediasi dialogis yang mempertemukan kedua belah pihak untuk mencapai kesepakatan damai, bukan menghindari akar masalah dengan memindahkan guru (D).'],

    [16, 'Manajerial', 'Dalam implementasi Kurikulum Merdeka, pengawas sekolah diharapkan untuk mengajak sekolah binaan dalam menerapkannya. Namun, dalam praktiknya, Anda melihat bahwa beberapa rekan pengawas sekolah lain menerapkan cara-cara instan yang tidak sesuai dengan tujuan Kurikulum Merdeka. Dalam situasi ini, tindakan yang seharusnya dilakukan sebagai Pengawas Sekolah adalah ....',
      'Menyampaikan apa yang dilihatnya kepada pengawas sekolah tersebut, dan memberikan masukan apa yang seharusnya dilakukan dalam mendorong sekolah binaan melaksanakan Kurikulum Merdeka.',
      'Menyampaikan masalah yang dilihatnya kepada pemangku kebijakan. Agar pemangku kebijakan melakukan tindakan ke depannya untuk perbaikan Implementasi Kurikulum Merdeka.',
      'Membuat sesi refleksi dalam forum pengawas sekolah dan memberikan rekomendasi mana yang perlu dihentikan, dilanjutkan dan dikembangkan dalam pendampingan di sekolah binaan.',
      'Mengingatkan pengawas sekolah tersebut bahwa cara-cara instan tersebut kurang relevan dalam Implementasi Kurikulum Merdeka.',
      'C', 'Menggunakan forum refleksi sejawat untuk mengevaluasi praktik pendampingan secara objektif (mengidentifikasi mana yang perlu dihentikan, dilanjutkan, dan dikembangkan) menjaga kolektivitas sekaligus meluruskan miskonsepsi tanpa konfrontasi.'],

    [17, 'Manajerial', 'Dalam mendampingi sekolah binaan, Anda mendapatkan sebuah praktik baik yang kemungkinan ketika diimplementasikan oleh pengawas sekolah lain akan berdampak juga pada sekolah binaan mereka. Tindakan yang seharusnya Anda lakukan sebagai rekan sejawat pengawas sekolah agar praktik baik Anda diimplementasikan oleh pengawas sekolah lain adalah ....',
      'Mengajak pengawas sekolah lain untuk menggunakan konsep dan pemikirannyam karena hasilnya sudah terlihat baik di sekolah binaan.',
      'Dalam sesi pertemuan antarpengawas di daerah, menyampaikan keberhasilan dan cara apa yang sudah dilakukan dalam mendampingi sekolah binaan.',
      'Memberi tahu praktik yang dilakukan pengawas lain masih perlu direfleksikan, dan menyampaikan praktik baik yang sudah terbukti berhasil.',
      'Menuliskan praktik baik langkah-langkah dalam mendampingi sekolah binaan untuk dipublikasikan dan dibaca oleh pengawas lain sebagai sumber inspirasi dan bisa direplikasi dan modifikasi di sekolah binaanya.',
      'D', 'Pendokumentasian dan publikasi langkah-langkah sistematis praktik baik menciptakan aset pengetahuan tertulis yang dapat dipelajari, direplikasi, serta dimodifikasi secara luas oleh rekan pengawas sesuai konteks sekolah masing-masing.'],

    [18, 'Manajerial', 'Sekolah binaan Anda beberapa kali mengalami perselisihan di antara warga sekolah dan pada beberapa sekolah masalah yang ada bisa diselesaikan antara guru yang bermasalah, sedangkan di beberapa sekolah lain masalah di tingkatan guru dibawa sampai ke Kepala Sekolah untuk diselesaikan. Beberapa masalah bahkan memakan waktu yang lama sehingga mengganggu kinerja para guru yang ada. Solusi yang Anda lakukan untuk masalah ini adalah ....',
      'Memberikan peringatan bagi guru yang bermasalah dengan harapan bisa memberikam efek jera dan contoh bagi guru yang lain agar tidak membuat masalah di masa mendatang.',
      'Semua masalah yang berada di bawah Kepala Sekolah Anda serahkan penyelesaiannya kepada Kepala Sekolah sehingga Anda hanya mengurusi masalah pada tingkatan Kepala Sekolah.',
      'Membuat sistem yang berjenjang sesuai dengan jenjang jabatan, sehingga masalah hanya akan melibatkan 1 jenjang jabatan di atas jenjang jabatan pihak yang bermasalah.',
      'Anda menangani langsung masalah yang terjadi pada warga sekolah Anda jika masalah yang terjadi sudah terlalu lama',
      'C', 'Membangun tata kelola penyelesaian konflik yang terstruktur dan berjenjang (eskalasi maksimal satu tingkat di atas pihak yang bermasalah) mencegah penumpukan masalah di pimpinan puncak dan mempercepat pemulihan kinerja.'],

    [19, 'Manajerial', 'Anda sebagai Pengawas Sekolah bersama dengan Kepala Sekolah akan memasuki tahun ajaran baru dengan beberapa guru yang baru bergabung di tahun ajaran tersebut, agar kompetensi guru dapat berkembang, yang Anda lakukan adalah ....',
      'Berdiskusi bersama dengan semua warga sekolah untuk menetapkan pengembangan kompetensi yang dibutuhkan oleh guru.',
      'Menghimbau semua guru mengikuti pelatihan agar setiap guru memiliki keterampilan minimal yang setara sebelum tahun ajaran dimulai.',
      'Merancang sendiri program pengembangan dan pelatihan guru lalu membagikannya kepada Kepala Sekolah untuk dilaksanakan.',
      'Mengadakan pertemuan sebelum tahun ajaran baru dimulai untuk perbaikan kompetensi guru berdasarkan evaluasi tahun ajaran sebelumnya.',
      'D', 'Perencanaan pengembangan kompetensi guru pada awal tahun ajaran baru harus berpijak pada data hasil evaluasi kinerja dan pembelajaran tahun ajaran sebelumnya agar program intervensi tepat sasaran.'],

    [20, 'Manajerial', 'Dinas Pendidikan melakukan sosialisasi strategi dalam penggunaan Platform Merdeka Mengajar. Namun saat sosialisasi banyak hal yang menurut Anda kurang sesuai jika dijalankan. Pada situasi tersebut, tindakan yang seharusnya Anda lakukan sebagai Pengawas Sekolah adalah ....',
      'Melakukan terlebih dahulu apa yang sudah dipikirkan dengan matang oleh Dinas Pendidikan. Untuk di lapangan dimodifikasi sesuai kesiapan belajar masing-masing sekolah binaan.',
      'Memberikan umpan balik tentang Platform Merdeka Mengajar, dan meminta Dinas Pendidikan mereview ulang dan mendiskusikan kembali strategi yang tepat untuk mencapai esensi.',
      'Mengikuti terlebih dahulu arahan dari Dinas Pendidikan dan Kebudayaan, jika pelaksanaan di lapangan tidak sesuai, bisa direfleksikan kemudian hari.',
      'Menyampaikan keberatan pada poin-poin yang disampaikan kepada Dinas Pendidikan. Sehingga dinas bisa mempertimbangkan strategi lainnya.',
      'A', 'Menunjukkan keluwesan implementasi kebijakan (adaptive leadership): menghormati dan menjalankan arah kebijakan Dinas Pendidikan sambil melakukan penyesuaian/modifikasi teknis di lapangan sesuai tingkat kesiapan belajar sekolah binaan.'],

    [21, 'Manajerial', 'Salah satu kompetensi dalam rapor pendidikan yang perlu dimiliki oleh peserta didik ialah kompetensi literasi dan numerasi. Upaya Anda sebagai pengawas sekolah yang mendorong sekolah binaan dalam meningkatkan kompetensi literasi dan numerasi peserta didik adalah ....',
      'Menyampaikan target kompetensi dan literasi peserta didik, dan mendiskusi cara-cara untuk melampaui target tersebut.',
      'Menyampaikan target yang perlu dikejar sekolah dalam hal literasi dan numerasi, dan mendorong sekolah melampaui target tersebut, dalam kurun waktu tertentu.',
      'Memberikan rekomendasi-rekomendasi program literasi dan numerasi yang sudah terbukti berhasil untuk dilaksanakan di sekolah binaan.',
      'Menyampaikan indikator kompetensi literasi dan numerasi yang ada di fase sekolah binaan, dan mengajak guru untuk mengintegrasikan literasi dan numerasi pada setiap pelajaran.',
      'D', 'Peningkatan literasi dan numerasi tidak dicapai melalui tekanan target angka semata (A dan B), melainkan dengan memahamkan indikator capaian fase dan mendorong guru mengintegrasikannya ke dalam seluruh proses pembelajaran lintas mata pelajaran.'],

    [22, 'Manajerial', 'Kementrian Pendidikan dan Kebudayaan mengadakan perubahan kurikulum untuk tahun ajaran baru dan harus diikuti oleh seluruh lembaga pendidikan se-Indonesia. Sebagai Pengawas Sekolah, untuk memastikan sekolah binaan Anda mengikuti kurikulum yang baru, yang Anda lakukan adalah ....',
      'Anda mengingatkan Kepala Sekolah untuk menghimbau para guru mengenai perubahan kuruikulum.',
      'Mengumpulkan Kepala Sekolah dan guru untuk memberikan arahan secara langsung, serta melakukan pemantauan langsung di interval waktu tertentu.',
      'Menghimbau secara langsung kepada guru dan meminta Kepala Sekolah untuk mengirimkan laporan kepada Anda setiap pertengahan semester.',
      'Menugaskan Kepala Sekolah untuk memberitahukan ke guru dan melaporkan kepada Anda secara berkala.',
      'B', 'Memastikan keberhasilan transisi kurikulum nasional membutuhkan penyamaan persepsi secara langsung kepada kepala sekolah dan guru serta pendampingan/pemantauan berkelanjutan dalam interval waktu tertentu, bukan sekadar pendelegasian pasif.'],

    [23, 'Manajerial', 'Sekolah binaan Anda mengalami kendala dalam menjalankan program sekolah, karena tidak memiliki jaringan mitra untuk diajak kerjasama dalam menyukseskan program tersebut. Pada situasi tersebut, tindakan yang seharusnya Anda lakukan sebagai Pengawas Sekolah adalah ....',
      'Menghubungi jaringan mitra yang dibutuhkan sekolah, meminta sekolah untuk menerima mitra tersebut agar program sekolah bisa berjalan dengan baik.',
      'Memberikan alternatif jaringan mitra dan mendiskusikan kekurangan dan kelebihan pada jaringan mitra tersebut.',
      'Memberikan kontak beberapa jaringan mitra yang dimiliki kepada kepala sekolah, agar sekolah bisa menjajaki kerjasama.',
      'Meminta jaringan mitra yang dikenal, untuk berkolaborasi dengan sekolah.',
      'B', 'Pengawas berperan sebagai fasilitator yang memberdayakan kemandirian sekolah dengan menyajikan beberapa alternatif mitra disertai analisis kelebihan dan kekurangannya agar sekolah dapat mengambil keputusan yang paling tepat.'],

    [24, 'Manajerial', 'Tugas pengawas adalah memantau dan mengevaluasi hasil kerja sekolah binaan terutama dalam meningkatkan kompetensi literasi dan numerasi peserta didik. Berikut ini yang merupakan tindakan yang paling tepat oleh Pengawas Sekolah terkait aktivitas memantau dan mengevaluasi praktik peningkatan kompetensi literasi dan numerasi di sekolah adalah ....',
      'Mengajak kepala sekolah untuk memiliki alternatif cara untuk mencapai tujuan peningkatan nilai literasi dan numerasi.',
      'Mengajak kepala sekolah dan guru di sekolah binaan melakukan refleksi dan mendiskusikan rencana aksi yang akan dilakukan.',
      'Melakukan monitoring dan pembinaan kepada kepala sekolah dan guru.',
      'Melakukan mentoring terhadap guru tentang pembelajaran literasi dan numerasi yang kontekstual terhadap peserta didik.',
      'B', 'Evaluasi yang memberdayakan selalu melibatkan kepala sekolah dan guru secara aktif untuk merefleksikan capaian praktik pembelajaran dan merumuskan rencana aksi perbaikan bersama.'],

    [25, 'Manajerial', 'Salah satu yang diharapkan dalam Kurikulum Merdeka adalah pembelajaran diferensiasi. Saat melakukan asesmen awal, ternyata sekolah-sekolah binaan Anda, banyak yang belum mengetahui praktik pembelajaran diferensiasi. Upaya yang Anda lakukan sebagai pengawas sekolah ketika melihat masalah tersebut adalah ....',
      'Mencarikan narasumber yang kompeten untuk menjadi pelatih ahli pada setiap sekolah binaan.',
      'Mengadakan pelatihan atau workshop tentang praktik pembelajaran diferensiasi untuk guru-guru sekolah binaan. Dalam pelatihan tersebut, pengawas dapat memberikan penjelasan mengenai konsep dan manfaat dari pembelajaran diferensiasi serta memberikan contoh-contoh praktik diferensiasi yang dapat diaplikasikan dalam kelas',
      'Memetakan praktik baik pembelajaran diferensiasi pada sekolah binaan, menginisiasi sesi saling berbagi praktik baik pembelajaran diferensiasi antar sekolah binaan.',
      'Menjadi narasumber pembelajaran diferensiasi bagi sekolah-sekolah binaan, agar memiliki pemahaman yang sama.',
      'B', 'Karena asesmen awal menunjukkan mayoritas sekolah belum mengetahui praktik diferensiasi (sehingga sesi berbagi antar sekolah pada opsi C belum memiliki basis praktik yang memadai), penyelenggaraan lokakarya yang memadukan konsep, manfaat, dan contoh aplikatif di kelas menjadi intervensi paling relevan.'],

    [26, 'Manajerial', 'Daerah salah satu sekolah binaan Anda rawan terhadap penculikan anak. Siswa-siswa sekolah binaan Anda terancam keselamatannya terutama saat-saat jam pulang sekolah. Untuk menjaga siswa sekolah Anda yang Anda lakukan adalah ....',
      'Anda dan seluruh warga sekolah mensosialisasikan kepada siswa langkah-langkah pencegahan yang sudah Anda siapkan dan disetujui oleh seluruh warga sekolah.',
      'Menghimbau siswa untuk lebih awas dan berhati-hati saat pulang untuk menurunkan resiko terjadinya penculikan.',
      'Mengumpulkan warga sekolah dan mendiskusikan bersama tindakan berjaga-jaga yang sebaiknya dilakukan.',
      'Meminta partisipasi seluruh warga sekolah untuk lebih berhati-hati,teliti, dan mengawasi siswa selama jam sekolah terutama di jam pulang.',
      'A', 'Mitigasi risiko keselamatan peserta didik memerlukan langkah nyata yang terstruktur: menyepakati protokol pencegahan bersama seluruh warga sekolah dan langsung mensosialisasikannya kepada siswa agar dapat diterapkan segera.'],

    [27, 'Manajerial', 'Salah satu sekolah binaan Anda berada di daerah pedalaman yang jauh dari perkotaan, sehingga pengetahuan, pola pikir, dan metode pengajaran guru di sekolah binaan Anda menurut Anda "ketinggalan jaman". Tindakan Anda untuk mengembangkan sekolah ini adalah ....',
      'Melakukan kunjungan ke sekolah lain untuk memelajari metode pengajaran dari sesama guru di sekolah lain.',
      'Tidak melakukan apapun dan melanjutkan saja yang sudah berjalan.',
      'Melaksanakan seminar untuk membagikan pengetahuan terbaru kepada sekolaah binaan.',
      'Mengatur jadwal dan menyelenggarakan pendidikan dan latihan rutin yang diikuti oleh semua guru.',
      'A', 'Untuk mengubah pola pikir (mindset) guru di daerah terpencil, pembelajaran berbasis pengalaman langsung melalui studi tiru (benchmarking) antarsejawat di sekolah lain memberikan gambaran konkret sehingga lebih mudah diyakini dan diadaptasi dibandingkan seminar satu arah (C).'],

    [28, 'Sosial Kultural', 'Dinas Pendidikan dan Kebudayaan mengadakan perlombaan antar sekolah dengan memilih dekorasi sekolah terbaik per wilayah. Sekolah binaan Anda juga ikut berpartisipasi dalam lomba tersebut. Lomba ini bertemakan Bhinneka Tunggal Ika. Kebetulan sekolah binaan Anda terdiri atas guru dengan latar yang berbeda-beda. Anda ingin agar sekolah binaan Anda mendapatkan penghargaan dari perlombaan tersebut. Tindakan yang Anda ambil sebagai Pengawas Sekolah adalah ....',
      'Berusaha menampilkan semua suku, agama, dengan adil pada dekorasi sekolah.',
      'Menganjurkan untuk menggunakan tema dekorasi kelompok mayoritas dengan alasan kemudahan dan ketersediaan pernak pernik dekorasi.',
      'Berdiskusi dengan Kepala Sekolah untuk menentukan tema dekorasi yang diinginkan.',
      'Memilih dekorasi kelompok yang minoritas di sekolah Anda untuk menunjukkan kepedulian terhadap kelompok minoritas.',
      'A', 'Mewujudkan kompetensi Perekat Bangsa secara autentik dengan memastikan seluruh unsur suku dan agama terwakili secara adil dan setara, tanpa bias terhadap kelompok mayoritas (B) maupun simbolisasi sepihak pada minoritas (D).'],

    [29, 'Sosial Kultural', 'Selama satu tahun ajaran, tentunya terdapat beberapa hari yang bertepatan dengan hari peringatan nasional seperti hari Kemerdekaan Indonesia dan hari Kartini. Untuk ikut meramaikan perayaan hari nasional tetapi juga menanamkan nilai kebangsaan pada setiap warga sekolah Anda, selain kewajiban mengikuti upacara, yang akan Anda lakukan adalah ....',
      'Mengadakan team building activity secara berkala dengan tujuan meningkatkan kesadaran akan keberagaman serta toleransi satu dengan yang lain.',
      'Menghimbau seluruh warga sekolah untuk menggunakan atribut adat salah satu provinsi di Indonesia setiap dua kali sebulan pada minggu genap hari Kamis.',
      'Pada hari nasional, Anda memberikan ketentuan berbusana daerah yang harus diikuti oleh warga sekolah Anda.',
      'Pada setiap hari kerja yang bertepatan dengan hari nasional, setiap warga sekolah menggunakan atribut penunjuk identitas suku asli mereka.',
      'A', 'Internalisasi nilai kebangsaan dan toleransi jauh lebih mendalam apabila dibangun melalui aktivitas interaksi kolaboratif (team building) yang menumbuhkan empati dan kebersamaan lintas latar belakang, dibandingkan sekadar simbolisasi atribut fisik (B, C, D).'],

    [30, 'Sosial Kultural', 'Sekolah-sekolah di bawah binaan Anda, semakin sering membuat program yang seragam, program tersebut identik dengan identitas sosial tertentu, sehingga tampak tidak memberi ruang pada identitas sosail lainnya. Kebanyakan beralasan sekolah dengan program identitas sosial tersebut lebih menarik minat masyarakat. Anda mendorong sekolah memiliki program yang lebih plural dengan ....',
      'Menyediakan saran kepada sekolah binaan berupa daftar program, sehingga mampu memilih menyelenggarakan yang baru.',
      'Menyarankan sekolah-sekolah binaan untuk terus merancang program-program baru yang berbeda dengan sekolah lainnya.',
      'Mengajak sekolah menganalisis kebutuhan masyarakat, minat & bakat murid demi menghasilkan beragam program sekolah.',
      'Mengajak sekolah merefleksikan profil lulusan masing-masing, lalu merancang program sekolah yang sesuai dengan itu.',
      'C', 'Mendorong inklusivitas dan keberagaman program sekolah melalui analisis objektif atas kebutuhan seluruh lapisan masyarakat serta minat dan bakat peserta didik yang beragam, bukan menyeragamkan identitas sosial tertentu.'],

    [31, 'Sosial Kultural', 'Muncul anggapan negatif di masyarakat bahwa semua sekolah binaan Anda hanya bisa dikepalai oleh orang dengan identitas sosial tertentu. Agar isu tersebut diatasi hal yang akan dilakukan adalah ....',
      'Menanggapinya dengan memberikan dorongan agar segera muncul calon kepala sekolah baru yang berasal dari latar belakang sosial berbeda.',
      'Menepis anggapan tersebut dengan menunjukkan kelebihan-kelebihan para kepala sekolah yang sekarang sedang menjabat di sekolah binaan.',
      'Membuat sebuah gerakan yang mendukung pendidik dengan identitas sosial berbeda untuk bisa mengembangkan karier sebagai kepala sekolah.',
      'Mengkomunikasikan tentang kriteria kepala sekolah yang digunakan, lalu mengajak pendidik dari berbagai kalangan mengembangkan karier.',
      'D', 'Pengangkatan kepala sekolah dalam birokrasi ASN tunduk pada sistem merit, kompetensi, dan regulasi kualifikasi—bukan kuota atau afirmasi identitas semata (A/C). Mengomunikasikan kriteria objektif secara transparan sekaligus mengajak pendidik dari semua kalangan mempersiapkan diri mengembangkan karier (D) menegakkan profesionalisme sekaligus inklusivitas.'],

    [32, 'Sosial Kultural', 'Sekolah binaan Anda terdiri atas guru yang berasal dari berbagai latar belakang suku, agama, dan ras. Hal ini sudah berlangsung selama beberapa tahun dan tidak ada masalah yang signifikan selama beberapa tahun belakangan ini. Untuk meningkatkan kedekatan diantara guru yang Anda lakukan adalah ....',
      'Menganjurkan Kepala Sekolah untuk mengadakan program pelatihan setiap semester untuk meningkatkan ikatan dan kerjasama antar guru',
      'Senantiasa menanamkan nilai Bhinneka Tunggal Ika di setiap kesempatan yang ada.',
      'Mendorong Kepala Sekolah melakukan workshop bertemakan Bhinneka Tunggal Ika untuk meningkatkan toleransi antar guru.',
      'Mendorong Kepala Sekolah untuk menyelenggarakan program khusus bagi kaum minoritas agar mereka merasa dihargai.',
      'B', 'Memelihara dan mempererat harmoni sosial yang sudah kondusif dilakukan melalui pembiasaan dan penanaman nilai-nilai Bhinneka Tunggal Ika secara alami di setiap kesempatan interaksi sehari-hari.'],

    [33, 'Sosial Kultural', 'Terdapat persaingan sekolah yang menempatkan sekolah dengan identitas minoritas menjadi pemenang, sehingga muncul anggapan sinis terhadap sekolah tersebut. Padahal kompetisi yang ada sudah tampak adil. Tindakan yang Anda lakukan adalah ....',
      'Menyelenggarakan kompetisi baru yang dirancang secara khusus, hingga mampu menempatkan sekolah dengan identitas mayoritas sebagai pemenang.',
      'Mengajak sekolah berefleksi terhadap capaian, sekaligus mengupayakan penghapusan kompetisi yang menimbulkan sinis pada identitas tertentu.',
      'Mengajak sekolah berefleksi terhadap capaian, hingga memperbanyak ragam program agar sekolah dengan beragam identitas berkesempatan tampil.',
      'Menyelenggarakan program khusus, sebagai ruang bagi masing-masing identitas sosial, agar dapat berkesempatan tampil di ruang masing-masing.',
      'C', 'Meredam sentimen negatif tanpa mengorbankan keadilan kompetisi (A) atau melakukan segregasi ruang (D). Refleksi capaian yang diikuti perluasan ragam program membuka peluang aktualisasi yang setara bagi seluruh sekolah dengan latar belakang beragam.'],

    [34, 'Teknis', 'Para kepala sekolah mulai membanding-bandingkan kualitas antarpengawas sekolah. Ada yang kecewa hingga tidak suka, adapula yang menganggap Anda lebih baik dari pengawas sekolah lainnya. Sebagian rekan tidak nyaman dengan kondisi ini terlebih karena Anda dianggap lebih baik, maka yang Anda lakukan adalah ....',
      'Memposisikan diri sebagai rekan kepala sekolah untuk menyuarakan kekecewaan pada kinerja pengawas sekolah.',
      'Muncul di sesi-sesi pelatihan maupun seminar sebagai narasumber agar menjadi teladan pengawas sekolah lain.',
      'Mencari beberapa pengawas sekolah yang paling mudah diajak berefleksi, agar mereka menularkan ke yang lain.',
      'Meminta arahan langsung kepada kepala dinas, agar segera melakukan peningkatan kapasitas pengawas sekolah.',
      'C', 'Strategi difusi inovasi dan manajemen perubahan di tengah kecemburuan sejawat: menggandeng beberapa rekan pengawas yang paling terbuka untuk berefleksi sebagai agen perubahan awal (early adopters) guna menularkan praktik baik secara organik tanpa menonjolkan ego pribadi (B).'],

    [35, 'Teknis', 'Anda ingin sekali terjadi kolaborasi antarpengawas sekolah, agar pendampingan sekolah/madrasah di daerah Anda semakin baik. Tiap kali melakukan ajakan, muncul pertanyaan balik "Apa yang bisa dikolaborasikan?" Maka anda menanggapinya dengan ....',
      'Menggunakan berkas kerja bersama untuk berbagi proses pendampingan, sehingga mempermudah menemukan kesempatan kolaborasi.',
      'Menggunakan grup percakapan bersama untuk saling berinteraksi antarpengawas sekolah, sehingga dapat saling berkolaborasi.',
      'Membuat daftar berbagai program untuk ditawarkan kepada Pengawas Sekolah lainnya sehingga menjadi pintu awal kolaborasi.',
      'Membuat lomba yang berisi aktivitas kolaborasi antarsekolah, sehingga diharapkan pengawas sekolahnya juga berkolaborasi.',
      'A', 'Ketika rekan sejawat bertanya secara konkret "Apa yang bisa dikolaborasikan?", penyediaan berkas kerja bersama yang memetakan proses, tantangan, dan kebutuhan pendampingan tiap pengawas (Opsi A) secara langsung memperlihatkan irisan masalah dan peluang kolaborasi nyata, jauh lebih terarah dibanding sekadar grup percakapan umum (B).'],

    [36, 'Teknis', 'Semakin banyak sekolah yang merasa membutuhkan narasumber terkait Implementasi Kurikulum Merdeka dengan sesi lebih intensif. Anda senang dengan semangat belajar ini, dan berinisiatif mengajak pengawas lain saling bergantian menjadi narasumber. Namun Pengawas Sekolah yang sudah membina penerapan IKM teralalu sibuk, sedangkan lainnya merasa belum menguasai topik tersebut. Jika kebutuhan narasumber tidak segera dipenuhi, Anda khawatir semangat belajar ini menurun. Maka tindakan anda adalah ....',
      'Mengirim para kepala sekolah ke daerah lain untuk belajar implementasi kurikulum merdeka.',
      'Mengakomodasi kebutuhan banyak sekolah sekaligus dengan menjadi narasumber di MKKS/K3S.',
      'Melibatkan kepala sekolah yang sudah menerapkan kurikulum merdeka sebagai narasumbernya.',
      'Membuat rekaman penjelasan Implementasi Kurikulum Merdeka, untuk dipelajari oleh sekolah.',
      'C', 'Pemberdayaan aset dan sumber daya lokal (asset-based thinking): melibatkan kepala sekolah penggerak/praktisi yang telah berhasil menerapkan Kurikulum Merdeka sebagai narasumber berbagi praktik baik menjaga momentum belajar tetap hidup sekaligus menumbuhkan kepemimpinan pembelajaran antarsekolah.'],

    [37, 'Teknis', 'Seringkali arahan yang Anda berikan tidak dilaksanakan oleh pihak sekolah, padahal arahan tersebut sudah sesuai dengan pengalaman Anda sebelumnya yang terbukti menghasilkan perkembangan. Tindakan Anda adalah...',
      'Menghentikan penggunaan arahan, mempelajari, kemudian menerapkan cara baru dalam memberikan umpan balik.',
      'Memberi umpan balik terus menerus kepada sekolah yang belum menerapkan arahan untuk segera melaksanakan.',
      'Merefleksikan pendampingan sekolah sebelumnya, mencari cara yang lebih sesuai dalam memberi umpan balik.',
      'Membuat media seperti kartu kendali pendampingan sekolah, agar memonitor terus perkembangan yang terjadi.',
      'C', 'Setiap satuan pendidikan memiliki karakteristik dan tingkat kesiapan berbeda. Pengawas yang reflektif mengevaluasi kembali pendekatan pendampingannya dan menyesuaikan gaya penyampaian umpan balik agar dapat diterima serta diimplementasikan oleh sekolah tersebut.'],

    [38, 'Teknis', 'Seorang rekan pengawas sekolah menyatakan keinginannya untuk dapat memberikan pendampingan lebih baik ke sekolah. Sebagai titik awal perbaikan dia sudah mempelajari program-program yang berhasil di sekolah lain, sayangnya sulit direplikasi di sekolah binaannya, di lain sisi ia takut dianggap tidak inovatif jika hanya meniru. Upaya yang Anda lakukan adalah ....',
      'Mengajak rekan-rekan pengawas sekolah membuat sesi refleksi bersama untuk mengidentifikasi kelebihan dan kelemahan diri. Hasilnya jadi pertimbangan untuk saling memberi umpan balik.',
      'Membuat bank data program-program sekolah yang pernah diterapkan. Data juga memuat langkah-langkah penyelenggaraan program agar memudahkan pengawas sekolah menemukan inovasi program.',
      'Memberikan saran kepada rekan pengawas sekolah untuk melaksanakan beberapa program. Program yang disarankan pernah dilakukan dan menurut Anda paling mudah diterapkan di sekolah binaan.',
      'Mengajak rekan-rekan pengawas membuat sesi berbagi cerita inovasi program yang berhasil di masing-masing sekolah binaan. kegiatan ini berisi pemaparan kemajuan yang dicapai sekolah.',
      'D', 'Sesi berbagi cerita proses inovasi di masing-masing sekolah binaan memperlihatkan bagaimana sebuah gagasan diadaptasi sesuai konteks lokal, sehingga menghilangkan kekhawatiran rekan sejawat bahwa modifikasi praktik baik berarti tidak inovatif.'],

    [39, 'Teknis', 'Anda merasa guru-guru dan kepala sekolah di daerah semakin banyak yang mendapat banyak pelatihan. Kemudian mulai muncul kekhawatiran pengawas akan bingung jika ditanyai mengenai topik tertentu yang sedang ramai diperbincangkan. Anda sempat mengajak Pengawas Sekolah untuk ikut belajar bersama guru, namun mereka menolak. Anda pun melakukan tindakan berupa ...',
      'Menunjukkan poster-poster pelatihan kepada para Pengawas Sekolah untuk mengajak mereka belajar lagi.',
      'Mengajak organisasi profesi Pengawas Sekolah untuk menyelenggarakan sesi-sesi peningkatan kompetensi.',
      'Memberi nasehat pada Pengawas Sekolah lain, jika tidak belajar bersama guru, tentu akan ketinggalan.',
      'Mengkomunikasikan kepada organisasi profesi Pengawas Sekolah terkait penyediaan pelatihan-pelatihan.',
      'B', 'Menghargai kenyamanan psikologis rekan sejawat yang enggan belajar dalam satu forum bersama guru dengan memfasilitasi sesi peningkatan kompetensi khusus antar-pengawas melalui organisasi profesi.'],

    [40, 'Teknis', 'Sekolah di bawah binaan Anda sudah sering menyelenggarakan pelatihan bagi guru, namun kepala sekolah merasa belum ada perkembangan berarti di SDM, padahal pelatihan tersebut adalah hasil rekomendasi Anda. Maka Anda mengambil tindakan berupa ....',
      'Mencari pelatihan-pelatihan dengan topik dan pelatih berbeda, hingga menemukan yang berdampak pada perkembangan SDM di sekolah.',
      'Membuat refleksi rutin terkait kesulitan yang dialami guru pada proses pembelajaran, untuk mengidentifikasi kebutuhan pelatihan.',
      'Memberi saran kepada kepala sekolah untuk tetap mengikuti rekomendasi pelatihan, namun mengganti narasumbernya dengan yang baru.',
      'Menggali alasan dari guru dan kepala sekolah, untuk menemukan penyebab pelatihan tersebut belum berdampak pada perkembangan SDM.',
      'B', 'Pelatihan sering kali tidak berdampak karena tidak berangkat dari masalah nyata di kelas. Membangun refleksi rutin atas kesulitan riil guru dalam pembelajaran menjadi dasar analisis kebutuhan pelatihan (Training Needs Analysis) yang tepat sasaran.'],

    [41, 'Teknis', 'Salah satu sekolah binaan Anda sudah jauh lebih maju menerapkan diferensiasi pembelajaran. Ketika meminta sekolah binaan lain untuk turut menerapkan, mereka menganggap tidak mungkin dilakukan karena murid lebih banyak dengan status sosial dan permasalahan berbeda. Cara Anda untuk memotivasi sekolah lain menerapkannya adalah ....',
      'Mensosialisasikan dampak-dampak yang akan didapatkan sekolah, setelah murid mengalami diferensiasi pembelajaran.',
      'Membangun proses diskusi antarsekolah terkait permasalahan murid yang teratasi berkat diferensiasi pembelajaran.',
      'Menyelenggarakan lomba publikasi video penerapan diferensiasi pembelajaran berhadiah yang menarik untuk sekolah.',
      'Memberi penghargaan khusus kepada sekolah yang menerapkan diferensiasi pembelajaran lebih dulu dibanding lainnya.',
      'B', 'Hambatan mental berupa anggapan "tidak mungkin diterapkan di kondisi sekolah kami" paling efektif diatasi melalui diskusi antarsekolah yang membedah bagaimana masalah konkret murid berhasil diselesaikan berkat pembelajaran berdiferensiasi.'],

    [42, 'Teknis', 'Permasalahan kualitas pembelajaran terjadi di beberapa sekolah sekaligus, masing-masing kepala sekolah merasa butuh pendampingan khusus untuk mengatasinya. Maka pilihan strategi pendampingan yang diambil Anda adalah ....',
      'Membagi waktu ke salah satu sekolah lebih dahulu agar fokus, sebelum lanjut untuk menyelesaikan masalah sekolah lainnya.',
      'Mencari permasalahan-permasalahan yang paling umum dimiliki sekolah-sekolah tersebut untuk diselesaikan terlebih dahulu.',
      'Mengidentifikasi masalah dan kebutuhan yang jika terpenuhi akan berpengaruh paling besar pada perbaikan beberapa sekolah.',
      'Mengidentifikasi kesamaan dari permasalahan dan kebutuhan, agar perbaikan dapat sekaligus dilakukan di beberapa sekolah.',
      'D', 'Strategi pendampingan berbasis klaster/komunitas: memetakan irisan kesamaan masalah dan kebutuhan agar intervensi perbaikan dapat dirancang dan dijalankan secara serentak, efektif, dan efisien di beberapa sekolah sekaligus.'],

    [43, 'Teknis', 'Anda sudah sering melakukan pendampingan dan mendapatkan perkembangan pada sekolah binaan. Namun perkembangan Anda belum dirasakan/dilihat oleh rekan kerja. Anda pun jadi mengurungkan niat untuk meminta umpan balik, meskipun merasa butuh, karena khawatir makin dianggap tidak memiliki perkembangan. Cara yang sebaiknya dilakukan adalah ....',
      'Membuat berkas monitoring pendampingan sekolah berisi bukti perkembangan agar rekan kerja bisa memberi masukan.',
      'Meminta umpan balik pada rekan kerja secara berkala dan konsisten, lalu mengajak mereka melakukan hal serupa.',
      'Mendorong sekolah binaan menyelenggarakan pameran karya sebagai bukti pendampingan dan media meminta masukan.',
      'Menyediakan media yang memungkinkan untuk memberikan umpan balik kepada Anda secara anonim (tanpa identitas).',
      'C', 'Pameran karya sekolah binaan menampilkan dampak nyata hasil pendampingan secara objektif di ruang publik, sekaligus menjadi sarana dialog terbuka untuk mendapatkan umpan balik konstruktif dari rekan sejawat tanpa terkesan membanggakan diri.'],

    [44, 'Teknis', 'Seorang rekan Pengawas Sekolah mudah tersinggung saat diajak untuk menerapkan inovasi program pendampingan yang sudah Anda lakukan. Ada anggapan Anda ingin tampil sebagai sosok pengawas baru yang ditiru oleh pengawas lain. Tindakan yang Anda lakukan adalah ...',
      'Mengkomunikasikan kepada pengawas sekolah tersebut, bahwa yang Anda lakukan bukan untuk mencari panggung.',
      'Menjaga jarak dan berupaya tidak bekerja bersama pengawas sekolah tersebut, untuk terhindar dari konflik.',
      'Meminta masukan kepada rekan kerja terkait sikap Anda, lalu berupaya mengenal pengawas sekolah tersebut.',
      'Mengkomunikasikan permintaan maaf pada pengawas sekolah tersebut dan menyampaikan tujuan Anda sebenarnya.',
      'C', 'Menyatakan secara langsung bahwa Anda "bukan untuk mencari panggung" (Opsi A) bernada defensif dan berisiko memperuncing ketegangan interpersonal. Meminta masukan rekan kerja terkait cara komunikasi/sikap Anda lalu berupaya mengenal pengawas tersebut secara personal (Opsi C) menunjukkan kematangan kecerdasan emosional dan kerendahan hati seorang pemimpin.'],

    [45, 'Teknis', 'Dalam implementasi kurikulum merdeka, beberapa Pengawas Sekolah saling membanggakan sekolah binaannya banyak yang menerapkan mandiri berubah dan berbagi. Karena saling iri, para pengawas memaksa sekolah binaannya untuk minimal memilih mandiri berubah. Padahal Anda mengetahui tujuan 3 pilihan tersebut agar sekolah memilih sesuai kesiapan dan kondisi masing-masing. Terhadap fenomena ini tindakan Anda adalah ....',
      'Menganggap kompetisi antarpengawas sekolah ini sebagai hal baik, agar implementasi kurikulum merdeka semakin cepat terjadi di banyak sekolah dan mencapai hal yang sama.',
      'Mengajak para pengawas sekolah mempelajari ulang tujuan implementasi kurikulum merdeka. Lalu menjelaskan bahwa implementasi kurikulum merdeka tidak memaksakan pilihan.',
      'Mengajak para pengawas sekolah untuk menghindari persaingan, memperbanyak kegiatan kolaborasi antarsekolah, agar menghindari saling iri terhadap kondisi sekolah binaan.',
      'Menginisiasi forum yang membahas fungsi serta kode etik kepengawasan dan manajemen diri pemimpin. Lalu mengajak diskusi tujuan dan proses Implementasi Kurikulum Merdeka.',
      'D', 'Akar masalah pemaksaan kategori IKM akibat rasa iri adalah persoalan integritas peran dan pengelolaan ego pemimpin. Menginisiasi forum yang menyentuh fungsi, kode etik kepengawasan, dan manajemen diri terlebih dahulu sebelum meluruskan esensi pilihan Implementasi Kurikulum Merdeka merupakan solusi paling mendasar dan komprehensif.']
  ];
}
