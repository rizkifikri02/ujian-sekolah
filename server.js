const express = require('express');
const sqlite3 = require('sqlite3').verbose();
const cors = require('cors');
const bodyParser = require('body-parser');
const exceljs = require('exceljs');
const path = require('path');
const multer = require('multer');
const mammoth = require('mammoth');
const XLSX = require('xlsx');
const pdfParse = require('pdf-parse');
const fs = require('fs');

const app = express();
const PORT = 3000;

// Konfigurasi Multer untuk Upload File
const upload = multer({ dest: 'uploads/' });

// Middleware
app.use(cors());
app.use(bodyParser.json());
app.use(express.static(path.join(__dirname, 'public')));

// Database SQLite
const db = new sqlite3.Database('./ujian.db', (err) => {
  if (err) console.error('Gagal membuka database:', err.message);
  else console.log('Database SQLite berhasil terhubung.');
});

// Tabel Utama
db.serialize(() => {
  db.run(`CREATE TABLE IF NOT EXISTS siswa (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT UNIQUE,
    nama TEXT,
    kelas TEXT,
    password TEXT
  )`);

  db.run(`CREATE TABLE IF NOT EXISTS soal (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    pertanyaan TEXT,
    opsi_a TEXT,
    opsi_b TEXT,
    opsi_c TEXT,
    opsi_d TEXT,
    kunci TEXT
  )`);

  db.run(`CREATE TABLE IF NOT EXISTS sesi_ujian (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    siswa_id INTEGER UNIQUE,
    waktu_mulai DATETIME,
    waktu_selesai DATETIME,
    pelanggaran INTEGER DEFAULT 0,
    status TEXT DEFAULT 'BERJALAN'
  )`);

  db.run(`CREATE TABLE IF NOT EXISTS jawaban_siswa (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    siswa_id INTEGER,
    soal_id INTEGER,
    jawaban_terpilih TEXT,
    UNIQUE(siswa_id, soal_id)
  )`);

  db.run(`CREATE TABLE IF NOT EXISTS nilai (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    siswa_id INTEGER UNIQUE,
    jumlah_benar INTEGER,
    jumlah_salah INTEGER,
    total_nilai REAL
  )`);
});

// Helper Fungsi Tambah Siswa Otomatis
function insertSiswaAuto(nama, kelas, callback) {
  if (!nama || !kelas) return callback(null);
  const prefixKelas = kelas.toString().replace(/[^a-zA-Z0-9]/g, '').toUpperCase();
  const randomNum = Math.floor(100 + Math.random() * 900);
  const username = `UJS-${prefixKelas}-${randomNum}`;
  const password = Math.floor(1000 + Math.random() * 9000).toString();

  db.run("INSERT OR IGNORE INTO siswa (username, nama, kelas, password) VALUES (?, ?, ?, ?)",
    [username, nama.toString().trim(), kelas.toString().trim(), password],
    function(err) {
      if (err) callback(null);
      else callback({ username, password });
    }
  );
}

// --- API SISWA ---

app.post('/api/login', (req, res) => {
  const { username, password } = req.body;
  db.get("SELECT * FROM siswa WHERE UPPER(username) = UPPER(?) AND password = ?", [username.trim(), password.trim()], (err, row) => {
    if (err) return res.status(500).json({ error: err.message });
    if (!row) return res.status(401).json({ message: 'Username atau Password salah!' });
    res.json({ message: 'Login berhasil', siswa: row });
  });
});

app.post('/api/mulai-ujian', (req, res) => {
  const { siswa_id, durasi_menit } = req.body;
  
  db.get("SELECT * FROM sesi_ujian WHERE siswa_id = ?", [siswa_id], (err, sesi) => {
    if (sesi && sesi.status === 'SELESAI') {
      return res.status(400).json({ message: 'Anda sudah menyelesaikan ujian ini!' });
    }

    const sekarang = new Date();
    const waktuSelesai = new Date(sekarang.getTime() + (durasi_menit || 30) * 60000);

    if (!sesi) {
      db.run("INSERT INTO sesi_ujian (siswa_id, waktu_mulai, waktu_selesai) VALUES (?, ?, ?)", 
        [siswa_id, sekarang.toISOString(), waktuSelesai.toISOString()]);
    }

    db.all("SELECT id, pertanyaan, opsi_a, opsi_b, opsi_c, opsi_d FROM soal", [], (err, rows) => {
      if (err) return res.status(500).json({ error: err.message });
      
      for (let i = rows.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [rows[i], rows[j]] = [rows[j], rows[i]];
      }

      res.json({ 
        soal: rows, 
        waktu_selesai: sesi ? sesi.waktu_selesai : waktuSelesai.toISOString() 
      });
    });
  });
});

app.post('/api/simpan-jawaban', (req, res) => {
  const { siswa_id, soal_id, jawaban } = req.body;
  const sql = `INSERT INTO jawaban_siswa (siswa_id, soal_id, jawaban_terpilih) 
               VALUES (?, ?, ?) 
               ON CONFLICT(siswa_id, soal_id) DO UPDATE SET jawaban_terpilih = ?`;
  
  db.run(sql, [siswa_id, soal_id, jawaban, jawaban], (err) => {
    if (err) return res.status(500).json({ error: err.message });
    res.json({ status: 'Tersimpan' });
  });
});

app.post('/api/catat-pelanggaran', (req, res) => {
  const { siswa_id } = req.body;
  db.run("UPDATE sesi_ujian SET pelanggaran = pelanggaran + 1 WHERE siswa_id = ?", [siswa_id], (err) => {
    if (err) return res.status(500).json({ error: err.message });
    res.json({ status: 'Pelanggaran dicatat' });
  });
});

app.post('/api/submit-ujian', (req, res) => {
  const { siswa_id } = req.body;

  db.all("SELECT s.id as soal_id, s.kunci, j.jawaban_terpilih FROM soal s LEFT JOIN jawaban_siswa j ON s.id = j.soal_id AND j.siswa_id = ?", 
  [siswa_id], (err, rows) => {
    if (err) return res.status(500).json({ error: err.message });

    let jumlahBenar = 0;
    let totalSoal = rows.length;

    rows.forEach(r => {
      if (r.jawaban_terpilih === r.kunci) {
        jumlahBenar++;
      }
    });

    let jumlahSalah = totalSoal - jumlahBenar;
    let totalNilai = totalSoal > 0 ? (jumlahBenar / totalSoal) * 100 : 0;

    db.run(`INSERT INTO nilai (siswa_id, jumlah_benar, jumlah_salah, total_nilai) 
            VALUES (?, ?, ?, ?) 
            ON CONFLICT(siswa_id) DO UPDATE SET total_nilai = ?`, 
            [siswa_id, jumlahBenar, jumlahSalah, totalNilai, totalNilai]);

    db.run("UPDATE sesi_ujian SET status = 'SELESAI' WHERE siswa_id = ?", [siswa_id]);

    res.json({ 
      message: 'Ujian selesai', 
      jumlah_benar: jumlahBenar, 
      jumlah_salah: jumlahSalah, 
      nilai: totalNilai 
    });
  });
});

// --- API ADMIN ---

app.get('/api/admin/monitoring', (req, res) => {
  const query = `
    SELECT 
      s.id, s.username, s.nama, s.kelas, s.password,
      COALESCE(su.status, 'BELUM MULAI') as status_ujian,
      COALESCE(su.pelanggaran, 0) as total_pelanggaran,
      COALESCE(n.total_nilai, 0) as nilai_akhir,
      COALESCE(n.jumlah_benar, 0) as jumlah_benar,
      COALESCE(n.jumlah_salah, 0) as jumlah_salah
    FROM siswa s
    LEFT JOIN sesi_ujian su ON s.id = su.siswa_id
    LEFT JOIN nilai n ON s.id = n.siswa_id
    ORDER BY s.kelas ASC, s.nama ASC
  `;
  db.all(query, [], (err, rows) => {
    if (err) return res.status(500).json({ error: err.message });
    res.json(rows);
  });
});

app.post('/api/admin/tambah-siswa', (req, res) => {
  const { nama, kelas } = req.body;
  insertSiswaAuto(nama, kelas, (result) => {
    if (!result) return res.status(500).json({ error: 'Gagal menambah siswa' });
    res.json({ message: 'Siswa berhasil ditambahkan', username: result.username, password: result.password });
  });
});

// API IMPORT DATA SISWA (EXCEL, WORD, PDF)
app.post('/api/admin/upload-siswa', upload.single('file_siswa'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'File tidak ditemukan' });

  const ext = path.extname(req.file.originalname).toLowerCase();
  let dataSiswaList = [];

  try {
    if (ext === '.xlsx' || ext === '.xls') {
      const workbook = XLSX.readFile(req.file.path);
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const json = XLSX.utils.sheet_to_json(sheet);

      json.forEach(row => {
        const nama = row['Nama'] || row['NAMA'] || row['Nama Siswa'] || Object.values(row)[0];
        const kelas = row['Kelas'] || row['KELAS'] || Object.values(row)[1] || '9A';
        if (nama) dataSiswaList.push({ nama, kelas });
      });

    } else if (ext === '.docx') {
      const result = await mammoth.convertToHtml({ path: req.file.path });
      const html = result.value;
      const rowRegex = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
      const cellRegex = /<td[^>]*>([\s\S]*?)<\/td>/gi;
      const rows = [...html.matchAll(rowRegex)];

      for (let i = 1; i < rows.length; i++) {
        const cells = [...rows[i][1].matchAll(cellRegex)].map(c => c[1].replace(/<[^>]+>/g, '').trim());
        if (cells.length >= 2) {
          dataSiswaList.push({ nama: cells[0], kelas: cells[1] });
        }
      }

    } else if (ext === '.pdf') {
      const dataBuffer = fs.readFileSync(req.file.path);
      const pdfData = await pdfParse(dataBuffer);
      const lines = pdfData.text.split('\n');

      lines.forEach(line => {
        const cleanLine = line.trim();
        if (cleanLine.length > 0) {
          if (cleanLine.includes(',')) {
            const parts = cleanLine.split(',');
            dataSiswaList.push({ nama: parts[0].trim(), kelas: parts[1].trim() });
          } else if (cleanLine.includes('\t')) {
            const parts = cleanLine.split('\t');
            dataSiswaList.push({ nama: parts[0].trim(), kelas: parts[1].trim() });
          }
        }
      });
    } else {
      fs.unlinkSync(req.file.path);
      return res.status(400).json({ error: 'Format file tidak didukung. Gunakan .xlsx, .docx, atau .pdf' });
    }

    fs.unlinkSync(req.file.path);

    if (dataSiswaList.length === 0) {
      return res.status(400).json({ error: 'Tidak dapat membaca data siswa dari file tersebut.' });
    }

    let suksesCount = 0;
    let completed = 0;

    dataSiswaList.forEach(s => {
      insertSiswaAuto(s.nama, s.kelas, (resSiswa) => {
        if (resSiswa) suksesCount++;
        completed++;

        if (completed === dataSiswaList.length) {
          res.json({ message: `Berhasil mengimpor ${suksesCount} data siswa dari file!` });
        }
      });
    });

  } catch (err) {
    if (fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
    res.status(500).json({ error: 'Gagal memproses file: ' + err.message });
  }
});

// API IMPORT SOAL DARI WORD
app.post('/api/admin/upload-word', upload.single('file_word'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'File tidak ditemukan' });

  try {
    const result = await mammoth.convertToHtml({ path: req.file.path });
    const html = result.value;
    fs.unlinkSync(req.file.path);

    const tableRegex = /<table[^>]*>([\s\S]*?)<\/table>/gi;
    const tables = [...html.matchAll(tableRegex)];

    if (tables.length === 0) {
      return res.status(400).json({ error: 'Tidak ada tabel yang ditemukan dalam file Word!' });
    }

    const rowRegex = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
    const cellRegex = /<td[^>]*>([\s\S]*?)<\/td>/gi;

    const rowsSoal = [...tables[0][1].matchAll(rowRegex)];
    let kunciMap = {};

    if (tables.length > 1) {
      const rowsKunci = [...tables[1][1].matchAll(rowRegex)];
      rowsKunci.slice(1).forEach(row => {
        const cells = [...row[1].matchAll(cellRegex)].map(c => c[1].replace(/<[^>]+>/g, '').trim());
        if (cells.length >= 2) {
          kunciMap[cells[0]] = cells[1].toUpperCase();
        }
      });
    }

    let totalDiimpor = 0;
    const stmt = db.prepare("INSERT INTO soal (pertanyaan, opsi_a, opsi_b, opsi_c, opsi_d, kunci) VALUES (?, ?, ?, ?, ?, ?)");

    for (let i = 1; i < rowsSoal.length; i++) {
      const cells = [...rowsSoal[i][1].matchAll(cellRegex)].map(c => c[1].replace(/<[^>]+>/g, '').trim());

      if (cells.length >= 6) {
        const no = cells[0];
        const pertanyaan = cells[1];
        const opsiA = cells[2];
        const opsiB = cells[3];
        const opsiC = cells[4];
        const opsiD = cells[5];
        let kunci = cells[6] ? cells[6].toUpperCase() : (kunciMap[no] || 'A');

        stmt.run([pertanyaan, opsiA, opsiB, opsiC, opsiD, kunci]);
        totalDiimpor++;
      }
    }

    stmt.finalize();
    res.json({ message: `Berhasil mengimpor ${totalDiimpor} soal dari Word!` });

  } catch (err) {
    res.status(500).json({ error: 'Gagal memproses file Word: ' + err.message });
  }
});

app.post('/api/admin/tambah-soal', (req, res) => {
  const { pertanyaan, opsi_a, opsi_b, opsi_c, opsi_d, kunci } = req.body;
  db.run("INSERT INTO soal (pertanyaan, opsi_a, opsi_b, opsi_c, opsi_d, kunci) VALUES (?, ?, ?, ?, ?, ?)",
    [pertanyaan, opsi_a, opsi_b, opsi_c, opsi_d, kunci],
    function(err) {
      if (err) return res.status(500).json({ error: err.message });
      res.json({ message: 'Soal berhasil disimpan' });
    }
  );
});

app.get('/api/admin/bank-soal', (req, res) => {
  db.all("SELECT * FROM soal ORDER BY id DESC", [], (err, rows) => {
    if (err) return res.status(500).json({ error: err.message });
    res.json(rows);
  });
});

app.get('/api/rekap-excel', async (req, res) => {
  const workbook = new exceljs.Workbook();
  const worksheet = workbook.addWorksheet('Hasil Ujian');

  worksheet.columns = [
    { header: 'No', key: 'no', width: 5 },
    { header: 'Username', key: 'username', width: 15 },
    { header: 'Nama Siswa', key: 'nama', width: 30 },
    { header: 'Kelas', key: 'kelas', width: 12 },
    { header: 'Password', key: 'password', width: 12 },
    { header: 'Status', key: 'status', width: 15 },
    { header: 'Pelanggaran', key: 'pelanggaran', width: 15 },
    { header: 'Benar', key: 'jumlah_benar', width: 10 },
    { header: 'Salah', key: 'jumlah_salah', width: 10 },
    { header: 'Total Nilai', key: 'total_nilai', width: 15 }
  ];

  const query = `
    SELECT 
      s.username, s.nama, s.kelas, s.password,
      COALESCE(su.status, 'BELUM MULAI') as status,
      COALESCE(su.pelanggaran, 0) as pelanggaran,
      COALESCE(n.jumlah_benar, 0) as jumlah_benar,
      COALESCE(n.jumlah_salah, 0) as jumlah_salah,
      COALESCE(n.total_nilai, 0) as total_nilai
    FROM siswa s
    LEFT JOIN sesi_ujian su ON s.id = su.siswa_id
    LEFT JOIN nilai n ON s.id = n.siswa_id
  `;

  db.all(query, [], async (err, rows) => {
    if (err) return res.status(500).send('Gagal membuat rekap Excel');

    rows.forEach((row, index) => {
      worksheet.addRow({
        no: index + 1,
        username: row.username,
        nama: row.nama,
        kelas: row.kelas,
        password: row.password,
        status: row.status,
        pelanggaran: row.pelanggaran,
        jumlah_benar: row.jumlah_benar,
        jumlah_salah: row.jumlah_salah,
        total_nilai: row.total_nilai
      });
    });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=Rekap_Nilai_Ujian.xlsx');

    await workbook.xlsx.write(res);
    res.end();
  });
});

// Server Listening
app.listen(PORT, () => {
  console.log(`===================================================`);
  console.log(` Server Ujian + Import Siswa Multi-Format Aktif!`);
  console.log(` Akses Siswa  : http://localhost:${PORT}`);
  console.log(` Akses Guru   : http://localhost:${PORT}/admin.html`);
  console.log(`===================================================`);
});