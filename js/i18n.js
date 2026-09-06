(function () {
  "use strict";

  var STORE = {
    id: {
      "nav.home": "Beranda",
      "nav.docs": "Dokumentasi",
      "nav.download": "Unduh",

      "hero.badge": "Rilis Resmi v1.0.0",
      "hero.title": "Universal Media & Document Suite",
      "hero.subtitle": "Utilitas desktop lokal serba bisa yang memproses foto RAW, video, audio, dokumen, dan unduhan media — tanpa mengirim data Anda ke internet.",
      "hero.cta.download": "Unduh Aplikasi",
      "hero.cta.docs": "Baca Dokumentasi",

      "hero.stat.1.t": "Isolasi Operasi",
      "hero.stat.1.v": "100% Offline Core",
      "hero.stat.2.t": "Backend Engine",
      "hero.stat.2.v": "FFmpeg + LibRaw",
      "hero.stat.3.t": "Interaksi Berkas",
      "hero.stat.3.v": "Drag-and-Drop (DND)",
      "hero.stat.4.t": "Penanganan Error",
      "hero.stat.4.v": "Log File Otomatis",

      "dl.title": "Unduh untuk Platform Anda",
      "dl.badge": "Rilis Terbaru",
      "dl.platform.detect": "Platform terdeteksi",
      "dl.auto": "Dari peramban Anda. Bisa diganti kapan saja.",
      "dl.win.name": "Windows",
      "dl.linux.name": "Linux",
      "dl.mac.name": "macOS",
      "dl.win.spec": "Windows 10 / 11 · 64-bit",
      "dl.linux.spec": "Linux x86_64 · glibc 2.31+",
      "dl.mac.spec": "macOS 12+ · Apple Silicon & Intel",
      "dl.win.size": "± 164 MB",
      "dl.linux.size": "± 151 MB",
      "dl.mac.size": "± 140 MB",
      "dl.btn.win": "Unduh .EXE",
      "dl.btn.linux": "Unduh .AppImage",
      "dl.btn.mac": "Unduh .DMG",
      "dl.note": "Unduhan disajikan langsung dari GitHub Release. Simpan file lalu jalankan.",
      "dl.checksum": "Periksa checksum SHA-256",
      "dl.other": "Pilih platform lain",

      "feat.title": "Mengapa Memilih Suite Ini?",
      "feat.1.t": "Isolasi I/O Lokal",
      "feat.1.d": "Semua konversi berjalan di perangkat Anda. Dokumen, foto, dan media tidak pernah dikirim ke server awan mana pun.",
      "feat.2.t": "Worker Threads Mandiri",
      "feat.2.d": "Proses berat dijalankan pada tread terpisah sehingga antarmuka tetap responsif dan tidak pernah membeku.",
      "feat.3.t": "Ketahanan Batch (Fault Tolerant)",
      "feat.3.d": "File yang korup tidak menghentikan antrean. Setiap kegagalan dicatat ke file log lokal untuk audit mudah.",

      "mod.title": "Modul Utama",
      "mod.subtitle": "Lima modul terintegrasi dalam satu aplikasi desktop.",
      "mod.1.t": "Foto & Kamera RAW",
      "mod.1.e": "Pillow + rawpy (LibRaw)",
      "mod.1.d": "Baca file mentah kamera profesional (CR2, NEF, ARW, DNG, dll) lalu konversi massal ke JPG, PNG, WEBP, atau PDF dengan opsi resize.",
      "mod.2.t": "Video Transcoder",
      "mod.2.e": "FFmpeg CLI Pipeline",
      "mod.2.d": "Perkecil resolusi ratusan video sekaligus dengan preset 1080p/720p/480p/360p atau dimensi custom, tanpa merusak rasio aspek.",
      "mod.3.t": "Ekstraktor Audio",
      "mod.3.e": "FFmpeg Demuxer",
      "mod.3.d": "Ekstrak trek audio dari video ke MP3, AAC, M4A, WAV, FLAC, atau OGG dengan bitrate hingga 320 kbps.",
      "mod.4.t": "Dokumen & PDF",
      "mod.4.e": "Win32 COM / LibreOffice / pypdf",
      "mod.4.d": "Konversi dokumen Office, gabung & pisah PDF, dan ekstrak teks — sepenuhnya lokal.",
      "mod.5.t": "Media Downloader",
      "mod.5.e": "yt-dlp + FFmpeg",
      "mod.5.d": "Unduh video/audio dari YouTube, TikTok, Instagram, X, dan SoundCloud langsung ke folder Anda.",
      "mod.more": "Baca dokumentasi lengkap",

      "matrix.title": "Matriks Kompatibilitas Sistem Operasi",
      "matrix.subtitle": "Rancangan lintas platform",
      "matrix.c0": "Komponen Arsitektur",
      "matrix.c1": "Windows (x64)",
      "matrix.c2": "Linux (x86_64)",
      "matrix.c3": "macOS",
      "matrix.r1": "Format Distribusi",
      "matrix.r1.w": "Setup.exe / Portable",
      "matrix.r1.linux": ".AppImage / Standalone",
      "matrix.r1.mac": ".DMG / .app Bundle",
      "matrix.r2": "Kebutuhan FFmpeg",
      "matrix.r2.w": "Bundled (otomatis ikut)",
      "matrix.r2.linux": "Paket sistem (apt install ffmpeg)",
      "matrix.r2.mac": "Homebrew / ikut aplikasi",
      "matrix.r3": "Engine Dokumen Office",
      "matrix.r3.w": "MS Office COM Dispatch",
      "matrix.r3.linux": "LibreOffice Headless",
      "matrix.r3.mac": "LibreOffice Headless",
      "matrix.r4": "Buka Folder Hasil",
      "matrix.r4.w": "os.startfile(dir)",
      "matrix.r4.linux": "xdg-open [dir]",
      "matrix.r4.mac": "open [dir]",

      "verify.title": "Verifikasi Integritas Binary",
      "verify.subtitle": "Checksum SHA-256",
      "verify.desc": "Bandingkan nilai checksum SHA-256 sebelum menjalankan file hasil unduhan untuk memastikan integritas rilis.",
      "verify.copy": "Salin Perintah",
      "verify.copied": "Perintah checksum disalin ke clipboard.",

      "help.title": "Pemecahan Masalah",
      "help.subtitle": "Panduan cepat mengatasi kendala umum saat menjalankan aplikasi.",
      "help.1.t": "Windows SmartScreen Block",
      "help.1.d": "Peringatan muncul karena binary dikompilasi independen (belum ditandatangani sertifikat EV berbayar). Klik More info lalu pilih Run anyway.",
      "help.2.t": "Linux — Izin Eksekusi",
      "help.2.d": "Jika AppImage tidak merespons, buka terminal dan tambahkan izin eksekusi dengan perintah di bawah.",
      "help.3.t": "macOS — Gatekeeper Quarantine",
      "help.3.d": "Jika timbul pesan App is damaged and cannot be opened, hapus atribut karantina melalui Terminal.",

      "faq.title": "Pertanyaan Umum",
      "faq.1.q": "Apakah aplikasi membutuhkan koneksi internet?",
      "faq.1.a": "Hanya untuk modul Media Downloader. Konversi foto, video, audio, dan dokumen berjalan sepenuhnya offline.",
      "faq.2.q": "Mengapa versi Linux/macOS perlu FFmpeg tambahan?",
      "faq.2.a": "Untuk menjaga ukuran binary tetap kecil, FFmpeg pada Linux/macOS diambil dari sistem. Windows sudah menyertakan FFmpeg.",
      "faq.3.q": "Apakah data saya dikirim ke server?",
      "faq.3.a": "Tidak. Seluruh proses berjalan lokal di perangkat Anda, kecuali Anda aktif mengunduh dari internet melalui modul downloader.",
      "faq.4.q": "Bagaimana cara memperbarui aplikasi?",
      "faq.4.a": "Unduh rilis terbaru dari halaman ini — tautan unduhan selalu mengarah ke rilis terbaru di GitHub Releases.",

      "footer.tagline": "Media, dokumen, dan unduhan dalam satu aplikasi desktop — lokal dan aman.",
      "footer.engine": "Local-First Engine",
      "footer.track": "Zero External Tracking",
      "footer.rights": "© 2026 Universal Media & Document Suite. Hak cipta dilindungi.",

      "docs.title": "Dokumentasi",
      "docs.subtitle": "Panduan lengkap penggunaan Universal Media & Document Suite di semua platform.",
      "docs.cta": "Unduh Aplikasi",
      "docs.back": "Kembali ke Beranda",
      "docs.toc.start": "Mulai",
      "docs.toc.quick": "Panduan Cepat",
      "docs.toc.req": "Persyaratan Sistem",
      "docs.toc.modules": "Dokumentasi Modul",
      "docs.toc.verify": "Verifikasi Checksum",
      "docs.toc.help": "Pemecahan Masalah",
      "docs.toc.faq": "FAQ",

      "dqs.title": "Panduan Cepat",
      "dqs.step.1.t": "1 · Unduh binary",
      "dqs.step.1.d": "Pilih platform Anda lalu unduh berkas rilis terbaru sesuai sistem operasi.",
      "dqs.step.2.t": "2 · Siapkan FFmpeg",
      "dqs.step.2.d": "Windows sudah ter-bundle. Linux dan macOS perlu FFmpeg dari sistem agar modul video/audio berfungsi.",
      "dqs.step.3.t": "3 · Izinkan eksekusi",
      "dqs.step.3.d": "Linux: chmod +x. macOS: hapus atribut karantina bila perlu. Windows: pilih Run anyway jika SmartScreen muncul.",
      "dqs.step.4.t": "4 · Jalankan aplikasi",
      "dqs.step.4.d": "Buka aplikasi, tarik file ke area drag-and-drop, pilih format tujuan, lalu klik tombol mulai.",

      "dreq.title": "Persyaratan Sistem",
      "dreq.desc": "Kebutuhan minimum agar seluruh modul berfungsi penuh.",
      "dreq.1": "Windows 10/11 64-bit, Linux x86_64 (glibc 2.31+), atau macOS 12+.",
      "dreq.2": "LibreOffice (Linux/macOS) atau Microsoft Office (Windows) untuk konversi dokumen Office.",
      "dreq.3": "FFmpeg tersedia di PATH sistem (khusus Linux/macOS).",
      "dreq.4": "RAM 2 GB (disarankan 4 GB) dan ruang disk yang cukup untuk file hasil.",

      "dmod.title": "Dokumentasi Modul",
      "dmod.1.t": "01 · Foto & Kamera RAW",
      "dmod.1.ov": "Mengekstrak data piksel mentah (bayer pattern) dari sensor kamera digital tanpa kompresi JPEG lossy bawaan vendor.",
      "dmod.1.in": "Input: CR2, CR3, CRW, NEF, NRW, ARW, SRF, SR2, RAF, ORF, RW2, PEF, PTX, DNG, X3F, ERF, MEF, MOS, KDC, DCR (RAW) + JPG, PNG, WEBP, BMP, TIFF, GIF, ICO, TGA, PPM, PCX, HEIC, HEIF, PSD (standar).",
      "dmod.1.out": "Output: JPG, PNG, WEBP, PDF, TIFF, BMP, ICO (otomatis dibatasi ≤ 256px), TGA.",
      "dmod.1.note": "Format tanpa transparansi (JPG/BMP/PDF) otomatis mengonversi RGBA → RGB untuk mencegah error. Resize memakai interpolasi LANCZOS untuk ketajaman maksimal.",

      "dmod.2.t": "02 · Video Transcoder & Resizer",
      "dmod.2.ov": "Transkoding video massal untuk mengecilkan dimensi dan bitrate sambil menjaga rasio aspek.",
      "dmod.2.codec": "Video: H.264 (libx264) · PixFmt yuv420p · Preset ultrafast · CRF 24. Audio: AAC stereo 128 kbps.",
      "dmod.2.note": "Filter pad menjaga dimensi selalu genap agar kompatibel dengan codec H.264.",

      "dmod.3.t": "03 · Audio Stream Extractor",
      "dmod.3.ov": "Mengekstrak audio langsung dari kontainer video (MP4, MKV, MOV, WebM) dengan memangkas stream video (-vn).",
      "dmod.3.fmt": "Format: MP3 (libmp3lame), AAC, M4A, WAV (pcm_s16le), FLAC, OGG.",
      "dmod.3.note": "Bitrate: 320/192/128/64 kbps. Mode 64 kbps ideal untuk voice note (hemat hingga 70% ukuran).",

      "dmod.4.t": "04 · Dokumen & PDF Tools",
      "dmod.4.ov": "Konversi silang dokumen Office, plus alat gabung dan pisah PDF yang berjalan sepenuhnya lokal.",
      "dmod.4.win": "Windows: mengontrol Microsoft Office headless lewat Win32 COM (Word, Excel, PowerPoint).",
      "dmod.4.lo": "Linux/macOS: automasi LibreOffice headless (soffice --headless --convert-to).",
      "dmod.4.pdf": "PDF → DOCX memakai engine pdf2docx. Gabung/pisah halaman memakai pypdf dengan sintaks rentang 1-5 atau 2,4,7.",

      "dmod.5.t": "05 · Network Media Downloader",
      "dmod.5.ov": "Mengunduh media publik dari YouTube, TikTok, Instagram, X, dan SoundCloud tanpa situs downloader pihak ketiga.",
      "dmod.5.note": "Video resolusi tinggi dan audio digabung (mux) via FFmpeg lokal. Mode musik mengekstrak langsung ke MP3 192 kbps.",

      "dverify.title": "Verifikasi Checksum",
      "dverify.subtitle": "SHA-256",
      "dverify.desc": "Bandingkan nilai SHA-256 berkas yang Anda unduh dengan yang diterbitkan pada rilis.",

      "dhelp.title": "Pemecahan Masalah",
      "dhelp.subtitle": "Panduan lengkap mengatasi galat eksekusi di setiap platform.",

      "err.404.title": "404 — Halaman tidak ditemukan",
      "err.404.desc": "Halaman yang Anda cari tidak ada atau telah dipindahkan.",
      "err.404.home": "Kembali ke Beranda"
    },

    en: {
      "nav.home": "Home",
      "nav.docs": "Documentation",
      "nav.download": "Download",

      "hero.badge": "Official Release v1.0.0",
      "hero.title": "Universal Media & Document Suite",
      "hero.subtitle": "An all-in-one local desktop utility that processes RAW photos, video, audio, documents, and media downloads — without sending your data over the internet.",
      "hero.cta.download": "Download App",
      "hero.cta.docs": "Read the Docs",

      "hero.stat.1.t": "Operation Isolation",
      "hero.stat.1.v": "100% Offline Core",
      "hero.stat.2.t": "Backend Engine",
      "hero.stat.2.v": "FFmpeg + LibRaw",
      "hero.stat.3.t": "File Interaction",
      "hero.stat.3.v": "Drag-and-Drop (DND)",
      "hero.stat.4.t": "Error Handling",
      "hero.stat.4.v": "Automatic Log Files",

      "dl.title": "Download for Your Platform",
      "dl.badge": "Latest Release",
      "dl.platform.detect": "Detected platform",
      "dl.auto": "Detected from your browser. You can switch at any time.",
      "dl.win.name": "Windows",
      "dl.linux.name": "Linux",
      "dl.mac.name": "macOS",
      "dl.win.spec": "Windows 10 / 11 · 64-bit",
      "dl.linux.spec": "Linux x86_64 · glibc 2.31+",
      "dl.mac.spec": "macOS 12+ · Apple Silicon & Intel",
      "dl.win.size": "± 164 MB",
      "dl.linux.size": "± 151 MB",
      "dl.mac.size": "± 140 MB",
      "dl.btn.win": "Download .EXE",
      "dl.btn.linux": "Download .AppImage",
      "dl.btn.mac": "Download .DMG",
      "dl.note": "Downloads are served directly from GitHub Releases. Save the file and run it.",
      "dl.checksum": "Check SHA-256 checksum",
      "dl.other": "Choose another platform",

      "feat.title": "Why Choose This Suite?",
      "feat.1.t": "Local I/O Isolation",
      "feat.1.d": "Every conversion runs on your device. Documents, photos, and media are never sent to any cloud server.",
      "feat.2.t": "Dedicated Worker Threads",
      "feat.2.d": "Heavy tasks run on separate threads so the interface stays responsive and never freezes.",
      "feat.3.t": "Fault-tolerant Batch",
      "feat.3.d": "A corrupt file does not stop the queue. Each failure is written to a local log file for easy auditing.",

      "mod.title": "Core Modules",
      "mod.subtitle": "Five modules integrated into a single desktop application.",
      "mod.1.t": "Photo & Camera RAW",
      "mod.1.e": "Pillow + rawpy (LibRaw)",
      "mod.1.d": "Read professional camera raw files (CR2, NEF, ARW, DNG, etc.) and batch-convert to JPG, PNG, WEBP, or PDF with resize options.",
      "mod.2.t": "Video Transcoder",
      "mod.2.e": "FFmpeg CLI Pipeline",
      "mod.2.d": "Shrink hundreds of videos at once with 1080p/720p/480p/360p presets or custom dimensions, without distorting the aspect ratio.",
      "mod.3.t": "Audio Extractor",
      "mod.3.e": "FFmpeg Demuxer",
      "mod.3.d": "Extract audio tracks from videos to MP3, AAC, M4A, WAV, FLAC, or OGG with bitrates up to 320 kbps.",
      "mod.4.t": "Documents & PDF",
      "mod.4.e": "Win32 COM / LibreOffice / pypdf",
      "mod.4.d": "Convert Office documents, merge & split PDFs, and extract text — fully local.",
      "mod.5.t": "Media Downloader",
      "mod.5.e": "yt-dlp + FFmpeg",
      "mod.5.d": "Download video/audio from YouTube, TikTok, Instagram, X, and SoundCloud straight to your folder.",
      "mod.more": "Read the full documentation",

      "matrix.title": "Operating System Compatibility Matrix",
      "matrix.subtitle": "Cross-platform blueprint",
      "matrix.c0": "Architecture Component",
      "matrix.c1": "Windows (x64)",
      "matrix.c2": "Linux (x86_64)",
      "matrix.c3": "macOS",
      "matrix.r1": "Distribution Format",
      "matrix.r1.w": "Setup.exe / Portable",
      "matrix.r1.linux": ".AppImage / Standalone",
      "matrix.r1.mac": ".DMG / .app Bundle",
      "matrix.r2": "FFmpeg Requirement",
      "matrix.r2.w": "Bundled (included)",
      "matrix.r2.linux": "System package (apt install ffmpeg)",
      "matrix.r2.mac": "Homebrew / bundled",
      "matrix.r3": "Office Document Engine",
      "matrix.r3.w": "MS Office COM Dispatch",
      "matrix.r3.linux": "LibreOffice Headless",
      "matrix.r3.mac": "LibreOffice Headless",
      "matrix.r4": "Open Output Folder",
      "matrix.r4.w": "os.startfile(dir)",
      "matrix.r4.linux": "xdg-open [dir]",
      "matrix.r4.mac": "open [dir]",

      "verify.title": "Binary Integrity Verification",
      "verify.subtitle": "SHA-256 checksum",
      "verify.desc": "Compare the SHA-256 checksum before running the downloaded file to ensure release integrity.",
      "verify.copy": "Copy Command",
      "verify.copied": "Checksum command copied to your clipboard.",

      "help.title": "Troubleshooting",
      "help.subtitle": "Quick guide to common issues while running the app.",
      "help.1.t": "Windows SmartScreen Block",
      "help.1.d": "The warning appears because the binary is independently compiled (not signed with a paid EV certificate). Click More info, then choose Run anyway.",
      "help.2.t": "Linux — Execution Permission",
      "help.2.d": "If the AppImage does not respond, open a terminal and add execute permission with the command below.",
      "help.3.t": "macOS — Gatekeeper Quarantine",
      "help.3.d": "If the App is damaged and cannot be opened message appears, clear the quarantine attribute through Terminal.",

      "faq.title": "Frequently Asked Questions",
      "faq.1.q": "Does the app require an internet connection?",
      "faq.1.a": "Only for the Media Downloader module. Photo, video, audio, and document conversion run fully offline.",
      "faq.2.q": "Why do the Linux/macOS builds need extra FFmpeg?",
      "faq.2.a": "To keep the binary size small, FFmpeg on Linux/macOS is taken from the system. Windows ships with FFmpeg included.",
      "faq.3.q": "Is my data sent to any server?",
      "faq.3.a": "No. Everything runs locally on your device, unless you actively download from the internet using the downloader module.",
      "faq.4.q": "How do I update the app?",
      "faq.4.a": "Download the newest release from this page — the download links always point to the latest release on GitHub Releases.",

      "footer.tagline": "Media, documents, and downloads in one desktop app — local and secure.",
      "footer.engine": "Local-First Engine",
      "footer.track": "Zero External Tracking",
      "footer.rights": "© 2026 Universal Media & Document Suite. All rights reserved.",

      "docs.title": "Documentation",
      "docs.subtitle": "Complete guide to using Universal Media & Document Suite on every platform.",
      "docs.cta": "Download App",
      "docs.back": "Back to Home",
      "docs.toc.start": "Getting Started",
      "docs.toc.quick": "Quick Start",
      "docs.toc.req": "System Requirements",
      "docs.toc.modules": "Module Documentation",
      "docs.toc.verify": "Checksum Verification",
      "docs.toc.help": "Troubleshooting",
      "docs.toc.faq": "FAQ",

      "dqs.title": "Quick Start",
      "dqs.step.1.t": "1 · Download the binary",
      "dqs.step.1.d": "Pick your platform and download the latest release binary for your operating system.",
      "dqs.step.2.t": "2 · Set up FFmpeg",
      "dqs.step.2.d": "Windows is bundled. Linux and macOS need system FFmpeg for the video/audio modules to work.",
      "dqs.step.3.t": "3 · Allow execution",
      "dqs.step.3.d": "Linux: chmod +x. macOS: clear the quarantine attribute if needed. Windows: choose Run anyway if SmartScreen appears.",
      "dqs.step.4.t": "4 · Launch the app",
      "dqs.step.4.d": "Open the app, drag files into the drop zone, choose the target format, then click the start button.",

      "dreq.title": "System Requirements",
      "dreq.desc": "Minimum requirements for every module to work fully.",
      "dreq.1": "Windows 10/11 64-bit, Linux x86_64 (glibc 2.31+), or macOS 12+.",
      "dreq.2": "LibreOffice (Linux/macOS) or Microsoft Office (Windows) for Office document conversion.",
      "dreq.3": "FFmpeg available on the system PATH (Linux/macOS only).",
      "dreq.4": "2 GB RAM (4 GB recommended) and enough disk space for output files.",

      "dmod.title": "Module Documentation",
      "dmod.1.t": "01 · Photo & Camera RAW",
      "dmod.1.ov": "Extracts raw bayer-pattern pixel data from professional camera sensors without vendor lossy JPEG compression.",
      "dmod.1.in": "Input: CR2, CR3, CRW, NEF, NRW, ARW, SRF, SR2, RAF, ORF, RW2, PEF, PTX, DNG, X3F, ERF, MEF, MOS, KDC, DCR (RAW) + JPG, PNG, WEBP, BMP, TIFF, GIF, ICO, TGA, PPM, PCX, HEIC, HEIF, PSD (standard).",
      "dmod.1.out": "Output: JPG, PNG, WEBP, PDF, TIFF, BMP, ICO (auto-clamped ≤ 256px), TGA.",
      "dmod.1.note": "Formats without transparency (JPG/BMP/PDF) automatically convert RGBA → RGB to prevent errors. Resize uses LANCZOS interpolation for maximum sharpness.",

      "dmod.2.t": "02 · Video Transcoder & Resizer",
      "dmod.2.ov": "Batch video transcoding to reduce dimensions and bitrate while preserving the aspect ratio.",
      "dmod.2.codec": "Video: H.264 (libx264) · PixFmt yuv420p · Preset ultrafast · CRF 24. Audio: AAC stereo 128 kbps.",
      "dmod.2.note": "The pad filter keeps dimensions always even so they stay compatible with the H.264 codec.",

      "dmod.3.t": "03 · Audio Stream Extractor",
      "dmod.3.ov": "Extracts audio straight from video containers (MP4, MKV, MOV, WebM) by dropping the video stream (-vn).",
      "dmod.3.fmt": "Formats: MP3 (libmp3lame), AAC, M4A, WAV (pcm_s16le), FLAC, OGG.",
      "dmod.3.note": "Bitrate: 320/192/128/64 kbps. The 64 kbps mode is ideal for voice notes (up to 70% smaller).",

      "dmod.4.t": "04 · Documents & PDF Tools",
      "dmod.4.ov": "Cross-convert Office documents, plus fully local PDF merge and split tools.",
      "dmod.4.win": "Windows: drives Microsoft Office headless through Win32 COM (Word, Excel, PowerPoint).",
      "dmod.4.lo": "Linux/macOS: LibreOffice headless automation (soffice --headless --convert-to).",
      "dmod.4.pdf": "PDF → DOCX uses the pdf2docx engine. Merge/split pages uses pypdf with range syntax 1-5 or 2,4,7.",

      "dmod.5.t": "05 · Network Media Downloader",
      "dmod.5.ov": "Downloads public media from YouTube, TikTok, Instagram, X, and SoundCloud without third-party downloader sites.",
      "dmod.5.note": "High-resolution video and audio are merged (muxed) via local FFmpeg. Music mode extracts straight to MP3 192 kbps.",

      "dverify.title": "Checksum Verification",
      "dverify.subtitle": "SHA-256",
      "dverify.desc": "Compare the SHA-256 of the file you downloaded with the one published in the release.",

      "dhelp.title": "Troubleshooting",
      "dhelp.subtitle": "Complete guide to execution errors on every platform.",

      "err.404.title": "404 — Page not found",
      "err.404.desc": "The page you are looking for does not exist or has been moved.",
      "err.404.home": "Back to Home"
    }
  };

  var storageKey = "ums-lang";

  function getLang() {
    var saved = null;
    try { saved = localStorage.getItem(storageKey); } catch (e) { saved = null; }
    if (saved === "id" || saved === "en") return saved;
    var nav = (navigator.language || "en").toLowerCase();
    return nav.indexOf("id") === 0 ? "id" : "en";
  }

  function t(key) {
    var lang = getLang();
    if (STORE[lang] && Object.prototype.hasOwnProperty.call(STORE[lang], key)) {
      return STORE[lang][key];
    }
    var other = lang === "id" ? "en" : "id";
    if (STORE[other] && Object.prototype.hasOwnProperty.call(STORE[other], key)) {
      return STORE[other][key];
    }
    return key;
  }

  function applyTranslations() {
    var nodes = document.querySelectorAll("[data-i18n]");
    for (var i = 0; i < nodes.length; i++) {
      nodes[i].textContent = t(nodes[i].getAttribute("data-i18n"));
    }
    var htmlNodes = document.querySelectorAll("[data-i18n-html]");
    for (var j = 0; j < htmlNodes.length; j++) {
      htmlNodes[j].innerHTML = t(htmlNodes[j].getAttribute("data-i18n-html"));
    }
    var lang = getLang();
    var btns = document.querySelectorAll("[data-lang-btn]");
    for (var k = 0; k < btns.length; k++) {
      var isActive = btns[k].getAttribute("data-lang-btn") === lang;
      if (isActive) {
        btns[k].classList.add("bg-[#21262d]", "text-white");
        btns[k].classList.remove("text-[#8b949e]");
      } else {
        btns[k].classList.remove("bg-[#21262d]", "text-white");
        btns[k].classList.add("text-[#8b949e]");
      }
    }
    document.documentElement.lang = lang === "id" ? "id" : "en";
  }

  function setLang(lang) {
    if (lang !== "id" && lang !== "en") lang = "en";
    try { localStorage.setItem(storageKey, lang); } catch (e) { /* ignore */ }
    applyTranslations();
    document.dispatchEvent(new CustomEvent("langchange", { detail: { lang: lang } }));
  }

  window.UMS = {
    t: t,
    getLang: getLang,
    setLang: setLang
  };

  document.addEventListener("DOMContentLoaded", function () {
    var btns = document.querySelectorAll("[data-lang-btn]");
    for (var i = 0; i < btns.length; i++) {
      btns[i].addEventListener("click", function () {
        setLang(this.getAttribute("data-lang-btn"));
      });
    }
    setLang(getLang());
  });
})();