# Transkrip YouTube Otomatis

Tool CLI berbasis Python untuk mendownload audio dari video/playlist YouTube dan mentranskripsinya menjadi teks menggunakan **OpenAI Whisper AI**.

## Fitur

- Mendukung **video tunggal** dan **playlist** YouTube
- Model Whisper yang bisa dipilih: `tiny`, `base`, `small`, `medium`, `large`
- Output multi-format:
  - **Excel (.xlsx)** — 3 sheet (Segmen, Paragraf, Statistik) + style formatting
  - **TXT** — format segmen, paragraf, dan teks kontinyu
  - **CSV** — data segmen
  - **SRT** — subtitle
- Folder output terorganisir dengan timestamp dan judul video
- Progress logging berwarna dengan emoji (Bahasa Indonesia)

## Persyaratan Sistem

| Kebutuhan | Detail |
|---|---|
| Python | 3.8 atau lebih baru |
| FFmpeg | Harus terinstall dan tersedia di PATH |
| GPU (opsional) | NVIDIA CUDA untuk transkripsi lebih cepat |
| Internet | Untuk download YouTube dan model Whisper |
| Disk | ~3 GB untuk model Whisper `large` |

---

## Cara Install & Menjalankan

### A. Menggunakan Anaconda (Direkomendasikan)

Buka **Anaconda Prompt**, lalu jalankan:

```bash
# 1. Clone atau masuk ke folder proyek
cd D:\node-rest-api\proyek_transkrip_yt

# 2. Buat environment Conda baru
conda create -n transkrip python=3.10 -y

# 3. Aktifkan environment
conda activate transkrip

# 4. Install Python dependencies
pip install -r requirements.txt

# 5. Install FFmpeg lewat Conda
conda install -c conda-forge ffmpeg -y

# 6. Jalankan
python transkrip_auto.py
```

### B. Menggunakan pip + venv

Buka **Command Prompt** atau **PowerShell**:

```bash
# 1. Masuk ke folder proyek
cd D:\node-rest-api\proyek_transkrip_yt

# 2. Buat virtual environment
python -m venv venv

# 3. Aktifkan
venv\Scripts\activate

# 4. Install dependencies
pip install -r requirements.txt

# 5. Install FFmpeg secara manual
#    Download dari https://ffmpeg.org/download.html
#    Tambahkan folder bin\ ke PATH Windows

# 6. Jalankan
python transkrip_auto.py
```

---

## Cara Pakai

Jalankan `python transkrip_auto.py`, lalu ikuti prompt:

1. **Pilih model Whisper**

   | Model | Ukuran | Kecepatan | Akurasi |
   |---|---|---|---|
   | `tiny` | ~75 MB | Sangat cepat | Rendah |
   | `base` | ~150 MB | Cepat | Cukup |
   | `small` | ~500 MB | Sedang | Baik |
   | `medium` | ~1.5 GB | Lambat | Sangat baik |
   | `large` | ~3 GB | Sangat lambat | Terbaik |

   > Model akan otomatis didownload saat pertama kali digunakan.

2. **Masukkan URL YouTube** — bisa video tunggal atau playlist

3. Script akan otomatis:
   - Mendeteksi apakah URL adalah playlist atau video tunggal
   - Mendownload audio
   - Mentranskripsi dengan Whisper
   - Menyimpan hasil di folder `transcripts/YYYYMMDD_HHMMSS_JudulVideo/`

---

## Output

Setiap video menghasilkan file di folder `transcripts/<timestamp>_<judul>/`:

```
transcripts/
└── 20250708_143025_Judul_Video/
    ├── Judul_Video.xlsx          # 3 sheet: Segmen, Paragraf, Statistik
    ├── Judul_Video.txt            # Format segmen (timestamp + teks)
    ├── Judul_Video_paragraf.txt   # Format paragraf
    ├── Judul_Video_kontinyu.txt   # Teks kontinyu tanpa timestamp
    ├── Judul_Video.csv            # CSV
    ├── Judul_Video.srt            # Subtitle
    └── run_info.txt               # Info proses (model, durasi, dll)
```

Untuk **playlist**, ada tambahan file **Ringkasan_Playlist.xlsx** berisi 4 sheet gabungan seluruh video.

---

## Struktur Proyek

```
proyek_transkrip_yt/
├── transkrip_auto.py      # Versi utama (single video + playlist)
├── transkrip_kontinyu.py  # Versi sederhana (teks kontinyu saja)
├── requirements.txt       # Python dependencies
├── .gitignore
└── README.md
```

---

## Dependency

| Package | Versi Minimal |
|---|---|
| `yt-dlp` | 2024.0.0 |
| `openai-whisper` | 20231117 |
| `torch` | 2.0.0 |
| `pandas` | 2.0.0 |
| `openpyxl` | 3.1.0 |
| `ffmpeg` | (system) |
