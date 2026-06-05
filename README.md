# Moodle Readability Analyzer v3.0

Alat analisis **readability kode PHP** untuk proyek Moodle secara otomatis. Menggunakan [lizard](https://github.com/terryyin/lizard) untuk ekstraksi fungsi yang akurat, lalu menghitung skor keterbacaan berdasarkan dua metrik utama: **panjang fungsi (LOC)** dan **kepadatan komentar (Comment Density)** dengan bobot **60:40**.

---

## Fitur

- ✅ Scan seluruh folder Moodle secara rekursif
- ✅ Filter otomatis folder/file yang tidak relevan (`vendor`, `tests`, `node_modules`, dll.)
- ✅ Dukungan sampling acak per modul (reproducible via `--seed`)
- ✅ Export hasil ke **Excel (.xlsx)** dengan 4 sheet: Results, Summary, Module Comparison, Low Readability
- ✅ Warna otomatis pada sel kategori (🟢 High / 🟡 Medium / 🔴 Low)
- ✅ Generate **8 grafik visualisasi** dalam satu perintah
- ✅ Auto-install dependensi Python yang belum terpasang

---

## Formula Skor

Skor readability dihitung dari dua komponen dengan bobot:

| Komponen | Bobot | Keterangan |
|---|---|---|
| Function Length Score (LOC) | **60%** | Skor 100 jika LOC ≤ 20; 0 jika LOC ≥ 100 |
| Comment Density Score (CD) | **40%** | Skor 100 jika CD ≥ 25%; proporsional di bawahnya |

```
Readability Score = 0.6 × LOC_Score + 0.4 × CD_Score
```

### Kategori

| Skor | Kategori |
|---|---|
| ≥ 80 | 🟢 High |
| 50 – 79 | 🟡 Medium |
| < 50 | 🔴 Low |

---

## Instalasi

Direkomendasikan menggunakan Python 3.8+.

```bash
# Clone repositori
git clone https://github.com/username/moodle-readability-analyzer.git
cd moodle-readability-analyzer

# (Opsional) Buat virtual environment
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate

# Install dependensi
pip install pandas matplotlib seaborn openpyxl lizard
```

> **Catatan:** Script juga dapat menginstall dependensi secara otomatis saat pertama kali dijalankan.

---

## Cara Penggunaan

### Analisis semua folder (tanpa sampling)
```bash
python analisis_bobot_60_40.py --path "C:/xampp/htdocs/moodle/public"
```

### Analisis modul tertentu dengan sampling
```bash
python analisis_bobot_60_40.py \
  --path "C:/xampp/htdocs/moodle/public" \
  --modules lib mod search \
  --sample 100
```

### Analisis modul tertentu tanpa sampling
```bash
python analisis_bobot_60_40.py \
  --path "C:/xampp/htdocs/moodle/public" \
  --modules lib mod
```

### Lewati pembuatan grafik
```bash
python analisis_bobot_60_40.py --path "..." --no-viz
```

### Tentukan nama file output dan folder grafik
```bash
python analisis_bobot_60_40.py \
  --path "..." \
  --output hasil_analisis.xlsx \
  --viz-dir grafik_output
```

---

## Parameter CLI

| Parameter | Singkat | Default | Deskripsi |
|---|---|---|---|
| `--path` | `-p` | *(wajib)* | Path ke folder `moodle/public` |
| `--modules` | `-m` | semua subfolder | Modul spesifik yang ingin dianalisis |
| `--sample` | `-s` | semua file | Jumlah file yang disampling acak per modul |
| `--seed` | | `42` | Random seed untuk reproduksibilitas |
| `--output` | `-o` | `moodle_readability_results.xlsx` | Nama file Excel output |
| `--viz-dir` | `-v` | `visualizations` | Folder penyimpanan grafik |
| `--no-viz` | | `False` | Lewati pembuatan grafik |

---

## Output

### File Excel (`.xlsx`)
Terdiri dari 4 sheet:

| Sheet | Isi |
|---|---|
| **Results** | Data lengkap per fungsi dengan warna kategori |
| **Summary** | Statistik ringkasan keseluruhan |
| **Module Comparison** | Perbandingan rata-rata skor antar modul |
| **Low Readability** | Daftar fungsi dengan kategori Low |

### Kolom pada sheet Results

`Module`, `File`, `File Path`, `Function`, `Start Line`, `End Line`, `Total Lines`, `Code Lines (LOC)`, `Comment Lines`, `Blank Lines`, `Comment Density (%)`, `Function Length Score`, `Comment Density Score`, `Readability Score`, `Category`

### Grafik Visualisasi (`visualizations/`)

| File | Deskripsi |
|---|---|
| `1_readability_distribution.png` | Histogram distribusi Readability Score |
| `2_readability_by_module.png` | Bar chart rata-rata skor per modul |
| `3_category_distribution.png` | Pie chart proporsi kategori High/Medium/Low |
| `4_loc_vs_comments.png` | Scatter plot LOC vs Comment Density |
| `5_module_boxplot.png` | Box plot variasi skor per modul |
| `6_function_length_distribution.png` | Histogram distribusi panjang fungsi |
| `7_comment_density_distribution.png` | Histogram distribusi comment density |
| `8_category_stacked_by_module.png` | Stacked bar proporsi kategori per modul |

---

## Folder yang Di-skip

Script secara otomatis mengabaikan folder dan file berikut:

`vendor`, `node_modules`, `tests`, `yui`, `amd`, `cache`, `local`, `behat`, `fixtures`, `coverage`, `pix`, `lang`, `theme`, `userpix`, `.git`

File PHP dengan nama diawali `test_` atau diakhiri `_test.php` juga di-skip.

---

## Dependensi

| Library | Kegunaan |
|---|---|
| [pandas](https://pandas.pydata.org/) | Manipulasi data & export Excel |
| [matplotlib](https://matplotlib.org/) | Pembuatan grafik |
| [seaborn](https://seaborn.pydata.org/) | Styling grafik |
| [openpyxl](https://openpyxl.readthedocs.io/) | Formatting Excel |
| [lizard](https://github.com/terryyin/lizard) | Ekstraksi metrik fungsi PHP |

---

## Lisensi

MIT License — bebas digunakan dan dimodifikasi untuk keperluan riset maupun pengembangan.
