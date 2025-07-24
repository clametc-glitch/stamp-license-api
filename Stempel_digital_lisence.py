import tkinter as tk
from tkinter import filedialog, colorchooser, Toplevel, messagebox, ttk
from PIL import Image, ImageDraw, ImageFont, ImageTk
import os
import math
import json
import win32com.client as win32
import tempfile
import datetime
import requests
import uuid
import sys
import time


# Pilihan font default (bisa dikembangkan ke list font sistem)
FONT_OPTIONS = ["arial.ttf", "times.ttf", "calibri.ttf", "comic.ttf"]

MAX_ATTEMPTS = 3
BLOCK_URL = "https://individuals-skiing-nj-addition.trycloudflare.com/api/block"
VERIFY_URL = "https://individuals-skiing-nj-addition.trycloudflare.com/api/verify"
BLOCKED_DEVICES_FILE = "blocked_devices.json"

def get_device_id():
    return str(uuid.getnode())

def load_blocked_devices():
    if os.path.exists(BLOCKED_DEVICES_FILE):
        with open(BLOCKED_DEVICES_FILE, "r") as f:
            return json.load(f)
    return []

def save_blocked_device(device_id):
    blocked_devices = load_blocked_devices()
    if device_id not in blocked_devices:
        blocked_devices.append(device_id)
        with open(BLOCKED_DEVICES_FILE, "w") as f:
            json.dump(blocked_devices, f)

def is_device_blocked(device_id):
    return device_id in load_blocked_devices()

def block_device_remotely(device_id):
    try:
        response = requests.post(BLOCK_URL, json={"device_id": device_id})
        if response.status_code == 200:
            print("✅ Device berhasil diblokir.")
        else:
            print("⚠️ Gagal memblokir device di server:", response.text)
    except Exception as e:
        print("⚠️ Gagal mengirim permintaan blokir:", e)

def verifikasi_lisensi(license_key, device_id):
    try:
        response = requests.post(VERIFY_URL, json={"license_key": license_key, "device_id": device_id})
        if response.status_code == 200:
            data = response.json()
            return data.get("status") == "valid"
        elif response.status_code == 403:
            print("🚫 Terlalu banyak percobaan gagal. Device diblokir.")
            save_blocked_device(device_id)
            block_device_remotely(device_id)
        else:
            print("❌ Lisensi tidak valid:", response.text)
    except Exception as e:
        print("❌ Gagal terhubung ke server lisensi:", e)
    return False

# ====== PROGRAM UTAMA ======
device_id = get_device_id()

if is_device_blocked(device_id):
    print("🚫 Akses ditolak. Device ini telah diblokir.")
    exit()

for attempt in range(1, MAX_ATTEMPTS + 1):
    print("=" * 40)
    print(f"Percobaan {attempt} dari {MAX_ATTEMPTS}")
    print("=" * 40)
    license_key = input("Masukkan kode lisensi: ")

    if verifikasi_lisensi(license_key, device_id):
        print("✅ Lisensi valid. Menjalankan aplikasi...")
        # Panggil fungsi atau buka aplikasi utama di sini
        root = tk.Tk()
        root.title("Aplikasi Stempel Digital")
        tk.Label(root, text="Selamat datang di Aplikasi Stempel!").pack(padx=20, pady=20)
        root.mainloop()
        break
    else:
        print("❌ Lisensi tidak valid.")

    if attempt == MAX_ATTEMPTS:
         print("🚫 Terlalu banyak percobaan gagal. Device diblokir.")
         # Blokir lokal terlebih dahulu, jangan tunggu remote
         save_blocked_device(device_id)
         try:
             block_device_remotely(device_id)
         except:
             print("⚠️ Tidak bisa menghubungi server untuk blokir remote.")
         exit()  # Stop paksa agar aplikasi tidak terbuka


def load_preferensi():
    try:
        with open("pref_stempel.json", "r") as file:
            pref = json.load(file)
            ent_teks.insert(0, pref.get("teks", ""))
            ent_nama.insert(0, pref.get("nama", ""))
            ent_tanggal.insert(0, pref.get("tanggal", ""))
            var_bentuk.set(pref.get("bentuk", "Bulat"))
            var_bentuk_teks.set(pref.get("bentuk_teks", "Lurus"))
            var_posisi.set(pref.get("posisi", "Atas"))
            ent_warna.insert(0, pref.get("warna", "#000000"))
            cb_font.set(pref.get("font_path", "arial.ttf"))
            ent_ukuran_font_melingkar.insert(0, str(pref.get("ukuran_font_melingkar", 24)))
            ent_ukuran_font_nama_tgl.insert(0, str(pref.get("ukuran_font_nama_tgl", 20)))
            ent_tebal_garis.insert(0, str(pref.get("tebal_garis", 10)))
            ent_tebal_font.insert(0, str(pref.get("tebal_font", 1)))
            ent_ttd.insert(0, pref.get("file_ttd", ""))
            var_ukuran.set(pref.get("ukuran_stempel", "Sedang"))
            ent_rotasi_teks.insert(0, str(pref.get("rotasi_teks", 0)))
            ent_warna_teks.insert(0, pref.get("warna_teks", "#000000"))
            ent_tebal_font_nama_tgl.insert(0, str(pref.get("tebal_font_nama_tgl", 1)))
            ent_jarak_nama_tanggal.insert(0, str(pref.get("jarak_nama_tanggal", 40)))
            var_arah_melingkar.set(pref.get("arah_melingkar", "Menghadap Luar"))

            # Tampilkan preview jika ada file tanda tangan
            if os.path.exists(ent_ttd.get()):
                tampilkan_preview_ttd(ent_ttd.get())
            else:
                hapus_file_ttd()

    except FileNotFoundError:
        pass

# Fungsi untuk menyimpan preferensi
def simpan_preferensi():
    pref = {
        "teks": ent_teks.get(),
        "nama": ent_nama.get(),
        "tanggal": ent_tanggal.get(),
        "bentuk": var_bentuk.get(),
        "warna": ent_warna.get(),
        "font_path": cb_font.get(),
        "bentuk_teks": var_bentuk_teks.get(),
        "posisi": var_posisi.get(),
        "ukuran_font": int(ent_ukuran_font.get()),
        "tebal_garis": int(ent_tebal_garis.get()),
        "tebal_font": int(ent_tebal_font.get()),
        "file_ttd": ent_ttd.get() or None,
        "ukuran_stempel": var_ukuran.get(),
        "rotasi_teks": int(ent_rotasi_teks.get()),
        "warna_teks": ent_warna_teks.get(),
        "tebal_font_nama_tgl": int(ent_tebal_font_nama_tgl.get()),
        "jarak_nama_tanggal": int(ent_jarak_nama_tanggal.get())
    }
    with open("pref_stempel.json", "w") as file:
        json.dump(pref, file)


def buat_stempel(teks, nama, tanggal, bentuk, warna, font_path, bentuk_teks, posisi,
                 ukuran_font_melingkar, ukuran_font_nama_tgl, tebal_garis, tebal_font,
                 warna_teks="#000000", tebal_font_nama_tgl=1, jarak_nama_tanggal=40,
                 file_ttd=None, preview=False, ekspor_pdf=False,
                 ukuran_stempel=(500, 500), arah_melingkar="Menghadap Luar", rotasi_teks=0):
    
    bg = (255, 255, 255, 0)
    gambar = Image.new('RGBA', ukuran_stempel, bg)
    draw = ImageDraw.Draw(gambar)

    # Bentuk stempel
    if bentuk == "Bulat":
        draw.ellipse([(20, 20), (ukuran_stempel[0]-20, ukuran_stempel[1]-20)], outline=warna, width=tebal_garis)
    elif bentuk == "Oval":
        draw.ellipse([(50, 100), (ukuran_stempel[0]-50, ukuran_stempel[1]-100)], outline=warna, width=tebal_garis)
    elif bentuk == "Kotak":
        draw.rectangle([(20, 20), (ukuran_stempel[0]-20, ukuran_stempel[1]-20)], outline=warna, width=tebal_garis)

    # Load font
    try:
        font_teks = ImageFont.truetype(font_path, ukuran_font_melingkar)
        font_nama = ImageFont.truetype(font_path, ukuran_font_nama_tgl + tebal_font_nama_tgl)
        font_tanggal = ImageFont.truetype(font_path, ukuran_font_nama_tgl - 4 + tebal_font_nama_tgl)
    except:
        messagebox.showerror("Error", "Font tidak ditemukan atau rusak!")
        return

    tengah_x = ukuran_stempel[0] // 2
    tengah_y = ukuran_stempel[1] // 2

    if teks:
        if bentuk_teks.lower() == "lurus":
            try:
                nama_bbox = font_nama.getbbox(nama)
                nama_width, nama_height = nama_bbox[2] - nama_bbox[0], nama_bbox[3] - nama_bbox[1]
                tanggal_bbox = font_tanggal.getbbox(tanggal)
                tanggal_width, tanggal_height = tanggal_bbox[2] - tanggal_bbox[0], tanggal_bbox[3] - tanggal_bbox[1]
            except AttributeError:
                nama_width, nama_height = draw.textsize(nama, font=font_nama)
                tanggal_width, tanggal_height = draw.textsize(tanggal, font=font_tanggal)

            total_height = nama_height + jarak_nama_tanggal + tanggal_height
            start_y = (ukuran_stempel[1] - total_height) / 2
            nama_pos = ((ukuran_stempel[0] - nama_width) / 2, start_y)
            tanggal_pos = ((ukuran_stempel[0] - tanggal_width) / 2, start_y + nama_height + jarak_nama_tanggal)

            # Gambar tanggal (dengan stroke)
            if tanggal:
                for dx in range(-tebal_font_nama_tgl, tebal_font_nama_tgl + 1):
                    for dy in range(-tebal_font_nama_tgl, tebal_font_nama_tgl + 1):
                        draw.text((tanggal_pos[0] + dx, tanggal_pos[1] + dy), tanggal, font=font_tanggal, fill=warna_teks)

            # Gambar nama (dengan stroke)
            if nama:
                for dx in range(-tebal_font_nama_tgl, tebal_font_nama_tgl + 1):
                    for dy in range(-tebal_font_nama_tgl, tebal_font_nama_tgl + 1):
                        draw.text((nama_pos[0] + dx, nama_pos[1] + dy), nama, font=font_nama, fill=warna_teks)


            if rotasi_teks != 0:
                img_text = Image.new('RGBA', ukuran_stempel, (255, 255, 255, 0))
                draw_text = ImageDraw.Draw(img_text)
                draw_text.text((10, 10), teks, font=font_teks, fill=warna)
                rotated = img_text.rotate(rotasi_teks, expand=1, resample=Image.BICUBIC)
                gambar.alpha_composite(rotated, ((ukuran_stempel[0] - rotated.width) // 2,
                                                 (ukuran_stempel[1] - rotated.height) // 2))
            else:
                try:
                    bbox = font_teks.getbbox(teks)
                    teks_width = bbox[2] - bbox[0]
                    teks_height = bbox[3] - bbox[1]
                except AttributeError:
                    teks_width, teks_height = draw.textsize(teks, font=font_teks)

                teks_pos = ((ukuran_stempel[0] - teks_width) / 2, (ukuran_stempel[1] - teks_height) / 2)
                draw.text(teks_pos, teks, font=font_teks, fill=warna)

        elif bentuk_teks == "Melingkar":
            radius = ukuran_stempel[0] // 2 - 40
            total_char = len(teks)
            sudut_awal = -90 - (total_char - 1) * 10 // 2
            for i, char in enumerate(teks):
                sudut = sudut_awal + i * 10
                angle_rad = math.radians(sudut)
                x = tengah_x + radius * math.cos(angle_rad)
                y = tengah_y + radius * math.sin(angle_rad)

                img_char = Image.new('RGBA', (100, 100), (255, 255, 255, 0))
                draw_char = ImageDraw.Draw(img_char)
                draw_char.text((50, 50), char, font=font_teks, fill=warna, anchor="mm")

                if arah_melingkar == "Menghadap Luar":
                    rotated = img_char.rotate(-sudut + 90 + 180, center=(50, 50), resample=Image.BICUBIC)
                else:
                    img_char = img_char.transpose(Image.FLIP_TOP_BOTTOM)
                    rotated = img_char.rotate(-sudut - 90 + 180, center=(50, 50), resample=Image.BICUBIC)

                gambar.alpha_composite(rotated, (int(x - 50), int(y - 50)))

    # === Blok tanda tangan ===

    if file_ttd:
        try:
            ttd = Image.open(file_ttd).convert('RGBA')
            ttd_width = ukuran_stempel[0] // 2
            ttd_height = ukuran_stempel[1] // 6
            ttd = ttd.resize((ttd_width, ttd_height), Image.Resampling.LANCZOS)

            # Posisi tanda tangan sedikit di atas nama
            ttd_x = (ukuran_stempel[0] - ttd_width) // 2
            ttd_y = tengah_y - 70  # naikkan sedikit di atas tengah

            gambar.alpha_composite(ttd, (ttd_x, ttd_y))
        except Exception as e:
            messagebox.showerror("Error", f"Gagal membuka tanda tangan!\n{e}")
            return

    # === Tetap gambar nama dan tanggal meskipun tanpa ttd ===
    h_nama = 0  # untuk referensi posisi tanggal

    if nama:
        bbox_nama = draw.textbbox((0, 0), nama, font=font_nama)
        w_nama, h_nama = bbox_nama[2] - bbox_nama[0], bbox_nama[3] - bbox_nama[1]
        for dx in range(-tebal_font_nama_tgl, tebal_font_nama_tgl+1):
            for dy in range(-tebal_font_nama_tgl, tebal_font_nama_tgl+1):
                draw.text(((ukuran_stempel[0]-w_nama)/2 + dx, tengah_y - h_nama//2 + dy),
                          nama, font=font_nama, fill=warna_teks)

    if tanggal:
        bbox_tgl = draw.textbbox((0, 0), tanggal, font=font_tanggal)
        w_tgl, h_tgl = bbox_tgl[2] - bbox_tgl[0], bbox_tgl[3] - bbox_tgl[1]
        for dx in range(-tebal_font_nama_tgl, tebal_font_nama_tgl+1):
            for dy in range(-tebal_font_nama_tgl, tebal_font_nama_tgl+1):
                draw.text(((ukuran_stempel[0]-w_tgl)/2 + dx,
                           tengah_y - h_nama//2 + jarak_nama_tanggal + dy),
                          tanggal, font=font_tanggal, fill=warna_teks)

    # === Bagian akhir: Preview / Simpan ===
    if preview:
        tampilkan_preview(gambar)
    elif ekspor_pdf:
        simpan_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if simpan_path:
            gambar.convert("RGB").save(simpan_path, "PDF")
    else:
        simpan_path = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG files", "*.png")])
        if simpan_path:
            gambar.save(simpan_path, "PNG", optimize=True)

    return gambar


def tampilkan_preview(img):
    preview_window = Toplevel(root)
    preview_window.title("Preview Stempel")
    img_resized = img.resize((250, 250))
    img_tk = ImageTk.PhotoImage(img_resized)

    label = tk.Label(preview_window, image=img_tk)
    label.image = img_tk
    label.pack()

def tampilkan_preview_ttd(path):
    try:
        img = Image.open(path)
        img = img.resize((150, 60))  # Ukuran preview
        img_tk = ImageTk.PhotoImage(img)
        preview_ttd_label.configure(image=img_tk)
        preview_ttd_label.image = img_tk  # Simpan referensi agar tidak dihapus
    except:
        preview_ttd_label.configure(image=None)
        preview_ttd_label.image = None

def hapus_file_ttd():
    ent_ttd.delete(0, tk.END)
    preview_ttd_label.configure(image=None)
    preview_ttd_label.image = None


def pilih_file_ttd():
    file_path = filedialog.askopenfilename(filetypes=[("Image files", "*.png;*.jpg;*.jpeg")])
    if file_path:
        ent_ttd.delete(0, tk.END)
        ent_ttd.insert(0, file_path)
        tampilkan_preview_ttd(file_path)
        
def pilih_warna():
    warna = colorchooser.askcolor(title="Pilih Warna Garis")[1]
    if warna:
        ent_warna.delete(0, tk.END)
        ent_warna.insert(0, warna)

def pilih_warna_teks(entry_field):
    warna = colorchooser.askcolor(title="Pilih Warna Teks")[1]
    if warna:
        entry_field.delete(0, tk.END)
        entry_field.insert(0, warna)


def get_ukuran_stempel():
    pilihan = var_ukuran.get()
    if pilihan == "Kecil":
        return (300, 300)
    elif pilihan == "Sedang":
        return (500, 500)
    elif pilihan == "Besar":
        return (700, 700)
    else:
        return (300, 300)  # fallback default

def ambil_input():
    return {
        "teks": ent_teks.get(),
        "nama": ent_nama.get(),
        "tanggal": ent_tanggal.get(),
        "bentuk": var_bentuk.get(),
        "warna": ent_warna.get(),
        "font_path": cb_font.get(),
        "bentuk_teks": var_bentuk_teks.get(),
        "posisi": var_posisi.get(),
        "ukuran_font_melingkar": int(ent_ukuran_font_melingkar.get()),
        "ukuran_font_nama_tgl": int(ent_ukuran_font_nama_tgl.get()),
        "tebal_garis": int(ent_tebal_garis.get()),
        "tebal_font": int(ent_tebal_font.get()),
        "file_ttd": ent_ttd.get() or None,
        "ukuran_stempel": get_ukuran_stempel(),
        "rotasi_teks": int(ent_rotasi_teks.get()),
        "warna_teks": ent_warna_teks.get(),
        "tebal_font_nama_tgl": int(ent_tebal_font_nama_tgl.get()),
        "jarak_nama_tanggal": int(ent_jarak_nama_tanggal.get()),
        "arah_melingkar": var_arah_melingkar.get()
        }

def buat():
    args = ambil_input()
    buat_stempel(**args)
    simpan_ke_history(args)

def preview():
    args = ambil_input()
    buat_stempel(**args, preview=True)

def ekspor_pdf():
    args = ambil_input()
    buat_stempel(**args, ekspor_pdf=True)

def kirim_stempel_ke_excel_aktif():
    data = ambil_input()
    img = buat_stempel(**data, preview=False)

    if img is None:
        return

    try:
        # Simpan gambar ke file sementara
        temp_path = os.path.join(tempfile.gettempdir(), "stempel_excel_embed.png")
        img.save(temp_path, "PNG", optimize=True)

        # Buka Excel dan ambil worksheet aktif
        excel = win32.Dispatch("Excel.Application")
        sheet = excel.ActiveSheet
        cell = excel.Selection

        # Koordinat dan ukuran sel (dalam points)
        left = cell.Left
        top = cell.Top
        width = cell.Width
        height = cell.Height

        # Ukuran sisi terkecil
        ukuran_min = min(width, height)

        # Posisi gambar di tengah sel
        center_left = left + (width - ukuran_min) / 2
        center_top = top + (height - ukuran_min) / 2

        embed = var_embed_excel.get()

        if embed:
            shape = sheet.Shapes.AddPicture(
                Filename=temp_path,
                LinkToFile=False,
                SaveWithDocument=True,
                Left=center_left,
                Top=center_top,
                Width=ukuran_min,
                Height=ukuran_min
            )
        else:
            folder_output = "stempel_excel_linked"
            os.makedirs(folder_output, exist_ok=True)
            filename = f"stempel_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
            final_path = os.path.join(folder_output, filename)
            img.convert("RGB").save(final_path)

            shape = sheet.Shapes.AddPicture(
                Filename=os.path.abspath(final_path),
                LinkToFile=True,
                SaveWithDocument=False,
                Left=center_left,
                Top=center_top,
                Width=ukuran_min,
                Height=ukuran_min
            )

        messagebox.showinfo("Berhasil", f"Stempel berhasil ditempel ke Excel di sel {cell.Address}")

    except Exception as e:
        messagebox.showerror("Gagal", f"Gagal mengirim stempel ke Excel:\n\n{e}")


def simpan_ke_history(data):
    file_history = "stempel_history.json"
    riwayat = []

    # Baca history jika sudah ada
    if os.path.exists(file_history):
        with open(file_history, "r") as f:
            try:
                riwayat = json.load(f)
            except:
                riwayat = []

    # Tambahkan timestamp
    data["waktu"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    riwayat.insert(0, data)

    # Batasi hanya 20 data terakhir
    riwayat = riwayat[:20]

    with open(file_history, "w") as f:
        json.dump(riwayat, f, indent=2)

def tampilkan_history():
    file_history = "stempel_history.json"
    if not os.path.exists(file_history):
        messagebox.showinfo("Kosong", "Belum ada history.")
        return

    with open(file_history, "r") as f:
        try:
            riwayat = json.load(f)
        except:
            messagebox.showerror("Error", "Gagal baca history.")
            return

    # Tampilkan popup list
    popup = Toplevel(root)
    popup.title("Riwayat Stempel")

    listbox = tk.Listbox(popup, width=80)
    listbox.pack(padx=10, pady=10)

    for i, item in enumerate(riwayat):
        teks = f"{i+1}. [{item['waktu']}] - {item['teks']} | {item['nama']} | {item['tanggal']}"
        listbox.insert(tk.END, teks)

    def load_terpilih():
        sel = listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        data = riwayat[idx]

        # Isi ulang form dengan data terpilih
        ent_teks.delete(0, tk.END)
        ent_teks.insert(0, data["teks"])

        ent_nama.delete(0, tk.END)
        ent_nama.insert(0, data["nama"])

        ent_tanggal.delete(0, tk.END)
        ent_tanggal.insert(0, data["tanggal"])

        var_bentuk.set(data["bentuk"])
        ent_warna.delete(0, tk.END)
        ent_warna.insert(0, data["warna"])

        cb_font.set(data["font_path"])
        var_bentuk_teks.set(data["bentuk_teks"])
        var_posisi.set(data["posisi"])

        ent_ukuran_font_melingkar.delete(0, tk.END)
        ent_ukuran_font_melingkar.insert(0, str(data.get("ukuran_font_melingkar", 24)))

        ent_ukuran_font_nama_tgl.delete(0, tk.END)
        ent_ukuran_font_nama_tgl.insert(0, str(data.get("ukuran_font_nama_tgl", 20)))

        ent_tebal_garis.delete(0, tk.END)
        ent_tebal_garis.insert(0, str(data["tebal_garis"]))

        ent_tebal_font.delete(0, tk.END)
        ent_tebal_font.insert(0, str(data["tebal_font"]))

        ent_ttd.delete(0, tk.END)
        ent_ttd.insert(0, data["file_ttd"] or "")

        var_ukuran.set(data["ukuran_stempel"])
        ent_rotasi_teks.delete(0, tk.END)
        ent_rotasi_teks.insert(0, str(data["rotasi_teks"]))

        ent_warna_teks.delete(0, tk.END)
        ent_warna_teks.insert(0, data["warna_teks"])

        ent_tebal_font_nama_tgl.delete(0, tk.END)
        ent_tebal_font_nama_tgl.insert(0, str(data["tebal_font_nama_tgl"]))

        ent_jarak_nama_tanggal.delete(0, tk.END)
        ent_jarak_nama_tanggal.insert(0, str(data["jarak_nama_tanggal"]))

        var_arah_melingkar.set(data.get("arah_melingkar", "Menghadap Luar"))

        # ➕ Tambahkan logika preview tanda tangan di sini
        if os.path.exists(data["file_ttd"] or ""):
            tampilkan_preview_ttd(data["file_ttd"])
        else:
            hapus_file_ttd()

        popup.destroy()

    tk.Button(popup, text="Gunakan Data Ini", command=load_terpilih).pack(pady=5)



# GUI
root = tk.Tk()
root.title("Stampel App Clamtech")

frame = tk.Frame(root)
frame.pack(padx=10, pady=10)

# Form Input
tk.Label(frame, text="Teks Stempel:").grid(row=0, column=0, sticky="e")
ent_teks = tk.Entry(frame, width=40)
ent_teks.grid(row=0, column=1)

tk.Label(frame, text="Nama:").grid(row=1, column=0, sticky="e")
ent_nama = tk.Entry(frame, width=40)
ent_nama.grid(row=1, column=1)

tk.Label(frame, text="Tanggal:").grid(row=2, column=0, sticky="e")
ent_tanggal = tk.Entry(frame, width=40)
ent_tanggal.grid(row=2, column=1)

tk.Label(frame, text="Bentuk Stempel:").grid(row=3, column=0, sticky="e")
var_bentuk = tk.StringVar(value="Bulat")
tk.OptionMenu(frame, var_bentuk, "Bulat", "Oval", "Kotak").grid(row=3, column=1, sticky="w")

tk.Label(frame, text="Bentuk Teks:").grid(row=4, column=0, sticky="e")
var_bentuk_teks = tk.StringVar(value="Lurus")
tk.OptionMenu(frame, var_bentuk_teks, "Lurus", "Melingkar").grid(row=4, column=1, sticky="w")

tk.Label(frame, text="Posisi Teks:").grid(row=5, column=0, sticky="e")
var_posisi = tk.StringVar(value="Atas")
tk.OptionMenu(frame, var_posisi, "Atas", "Tengah", "Bawah").grid(row=5, column=1, sticky="w")

tk.Label(frame, text="Warna:").grid(row=6, column=0, sticky="e")
ent_warna = tk.Entry(frame)
ent_warna.insert(0, "#FF0000")  # merah
ent_warna.grid(row=6, column=1, sticky="w")
tk.Button(frame, text="Pilih", command=pilih_warna).grid(row=6, column=2)

tk.Label(frame, text="Tipe Font:").grid(row=7, column=0, sticky="e")
cb_font = ttk.Combobox(frame, values=FONT_OPTIONS, width=37)
cb_font.set("arial.ttf")
cb_font.grid(row=7, column=1)

# Baris baru untuk ukuran font stempel melingkar
tk.Label(frame, text="Ukuran Font Melingkar:").grid(row=8, column=0, sticky="e")
ent_ukuran_font_melingkar = tk.Entry(frame, width=10)
ent_ukuran_font_melingkar.insert(0, "24")
ent_ukuran_font_melingkar.grid(row=8, column=1, sticky="w")

# Baris baru untuk ukuran font nama/tanggal
tk.Label(frame, text="Ukuran Font Nama/Tgl:").grid(row=9, column=0, sticky="e")
ent_ukuran_font_nama_tgl = tk.Entry(frame, width=10)
ent_ukuran_font_nama_tgl.insert(0, "20")
ent_ukuran_font_nama_tgl.grid(row=9, column=1, sticky="w")

tk.Label(frame, text="Tebal Garis:").grid(row=10, column=0, sticky="e")
ent_tebal_garis = tk.Entry(frame, width=10)
ent_tebal_garis.insert(0, "10")
ent_tebal_garis.grid(row=10, column=1, sticky="w")

tk.Label(frame, text="Tebal Font:").grid(row=11, column=0, sticky="e")
ent_tebal_font = tk.Entry(frame, width=10)
ent_tebal_font.insert(0, "1")
ent_tebal_font.grid(row=11, column=1, sticky="w")

tk.Label(frame, text="Tanda Tangan:").grid(row=18, column=0, sticky="e")
ent_ttd = tk.Entry(frame, width=30)
ent_ttd.grid(row=18, column=1, sticky="w")

tk.Button(frame, text="Pilih", command=pilih_file_ttd).grid(row=18, column=2)
tk.Button(frame, text="Hapus", command=hapus_file_ttd).grid(row=18, column=3)

preview_ttd_label = tk.Label(frame)
preview_ttd_label.grid(row=19, column=1, columnspan=2, pady=5)

tk.Label(frame, text="Ukuran Stempel:").grid(row=12, column=0, sticky="e")
var_ukuran = tk.StringVar(value="Sedang")
tk.OptionMenu(frame, var_ukuran, "Kecil", "Sedang", "Besar").grid(row=12, column=1, sticky="w")

tk.Label(frame, text="Rotasi Teks (derajat):").grid(row=13, column=0, sticky="e")
ent_rotasi_teks = tk.Entry(frame, width=10)
ent_rotasi_teks.insert(0, "0")
ent_rotasi_teks.grid(row=13, column=1, sticky="w")

tk.Label(frame, text="Warna Teks:").grid(row=14, column=0, sticky="e")
ent_warna_teks = tk.Entry(frame)
ent_warna_teks.insert(0, "#FF0000")
ent_warna_teks.grid(row=14, column=1, sticky="w")
tk.Button(frame, text="Pilih", command=lambda: pilih_warna_teks(ent_warna_teks)).grid(row=14, column=2)

tk.Label(frame, text="Jarak Nama & Tanggal:").grid(row=15, column=0, sticky="e")
ent_jarak_nama_tanggal = tk.Entry(frame, width=10)
ent_jarak_nama_tanggal.insert(0, "40")
ent_jarak_nama_tanggal.grid(row=15, column=1, sticky="w")

tk.Label(frame, text="Tebal Font Nama/Tanggal:").grid(row=16, column=0, sticky="e")
ent_tebal_font_nama_tgl = tk.Entry(frame, width=10)
ent_tebal_font_nama_tgl.insert(0, "1")
ent_tebal_font_nama_tgl.grid(row=16, column=1, sticky="w")

tk.Label(frame, text="Arah Teks Melingkar:").grid(row=17, column=0, sticky="e")
var_arah_melingkar = tk.StringVar(value="Menghadap Luar")
tk.OptionMenu(frame, var_arah_melingkar, "Menghadap Luar", "Menghadap Dalam").grid(row=17, column=1, sticky="w")


# Tombol Aksi
frame_btn = tk.Frame(root)
frame_btn.pack(pady=10)
tk.Button(frame_btn, text="Preview", command=preview).pack(side="left", padx=10)
tk.Button(frame_btn, text="Simpan PNG", command=buat).pack(side="left", padx=10)
tk.Button(frame_btn, text="Ekspor PDF", command=ekspor_pdf).pack(side="left", padx=10)
# Tambahkan ini sebelum tombol 'Kirim ke Excel Aktif'
var_embed_excel = tk.BooleanVar(value=True)
tk.Checkbutton(frame_btn, text="Embed permanen ke Excel", variable=var_embed_excel).pack(side="left", padx=10)


tk.Button(frame_btn, text="Kirim ke Excel Aktif", command=kirim_stempel_ke_excel_aktif).pack(side="left", padx=10)
tk.Button(frame_btn, text="History", command=tampilkan_history).pack(side="left", padx=10)



load_preferensi()

root.mainloop()
