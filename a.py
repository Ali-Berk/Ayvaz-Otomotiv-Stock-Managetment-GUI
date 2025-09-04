import tkinter as tk
from tkinter import ttk, messagebox
import sqlite3
import pandas as pd
from pandastable import Table
import requests
import xml.etree.ElementTree as ET
import os
from PIL import Image, ImageTk
from tkinter import filedialog
from tkinter import Label
import io
from fpdf import FPDF 
import datetime
from datetime import datetime

# === ÖN AYARLAR ===
url = "https://www.tcmb.gov.tr/kurlar/today.xml"
base_path = os.path.dirname(os.path.abspath(__file__))
ilceler_df = pd.read_excel(base_path+"/ilce-listesi.xlsx")
iller_list = ilceler_df["IL_ADI"].drop_duplicates().tolist()

# === VERİTABANI BAĞLANTILARI ===
conn_urun = sqlite3.connect("urunler.db")
conn_musteri = sqlite3.connect("musteriler.db")
conn_siparis = sqlite3.connect("siparisler.db")

cur_musteri = conn_musteri.cursor()
cur_siparis = conn_siparis.cursor()
cursor = conn_urun.cursor()

# === TABLOLAR ===

cursor.execute("""
    CREATE TABLE IF NOT EXISTS urunler (
    urunid TEXT PRIMARY KEY,
    urun_grubu TEXT,
    urun_model TEXT,
    fiyat REAL,
    stok INTEGER,
    resim TEXT,
    ebat TEXT,
    agirlik TEXT,
    sonstoktarihi DATE,
    renk TEXT)
    """)
conn_urun.commit()

cur_musteri.execute("""
CREATE TABLE IF NOT EXISTS musteriler (
    musteri_id INTEGER PRIMARY KEY AUTOINCREMENT,
    telefon TEXT UNIQUE,
    ad TEXT,
    soyad TEXT,
    il TEXT,
    ilce TEXT,
    adres TEXT,
    aciklama TEXT)                
""")
conn_musteri.commit()

cur_siparis.execute("""
CREATE TABLE IF NOT EXISTS siparisler (
    siparis_id  INTEGER PRIMARY KEY AUTOINCREMENT,
    telefon TEXT,
    urun_grubu TEXT,
    marka TEXT,
    urun_id TEXT,
    urun_model TEXT,
    adet INTEGER,
    kargo TEXT,
    tarih DATE,
    gonderimadresi TEXT,
    renk TEXT,
    odemesekli TEXT,
    durumu TEXT,
    satisftl REAL,
    kur REAL,
    tutar REAL,
    onay BOOL,
    fatura_no TEXT,
    FOREIGN KEY (telefon) REFERENCES musteriler(telefon),
    FOREIGN KEY (urun_id) REFERENCES urunler(urunid))
""")
conn_siparis.commit()

#=== FONKSİYONLAR ===
#Dolar kurunu internetten çekme ve hataların önüne geçme
try:
    response = requests.get(url, timeout=5)
    response.encoding = "utf-8"
    root = ET.fromstring(response.text)
    usd = root.find(".//Currency[@Kod = 'USD']/BanknoteSelling").text
except Exception:
    cur_siparis.execute("SELECT kur FROM siparisler ORDER BY siparis_id DESC LIMIT 1")
    last_kur = cur_siparis.fetchone()
    usd = last_kur[0]
    messagebox.showwarning("Uyarı", "İnternet olmadığından EN SON SİPARİŞİN dolar kuru kullanıldı.")

def Get():
    try:
        global table_siparis1
        phone_number = entry_phone.get().strip()
        if not phone_number:
            messagebox.showwarning("Uyarı", "Telefon numarası giriniz.")
            return

        # SQL'den dataframe çek
        df_siparis = pd.read_sql_query(
            "SELECT * FROM siparisler WHERE telefon = ?",
            conn_siparis,
            params=(phone_number,)
        )

        if df_siparis.empty:
            messagebox.showinfo("Bilgi", "Kayıt bulunamadı.")
            return

        # Mevcut tabloyu temizle ve yeniden yükle
        frame_right = table_siparis1.parentframe
        for widget in frame_right.winfo_children():
            widget.destroy()

        table_siparis1 = Table(frame_right, dataframe=df_siparis, showtoolbar=True, showstatusbar=True)
        table_siparis1.show()

    except sqlite3.Error as e:
        messagebox.showerror("Veritabanı Hatası", str(e))
    except Exception as e:
        messagebox.showerror("Hata", str(e))

table_urun2 = None  
frame2_table = None
def GetProduct():
    global table_urun2
    try:
        productGroup = entry_productGroup.get().strip()
        productID = entry_productID.get().strip()

        if not productID and not productGroup:
            messagebox.showerror("Uyarı", "Ürün Grubu veya Ürün ID giriniz.")
            return

        query = "SELECT urunid, urun_grubu, urun_model, fiyat, stok, ebat, agirlik, renk FROM urunler"
        params = []

        if productID and not productGroup:
            query += " WHERE urunid = ?"
            params.append(productID)
        elif not productID and productGroup:
            query += " WHERE urun_grubu = ?"
            params.append(productGroup)
        elif productID and productGroup:
            query += " WHERE urunid = ? AND urun_grubu = ?"
            params.extend([productID, productGroup])

        df_urun = pd.read_sql_query(query, conn_urun, params=params)

        if df_urun.empty:
            messagebox.showinfo("Bilgi", "Ürün bulunamadı.")
            return

        # Eski tabloyu sil ve yenisini oluştur
        frame2_table = table_urun2.parentframe
        for widget in frame2_table.winfo_children():
            widget.destroy()

        table_urun2 = Table(frame2_table, dataframe=df_urun, showtoolbar=True, showstatusbar=True)
        table_urun2.show()
        
    except sqlite3.Error as e:
        messagebox.showerror("Veritabanı Hatası", str(e))
    except Exception as e:
        messagebox.showerror("Hata", str(e))

def Add():
    global table_siparis1, table_urun2

    phone_number = entry_phone.get().strip()
    ship = entry_ship.get().strip()
    qty = entry_qty.get().strip()
    payment = entry_payment.get().strip()

    if not phone_number:
        messagebox.showwarning("Uyarı", "Müşteri Listesinden Müşteri bilgilerini giriniz.")
        return

    try:
        # Müşteri bilgisi al
        cur_musteri.execute("SELECT telefon, il, ilce, adres FROM musteriler WHERE telefon = ?", (phone_number,))
        musteri = cur_musteri.fetchone()
        if not musteri:
            messagebox.showerror("Hata", "Bu telefon numarasına ait müşteri bulunmamaktadır.")
            return
        telefon, il, ilce, adres = musteri

        # Pandastable üzerinden seçili ürün satırını al
        row_index = table_urun2.getSelectedRow()
        if row_index is None or row_index < 0:
            messagebox.showwarning("Uyarı", "Lütfen bir ürün seçiniz.")
            return

        df_selected = table_urun2.model.df.iloc[row_index]
        urunid = df_selected["urunid"]
        urun_grubu = df_selected["urun_grubu"]
        urun_model = df_selected["urun_model"]
        fiyat = df_selected["fiyat"]
        renk = df_selected["renk"] if "renk" in df_selected.index else "Belirtilmedi"
        stok = df_selected["stok"] if "stok" in df_selected.index else 0
        status = combobox_status.get()

        # Eksik alanları son siparişten doldur
        if not ship or not payment or not qty:
            cur_siparis.execute("""
                SELECT odemesekli, kargo
                FROM siparisler 
                WHERE telefon = ?
                ORDER BY siparis_id DESC
                LIMIT 1
            """, (phone_number,))
            last_order = cur_siparis.fetchone()
            if last_order:
                last_odeme, last_kargo = last_order
                ship = ship if ship else (last_kargo or "Belirtilmedi")
                payment = payment if payment else (last_odeme or "Belirtilmedi")
            qty = int(qty) if qty else 1
        else:
            qty = int(qty)
        
        if qty <= 0:
            return
        if stok-qty < 10:
            messagebox.showwarning("Uyarı", "Stok azalmış, lütfen stok kontrolü yapınız.")
        
        # Stok kontrolü
        cursor.execute("SELECT stok FROM urunler WHERE urunid = ?", (urunid,))
        stok = cursor.fetchone()
        if not stok or stok[0] < qty:
            messagebox.showerror("Hata", "Yeterli stok bulunmamaktadır.")
            return

        # Siparişi ekle
        cur_siparis.execute("""
            INSERT INTO siparisler 
            (telefon, urun_grubu, urun_id, urun_model, adet, kargo, tarih, gonderimadresi, renk, odemesekli, durumu, satisftl, tutar, onay, kur)
            VALUES (?, ?, ?, ?, ?, ?, DATE('now'), ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            telefon, urun_grubu, urunid, urun_model, qty, ship,
            f"{il} {ilce} {adres}", renk, payment, status ,
            fiyat, int(fiyat) * qty, 0, usd
        ))
        conn_siparis.commit()

        # Güncel tabloyu yeniden yükle
        df_siparis = pd.read_sql_query("SELECT * FROM siparisler WHERE telefon = ?", conn_siparis, params=(phone_number,))
        for widget in table_siparis1.parentframe.winfo_children():
            widget.destroy()

        table_siparis1 = Table(table_siparis1.parentframe, dataframe=df_siparis, showtoolbar=True, showstatusbar=True)
        table_siparis1.show()
        cursor.execute("UPDATE urunler SET stok = stok - ? WHERE urunid = ?", (qty, urunid))
        conn_urun.commit()
        refresh_all_tables()
        messagebox.showinfo("Başarılı", "Ürün müşterinin siparişlerine eklendi.")

    except Exception as e:
        messagebox.showerror("Hata", str(e))

def Delete():
    global table_siparis1

    try:
        # Pandastable'dan seçilen satırın indeksini al
        row_index = table_siparis1.getSelectedRow()
        if row_index is None or row_index < 0:
            messagebox.showwarning("Uyarı", "Sileceğiniz siparişi seçiniz.")
            return

        # Seçilen satırın verilerini DataFrame'den çek
        df_selected = table_siparis1.model.df.iloc[row_index]
        siparis_id = int(df_selected["siparis_id"])

        # Veritabanından sil
        cur_siparis.execute("DELETE FROM siparisler WHERE siparis_id = ?", (siparis_id,))
        conn_siparis.commit()

        messagebox.showinfo("Başarılı", f"Sipariş (ID: {siparis_id}) silindi.")

        # Tabloyu güncelle
        Get()
        refresh_all_tables()
    except Exception as e:
        messagebox.showerror("Hata", str(e))

def newCustomer():
    try:
        phone_number = entry_phone.get()
        Name = entry_name.get()
        Surname = entry_surname.get()
        state = combo_state.get()
        district = combo_district.get()
        address = entry_address.get()
        note = entry_note.get()
        if not Name or not Surname or not phone_number:
            messagebox.showwarning("Uyarı", "Ad, Soyad ve Telefon alanları zorunludur.")
            return

        cur_musteri.execute("""
            INSERT INTO musteriler (ad, soyad, telefon, il, ilce, adres, aciklama)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (Name, Surname, phone_number, state, district, address, note))
        conn_musteri.commit()

        messagebox.showinfo("Başarılı", "Müşteri kaydı başarıyla eklendi.")
        
        entry_name.delete(0, tk.END)
        entry_surname.delete(0, tk.END)
        entry_phone.delete(0, tk.END)
        entry_address.delete(0, tk.END)
        entry_note.delete(0, tk.END)
        combo_state.set("")
        combo_district.set("")
    except sqlite3.IntegrityError:
        messagebox.showerror("Hata", "Bu telefon numarası zaten kayıtlı.")
    except Exception as e:
        messagebox.showerror("Hata", str(e))
    finally:
        refresh_all_tables()

def refresh_urun_table():
    global table_urun, table_urun2
    df_urun = pd.read_sql_query("SELECT * FROM urunler", conn_urun)
    for widget in table_urun.parentframe.winfo_children():
        widget.destroy()
    table_urun = EditableTable(table_urun.parentframe, dataframe=df_urun, table_name="urunler", id_column="urunid", conn=conn_urun, showtoolbar=True, showstatusbar=True)
    table_urun.show()
    df_urun2 = pd.read_sql_query("SELECT * FROM urunler", conn_urun)
    for widget in table_urun2.parentframe.winfo_children():
        widget.destroy()
    table_urun2 = EditableTable(table_urun2.parentframe, dataframe=df_urun2, table_name="urunler", id_column="urunid", conn=conn_urun, showtoolbar=True, showstatusbar=True)
    table_urun2.show()

def refresh_siparis_table():
    global table_siparis1
    df_siparis = pd.read_sql_query("SELECT * FROM siparisler", conn_siparis)
    for widget in table_siparis1.parentframe.winfo_children():
        widget.destroy()
    table_siparis1 = EditableTable(table_siparis1.parentframe, dataframe=df_siparis, table_name="siparisler", id_column="siparis_id", conn=conn_siparis, showtoolbar=True, showstatusbar=True)
    table_siparis1.show()

def refresh_customer_list():
    global table_musteri
    df_musteri = pd.read_sql_query("SELECT * FROM musteriler", conn_musteri)
    for widget in table_musteri.parentframe.winfo_children():
        widget.destroy()
    table_musteri = EditableTable(table_musteri.parentframe, dataframe=df_musteri, table_name="musteriler", id_column="musteri_id", conn=conn_musteri, showtoolbar=True, showstatusbar=True)
    table_musteri.show()

def refresh_all_tables():
    refresh_urun_table()
    refresh_siparis_table()
    refresh_customer_list()

def update_districts(event):
    selected_state = combo_state.get()
    filtered_districts = ilceler_df[ilceler_df["IL_ADI"] == selected_state]["AD"].tolist()
    combo_district["values"] = filtered_districts
    combo_district.set("")

class EditableTable(Table):
    def __init__(self, parent=None, dataframe=None, table_name=None, id_column=None, conn=None, **kwargs):
        Table.__init__(self, parent=parent, dataframe=dataframe, **kwargs)
        self.id_column = id_column
        self.table_name = table_name
        self.conn = conn
        self.cursor = self.conn.cursor()
    def handleCellEntry(self, row, col):
        super().handleCellEntry(row, col)
        try:
            new_value = self.model.getValueAt(row, col)
            try:
                new_value = int(new_value)
            except ValueError:
                new_value = str(new_value)
                       
            col_name = self.model.df.columns[col]
            try:
                record_id = int(self.model.df.iloc[row][self.id_column])
            except ValueError:
                record_id = str(self.model.df.iloc[row][self.id_column])
            # SQL sorgusunu oluştur
            sql_query = f"UPDATE {self.table_name} SET {col_name} = ? WHERE {self.id_column} = ?"
            self.cursor.execute(sql_query, (new_value, record_id))
            self.conn.commit()
            messagebox.showinfo("Başarılı", f"{self.table_name} tablosu güncellendi.")
        except Exception as e:
            messagebox.showerror("Hata", str(e))

def Add_stock():
    global image_path
    stock_id = entry_stock_id.get().strip()
    stock_group = entry_stock_group.get().strip()
    stock_model = entry_stock_model.get().strip()
    stock_qty = entry_stock_qty.get().strip()
    stock_mass = entry_stock_mass.get().strip()
    stock_cost = entry_stock_cost.get().strip()
    stock_price = entry_stock_price.get().strip()
    date = datetime.now().date()

    try:
        # Stok integer olsun
        stock_qty = int(stock_qty) if stock_qty else 0

        # Ürün var mı kontrol et
        cursor.execute("SELECT * FROM urunler WHERE urunid = ?", (stock_id,))
        eski_kayit = cursor.fetchone()

        if eski_kayit:
            # Tablo kolonlarının sırasını öğrenelim
            kolonlar = [desc[0] for desc in cursor.description]
            eski_veri = dict(zip(kolonlar, eski_kayit))

            # Kullanıcıdan gelen veriler (boş olabilir)
            yeni_veri = {
                "urunid": stock_id,
                "urun_grubu": stock_group,
                "urun_model": stock_model,
                "stok": (eski_veri["stok"] + stock_qty) if stock_qty else eski_veri["stok"],
                "agirlik": stock_mass,
                "fiyat": float(stock_price) if stock_price else eski_veri["fiyat"],
                "resim": image_path,
                "sonstoktarihi": date
            }

            # Boş gelenleri eski verilerle doldur
            for kolon, deger in yeni_veri.items():
                if not deger:  # "" veya None ise
                    yeni_veri[kolon] = eski_veri[kolon]

            # Güncelleme sorgusu
            cursor.execute(f"""
                UPDATE urunler
                SET urun_grubu=?, urun_model=?, stok=?, agirlik=?, fiyat=?, resim=?, sonstoktarihi=?
                WHERE urunid=?
            """, (
                yeni_veri["urun_grubu"],
                yeni_veri["urun_model"],
                yeni_veri["stok"],
                yeni_veri["agirlik"],
                yeni_veri["fiyat"],
                yeni_veri["resim"],
                yeni_veri["sonstoktarihi"],
                yeni_veri["urunid"]
            ))
            conn_urun.commit()
            messagebox.showinfo("Başarılı", f"Stok güncellendi. Yeni stok: {yeni_veri['stok']}")

        else:
            # Yeni ürün ekle
            fiyat = float(stock_price) if stock_price else 0.0
            cursor.execute("""
                INSERT INTO urunler (urunid, urun_grubu, urun_model, stok, agirlik, fiyat, resim, sonstoktarihi)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (stock_id, stock_group, stock_model, stock_qty, stock_mass, fiyat, image_path, date))
            conn_urun.commit()
            messagebox.showinfo("Başarılı", "Yeni ürün eklendi.")

    except ValueError:
        messagebox.showerror("Hata", "Lütfen stok ve fiyat değerlerini doğru giriniz.")
    except Exception as e:
        messagebox.showerror("Hata", str(e))
    finally:
        entry_stock_id.delete(0, tk.END)
        entry_stock_group.delete(0, tk.END)
        entry_stock_model.delete(0, tk.END)
        entry_stock_qty.delete(0, tk.END)
        entry_stock_mass.delete(0, tk.END)
        entry_stock_cost.delete(0, tk.END)
        entry_stock_price.delete(0, tk.END)
        refresh_all_tables()


def on_text_change(*args):
    active_num.config(text=f"AKTİF NUMARA: {phone_var.get()}")

image_path = None
def img_to_bytes():
    global image_path
    image_path = filedialog.askopenfilename(title="Resim Seç", filetypes=[("Image files", "*.jpg *.jpeg *.png")])
#Seçili satırı al> event ile seçim değiştiğinde çağır>label configden resmi güncelle.
def on_row_selected(event):
    selected_row = table_urun2.getSelectedRow()
    if selected_row is not None:
        df_selected = table_urun2.model.df.iloc[selected_row]
        urunid = df_selected["urunid"]

        result = cursor.execute("SELECT resim FROM urunler WHERE urunid = ?", (urunid,)).fetchone()
        if result and result[0]:
            image_path = result[0]
            try:
                img = Image.open(image_path)
                img = img.resize((200, 200), Image.Resampling.LANCZOS)
                photo = ImageTk.PhotoImage(img)
                image_box.config(image=photo, height=200, width=200)
                image_box.image = photo  # REFERANS tutulmalı!
            except FileNotFoundError:
                image_box.config(image='', text="Resim bulunamadı", bg="white")
            except Exception as e:
                image_box.config(image='', text="Hata oluştu", bg="white")
        else:
            image_box.config(image='',text="Resim bulunamadı", bg="white", height=25, width=25)

def cargoBill():
    selected_customer = table_siparis1.getSelectedRow()

    if selected_customer is None:
        messagebox.showwarning("Uyarı", "Lütfen bir sipariş seçiniz.")
        return

    df_Selected = table_siparis1.model.df.iloc[selected_customer].to_list()
    siparis_id = int(df_Selected[0])
    result = cur_siparis.execute("SELECT * FROM siparisler WHERE siparis_id = ?", (siparis_id,)).fetchone()
    if result is None:
        messagebox.showerror("Hata", "Sipariş bulunamadı.")
        return

    telefon = result[1]

    customer_info = cur_musteri.execute("SELECT * FROM musteriler WHERE telefon = ?", (telefon,)).fetchone()
    if customer_info is None:
        messagebox.showerror("Hata", f"Telefon {telefon} ile müşteri bulunamadı.")
        return

    musteri_ad = customer_info[2] 
    musteri_soyad = customer_info[3]  
    musteri_adres = f"{customer_info[5]} {customer_info[4]} {customer_info[6]}" 
    musteri_tel = customer_info[1] 
    urun_resmi = cursor.execute("SELECT resim FROM urunler WHERE urunid = ?", (result[4],)).fetchone()
    new_Form = tk.Toplevel(root)
    new_Form.title("Kargo Fişi")
    new_Form.geometry("700x500")
    new_Form.configure(bg="white")
    container = tk.Frame(new_Form,bg="white")
    container.pack(fill="both",expand=True)
    tk.Label(container, text="Adı Soyadı:", bg="white", font=("Arial", 20 )).grid(row=0, column=0, sticky="w")
    tk.Label(container, text=f"{musteri_ad} {musteri_soyad}", bg="white", font=("Arial", 20, "bold")).grid(row=0, column=1, sticky="w")
    tk.Label(container, text="İşletme:", bg="white", font=("Arial", 20)).grid(row=1, column=0, sticky="w")

    tk.Label(container, text="Adres:", bg="white", font=("Arial", 20)).grid(row=2, column=0, sticky="w")
    tk.Label(container, text=musteri_adres, bg="white", font=("Arial", 16)).grid(row=2, column=1, sticky="w")

    tk.Label(container, text="Telefon:", bg="white", font=("Arial", 20)).grid(row=3, column=0, sticky="w")
    tk.Label(container, text=musteri_tel, bg="white", font=("Arial", 20, "bold")).grid(row=3, column=1, sticky="w")

    tk.Label(container, text="Ürün:", bg="white", font=("Arial", 20)).grid(row=4, column=0, sticky="w")
    tk.Label(container, text=result[5], bg="white", font=("Arial", 20)).grid(row=4, column=1, sticky="w")

    tk.Button(new_Form, text="Yazdır", bg="#4CAF50", fg="white", font=("Arial", 16), command=lambda: print_cargo_bill(musteri_ad, musteri_soyad, musteri_adres, musteri_tel, result[5], urun_resmi)).pack(pady=20)

def print_cargo_bill(musteri_ad, musteri_soyad, musteri_adres, musteri_tel, urun, urun_resmi):
    try:
        
        pdf = FPDF("p", "mm", "A4")
        pdf.add_page()
        pdf.add_font("Arial", "", f"{base_path}/fonts/arial.ttf", uni=True)
        pdf.add_font("Arial", "B", f"{base_path}/fonts/Arial-bold.ttf", uni=True)
        pdf.set_font("Arial", size=20)
        pdf.cell(0,10,"Ayvaz Otomotiv", align="C", ln=1)
        pdf.cell(40,10,"AD SOYAD:", align="L")
        pdf.cell(0,10,f"{musteri_ad} {musteri_soyad}", align="L", ln=1)
        pdf.cell(40,10,"ADRES:", align="L")
        pdf.cell(0,10,f"{musteri_adres}", align="L", ln=1)
        pdf.cell(40,10,"TELEFON:", align="L")
        pdf.cell(0,10,f"{musteri_tel}", align="L", ln=1)
        pdf.cell(40,10,"ÜRÜN:", align="L")
        pdf.cell(0,10,f"{urun}", align="L", ln=1)
        if urun_resmi[0] != None:
            pdf.image(f"{urun_resmi[0]}")
        pdf.output("deneme.pdf")

    except Exception as e:
        messagebox.showerror("Hata", str(e))

def run_query():
    sql = sql_entry.get("1.0", tk.END).strip()
    if not sql:
        messagebox.showwarning("Uyarı", "Lütfen bir SQL sorgusu giriniz.")
        return
    selected_db = sql_combo.get()
    db_path = databases[selected_db]

    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        cursor.execute(sql)
        
        if sql.lower().startswith("select"):
            rows = cursor.fetchall()
            cols = [desc[0] for desc in cursor.description]

            tree.delete(*tree.get_children())
            tree["column"] = cols

            for col in cols:
                tree.heading(col, text=col)
                tree.column(col, width=120)

                for row in rows:
                    tree.insert("", tk.END, values=row)

            else:
                conn.commit()
                messagebox.showinfo("Başarılı", f"{selected_db} veritabanında sorgu çalıştırıldı.")

            conn.close()

    except Exception as e:
        messagebox.showerror("SQL Hatası", str(e))
# === ANA EKRAN ===
root = tk.Tk()
root.title("Müşteri / Sipariş Yönetim Sistemi")
root.geometry("1500x800")
root.configure(bg="#f5f5f5")

style = ttk.Style()
style.theme_use("clam")
style.configure("TNotebook", background="#f5f5f5", borderwidth=0)
style.configure("TFrame", background="#f5f5f5")
style.configure("TLabel", background="#f5f5f5", font=("Arial", 12))
style.configure("TButton", font=("Arial", 12), padding=6)
style.configure("TEntry", font=("Arial", 12))
style.configure("TCombobox", font=("Arial", 12))

notebook = ttk.Notebook(root)
notebook.pack(expand=True, fill="both", padx=10, pady=10)

# Sekmeler
frame1 = ttk.Frame(notebook, padding=20)  
frame2 = ttk.Frame(notebook, padding=20)  
frame3 = ttk.Frame(notebook, padding=20)  
frame4 = ttk.Frame(notebook, padding=20)  
frame5 = ttk.Frame(notebook, padding=20)  
frame6 = ttk.Frame(notebook, padding=20) 
frame7 = ttk.Frame(notebook, padding=20)

notebook.add(frame1, text="Müşteri Sipariş Kontrol")
notebook.add(frame2, text="Ürün Bul")
notebook.add(frame3, text="Depo Giriş")
notebook.add(frame4, text="Müşteri Listesi")
notebook.add(frame5, text="Sipariş Listesi")
notebook.add(frame6, text="Ürün Listesi")
notebook.add(frame7, text="Sorgular")

form_container = tk.Frame(root, bg="#f5f5f5")

# === FRAME1: MÜŞTERİ SİPARİŞ KONTROL ===
frame_left = tk.Frame(frame1, padx=20, pady=20, bg="#ffffff", relief="groove", bd=2)
frame_left.pack(side="left", fill="y", padx=10, pady=10)

phone_var = tk.StringVar()
tk.Label(frame_left, text="Telefon:", anchor="w", font=("Arial", 13, "bold"), bg="#ffffff").grid(row=0, column=0, sticky="w", pady=8)
entry_phone = tk.Entry(frame_left, width=30, textvariable=phone_var, font=("Arial", 12))
entry_phone.grid(row=0, column=1, pady=8)

tk.Label(frame_left, text="Ad:", anchor="w", font=("Arial", 13, "bold"), bg="#ffffff").grid(row=1, column=0, sticky="w", pady=8)
entry_name = tk.Entry(frame_left, width=30, font=("Arial", 12))
entry_name.grid(row=1, column=1, pady=8)

tk.Label(frame_left, text="Soyad:", anchor="w", font=("Arial", 13, "bold"), bg="#ffffff").grid(row=2, column=0, sticky="w", pady=8)
entry_surname = tk.Entry(frame_left, width=30, font=("Arial", 12))
entry_surname.grid(row=2, column=1, pady=8)

tk.Label(frame_left, text="İl:", anchor="w", font=("Arial", 13, "bold"), bg="#ffffff").grid(row=3, column=0, sticky="w", pady=8)
combo_state = ttk.Combobox(frame_left, values=iller_list, state="readonly", width=27)
combo_state.grid(row=3, column=1, pady=8)
combo_state.bind("<<ComboboxSelected>>", update_districts)

tk.Label(frame_left, text="İlçe:", anchor="w", font=("Arial", 13, "bold"), bg="#ffffff").grid(row=4, column=0, sticky="w", pady=8)
combo_district = ttk.Combobox(frame_left, values=[], state="readonly", width=27)
combo_district.grid(row=4, column=1, pady=8)

tk.Label(frame_left, text="Adres:", anchor="w", font=("Arial", 13, "bold"), bg="#ffffff").grid(row=5, column=0, sticky="w", pady=8)
entry_address = tk.Entry(frame_left, width=30, font=("Arial", 12))
entry_address.grid(row=5, column=1, pady=8)

tk.Label(frame_left, text="Not:", anchor="w", font=("Arial", 13, "bold"), bg="#ffffff").grid(row=6, column=0, sticky="w", pady=8)
entry_note = tk.Entry(frame_left, width=30, font=("Arial", 12))
entry_note.grid(row=6, column=1, pady=8)

btn_newCustomer = tk.Button(frame_left, text="Müşteri Kaydet", width=20, bg="#4CAF50", fg="white", font=("Arial", 12, "bold"), command=newCustomer)
btn_newCustomer.grid(row=7, column=0, columnspan=2, pady=12)

btn_get = tk.Button(frame_left, text="Getir", width=20, bg="#2196F3", fg="white", font=("Arial", 12, "bold"), command=Get)
btn_get.grid(row=8, column=0, columnspan=2, pady=12)

btn_deleteOrder = tk.Button(frame_left, text="Seçili siparişi sil", width=20, bg="#FF9800", fg="white", font=("Arial", 12, "bold"), command=Delete)
btn_deleteOrder.grid(row=9, column=0, columnspan=2, pady=12)

btn_cargoBill = tk.Button(frame_left, text="Kargo Fişi Yazdır",width=20,bg="#9C27B0",fg="white",font=("Arial", 12, "bold"))
btn_cargoBill.grid(row=10, columnspan=2, pady=12)
btn_cargoBill.config(command=cargoBill)

# Sipariş tablosu sağda
frame_right = tk.Frame(frame1, bg="#f5f5f5")
frame_right.pack(side="right", fill="both", expand=True, padx=10, pady=10)
df_siparis = pd.read_sql_query("SELECT * FROM siparisler", conn_siparis)
table_siparis1 = EditableTable(frame_right,dataframe=df_siparis, table_name="siparisler", id_column="siparis_id", conn=conn_siparis, showtoolbar=True, showstatusbar=True)
table_siparis1.show()

# === FRAME2: ÜRÜN BUL ===
frame2_left = tk.Frame(frame2, padx=20, pady=20, bg="#ffffff", relief="groove", bd=2)
frame2_left.pack(side="left", fill="y", padx=10, pady=10)

active_num = tk.Label(frame2_left, textvariable=phone_var, font=("Arial", 13, "bold"), bg="#ffffff")
active_num.grid(row=0, column=0, columnspan=2, pady=12)
tk.Label(frame2_left, text="Ürün Grubu:", font=("Arial", 13, "bold"), bg="#ffffff").grid(row=1, column=0, sticky="w", pady=8)
entry_productGroup = tk.Entry(frame2_left, width=30, font=("Arial", 12))
entry_productGroup.grid(row=1, column=1, pady=8)

tk.Label(frame2_left, text="Ürün ID:", font=("Arial", 13, "bold"), bg="#ffffff").grid(row=2, column=0, sticky="w", pady=8)
entry_productID = tk.Entry(frame2_left, width=30, font=("Arial", 12))
entry_productID.grid(row=2, column=1, pady=8)

tk.Label(frame2_left, text="Kargo:", font=("Arial", 13, "bold"), bg="#ffffff").grid(row=3, column=0, sticky="w", pady=8)
entry_ship = tk.Entry(frame2_left, width=30, font=("Arial", 12))
entry_ship.grid(row=3, column=1, pady=8)

tk.Label(frame2_left, text="Adet:", font=("Arial", 13, "bold"), bg="#ffffff").grid(row=4, column=0, sticky="w", pady=8)
entry_qty = tk.Entry(frame2_left, width=30, font=("Arial", 12))
entry_qty.grid(row=4, column=1, pady=8)

tk.Label(frame2_left, text="Ödeme:", font=("Arial", 13, "bold"), bg="#ffffff").grid(row=5, column=0, sticky="w", pady=8)
entry_payment = tk.Entry(frame2_left, width=30, font=("Arial", 12))
entry_payment.grid(row=5, column=1, pady=8)

tk.Label(frame2_left, text="Durum:", font=("Arial", 13, "bold"), bg="#ffffff").grid(row=6, column=0, sticky="w", pady=8)
combobox_status = ttk.Combobox(frame2_left, 
    values=["Sipariş verildi", "Hazırlanıyor", "Kargoya verildi", "Teslim edildi"],
    state="readonly", width=27)
combobox_status.grid(row=6, column=1, pady=8)
combobox_status.set("Sipariş verildi")

btn_get2 = tk.Button(frame2_left, text="Ürünü Getir", width=20, bg="#2196F3", fg="white", font=("Arial", 12, "bold"), command=GetProduct)
btn_get2.grid(row=7, column=0, columnspan=2, pady=12)

btn_add = tk.Button(frame2_left, text="Seçili Müşteriye Siparişi Ekle", width=25, bg="#4CAF50", fg="white", font=("Arial", 12, "bold"), command=Add)
btn_add.grid(row=8, column=0, columnspan=2, pady=12)

image_box = tk.Label(frame2_left, width=50,height=50, bg="#eeeeee", relief="ridge")
image_box.grid(row=9, column=0, columnspan=2, pady=12)

frame2_table = tk.Frame(frame2, bg="#f5f5f5")
frame2_table.pack(side="right", fill="both", expand=True, padx=10, pady=10)
df_urun = pd.read_sql_query("SELECT * FROM urunler", conn_urun)
df_urun.drop(columns=["resim"], inplace=True)
table_urun2 = EditableTable(frame2_table, dataframe=df_urun, table_name="urunler", id_column="urunid", conn=conn_urun ,showtoolbar=True, showstatusbar=True)
table_urun2.show()
table_urun2.bind("<ButtonRelease-1>", on_row_selected)

# === FRAME3: DEPO GİRİŞ ===
frame3_table = tk.Frame(frame3, padx=20, pady=20, bg="#ffffff", relief="groove", bd=2)
frame3_table.pack(fill="both", expand=True, padx=10, pady=10)
for i in range(6):
    frame3_table.grid_columnconfigure(i, weight=1)
tk.Label(frame3_table, text="Ürün ID:", font=("Arial", 13, "bold"), bg="#ffffff").grid(row=0, column=0, sticky="w", padx=5, pady=8)
entry_stock_id = tk.Entry(frame3_table, width=30, font=("Arial", 12))
entry_stock_id.grid(row=0, column=1, padx=10, pady=8)

tk.Label(frame3_table, text="Ürün Grubu:", font=("Arial", 13, "bold"), bg="#ffffff").grid(row=1, column=0, sticky="w", padx=5, pady=8)
entry_stock_group = tk.Entry(frame3_table, width=30, font=("Arial", 12))
entry_stock_group.grid(row=1, column=1, padx=10, pady=8)

tk.Label(frame3_table, text="Ürün Model:", font=("Arial", 13, "bold"), bg="#ffffff").grid(row=2, column=0, sticky="w", padx=5, pady=8)
entry_stock_model = tk.Entry(frame3_table, width=30, font=("Arial", 12))
entry_stock_model.grid(row=2, column=1, padx=10, pady=8)

tk.Label(frame3_table, text="Ağırlık:", font=("Arial", 13, "bold"), bg="#ffffff").grid(row=0, column=2, sticky="w", padx=5, pady=8)
entry_stock_mass = tk.Entry(frame3_table, width=30, font=("Arial", 12))
entry_stock_mass.grid(row=0, column=3, padx=10, pady=8)

tk.Label(frame3_table, text="Geliş Fiyatı TL:", font=("Arial", 13, "bold"), bg="#ffffff").grid(row=1, column=2, sticky="w", padx=5, pady=8)
entry_stock_tl = tk.Entry(frame3_table, width=30, font=("Arial", 12))
entry_stock_tl.grid(row=1, column=3, padx=10, pady=8)

tk.Label(frame3_table, text="Geliş Fiyatı USD:", font=("Arial", 13, "bold"), bg="#ffffff").grid(row=2, column=2, sticky="w", padx=5, pady=8)
entry_stock_usdf = tk.Entry(frame3_table, width=30, font=("Arial", 12))
entry_stock_usdf.grid(row=2, column=3, padx=10, pady=8)

tk.Label(frame3_table, text="Adet:", font=("Arial", 13, "bold"), bg="#ffffff").grid(row=0, column=4, sticky="w", padx=5, pady=8)
entry_stock_qty = tk.Entry(frame3_table, width=30, font=("Arial", 12))
entry_stock_qty.grid(row=0, column=5, padx=10, pady=8)

tk.Label(frame3_table, text="Güncel Maliyet:", font=("Arial", 13, "bold"), bg="#ffffff").grid(row=1, column=4, sticky="w", padx=5, pady=8)
entry_stock_cost = tk.Entry(frame3_table, width=30, font=("Arial", 12))
entry_stock_cost.grid(row=1, column=5, padx=10, pady=8)

tk.Label(frame3_table, text="Satış Fiyatı TL:", font=("Arial", 13, "bold"), bg="#ffffff").grid(row=2, column=4, sticky="w", padx=5, pady=8)
entry_stock_price = tk.Entry(frame3_table, width=30, font=("Arial", 12))
entry_stock_price.grid(row=2, column=5, padx=10, pady=8)

btn_choose_img = tk.Button(frame3_table, text="Resim Seç",width=20, bg="#2196F3", fg="white", font=("Arial", 12, "bold"), command=img_to_bytes)
btn_choose_img.grid(row=0, column=6, padx=10, pady=8)

btn_add_stock = tk.Button(frame3_table,text="Stok girişini yap",width=20,bg="#4CAF50",fg="white",font=("Arial", 12, "bold"),command=Add_stock)
btn_add_stock.grid(row=4,column=0,padx=10,pady=12)

# === FRAME4: MÜŞTERİ TABLOSU ===
frame4_table = tk.Frame(frame4)
frame4_table.pack(fill="both", expand=True)
df_musteri = pd.read_sql_query("SELECT * FROM musteriler", conn_musteri)
table_musteri = EditableTable(frame4_table, dataframe=df_musteri, table_name="musteriler", id_column="musteri_id", conn=conn_musteri, showtoolbar=True, showstatusbar=True)
table_musteri.show()

# === FRAME5: SİPARİŞ TABLOSU ===
frame5_table = tk.Frame(frame5)
frame5_table.pack(fill="both", expand=True)
df_siparis = pd.read_sql_query("SELECT * FROM siparisler", conn_siparis)
table_siparis = EditableTable(frame5_table, dataframe=df_siparis, table_name="siparisler", id_column="siparis_id", conn=conn_siparis, showtoolbar=True, showstatusbar=True)
table_siparis.show()

# === FRAME6: ÜRÜN TABLOSU ===
frame6_table = tk.Frame(frame6)
frame6_table.pack(fill="both", expand=True)
df_urun = pd.read_sql_query("SELECT * FROM urunler", conn_urun)
df_urun.drop(columns=["resim"], inplace=True)
table_urun = EditableTable(frame6_table, dataframe=df_urun, table_name="urunler", id_column="urunid", conn=conn_urun, showtoolbar=True, showstatusbar=True)
table_urun.show()
phone_var.trace("w", on_text_change)

# === FRAME7: SORGULAR ===
sql_entry = tk.Text(frame7, height=4, width=80)
sql_entry.pack(padx=10, pady=5)
databases ={
    "Ürünler": "urunler.db",
    "Müşteriler": "musteriler.db",
    "Siparişler": "siparisler.db"
}
db_var = tk.StringVar(value=list(databases.keys())[0])
sql_combo = ttk.Combobox(frame7, textvariable=db_var)
sql_combo.pack(padx=10, pady=10)
sql_button = tk.Button(frame7, text="Çalıştır", command=run_query)
sql_button.pack(pady=5)

tree = ttk. Treeview(frame7, show="headings")
tree.pack(anchor="s", fill="both", padx=10, pady=10)

root.mainloop()