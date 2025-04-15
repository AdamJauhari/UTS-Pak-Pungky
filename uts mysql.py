import re
import datetime
try:
    import mysql.connector
except ImportError:
    print("Error: Modul 'mysql-connector-python' belum terinstall.")
    print("Silakan install terlebih dahulu dengan perintah:")
    print("pip install mysql-connector-python")
    exit()
from mysql.connector import Error
try:
    import pandas as pd
except ImportError:
    print("Error: Modul 'pandas' belum terinstall.")
    print("Silakan install terlebih dahulu dengan perintah:")
    print("pip install pandas")
    exit()

# ==============================================
# FUNGSI UTILITAS DAN VALIDASI INPUT
# ==============================================

def get_int_input(prompt):
    """Mendapatkan input integer dengan validasi"""
    while True:
        try:
            return int(input(prompt))
        except ValueError:
            print("Masukkan harus berupa bilangan bulat. Silakan coba lagi.")

def get_float_input(prompt):
    """Mendapatkan input float dengan validasi"""
    while True:
        try:
            return float(input(prompt))
        except ValueError:
            print("Masukkan harus berupa angka. Silakan coba lagi.")

def get_date_input(prompt):
    """Mendapatkan input tanggal dengan validasi format YYYY-MM-DD"""
    while True:
        date_str = input(prompt)
        if re.match(r'^\d{4}-\d{2}-\d{2}$', date_str):
            try:
                datetime.datetime.strptime(date_str, '%Y-%m-%d')
                return date_str
            except ValueError:
                print("Tanggal tidak valid. Gunakan format YYYY-MM-DD.")
        else:
            print("Format tanggal tidak valid. Gunakan format YYYY-MM-DD.")

def validate_non_empty(prompt):
    """Validasi input tidak boleh kosong"""
    while True:
        value = input(prompt).strip()
        if value:
            return value
        print("Input tidak boleh kosong. Silakan coba lagi.")

# ==============================================
# FUNGSI DATABASE DAN TABEL
# ==============================================

def create_database_connection(database="zakat"):
    """Membuat koneksi ke database MySQL"""
    try:
        connection = mysql.connector.connect(
            host="localhost",
            user="root",
            password="",
            database=database
        )
        return connection
    except Error as e:
        print(f"Error connecting to MySQL: {e}")
        return None

def create_tables():
    """Membuat database dan tabel jika belum ada"""
    try:
        # Buat koneksi ke MySQL (tanpa spesifik database)
        connection = mysql.connector.connect(
            host="localhost",
            user="root",
            password=""
        )
        
        cursor = connection.cursor()
        
        # Buat database jika belum ada
        cursor.execute("CREATE DATABASE IF NOT EXISTS zakat")
        cursor.execute("USE zakat")
        
        # Buat tabel zakat_data
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS zakat_data (
            id INT AUTO_INCREMENT PRIMARY KEY,
            nama VARCHAR(100) NOT NULL,
            jenis_zakat ENUM('Fitrah', 'Mal') NOT NULL,
            jumlah DECIMAL(10, 2) NOT NULL,
            tanggal DATE NOT NULL
        )
        """)
        
        # Buat tabel master_beras
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS master_beras (
            id INT AUTO_INCREMENT PRIMARY KEY,
            nama_beras VARCHAR(50) NOT NULL,
            harga_per_kg DECIMAL(10, 2) NOT NULL
        )
        """)
        
        # Buat tabel transaksi_zakat
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS transaksi_zakat (
            id INT AUTO_INCREMENT PRIMARY KEY,
            id_zakat INT NOT NULL,
            id_beras INT NOT NULL,
            jumlah_beras DECIMAL(10, 2) NOT NULL,
            total_harga DECIMAL(10, 2) NOT NULL,
            tanggal DATE NOT NULL,
            FOREIGN KEY (id_zakat) REFERENCES zakat_data(id),
            FOREIGN KEY (id_beras) REFERENCES master_beras(id)
        )
        """)
        
        # Tambahkan data default jika tabel master_beras kosong
        cursor.execute("SELECT COUNT(*) FROM master_beras")
        if cursor.fetchone()[0] == 0:
            cursor.execute("""
            INSERT INTO master_beras (nama_beras, harga_per_kg) 
            VALUES 
                ('Beras Premium', 15000.00),
                ('Beras Medium', 12000.00),
                ('Beras Standard', 10000.00)
            """)
        
        connection.commit()
        print("Database dan tabel berhasil dibuat/diperiksa")
        return True
        
    except Error as e:
        print(f"Error creating database and tables: {e}")
        return False
    finally:
        if connection.is_connected():
            cursor.close()
            connection.close()

# ==============================================
# FUNGSI OPERASI DATABASE
# ==============================================

def add_zakat(nama, jenis_zakat, jumlah, tanggal):
    """Menambahkan data pembayar zakat baru"""
    conn = create_database_connection()
    if not conn:
        return False
    cursor = None
    try:
        cursor = conn.cursor()
        query = "INSERT INTO zakat_data (nama, jenis_zakat, jumlah, tanggal) VALUES (%s, %s, %s, %s)"
        cursor.execute(query, (nama, jenis_zakat, jumlah, tanggal))
        conn.commit()
        return True
    except Error as err:
        print(f"Error database: {err}")
        conn.rollback()
        return False
    finally:
        if cursor: cursor.close()
        conn.close()

def update_zakat(id, nama, jenis_zakat, jumlah, tanggal):
    """Memperbarui data pembayar zakat"""
    conn = create_database_connection()
    if not conn:
        return False
    cursor = None
    try:
        cursor = conn.cursor()
        query = """UPDATE zakat_data 
                SET nama = %s, jenis_zakat = %s, jumlah = %s, tanggal = %s 
                WHERE id = %s"""
        cursor.execute(query, (nama, jenis_zakat, jumlah, tanggal, id))
        conn.commit()
        return cursor.rowcount > 0
    except Error as err:
        print(f"Error database: {err}")
        conn.rollback()
        return False
    finally:
        if cursor: cursor.close()
        conn.close()

def delete_zakat(id):
    """Menghapus data pembayar zakat"""
    conn = create_database_connection()
    if not conn:
        return False
    cursor = None
    try:
        cursor = conn.cursor()
        
        # Cek apakah ada transaksi terkait
        cursor.execute("SELECT COUNT(*) FROM transaksi_zakat WHERE id_zakat = %s", (id,))
        if cursor.fetchone()[0] > 0:
            print("Tidak bisa menghapus. Data memiliki transaksi terkait.")
            return False
            
        cursor.execute("DELETE FROM zakat_data WHERE id = %s", (id,))
        conn.commit()
        return cursor.rowcount > 0
    except Error as err:
        print(f"Error database: {err}")
        conn.rollback()
        return False
    finally:
        if cursor: cursor.close()
        conn.close()

def add_beras(nama_beras, harga_per_kg):
    """Menambahkan data master beras"""
    conn = create_database_connection()
    if not conn:
        return False
    cursor = None
    try:
        cursor = conn.cursor()
        query = "INSERT INTO master_beras (nama_beras, harga_per_kg) VALUES (%s, %s)"
        cursor.execute(query, (nama_beras, harga_per_kg))
        conn.commit()
        return True
    except Error as err:
        print(f"Error database: {err}")
        conn.rollback()
        return False
    finally:
        if cursor: cursor.close()
        conn.close()

def view_master_beras():
    """Menampilkan data master beras"""
    conn = create_database_connection()
    if not conn:
        return
    cursor = None
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM master_beras")
        results = cursor.fetchall()
        
        if not results:
            print("\nBelum ada data master beras")
            return
            
        print("\n{:<5} {:<20} {:<15}".format("ID", "Nama Beras", "Harga per Kg"))
        print("-"*45)
        for row in results:
            print("{:<5} {:<20} Rp{:<10,.2f}".format(row[0], row[1], row[2]))
    except Error as err:
        print(f"Error database: {err}")
    finally:
        if cursor: cursor.close()
        conn.close()

def add_transaksi_zakat(id_zakat, id_beras, jumlah_beras, tanggal):
    """Menambahkan transaksi zakat beras"""
    conn = create_database_connection()
    if not conn:
        return False
    cursor = None
    try:
        cursor = conn.cursor()
        
        # Validasi ID zakat
        cursor.execute("SELECT id FROM zakat_data WHERE id = %s", (id_zakat,))
        if not cursor.fetchone():
            print("ID zakat tidak valid!")
            return False
            
        # Validasi dan ambil harga beras
        cursor.execute("SELECT harga_per_kg FROM master_beras WHERE id = %s", (id_beras,))
        result = cursor.fetchone()
        if not result:
            print("ID beras tidak valid!")
            return False
            
        total_harga = result[0] * jumlah_beras
        query = """INSERT INTO transaksi_zakat 
                (id_zakat, id_beras, jumlah_beras, total_harga, tanggal)
                VALUES (%s, %s, %s, %s, %s)"""
        cursor.execute(query, (id_zakat, id_beras, jumlah_beras, total_harga, tanggal))
        conn.commit()
        return True
    except Error as err:
        print(f"Error database: {err}")
        conn.rollback()
        return False
    finally:
        if cursor: cursor.close()
        conn.close()

def view_transaksi_zakat():
    """Menampilkan data transaksi zakat"""
    conn = create_database_connection()
    if not conn:
        return
    cursor = None
    try:
        cursor = conn.cursor(dictionary=True)
        query = """SELECT tz.id, z.nama, z.jenis_zakat, mb.nama_beras, 
                tz.jumlah_beras, tz.total_harga, tz.tanggal
                FROM transaksi_zakat tz
                JOIN zakat_data z ON tz.id_zakat = z.id
                JOIN master_beras mb ON tz.id_beras = mb.id"""
        cursor.execute(query)
        
        results = cursor.fetchall()
        if not results:
            print("\nBelum ada data transaksi")
            return
            
        print("\n{:<5} {:<20} {:<15} {:<15} {:<10} {:<15} {:<10}".format(
            "ID", "Nama", "Jenis Zakat", "Beras", "Jumlah", "Total", "Tanggal"))
        print("-"*90)
        for row in results:
            print(f"{row['id']:<5} {row['nama']:<20} {row['jenis_zakat']:<15} "
                f"{row['nama_beras']:<15} {row['jumlah_beras']:<10} "
                f"Rp{row['total_harga']:<10,.2f} {row['tanggal']}")
    except Error as err:
        print(f"Error database: {err}")
    finally:
        if cursor: cursor.close()
        conn.close()

def export_to_excel():
    """Mengekspor data zakat ke file Excel"""
    try:
        conn = create_database_connection()
        if not conn:
            return
            
        # Ekspor data zakat
        df_zakat = pd.read_sql("SELECT * FROM zakat_data", conn)
        df_zakat.to_excel("data_zakat.xlsx", index=False)
        
        # Ekspor data transaksi
        df_transaksi = pd.read_sql("""
            SELECT tz.id, z.nama, z.jenis_zakat, mb.nama_beras, 
                   tz.jumlah_beras, tz.total_harga, tz.tanggal
            FROM transaksi_zakat tz
            JOIN zakat_data z ON tz.id_zakat = z.id
            JOIN master_beras mb ON tz.id_beras = mb.id
        """, conn)
        df_transaksi.to_excel("data_transaksi_zakat.xlsx", index=False)
        
        print("Data berhasil diekspor ke:")
        print("- data_zakat.xlsx (Data pembayar zakat)")
        print("- data_transaksi_zakat.xlsx (Data transaksi zakat)")
    except Exception as e:
        print(f"Error ekspor data: {e}")
    finally:
        if conn: conn.close()

# ==============================================
# FUNGSI MENU UTAMA
# ==============================================

def menu_tambah_zakat():
    """Menu untuk menambahkan data pembayar zakat"""
    print("\n=== TAMBAH DATA PEMBAYAR ZAKAT ===")
    nama = validate_non_empty("Nama lengkap: ")
    
    jenis_zakat = input("Jenis zakat (Fitrah/Mal): ").strip().capitalize()
    while jenis_zakat not in ["Fitrah", "Mal"]:
        print("Jenis zakat harus Fitrah atau Mal!")
        jenis_zakat = input("Jenis zakat (Fitrah/Mal): ").strip().capitalize()
        
    jumlah = get_float_input("Jumlah zakat (Rp): ")
    tanggal = get_date_input("Tanggal pembayaran (YYYY-MM-DD): ")
    
    if add_zakat(nama, jenis_zakat, jumlah, tanggal):
        print("\nData berhasil disimpan!")
    else:
        print("\nGagal menyimpan data!")

def menu_edit_zakat():
    """Menu untuk mengedit data pembayar zakat"""
    print("\n=== EDIT DATA PEMBAYAR ZAKAT ===")
    id_zakat = get_int_input("ID pembayar yang akan diedit: ")
    
    # Validasi ID
    conn = create_database_connection()
    if not conn:
        return
    
    cursor = conn.cursor(dictionary=True)
    cursor.execute("SELECT * FROM zakat_data WHERE id = %s", (id_zakat,))
    old_data = cursor.fetchone()
    conn.close()
    
    if not old_data:
        print("ID tidak ditemukan!")
        return
    
    print("\nData saat ini:")
    print(f"Nama: {old_data['nama']}")
    print(f"Jenis Zakat: {old_data['jenis_zakat']}")
    print(f"Jumlah: Rp{old_data['jumlah']:,.2f}")
    print(f"Tanggal: {old_data['tanggal']}")
    
    print("\nMasukkan data baru (kosongkan jika tidak ingin diubah):")
    nama = input(f"Nama [{old_data['nama']}]: ").strip() or old_data['nama']
    
    jenis_zakat = input(f"Jenis zakat (Fitrah/Mal) [{old_data['jenis_zakat']}]: ").strip().capitalize()
    while jenis_zakat and jenis_zakat not in ["Fitrah", "Mal"]:
        print("Jenis zakat harus Fitrah atau Mal!")
        jenis_zakat = input(f"Jenis zakat (Fitrah/Mal) [{old_data['jenis_zakat']}]: ").strip().capitalize()
    jenis_zakat = jenis_zakat or old_data['jenis_zakat']
    
    jumlah = input(f"Jumlah zakat (Rp) [{old_data['jumlah']}]: ").strip()
    jumlah = float(jumlah) if jumlah else old_data['jumlah']
    
    tanggal = input(f"Tanggal pembayaran (YYYY-MM-DD) [{old_data['tanggal']}]: ").strip()
    while tanggal and not re.match(r'^\d{4}-\d{2}-\d{2}$', tanggal):
        print("Format tanggal tidak valid. Gunakan format YYYY-MM-DD.")
        tanggal = input(f"Tanggal pembayaran (YYYY-MM-DD) [{old_data['tanggal']}]: ").strip()
    tanggal = tanggal or old_data['tanggal']
    
    if update_zakat(id_zakat, nama, jenis_zakat, jumlah, tanggal):
        print("\nData berhasil diperbarui!")
    else:
        print("\nGagal memperbarui data!")

def menu_hapus_zakat():
    """Menu untuk menghapus data pembayar zakat"""
    print("\n=== HAPUS DATA PEMBAYAR ZAKAT ===")
    id_zakat = get_int_input("ID pembayar yang akan dihapus: ")
    
    if delete_zakat(id_zakat):
        print("\nData berhasil dihapus!")
    else:
        print("\nGagal menghapus data. Pastikan tidak ada transaksi terkait.")

def menu_tambah_beras():
    """Menu untuk menambahkan data master beras"""
    print("\n=== TAMBAH DATA MASTER BERAS ===")
    nama_beras = validate_non_empty("Nama jenis beras: ")
    harga = get_float_input("Harga per kg (Rp): ")
    
    if add_beras(nama_beras, harga):
        print("\nData beras berhasil ditambahkan!")
    else:
        print("\nGagal menambahkan data beras!")

def menu_tambah_transaksi():
    """Menu untuk menambahkan transaksi zakat beras"""
    print("\n=== TAMBAH TRANSAKSI ZAKAT BERAS ===")
    
    # Tampilkan daftar pembayar zakat
    conn = create_database_connection()
    if not conn:
        return
    
    cursor = conn.cursor(dictionary=True)
    cursor.execute("SELECT id, nama FROM zakat_data")
    pembayar = cursor.fetchall()
    
    if not pembayar:
        print("Belum ada data pembayar zakat. Silakan tambahkan dulu.")
        conn.close()
        return
    
    print("\nDaftar Pembayar Zakat:")
    for p in pembayar:
        print(f"{p['id']}. {p['nama']}")
    
    id_zakat = get_int_input("\nPilih ID pembayar: ")
    if id_zakat not in [p['id'] for p in pembayar]:
        print("ID pembayar tidak valid!")
        conn.close()
        return
    
    # Tampilkan daftar beras
    cursor.execute("SELECT id, nama_beras, harga_per_kg FROM master_beras")
    beras = cursor.fetchall()
    conn.close()
    
    if not beras:
        print("Belum ada data master beras. Silakan tambahkan dulu.")
        return
    
    print("\nDaftar Beras:")
    for b in beras:
        print(f"{b['id']}. {b['nama_beras']} (Rp{b['harga_per_kg']:,.2f}/kg)")
    
    id_beras = get_int_input("\nPilih ID beras: ")
    if id_beras not in [b['id'] for b in beras]:
        print("ID beras tidak valid!")
        return
    
    jumlah_beras = get_float_input("Jumlah beras (kg): ")
    tanggal = get_date_input("Tanggal transaksi (YYYY-MM-DD): ")
    
    if add_transaksi_zakat(id_zakat, id_beras, jumlah_beras, tanggal):
        print("\nTransaksi berhasil dicatat!")
    else:
        print("\nGagal membuat transaksi.")

# ==============================================
# MAIN PROGRAM
# ==============================================

def main():
    # Cek dan buat database/tabel jika belum ada
    print("Memeriksa database dan tabel...")
    if not create_tables():
        print("Gagal mempersiapkan database. Aplikasi akan keluar.")
        return
    
    # Menu utama
    while True:
        print("\n===== SISTEM MANAJEMEN ZAKAT =====")
        print("1. Tambah Data Pembayar Zakat")
        print("2. Edit Data Pembayar Zakat")
        print("3. Hapus Data Pembayar Zakat")
        print("4. Lihat Master Beras")
        print("5. Tambah Master Beras")
        print("6. Buat Transaksi Zakat Beras")
        print("7. Lihat Transaksi Zakat")
        print("8. Ekspor Data ke Excel")
        print("9. Keluar")
        
        choice = input("\nPilih menu [1-9]: ").strip()
        
        if choice == "1":
            menu_tambah_zakat()
        elif choice == "2":
            menu_edit_zakat()
        elif choice == "3":
            menu_hapus_zakat()
        elif choice == "4":
            print("\n=== DAFTAR MASTER BERAS ===")
            view_master_beras()
        elif choice == "5":
            menu_tambah_beras()
        elif choice == "6":
            menu_tambah_transaksi()
        elif choice == "7":
            print("\n=== DAFTAR TRANSAKSI ZAKAT ===")
            view_transaksi_zakat()
        elif choice == "8":
            print("\n=== EKSPOR DATA ===")
            export_to_excel()
        elif choice == "9":
            print("\nTerima kasih telah menggunakan Sistem Manajemen Zakat.")
            break
        else:
            print("\nPilihan tidak valid. Silakan pilih 1-9.")

if __name__ == "__main__":
    main()