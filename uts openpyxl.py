try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.utils import get_column_letter
except ImportError:
    print("Error: Modul 'openpyxl' belum terinstall.")
    print("Silakan install terlebih dahulu dengan perintah:")
    print("pip install openpyxl")
    exit()

import os
from datetime import datetime
# Excel file paths
ZAKAT_DATA_FILE = "zakat_data.xlsx"
MASTER_BERAS_FILE = "master_beras.xlsx"
TRANSAKSI_ZAKAT_FILE = "transaksi_zakat.xlsx"

def initialize_files():
    """Initialize Excel files with headers if they don't exist"""
    if not os.path.exists(ZAKAT_DATA_FILE):
        wb = Workbook()
        ws = wb.active
        ws.title = "Zakat Data"
        ws.append(["ID", "Nama", "Jenis Zakat", "Jumlah", "Tanggal"])
        wb.save(ZAKAT_DATA_FILE)
    
    if not os.path.exists(MASTER_BERAS_FILE):
        wb = Workbook()
        ws = wb.active
        ws.title = "Master Beras"
        ws.append(["ID", "Nama Beras", "Harga per Kg"])
        wb.save(MASTER_BERAS_FILE)
    
    if not os.path.exists(TRANSAKSI_ZAKAT_FILE):
        wb = Workbook()
        ws = wb.active
        ws.title = "Transaksi Zakat"
        ws.append(["ID", "ID Zakat", "ID Beras", "Jumlah Beras", "Total Harga", "Tanggal"])
        wb.save(TRANSAKSI_ZAKAT_FILE)

def get_next_id(file_path):
    """Get the next available ID for a given Excel file"""
    if not os.path.exists(file_path):
        return 1
    
    wb = load_workbook(file_path)
    ws = wb.active
    max_id = 0
    
    for row in ws.iter_rows(min_row=2, max_col=1, values_only=True):
        if row[0] is not None and row[0] > max_id:
            max_id = row[0]
    
    return max_id + 1

def add_zakat(nama, jenis_zakat, jumlah, tanggal):
    """Add new zakat data to the Excel file"""
    try:
        wb = load_workbook(ZAKAT_DATA_FILE)
        ws = wb.active
        new_id = get_next_id(ZAKAT_DATA_FILE)
        
        ws.append([new_id, nama, jenis_zakat, jumlah, tanggal])
        wb.save(ZAKAT_DATA_FILE)
        return True
    except Exception as e:
        print(f"Error adding zakat: {e}")
        return False

def update_zakat(id, nama, jenis_zakat, jumlah, tanggal):
    """Update existing zakat data"""
    try:
        wb = load_workbook(ZAKAT_DATA_FILE)
        ws = wb.active
        found = False
        
        for row in ws.iter_rows(min_row=2):
            if row[0].value == id:
                row[1].value = nama
                row[2].value = jenis_zakat
                row[3].value = jumlah
                row[4].value = tanggal
                found = True
                break
        
        if found:
            wb.save(ZAKAT_DATA_FILE)
            return True
        else:
            print(f"ID {id} not found")
            return False
    except Exception as e:
        print(f"Error updating zakat: {e}")
        return False

def delete_zakat(id):
    """Delete zakat data by ID"""
    try:
        wb = load_workbook(ZAKAT_DATA_FILE)
        ws = wb.active
        rows_to_delete = []
        
        for idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
            if row[0].value == id:
                rows_to_delete.append(idx)
        
        for idx in sorted(rows_to_delete, reverse=True):
            ws.delete_rows(idx)
        
        if rows_to_delete:
            wb.save(ZAKAT_DATA_FILE)
            return True
        else:
            print(f"ID {id} not found")
            return False
    except Exception as e:
        print(f"Error deleting zakat: {e}")
        return False

def add_beras(nama_beras, harga_per_kg):
    """Add new beras data to the Excel file"""
    try:
        wb = load_workbook(MASTER_BERAS_FILE)
        ws = wb.active
        new_id = get_next_id(MASTER_BERAS_FILE)
        
        ws.append([new_id, nama_beras, harga_per_kg])
        wb.save(MASTER_BERAS_FILE)
        return True
    except Exception as e:
        print(f"Error adding beras: {e}")
        return False

def view_master_beras():
    """View all master beras data"""
    try:
        if not os.path.exists(MASTER_BERAS_FILE):
            print("No data available")
            return
        
        wb = load_workbook(MASTER_BERAS_FILE)
        ws = wb.active
        
        if ws.max_row <= 1:
            print("No data available")
            return
        
        print("\nMaster Data Beras:")
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0] is not None:
                print(f"ID: {row[0]}, Nama Beras: {row[1]}, Harga per Kg: {row[2]}")
    except Exception as e:
        print(f"Error viewing master beras: {e}")

def add_transaksi_zakat(id_zakat, id_beras, jumlah_beras, tanggal):
    """Add new zakat transaction"""
    try:
        # Check if zakat ID exists
        zakat_exists = False
        if os.path.exists(ZAKAT_DATA_FILE):
            wb_zakat = load_workbook(ZAKAT_DATA_FILE)
            ws_zakat = wb_zakat.active
            for row in ws_zakat.iter_rows(min_row=2, max_col=1, values_only=True):
                if row[0] == id_zakat:
                    zakat_exists = True
                    break
        
        if not zakat_exists:
            print("Error: ID zakat tidak ditemukan!")
            return False
        
        # Check if beras ID exists and get price
        beras_price = None
        if os.path.exists(MASTER_BERAS_FILE):
            wb_beras = load_workbook(MASTER_BERAS_FILE)
            ws_beras = wb_beras.active
            for row in ws_beras.iter_rows(min_row=2, values_only=True):
                if row[0] == id_beras:
                    beras_price = row[2]
                    break
        
        if beras_price is None:
            print("Error: ID beras tidak ditemukan!")
            return False
        
        total_harga = beras_price * jumlah_beras
        
        # Add transaction
        wb = load_workbook(TRANSAKSI_ZAKAT_FILE)
        ws = wb.active
        new_id = get_next_id(TRANSAKSI_ZAKAT_FILE)
        
        ws.append([new_id, id_zakat, id_beras, jumlah_beras, total_harga, tanggal])
        wb.save(TRANSAKSI_ZAKAT_FILE)
        print("Transaksi zakat berhasil ditambahkan!")
        return True
    except Exception as e:
        print(f"Error adding transaction: {e}")
        return False

def view_transaksi_zakat():
    """View all zakat transactions"""
    try:
        if not os.path.exists(TRANSAKSI_ZAKAT_FILE):
            print("No transaction data available")
            return
        
        wb_trans = load_workbook(TRANSAKSI_ZAKAT_FILE)
        ws_trans = wb_trans.active
        
        if ws_trans.max_row <= 1:
            print("No transaction data available")
            return
        
        # Load zakat data
        zakat_data = {}
        if os.path.exists(ZAKAT_DATA_FILE):
            wb_zakat = load_workbook(ZAKAT_DATA_FILE)
            ws_zakat = wb_zakat.active
            for row in ws_zakat.iter_rows(min_row=2, values_only=True):
                if row[0] is not None:
                    zakat_data[row[0]] = (row[1], row[2])  # (nama, jenis_zakat)
        
        # Load beras data
        beras_data = {}
        if os.path.exists(MASTER_BERAS_FILE):
            wb_beras = load_workbook(MASTER_BERAS_FILE)
            ws_beras = wb_beras.active
            for row in ws_beras.iter_rows(min_row=2, values_only=True):
                if row[0] is not None:
                    beras_data[row[0]] = row[1]  # nama_beras
        
        print("\nTransaksi Zakat:")
        for row in ws_trans.iter_rows(min_row=2, values_only=True):
            if row[0] is not None:
                zakat_info = zakat_data.get(row[1], ("Unknown", "Unknown"))
                beras_name = beras_data.get(row[2], "Unknown")
                
                print(f"ID Transaksi: {row[0]}, Nama Zakat: {zakat_info[0]}, Jenis Zakat: {zakat_info[1]}, "
                      f"Nama Beras: {beras_name}, Jumlah Beras: {row[3]}, "
                      f"Total Harga: {row[4]}, Tanggal: {row[5]}")
    except Exception as e:
        print(f"Error viewing transactions: {e}")

def export_to_excel():
    """Export zakat data to a new Excel file"""
    try:
        if not os.path.exists(ZAKAT_DATA_FILE):
            print("No data to export")
            return
        
        wb = load_workbook(ZAKAT_DATA_FILE)
        ws = wb.active
        
        if ws.max_row <= 1:
            print("No data to export")
            return
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"data_zakat_export_{timestamp}.xlsx"
        
        export_wb = Workbook()
        export_ws = export_wb.active
        export_ws.title = "Zakat Data Export"
        
        # Copy headers
        for row in ws.iter_rows(max_row=1):
            export_ws.append([cell.value for cell in row])
        
        # Copy data
        for row in ws.iter_rows(min_row=2):
            export_ws.append([cell.value for cell in row])
        
        export_wb.save(filename)
        print(f"Data zakat berhasil diekspor ke dalam file '{filename}'")
    except Exception as e:
        print(f"Error exporting data: {e}")

def input_master_beras():
    """Input new master beras data from user"""
    print("\nTambah Data Master Beras")
    nama_beras = input("Masukkan nama jenis beras: ")
    try:
        harga_per_kg = float(input("Masukkan harga per kg: "))
        if add_beras(nama_beras, harga_per_kg):
            print("Data master beras berhasil ditambahkan!")
    except ValueError:
        print("Harga harus berupa angka!")

def main():
    initialize_files()
    
    while True:
        print("\nMenu:")
        print("1. Tambah Data Zakat")
        print("2. Edit Data Zakat")
        print("3. Hapus Data Zakat")
        print("4. Lihat Data Master Beras")
        print("5. Tambah Data Master Beras")
        print("6. Tambah Transaksi Zakat")
        print("7. Lihat Transaksi Zakat")
        print("8. Ekspor Data Zakat ke Excel")
        print("9. Keluar")
        
        choice = input("Pilih opsi (1-9): ")
        
        if choice == "1":
            nama = input("Masukkan nama: ")
            jenis_zakat = input("Masukkan jenis zakat: ")
            try:
                jumlah = float(input("Masukkan jumlah zakat: "))
                tanggal = input("Masukkan tanggal (YYYY-MM-DD): ")
                if add_zakat(nama, jenis_zakat, jumlah, tanggal):
                    print("Data zakat berhasil ditambahkan.")
            except ValueError:
                print("Jumlah harus berupa angka!")
        
        elif choice == "2":
            try:
                id_zakat = int(input("Masukkan ID zakat yang ingin diubah: "))
                nama = input("Masukkan nama baru: ")
                jenis_zakat = input("Masukkan jenis zakat baru: ")
                try:
                    jumlah = float(input("Masukkan jumlah zakat baru: "))
                    tanggal = input("Masukkan tanggal baru (YYYY-MM-DD): ")
                    if update_zakat(id_zakat, nama, jenis_zakat, jumlah, tanggal):
                        print("Data zakat berhasil diperbarui.")
                except ValueError:
                    print("Jumlah harus berupa angka!")
            except ValueError:
                print("ID harus berupa angka!")
        
        elif choice == "3":
            try:
                id_zakat = int(input("Masukkan ID zakat yang ingin dihapus: "))
                if delete_zakat(id_zakat):
                    print("Data zakat berhasil dihapus.")
            except ValueError:
                print("ID harus berupa angka!")
        
        elif choice == "4":
            print("\nMaster Data Beras:")
            view_master_beras()
        
        elif choice == "5":
            input_master_beras()
        
        elif choice == "6":
            try:
                id_zakat = int(input("Masukkan ID zakat: "))
                id_beras = int(input("Masukkan ID beras: "))
                try:
                    jumlah_beras = float(input("Masukkan jumlah beras (kg): "))
                    tanggal = input("Masukkan tanggal (YYYY-MM-DD): ")
                    add_transaksi_zakat(id_zakat, id_beras, jumlah_beras, tanggal)
                except ValueError:
                    print("Jumlah beras harus berupa angka!")
            except ValueError:
                print("ID harus berupa angka!")
        
        elif choice == "7":
            print("\nTransaksi Zakat:")
            view_transaksi_zakat()
        
        elif choice == "8":
            export_to_excel()
        
        elif choice == "9":
            print("Keluar dari program.")
            break
        
        else:
            print("Pilihan tidak valid. Silakan coba lagi.")

if __name__ == "__main__":
    main()