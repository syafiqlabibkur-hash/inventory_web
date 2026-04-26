from flask import Flask, render_template, request, redirect, url_for, session
from datetime import datetime
from typing import List, Dict
import gspread
from google.oauth2.service_account import Credentials

app = Flask(__name__)


def format_rupiah(angka):
    try:
        return f"{int(angka):,}".replace(",", ".")
    except:
        return angka


app.jinja_env.filters['rupiah'] = format_rupiah
app.secret_key = "secret123"

SERVICE_ACCOUNT_FILE = "service_account.json"
SPREADSHEET_ID = "1yK-WxTeEFutwPMvAV6vMnsXgG4G7LAIv5oZQEqMQAuU"
SHEET_BARANG = "barang"
SHEET_TRANSAKSI = "transaksi"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


class SheetsDB:
    def __init__(self):
        creds = Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE,
            scopes=SCOPES,
        )
        self.client = gspread.authorize(creds)
        self.spreadsheet = self.client.open_by_key(SPREADSHEET_ID)

        self.ws_barang = self._get_or_create_worksheet(
            SHEET_BARANG,
            ["user", "kode", "nama", "kategori", "satuan", "stok", "updated_at"],
        )
        self.ws_transaksi = self._get_or_create_worksheet(
            SHEET_TRANSAKSI,
            [
                "user",
                "id_transaksi",
                "tanggal",
                "jenis",
                "kode_barang",
                "nama_barang",
                "jumlah",
                "total_harga",
                "setor",
                "sisa_setoran",
                "status",
                "keterangan",
                "updated_at",
            ],
        )

    def _get_or_create_worksheet(self, title: str, headers: List[str]):
        try:
            ws = self.spreadsheet.worksheet(title)
        except gspread.WorksheetNotFound:
            ws = self.spreadsheet.add_worksheet(
                title=title,
                rows=2000,
                cols=max(20, len(headers) + 2)
            )
            ws.append_row(headers)
            return ws

        current = ws.row_values(1)
        if current != headers:
            ws.clear()
            ws.append_row(headers)
        return ws

    def now_str(self) -> str:
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    def get_barang(self, current_user: str) -> List[Dict]:
        records = self.ws_barang.get_all_records()
        items = []

        for row in records:
            item = {
                "user": str(row.get("user", "")).strip(),
                "kode": str(row.get("kode", "")).strip(),
                "nama": str(row.get("nama", "")).strip(),
                "kategori": str(row.get("kategori", "")).strip(),
                "satuan": str(row.get("satuan", "")).strip(),
                "stok": int(row.get("stok", 0) or 0),
                "updated_at": str(row.get("updated_at", "")).strip(),
            }

            if item["kode"] and item["user"].lower() == current_user.lower():
                items.append(item)

        return items

    def tambah_barang(self, current_user: str, kode: str, nama: str, kategori: str, satuan: str, stok: int):
        data = self.get_barang(current_user)

        if any(x["kode"].lower() == kode.lower() for x in data):
            raise ValueError("Kode barang sudah ada untuk user ini.")

        self.ws_barang.append_row(
            [current_user, kode, nama, kategori, satuan, stok, self.now_str()]
        )

    def edit_barang(self, current_user: str, kode_lama: str, nama: str, kategori: str, satuan: str, stok: int):
        records = self.ws_barang.get_all_records()

        for idx, row in enumerate(records, start=2):
            user = str(row.get("user", "")).strip()
            kode = str(row.get("kode", "")).strip()

            if user.lower() == current_user.lower() and kode == kode_lama:
                self.ws_barang.update(
                    f"C{idx}:G{idx}",
                    [[nama, kategori, satuan, stok, self.now_str()]]
                )
                return

        raise ValueError("Barang tidak ditemukan.")


def hapus_barang(self, current_user: str, kode_barang: str):
    records = self.ws_barang.get_all_records()

    for idx, row in enumerate(records, start=2):
        user = str(row.get("user", "")).strip()
        kode = str(row.get("kode", "")).strip()

        if user.lower() == current_user.lower() and kode == kode_barang:
            self.ws_barang.delete_rows(idx)
            return

    raise ValueError("Barang tidak ditemukan.")

    def get_transaksi(self, current_user: str) -> List[Dict]:
        records = self.ws_transaksi.get_all_records()
        rows = []

        for row in records:
            trx = {
                "user": str(row.get("user", "")).strip(),
                "id_transaksi": str(row.get("id_transaksi", "")).strip(),
                "tanggal": str(row.get("tanggal", "")).strip(),
                "jenis": str(row.get("jenis", "")).strip(),
                "kode_barang": str(row.get("kode_barang", "")).strip(),
                "nama_barang": str(row.get("nama_barang", "")).strip(),
                "jumlah": int(row.get("jumlah", 0) or 0),
                "total_harga": int(row.get("total_harga", 0) or 0),
                "setor": int(row.get("setor", 0) or 0),
                "sisa_setoran": int(row.get("sisa_setoran", 0) or 0),
                "status": str(row.get("status", "")).strip(),
                "keterangan": str(row.get("keterangan", "")).strip(),
                "updated_at": str(row.get("updated_at", "")).strip(),
            }

            if trx["id_transaksi"] and trx["user"].lower() == current_user.lower():
                rows.append(trx)

        return rows

    def update_stok_barang(self, current_user: str, kode_barang: str, jumlah: int, jenis: str):
        records = self.ws_barang.get_all_records()

        for idx, row in enumerate(records, start=2):
            user = str(row.get("user", "")).strip()
            kode = str(row.get("kode", "")).strip()

            if user.lower() == current_user.lower() and kode == kode_barang:
                stok_lama = int(row.get("stok", 0) or 0)

                if jenis == "Masuk":
                    stok_baru = stok_lama + jumlah
                else:
                    stok_baru = stok_lama - jumlah
                    if stok_baru < 0:
                        raise ValueError(
                            "Stok tidak mencukupi untuk transaksi keluar.")

                self.ws_barang.update(
                    f"F{idx}:G{idx}", [[stok_baru, self.now_str()]])
                return

        raise ValueError("Barang tidak ditemukan.")

    def tambah_transaksi(
        self,
        current_user: str,
        tanggal: str,
        jenis: str,
        kode_barang: str,
        nama_barang: str,
        jumlah: int,
        total_harga: int,
        setor: int,
        keterangan: str,
    ):
        trx_id = f"TRX-{datetime.now().strftime('%Y%m%d%H%M%S')}"
        sisa_setoran = total_harga - setor

        if sisa_setoran <= 0:
            status = "Lunas"
            sisa_setoran = 0
        else:
            status = "Belum Lunas"

        self.update_stok_barang(current_user, kode_barang, jumlah, jenis)

        self.ws_transaksi.append_row([
            current_user,
            trx_id,
            tanggal,
            jenis,
            kode_barang,
            nama_barang,
            jumlah,
            total_harga,
            setor,
            sisa_setoran,
            status,
            keterangan,
            self.now_str(),
        ])


def parse_rupiah(text):
    text = str(text).strip()
    if not text:
        return 0
    return int(text.replace(".", "").replace(",", ""))


def get_db():
    return SheetsDB()


@app.route("/")
def home():
    return render_template("index.html")


@app.route("/login", methods=["POST"])
def login():
    username = request.form.get("username", "").strip()

    if not username:
        return "Username kosong!"

    session["user"] = username
    return redirect(url_for("dashboard"))


@app.route("/dashboard")
def dashboard():
    user = session.get("user")

    if not user:
        return redirect(url_for("home"))

    db = get_db()
    barang = db.get_barang(user)

    total_barang = len(barang)
    total_stok = sum(int(x["stok"]) for x in barang)

    return render_template(
        "dashboard.html",
        user=user,
        total_barang=total_barang,
        total_stok=total_stok,
    )


@app.route("/barang", methods=["GET", "POST"])
def barang():
    user = session.get("user")

    if not user:
        return redirect(url_for("home"))

    db = get_db()
    error = None

    if request.method == "POST":
        try:
            kode = request.form.get("kode", "").strip()
            nama = request.form.get("nama", "").strip()
            kategori = request.form.get("kategori", "").strip()
            satuan = request.form.get("satuan", "").strip()
            stok_text = request.form.get("stok", "").strip()

            if not all([kode, nama, kategori, satuan, stok_text]):
                raise ValueError("Semua field wajib diisi.")

            if not stok_text.isdigit():
                raise ValueError("Stok harus angka.")

            db.tambah_barang(
                current_user=user,
                kode=kode,
                nama=nama,
                kategori=kategori,
                satuan=satuan,
                stok=int(stok_text),
            )

            return redirect(url_for("barang"))

        except Exception as e:
            error = str(e)

    rows = db.get_barang(user)
    return render_template("barang.html", user=user, rows=rows, error=error)


@app.route("/transaksi", methods=["GET", "POST"])
def transaksi():
    user = session.get("user")

    if not user:
        return redirect(url_for("home"))

    db = get_db()
    error = None
    barang_list = db.get_barang(user)

    if request.method == "POST":
        try:
            tanggal = request.form.get("tanggal", "").strip()
            jenis = request.form.get("jenis", "").strip()
            kode_barang = request.form.get("kode_barang", "").strip()
            jumlah_text = request.form.get("jumlah", "").strip()
            total_harga_text = request.form.get("total_harga", "").strip()
            setor_text = request.form.get("setor", "").strip()
            keterangan = request.form.get("keterangan", "").strip()

            if not all([tanggal, jenis, kode_barang, jumlah_text]):
                raise ValueError(
                    "Tanggal, jenis, barang, dan jumlah wajib diisi.")

            jumlah = int(jumlah_text)
            total_harga = parse_rupiah(total_harga_text)
            setor = parse_rupiah(setor_text)

            selected_barang = None
            for item in barang_list:
                if item["kode"] == kode_barang:
                    selected_barang = item
                    break

            if not selected_barang:
                raise ValueError("Barang tidak ditemukan.")

            db.tambah_transaksi(
                current_user=user,
                tanggal=tanggal,
                jenis=jenis,
                kode_barang=selected_barang["kode"],
                nama_barang=selected_barang["nama"],
                jumlah=jumlah,
                total_harga=total_harga,
                setor=setor,
                keterangan=keterangan,
            )

            return redirect(url_for("transaksi"))

        except Exception as e:
            error = str(e)

    rows = db.get_transaksi(user)
    return render_template(
        "transaksi.html",
        user=user,
        barang_list=barang_list,
        rows=rows,
        error=error,
        today=datetime.now().strftime("%Y-%m-%d"),
    )


@app.route("/barang/edit/<kode>", methods=["POST"])
def edit_barang(kode):
    user = session.get("user")

    if not user:
        return redirect(url_for("home"))

    db = get_db()

    try:
        nama = request.form.get("nama", "").strip()
        kategori = request.form.get("kategori", "").strip()
        satuan = request.form.get("satuan", "").strip()
        stok = int(request.form.get("stok", 0))

        db.edit_barang(
            current_user=user,
            kode_lama=kode,
            nama=nama,
            kategori=kategori,
            satuan=satuan,
            stok=stok,
        )

    except Exception as e:
        print("ERROR EDIT:", e)

    return redirect(url_for("barang"))


@app.route("/barang/hapus/<kode>", methods=["POST"])
def hapus_barang(kode):
    user = session.get("user")

    if not user:
        return redirect(url_for("home"))

    db = get_db()

    try:
        db.hapus_barang(
            current_user=user,
            kode_barang=kode
        )

    except Exception as e:
        print("ERROR HAPUS:", e)

    return redirect(url_for("barang"))


@app.route("/logout")
def logout():
    session.pop("user", None)
    return redirect(url_for("home"))


if __name__ == "__main__":
    import os
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
