# -*- coding: utf-8 -*-
"""
AutoNyusunExcel V2
------------------
Alur:
  1. Pilih file Excel
  2. Pilih Sheet
  3. Pilih Kolom + Value
  4. Isi nama sheet baru (hasil copy)
  5. PROSES (ada progress bar hijau) / BUKA FILE

Install dulu:
    pip install wxPython openpyxl
Jalankan dengan double-click file .pyw (tanpa jendela console).
"""

import os
import sys
import copy
import wx
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

APP_TITLE = "AutoNyusunExcel V2"
HEADER_ROW = 1          # baris judul kolom
BLANK_LABEL = "(Kosong)"
INVALID_SHEET_CHARS = set('[]:*?/\\')
GREEN = wx.Colour(40, 167, 69)


class GreenBar(wx.Panel):
    """Progress bar hijau (digambar manual supaya warnanya pasti hijau)."""

    def __init__(self, parent):
        super().__init__(parent)
        self.SetMinSize((-1, 24))
        self.value = 0
        self.Bind(wx.EVT_PAINT, self.on_paint)
        self.Bind(wx.EVT_ERASE_BACKGROUND, lambda e: None)

    def SetValue(self, v):
        self.value = max(0, min(100, int(v)))
        self.Refresh()
        self.Update()

    def on_paint(self, _evt):
        dc = wx.PaintDC(self)
        w, h = self.GetClientSize()
        dc.SetPen(wx.TRANSPARENT_PEN)
        dc.SetBrush(wx.Brush(wx.Colour(222, 222, 222)))
        dc.DrawRectangle(0, 0, w, h)
        fw = int(w * self.value / 100)
        if fw > 0:
            dc.SetBrush(wx.Brush(GREEN))
            dc.DrawRectangle(0, 0, fw, h)
        txt = f"{self.value}%"
        tw, th = dc.GetTextExtent(txt)
        dc.SetTextForeground(wx.WHITE if fw > w // 2 else wx.Colour(60, 60, 60))
        dc.DrawText(txt, (w - tw) // 2, (h - th) // 2)


class MainFrame(wx.Frame):
    def __init__(self):
        super().__init__(None, title=APP_TITLE, size=(520, 540),
                         style=wx.DEFAULT_FRAME_STYLE & ~(wx.RESIZE_BORDER | wx.MAXIMIZE_BOX))

        self.file_path = ""
        self.wb = None          # workbook asli (rumus + style)
        self.wb_vals = None     # workbook nilai (hasil hitung rumus)
        self.headers = []       # list nama kolom
        self.value_map = {}     # label value -> nilai asli

        self._build_ui()
        self.Centre()
        self._reset_from_step(2)

    # ------------------------------------------------------------------ UI
    def _build_ui(self):
        self.panel = wx.Panel(self)
        panel = self.panel
        main = wx.BoxSizer(wx.VERTICAL)

        title = wx.StaticText(panel, label="AutoNyusunExcel")
        title.SetFont(wx.Font(16, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        main.Add(title, 0, wx.ALL | wx.ALIGN_CENTER, 12)

        # Langkah 1
        self.btn_pilih = wx.Button(panel, label="1. Pilih Excel")
        self.btn_pilih.SetMinSize((-1, 34))
        self.lbl_file = wx.StaticText(panel, label="Belum ada file dipilih")
        self.lbl_file.SetForegroundColour(wx.Colour(110, 110, 110))
        main.Add(self.btn_pilih, 0, wx.LEFT | wx.RIGHT | wx.EXPAND, 15)
        main.Add(self.lbl_file, 0, wx.LEFT | wx.RIGHT | wx.TOP, 15)
        main.AddSpacer(10)

        # Langkah 2
        main.Add(wx.StaticText(panel, label="2. Pilih Sheet"), 0, wx.LEFT | wx.RIGHT, 15)
        self.cb_sheet = wx.ComboBox(panel, style=wx.CB_READONLY)
        main.Add(self.cb_sheet, 0, wx.ALL | wx.EXPAND, 15)

        # Langkah 3
        main.Add(wx.StaticText(panel, label="3. Pilih Tabel Kolom"), 0, wx.LEFT | wx.RIGHT, 15)
        self.cb_kolom = wx.ComboBox(panel, style=wx.CB_READONLY)
        main.Add(self.cb_kolom, 0, wx.LEFT | wx.RIGHT | wx.TOP | wx.EXPAND, 15)
        main.Add(wx.StaticText(panel, label="    Pilih VALUE"), 0, wx.LEFT | wx.RIGHT | wx.TOP, 15)
        self.cb_value = wx.ComboBox(panel, style=wx.CB_READONLY)
        main.Add(self.cb_value, 0, wx.ALL | wx.EXPAND, 15)

        # Langkah 4
        main.Add(wx.StaticText(panel, label="4. Copy to Sheet (nama sheet baru)"),
                 0, wx.LEFT | wx.RIGHT, 15)
        self.tx_nama = wx.TextCtrl(panel)
        main.Add(self.tx_nama, 0, wx.ALL | wx.EXPAND, 15)

        # Langkah 5
        row = wx.BoxSizer(wx.HORIZONTAL)
        self.btn_proses = wx.Button(panel, label="PROSES")
        self.btn_buka = wx.Button(panel, label="BUKA FILE")
        self.btn_proses.SetMinSize((-1, 38))
        self.btn_buka.SetMinSize((-1, 38))
        row.Add(self.btn_proses, 1, wx.RIGHT, 8)
        row.Add(self.btn_buka, 1)
        main.Add(row, 0, wx.LEFT | wx.RIGHT | wx.EXPAND, 15)

        # Progress bar hijau
        self.bar = GreenBar(panel)
        main.Add(self.bar, 0, wx.LEFT | wx.RIGHT | wx.TOP | wx.EXPAND, 15)

        self.lbl_status = wx.StaticText(panel, label="")
        main.Add(self.lbl_status, 0, wx.ALL, 15)

        panel.SetSizer(main)

        self.btn_pilih.Bind(wx.EVT_BUTTON, self.on_pilih_excel)
        self.cb_sheet.Bind(wx.EVT_COMBOBOX, self.on_sheet)
        self.cb_kolom.Bind(wx.EVT_COMBOBOX, self.on_kolom)
        self.btn_proses.Bind(wx.EVT_BUTTON, self.on_proses)
        self.btn_buka.Bind(wx.EVT_BUTTON, self.on_buka)

    def _reset_from_step(self, step):
        """Kosongkan kontrol mulai dari langkah tertentu."""
        if step <= 2:
            self.cb_sheet.Clear()
            self.cb_sheet.Enable(False)
        if step <= 3:
            self.cb_kolom.Clear()
            self.cb_value.Clear()
            self.cb_kolom.Enable(False)
            self.cb_value.Enable(False)
            self.value_map = {}
            self.headers = []
        self.btn_proses.Enable(False)
        self.btn_buka.Enable(bool(self.file_path))

    def set_status(self, msg, ok=True):
        self.lbl_status.SetLabel(msg)
        self.lbl_status.SetForegroundColour(wx.Colour(0, 130, 0) if ok else wx.Colour(190, 0, 0))
        self.lbl_status.Wrap(470)
        self.panel.Layout()

    def set_progress(self, pct):
        self.bar.SetValue(pct)
        wx.GetApp().Yield()

    # ------------------------------------------------------------- Handlers
    def on_pilih_excel(self, _evt):
        with wx.FileDialog(self, "Pilih file Excel",
                           wildcard="File Excel (*.xlsx;*.xlsm)|*.xlsx;*.xlsm",
                           style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as dlg:
            if dlg.ShowModal() != wx.ID_OK:
                return
            path = dlg.GetPath()

        try:
            wx.BeginBusyCursor()
            keep_vba = path.lower().endswith(".xlsm")
            self.wb = load_workbook(path, keep_vba=keep_vba)
            self.wb_vals = load_workbook(path, data_only=True)
        except Exception as e:
            self.wb = self.wb_vals = None
            self.set_status(f"Gagal membuka file: {e}", ok=False)
            return
        finally:
            if wx.IsBusy():
                wx.EndBusyCursor()

        self.file_path = path
        self.lbl_file.SetLabel(os.path.basename(path))
        self.bar.SetValue(0)
        self._reset_from_step(2)
        self.cb_sheet.AppendItems(self.wb.sheetnames)
        self.cb_sheet.Enable(True)
        self.set_status("File dimuat. Silakan pilih sheet.")

    def on_sheet(self, _evt):
        self._reset_from_step(3)
        ws = self.wb_vals[self.cb_sheet.GetStringSelection()]
        self.headers = []
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=HEADER_ROW, column=c).value
            if v is not None and str(v).strip() != "":
                self.headers.append((c, str(v).strip()))

        if not self.headers:
            self.set_status(f"Baris {HEADER_ROW} di sheet ini kosong (tidak ada judul kolom).", ok=False)
            return

        self.cb_kolom.AppendItems([h for _, h in self.headers])
        self.cb_kolom.Enable(True)
        self.set_status("Pilih tabel kolom.")

    def on_kolom(self, _evt):
        idx = self.cb_kolom.GetSelection()
        col_idx = self.headers[idx][0]
        ws = self.wb_vals[self.cb_sheet.GetStringSelection()]

        self.value_map = {}
        seen = []
        for r in range(HEADER_ROW + 1, ws.max_row + 1):
            v = ws.cell(row=r, column=col_idx).value
            label = BLANK_LABEL if (v is None or str(v).strip() == "") else str(v).strip()
            if label not in self.value_map:
                self.value_map[label] = label
                seen.append(label)

        self.cb_value.Clear()
        self.cb_value.AppendItems(sorted(seen, key=lambda s: s.lower()))
        self.cb_value.Enable(True)
        self.btn_proses.Enable(True)
        self.set_status(f"{len(seen)} value unik ditemukan. Pilih value lalu isi nama sheet.")

    def _validate_sheet_name(self, name):
        if not name:
            return "Nama sheet baru belum diisi."
        if len(name) > 31:
            return "Nama sheet maksimal 31 karakter."
        if any(ch in INVALID_SHEET_CHARS for ch in name):
            return "Nama sheet tidak boleh mengandung: [ ] : * ? / \\"
        return None

    def on_proses(self, _evt):
        if self.cb_sheet.GetSelection() == wx.NOT_FOUND or \
           self.cb_kolom.GetSelection() == wx.NOT_FOUND or \
           self.cb_value.GetSelection() == wx.NOT_FOUND:
            self.set_status("Lengkapi dulu: Sheet, Kolom, dan Value.", ok=False)
            return

        nama_baru = self.tx_nama.GetValue().strip()
        err = self._validate_sheet_name(nama_baru)
        if err:
            self.set_status(err, ok=False)
            return

        sheet_name = self.cb_sheet.GetStringSelection()
        col_idx = self.headers[self.cb_kolom.GetSelection()][0]
        target = self.cb_value.GetStringSelection()

        if nama_baru == sheet_name:
            self.set_status("Nama sheet baru tidak boleh sama dengan sheet sumber.", ok=False)
            return

        if nama_baru in self.wb.sheetnames:
            ask = wx.MessageBox(f'Sheet "{nama_baru}" sudah ada. Timpa isinya?',
                                APP_TITLE, wx.YES_NO | wx.ICON_QUESTION)
            if ask != wx.YES:
                return
            del self.wb[nama_baru]

        # kunci tombol selama proses
        self.btn_pilih.Enable(False)
        self.btn_proses.Enable(False)
        self.btn_buka.Enable(False)
        self.bar.SetValue(0)
        self.set_status("Memproses...")
        wx.GetApp().Yield()

        try:
            jumlah = self._copy_rows(sheet_name, nama_baru, col_idx, target)
            self.set_progress(95)
            self.wb.save(self.file_path)
            self.set_progress(100)
            self.set_status(f'Selesai! {jumlah} baris disalin ke sheet "{nama_baru}".')
        except PermissionError:
            self.bar.SetValue(0)
            self.set_status("Gagal menyimpan. Tutup file Excel dulu kalau sedang terbuka, lalu coba lagi.", ok=False)
        except Exception as e:
            self.bar.SetValue(0)
            self.set_status(f"Terjadi error: {e}", ok=False)
        finally:
            self.btn_pilih.Enable(True)
            self.btn_proses.Enable(True)
            self.btn_buka.Enable(True)

        # refresh workbook nilai supaya sinkron dengan file terbaru
        try:
            self.wb_vals = load_workbook(self.file_path, data_only=True)
        except Exception:
            pass

    def _copy_rows(self, src_name, dst_name, col_idx, target):
        src = self.wb[src_name]
        src_vals = self.wb_vals[src_name]
        dst = self.wb.create_sheet(dst_name)

        max_col = src.max_column

        # lebar kolom
        for c in range(1, max_col + 1):
            letter = get_column_letter(c)
            if letter in src.column_dimensions:
                dst.column_dimensions[letter].width = src.column_dimensions[letter].width

        def copy_row(src_r, dst_r):
            for c in range(1, max_col + 1):
                s = src.cell(row=src_r, column=c)
                d = dst.cell(row=dst_r, column=c)
                val = src_vals.cell(row=src_r, column=c).value
                d.value = val if val is not None else s.value
                if s.has_style:
                    d._style = copy.copy(s._style)
            h = src.row_dimensions[src_r].height
            if h:
                dst.row_dimensions[dst_r].height = h

        # header
        copy_row(HEADER_ROW, 1)

        # baris data yang cocok
        total = max(1, src_vals.max_row - HEADER_ROW)
        out_r = 2
        last_pct = -1
        for r in range(HEADER_ROW + 1, src_vals.max_row + 1):
            v = src_vals.cell(row=r, column=col_idx).value
            label = BLANK_LABEL if (v is None or str(v).strip() == "") else str(v).strip()
            if label == target:
                copy_row(r, out_r)
                out_r += 1
            pct = int(5 + 85 * (r - HEADER_ROW) / total)
            if pct != last_pct:
                last_pct = pct
                self.set_progress(pct)

        dst.freeze_panes = "A2"
        return out_r - 2

    def on_buka(self, _evt):
        if not self.file_path or not os.path.exists(self.file_path):
            self.set_status("Belum ada file untuk dibuka.", ok=False)
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(self.file_path)
            elif sys.platform == "darwin":
                os.system(f'open "{self.file_path}"')
            else:
                os.system(f'xdg-open "{self.file_path}"')
        except Exception as e:
            self.set_status(f"Gagal membuka file: {e}", ok=False)


if __name__ == "__main__":
    app = wx.App(False)
    frame = MainFrame()
    frame.Show()
    app.MainLoop()
