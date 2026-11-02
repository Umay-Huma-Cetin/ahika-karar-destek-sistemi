# -*- coding: utf-8 -*-
"""
AHİKA Karar Destek Sistemi — Masaüstü Uygulaması (v6.0)
======================================================
Dosya yapısı:
  analiz.py   -> Arayüz (bu dosya)
  guvenlik.py -> Kimlik doğrulama, yetkilendirme, denetim günlüğü, yedekleme
  motor.py    -> Yapay zekâ modeli, kayıt sorguları, PDF/Excel raporları
"""

import os
import re
import sys
import time
import math
import socket
import sqlite3
import threading
import platform
from datetime import datetime

import tkinter as tk
from tkinter import ttk, messagebox, filedialog

import guvenlik as g
import motor as m

try:
    import tkintermapview
except Exception:  # harita kütüphanesi yoksa uygulama yine çalışır
    tkintermapview = None

SURUM = "6.0"
WINDOWS = platform.system() == "Windows"
FONT = "Segoe UI" if WINDOWS else "DejaVu Sans"

# =====================================================================================
# TASARIM SİSTEMİ
# =====================================================================================
R = {
    "sidebar": "#0f172a", "sidebar_hover": "#1e293b", "sidebar_metin": "#cbd5e1", "sidebar_soluk": "#64748b",
    "zemin": "#f1f5f9", "kart": "#ffffff", "kenar": "#e2e8f0", "kenar_koyu": "#cbd5e1",
    "metin": "#0f172a", "metin2": "#334155", "soluk": "#64748b", "cok_soluk": "#94a3b8",
    "vurgu": "#2563eb", "vurgu_koyu": "#1d4ed8", "vurgu_acik": "#eff6ff", "vurgu_orta": "#dbeafe",
    "basari": "#16a34a", "basari_koyu": "#15803d", "basari_acik": "#dcfce7",
    "tehlike": "#dc2626", "tehlike_koyu": "#b91c1c", "tehlike_acik": "#fee2e2",
    "uyari": "#d97706", "uyari_koyu": "#b45309", "uyari_acik": "#fef3c7",
    "zebra": "#f8fafc", "pasif_zemin": "#e2e8f0", "pasif_metin": "#94a3b8",
}
ROL_RENK = {"Ana Yönetici": "#7c3aed", "Yönetici": "#2563eb", "Uzman": "#16a34a", "Stajyer": "#64748b"}
AYLAR = ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"]
GUNLER = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "Cumartesi", "Pazar"]


def F(boyut=10, kalin=False):
    return (FONT, boyut, "bold" if kalin else "normal")


def seviye_renk(oran):
    if oran >= 70:
        return R["basari"], R["basari_acik"]
    if oran >= 50:
        return R["uyari"], R["uyari_acik"]
    return R["tehlike"], R["tehlike_acik"]


def kaynak_yolu(ad):
    base = getattr(sys, "_MEIPASS", g.APP_DIR)
    return os.path.join(base, ad)


def ikon_ata(win):
    try:
        win.iconbitmap(kaynak_yolu("logo.ico"))
    except Exception:
        pass


def ortala(win, w, h, parent=None):
    win.update_idletasks()
    w = min(w, win.winfo_screenwidth() - 40)
    h = min(h, win.winfo_screenheight() - 80)
    if parent is not None and parent.winfo_viewable():
        x = parent.winfo_rootx() + (parent.winfo_width() - w) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - h) // 2
    else:
        x = (win.winfo_screenwidth() - w) // 2
        y = int((win.winfo_screenheight() - h) / 2.3)
    win.geometry(f"{w}x{h}+{max(0, x)}+{max(0, y)}")


def stil_kur(root):
    st = ttk.Style(root)
    try:
        st.theme_use("clam")
    except tk.TclError:
        pass
    st.configure(".", font=F(10), background=R["kart"])
    st.configure("Modern.Treeview", background=R["kart"], fieldbackground=R["kart"], foreground=R["metin2"],
                 rowheight=34, font=F(10), borderwidth=0, relief="flat")
    st.configure("Modern.Treeview.Heading", background=R["zebra"], foreground=R["soluk"], font=F(9, True),
                 relief="flat", borderwidth=0, padding=(10, 9))
    st.map("Modern.Treeview.Heading", background=[("active", R["kenar"])])
    st.map("Modern.Treeview", background=[("selected", R["vurgu_orta"])], foreground=[("selected", R["metin"])])
    st.layout("Modern.Treeview", [("Modern.Treeview.treearea", {"sticky": "nswe"})])
    st.configure("TCombobox", padding=(8, 6), arrowsize=14, foreground=R["metin"], fieldbackground="white",
                 background="white", bordercolor=R["kenar_koyu"], lightcolor="white", darkcolor="white",
                 arrowcolor=R["soluk"])
    st.map("TCombobox", fieldbackground=[("readonly", "white")], foreground=[("readonly", R["metin"])],
           bordercolor=[("focus", R["vurgu"])], selectbackground=[("readonly", "white")],
           selectforeground=[("readonly", R["metin"])])
    st.configure("Hata.TCombobox", bordercolor=R["tehlike"])
    st.configure("Vertical.TScrollbar", background=R["kenar"], troughcolor=R["kart"], bordercolor=R["kart"],
                 arrowcolor=R["soluk"], lightcolor=R["kenar"], darkcolor=R["kenar"], gripcount=0, arrowsize=12)
    st.configure("Horizontal.TScrollbar", background=R["kenar"], troughcolor=R["kart"], bordercolor=R["kart"],
                 arrowcolor=R["soluk"], lightcolor=R["kenar"], darkcolor=R["kenar"], gripcount=0, arrowsize=12)
    root.option_add("*TCombobox*Listbox.font", F(10))
    root.option_add("*TCombobox*Listbox.selectBackground", R["vurgu"])

# =====================================================================================
# TEMEL BİLEŞENLER
# =====================================================================================

class Buton(tk.Label):
    """Tam kontrollü renk/hover/pasif durumu olan modern düz buton."""
    STILLER = {
        "birincil": (R["vurgu"], R["vurgu_koyu"], "white", None),
        "basari": (R["basari"], R["basari_koyu"], "white", None),
        "tehlike": (R["tehlike"], R["tehlike_koyu"], "white", None),
        "koyu": (R["sidebar"], R["sidebar_hover"], "white", None),
        "ikincil": ("white", R["zebra"], R["metin2"], R["kenar_koyu"]),
        "tehlike_cizgi": ("white", R["tehlike_acik"], R["tehlike"], "#fca5a5"),
        "hayalet": (None, R["kenar"], R["vurgu"], None),
    }

    def __init__(self, parent, text, command=None, stil="birincil", boyut=10, padx=16, pady=8, **kw):
        bg, self._hover, fg, cerceve = self.STILLER[stil]
        self._bg = bg or parent.cget("bg")
        self._fg = fg
        self._cmd = command
        self._etkin = True
        super().__init__(parent, text=text, bg=self._bg, fg=fg, font=F(boyut, True), padx=padx, pady=pady,
                         cursor="hand2", **kw)
        if cerceve:
            self.config(highlightthickness=1, highlightbackground=cerceve, highlightcolor=cerceve)
        self.bind("<Enter>", lambda e: self._etkin and self.config(bg=self._hover))
        self.bind("<Leave>", lambda e: self._etkin and self.config(bg=self._bg))
        self.bind("<ButtonRelease-1>", self._tikla)

    def _tikla(self, e):
        if self._etkin and self._cmd and 0 <= e.x <= self.winfo_width() and 0 <= e.y <= self.winfo_height():
            self._cmd()

    def etkin(self, durum=True):
        self._etkin = durum
        if durum:
            self.config(bg=self._bg, fg=self._fg, cursor="hand2")
        else:
            self.config(bg=R["pasif_zemin"], fg=R["pasif_metin"], cursor="arrow")

    def komut(self, cmd):
        self._cmd = cmd


class Baglanti(tk.Label):
    """Metin bağlantısı (link) görünümlü buton."""

    def __init__(self, parent, text, command, renk=None, boyut=9):
        super().__init__(parent, text=text, fg=renk or R["vurgu"], bg=parent.cget("bg"), font=F(boyut, True),
                         cursor="hand2")
        self.bind("<Button-1>", lambda e: command())
        self.bind("<Enter>", lambda e: self.config(font=(FONT, boyut, "bold underline")))
        self.bind("<Leave>", lambda e: self.config(font=F(boyut, True)))


class Girdi(tk.Entry):
    """Yer tutucu (placeholder) destekli giriş kutusu."""

    def __init__(self, parent, placeholder="", show=None, width=20, **kw):
        super().__init__(parent, font=F(10), relief="flat", bd=0, bg="white", fg=R["metin"], width=width,
                         insertbackground=R["metin"], highlightthickness=1, highlightbackground=R["kenar_koyu"],
                         highlightcolor=R["vurgu"], **kw)
        self._ph = placeholder
        self._show = show or ""
        self._ph_aktif = False
        self.bind("<FocusIn>", self._odak_al, add="+")
        self.bind("<FocusOut>", self._odak_kayip, add="+")
        self.bind("<Key>", lambda e: self.hata(False), add="+")
        self._ph_goster()

    def _ph_goster(self):
        if self._ph and not super().get():
            self._ph_aktif = True
            self.config(show="", fg=R["cok_soluk"])
            self.insert(0, self._ph)

    def _odak_al(self, e=None):
        if self._ph_aktif:
            self.delete(0, tk.END)
            self._ph_aktif = False
            self.config(show=self._show, fg=R["metin"])

    def _odak_kayip(self, e=None):
        if not super().get():
            self._ph_goster()

    def get(self):
        return "" if self._ph_aktif else super().get()

    def set(self, deger):
        self._odak_al()
        self.delete(0, tk.END)
        self.insert(0, str(deger))
        if not self._odakta():
            self._odak_kayip()

    def _odakta(self):
        try:
            return self.focus_get() is self
        except (KeyError, tk.TclError):
            return False

    def temizle(self):
        self._odak_al()
        self.delete(0, tk.END)
        if not self._odakta():
            self._ph_goster()

    def goster_gizle(self, goster):
        self._show = "" if goster else "•"
        if not self._ph_aktif:
            self.config(show=self._show)

    def hata(self, durum=True):
        self.config(highlightbackground=R["tehlike"] if durum else R["kenar_koyu"],
                    highlightcolor=R["tehlike"] if durum else R["vurgu"])


class Alan(tk.Frame):
    """Etiket + giriş (ya da açılır liste) + satır içi hata mesajı."""

    def __init__(self, parent, etiket, tur="entry", placeholder="", values=(), show=None, ipucu="", width=20):
        bg = parent.cget("bg")
        super().__init__(parent, bg=bg)
        ust = tk.Frame(self, bg=bg)
        ust.pack(fill="x")
        tk.Label(ust, text=etiket, font=F(9, True), bg=bg, fg=R["metin2"]).pack(side="left")
        self.ust = ust
        if ipucu:
            tk.Label(ust, text=ipucu, font=F(8), bg=bg, fg=R["cok_soluk"]).pack(side="right")
        self.tur = tur
        if tur == "combo":
            self.w = ttk.Combobox(self, values=list(values), state="readonly", font=F(10), width=width)
            self.w.pack(fill="x", pady=(4, 0))
            self.w.bind("<<ComboboxSelected>>", lambda e: self.hata(""), add="+")
        else:
            self.w = Girdi(self, placeholder=placeholder, show=show, width=width)
            self.w.pack(fill="x", pady=(4, 0), ipady=7)
        self.hata_lbl = tk.Label(self, text="", font=F(8), bg=bg, fg=R["tehlike"], anchor="w", justify="left")

    def get(self):
        return self.w.get().strip()

    def set(self, v):
        if self.tur == "combo":
            self.w.set(v)
        else:
            self.w.set(v)

    def temizle(self):
        self.hata("")
        if self.tur == "combo":
            self.w.set("")
        else:
            self.w.temizle()

    def hata(self, mesaj):
        self.hata_lbl.config(text=mesaj)
        if mesaj:
            self.hata_lbl.pack(fill="x", pady=(2, 0))
        else:
            self.hata_lbl.pack_forget()
        if self.tur == "combo":
            self.w.configure(style="Hata.TCombobox" if mesaj else "TCombobox")
        else:
            self.w.hata(bool(mesaj))

    def odak(self):
        self.w.focus_set()


def kart(parent, padding=0, **kw):
    f = tk.Frame(parent, bg=R["kart"], highlightthickness=1, highlightbackground=R["kenar"], **kw)
    if padding:
        ic = tk.Frame(f, bg=R["kart"])
        ic.pack(fill="both", expand=True, padx=padding, pady=padding)
        return f, ic
    return f


def kart_basligi(parent, baslik, alt=None):
    bg = parent.cget("bg")
    f = tk.Frame(parent, bg=bg)
    f.pack(fill="x")
    sol = tk.Frame(f, bg=bg)
    sol.pack(side="left", fill="x", expand=True)
    tk.Label(sol, text=baslik, font=F(12, True), bg=bg, fg=R["metin"]).pack(anchor="w")
    if alt:
        tk.Label(sol, text=alt, font=F(9), bg=bg, fg=R["soluk"]).pack(anchor="w", pady=(2, 0))
    sag = tk.Frame(f, bg=bg)
    sag.pack(side="right")
    return sag


def cip(parent, metin, fg, bg, boyut=8):
    return tk.Label(parent, text=metin, font=F(boyut, True), fg=fg, bg=bg, padx=8, pady=2)


class CubukGrafik(tk.Canvas):
    """Yuvarlatılmış yatay ilerleme çubuğu."""

    def __init__(self, parent, oran=0, renk=None, yukseklik=8, zemin=None):
        super().__init__(parent, height=yukseklik, bg=parent.cget("bg"), highlightthickness=0, bd=0)
        self.oran, self.renk, self.h = oran, renk or R["vurgu"], yukseklik
        self.zemin = zemin or R["kenar"]
        self.bind("<Configure>", lambda e: self.ciz())

    def ciz(self):
        self.delete("all")
        w = self.winfo_width()
        if w < 4:
            return
        h, r = self.h, self.h / 2
        self._yuvarlak(0, 0, w, h, r, self.zemin)
        dolu = max(0.0, min(1.0, self.oran / 100.0)) * w
        if dolu >= h:
            self._yuvarlak(0, 0, dolu, h, r, self.renk)

    def _yuvarlak(self, x1, y1, x2, y2, r, renk):
        self.create_oval(x1, y1, x1 + 2 * r, y2, fill=renk, outline=renk)
        self.create_oval(x2 - 2 * r, y1, x2, y2, fill=renk, outline=renk)
        self.create_rectangle(x1 + r, y1, x2 - r, y2, fill=renk, outline=renk)


class Tablo(tk.Frame):
    """Sıralanabilir başlıklar, zebra satırlar ve boş durum mesajı olan tablo."""

    def __init__(self, parent, sutunlar, secim="extended", height=10, bos_metin="Gösterilecek kayıt yok."):
        super().__init__(parent, bg=R["kart"])
        self.sutunlar = sutunlar  # [(anahtar, başlık, genişlik, hiza)]
        anahtarlar = [s[0] for s in sutunlar]
        self.tree = ttk.Treeview(self, columns=anahtarlar, show="headings", selectmode=secim,
                                 style="Modern.Treeview", height=height)
        vsb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self._ham = {}
        self._sira = (None, False)
        for a, b, w, hz in sutunlar:
            self.tree.heading(a, text=b, anchor=hz, command=lambda k=a: self._sirala(k))
            self.tree.column(a, width=w, minwidth=50, anchor=hz, stretch=True)
        self.tree.tag_configure("z", background=R["zebra"])
        self.tree.tag_configure("vurgu", background=R["vurgu_acik"])
        self.tree.tag_configure("iyi", background=R["basari_acik"])
        self.bos = tk.Label(self, text=bos_metin, font=F(10), fg=R["cok_soluk"], bg=R["kart"])

    def doldur(self, satirlar):
        """satirlar: [(iid, gösterilen değerler, ham değerler, ek_etiket)]"""
        self.tree.delete(*self.tree.get_children())
        self._ham = {}
        for i, (iid, goster, ham, *ek) in enumerate(satirlar):
            etiket = ("z",) if i % 2 else ()
            if ek and ek[0]:
                etiket = (ek[0],)
            self.tree.insert("", "end", iid=str(iid), values=goster, tags=etiket)
            self._ham[str(iid)] = ham
        if satirlar:
            self.bos.place_forget()
        else:
            self.bos.place(relx=0.5, rely=0.5, anchor="center")
        if self._sira[0]:
            k, ters = self._sira
            self._sira = (k, not ters)
            self._sirala(k)

    def _sirala(self, anahtar):
        idx = [s[0] for s in self.sutunlar].index(anahtar)
        ters = not self._sira[1] if self._sira[0] == anahtar else False
        self._sira = (anahtar, ters)

        def deger(iid):
            v = self._ham.get(iid, ())[idx] if idx < len(self._ham.get(iid, ())) else ""
            return (0, v) if isinstance(v, (int, float)) and v is not None else (1, m.tr_kucuk(v))

        sirali = sorted(self.tree.get_children(), key=deger, reverse=ters)
        for i, iid in enumerate(sirali):
            self.tree.move(iid, "", i)
            mevcut = [t for t in self.tree.item(iid, "tags") if t != "z"]
            if not mevcut or mevcut == []:
                self.tree.item(iid, tags=("z",) if i % 2 else ())
        for a, b, _, _ in self.sutunlar:
            ok = (" ▼" if ters else " ▲") if a == anahtar else ""
            self.tree.heading(a, text=b + ok)

    def secili(self):
        return list(self.tree.selection())


class KaydirilabilirAlan(tk.Frame):
    def __init__(self, parent, bg):
        super().__init__(parent, bg=bg)
        self.canvas = tk.Canvas(self, bg=bg, highlightthickness=0, bd=0)
        self.vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.ic = tk.Frame(self.canvas, bg=bg)
        self._win = self.canvas.create_window((0, 0), window=self.ic, anchor="nw")
        self.canvas.configure(yscrollcommand=self.vsb.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.vsb.pack(side="right", fill="y")
        self.ic.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfig(self._win, width=e.width))
        self.canvas.bind("<Enter>", self._bagla)
        self.canvas.bind("<Leave>", self._coz)

    def _bagla(self, e):
        self.canvas.bind_all("<MouseWheel>", self._tekerlek)
        self.canvas.bind_all("<Button-4>", lambda e: self.canvas.yview_scroll(-2, "units"))
        self.canvas.bind_all("<Button-5>", lambda e: self.canvas.yview_scroll(2, "units"))

    def _coz(self, e):
        self.canvas.unbind_all("<MouseWheel>")
        self.canvas.unbind_all("<Button-4>")
        self.canvas.unbind_all("<Button-5>")

    def _tekerlek(self, e):
        self.canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")

    def basa(self):
        self.canvas.yview_moveto(0)


class Diyalog(tk.Toplevel):
    def __init__(self, parent, baslik, w, h, alt_baslik=None, boyutlanabilir=False):
        super().__init__(parent)
        self.withdraw()
        self.title(f"{baslik} — AHİKA")
        self.configure(bg=R["kart"])
        ikon_ata(self)
        self.transient(parent.winfo_toplevel())
        self.resizable(boyutlanabilir, boyutlanabilir)
        ust = tk.Frame(self, bg=R["kart"])
        ust.pack(fill="x", padx=28, pady=(24, 0))
        tk.Label(ust, text=baslik, font=F(14, True), bg=R["kart"], fg=R["metin"]).pack(anchor="w")
        if alt_baslik:
            tk.Label(ust, text=alt_baslik, font=F(9), bg=R["kart"], fg=R["soluk"], justify="left",
                     wraplength=w - 60).pack(anchor="w", pady=(4, 0))
        self.alt = tk.Frame(self, bg=R["zebra"])
        self.alt.pack(fill="x", side="bottom")
        tk.Frame(self, bg=R["kenar"], height=1).pack(fill="x", side="bottom")
        self.alt_ic = tk.Frame(self.alt, bg=R["zebra"])
        self.alt_ic.pack(fill="x", padx=28, pady=14)
        self.govde = tk.Frame(self, bg=R["kart"])
        self.govde.pack(fill="both", expand=True, padx=28, pady=(16, 12))
        ortala(self, w, h, parent.winfo_toplevel())
        self.deiconify()
        self.bind("<Escape>", lambda e: self.destroy())
        self.after(10, self._modal)

    def sigdir(self, genislik=None):
        """İçeriğe göre yüksekliği ayarlar (gereksiz boşluk bırakmaz)."""
        self.update_idletasks()
        w = genislik or self.winfo_width()
        ortala(self, w, self.winfo_reqheight(), self.master.winfo_toplevel())

    def _modal(self):
        try:
            self.grab_set()
            self.focus_set()
        except tk.TclError:
            pass


class SifreGucCubugu(tk.Frame):
    def __init__(self, parent):
        bg = parent.cget("bg")
        super().__init__(parent, bg=bg)
        self.cubuklar = []
        satir = tk.Frame(self, bg=bg)
        satir.pack(fill="x")
        for i in range(4):
            c = tk.Frame(satir, bg=R["kenar"], height=4)
            c.pack(side="left", fill="x", expand=True, padx=(0 if i == 0 else 3, 0))
            self.cubuklar.append(c)
        self.lbl = tk.Label(self, text=f"En az {g.MIN_SIFRE_UZUNLUK} karakter, harf ve rakam içermeli.",
                            font=F(8), bg=bg, fg=R["soluk"], anchor="w")
        self.lbl.pack(fill="x", pady=(3, 0))

    def guncelle(self, sifre):
        puan, etiket = g.sifre_gucu(sifre)
        renkler = [R["tehlike"], R["tehlike"], R["uyari"], R["basari"], R["basari"]]
        for i, c in enumerate(self.cubuklar):
            c.config(bg=renkler[puan] if sifre and i < max(1, puan) else R["kenar"])
        if sifre:
            self.lbl.config(text=f"Şifre gücü: {etiket}", fg=renkler[puan])
        else:
            self.lbl.config(text=f"En az {g.MIN_SIFRE_UZUNLUK} karakter, harf ve rakam içermeli.", fg=R["soluk"])


def sifre_alani(parent, etiket, guc=False):
    a = Alan(parent, etiket, show="•")
    goster = {"v": False}

    def degistir():
        goster["v"] = not goster["v"]
        a.w.goster_gizle(goster["v"])
        lnk.config(text="Gizle" if goster["v"] else "Göster")

    lnk = Baglanti(a.ust, "Göster", degistir, renk=R["soluk"], boyut=8)
    lnk.pack(side="right")
    if guc:
        cubuk = SifreGucCubugu(a)
        cubuk.pack(fill="x", pady=(5, 0))
        a.w.bind("<KeyRelease>", lambda e: cubuk.guncelle(a.w.get()), add="+")
    return a


def kopyalanabilir_kutu(parent, metin, boyut=16):
    e = tk.Entry(parent, font=("Consolas" if WINDOWS else "DejaVu Sans Mono", boyut, "bold"), justify="center",
                 relief="flat", bg=R["vurgu_acik"], fg=R["vurgu_koyu"], readonlybackground=R["vurgu_acik"],
                 highlightthickness=1, highlightbackground=R["vurgu_orta"])
    e.insert(0, metin)
    e.config(state="readonly")
    return e

# =====================================================================================
# UYGULAMA
# =====================================================================================

class Uygulama:
    def __init__(self):
        self.root = tk.Tk()
        self.root.withdraw()
        self.root.title("AHİKA Karar Destek Sistemi")
        self.root.configure(bg=R["kart"])
        ikon_ata(self.root)
        stil_kur(self.root)
        self.root.protocol("WM_DELETE_WINDOW", self.cikis)

        self.kullanici = None
        self.son_hareket = time.time()
        self._oturum_after = None
        self._bildirimler = []
        self.sayfalar = {}
        self.nav = {}
        self.aktif_sayfa = None

        self.motor = m.KararMotoru()
        self.model_hatasi = None
        for olay in ("<Any-KeyPress>", "<Any-ButtonPress>", "<Motion>"):
            self.root.bind_all(olay, self._hareket, add="+")
        self._baslat()
        self.root.deiconify()

    def _baslat(self):
        try:
            g.veritabani_hazirla()
        except g.AnahtarEksik as e:
            self.anahtar_kurtarma_ekrani(str(e))
            return
        except Exception as e:
            messagebox.showerror("Veritabanı hatası", f"Veritabanı açılamadı:\n{e}")
            self.root.destroy()
            sys.exit(1)
        if not self.motor.hazir:
            self.model_yukle(g.ayar_al("veri_seti_yolu") or os.path.join(g.APP_DIR, "sentetik.xlsx"))
        self.giris_ekrani()

    def anahtar_kurtarma_ekrani(self, mesaj):
        self._temizle()
        self.root.resizable(False, False)
        ortala(self.root, 620, 400)
        f = tk.Frame(self.root, bg=R["kart"])
        f.pack(fill="both", expand=True, padx=40, pady=36)
        tk.Label(f, text="Veri şifreleme anahtarı gerekli", font=F(17, True), bg=R["kart"], fg=R["metin"]).pack(anchor="w")
        tk.Label(f, text=mesaj, font=F(10), bg=R["kart"], fg=R["soluk"], wraplength=530, justify="left").pack(anchor="w", pady=(8, 6))
        tk.Label(f, text="Kurtarma anahtarı, Ana Yönetici'nin daha önce kaydettiği 44 karakterlik metindir. "
                         "Anahtar olmadan başvuru sahiplerinin kişisel verileri açılamaz.",
                 font=F(9), bg=R["kart"], fg=R["soluk"], wraplength=530, justify="left").pack(anchor="w", pady=(0, 16))
        banner = self._banner(f)
        a = Alan(f, "Kurtarma anahtarı", placeholder="Kurtarma anahtarını yapıştırın")
        a.pack(fill="x")

        def yukle():
            try:
                g.kurtarma_anahtari_geri_yukle(a.get())
            except g.GuvenlikHatasi as e:
                self._banner_goster(banner, str(e), "hata", once=a)
                return
            self._baslat()
            self.bildirim("Anahtar geri yüklendi; veriler yeniden erişilebilir.")

        bt = tk.Frame(f, bg=R["kart"])
        bt.pack(fill="x", pady=(18, 0))
        Buton(bt, "Anahtarı geri yükle", yukle).pack(side="left")
        Buton(bt, "Programı kapat", self.root.destroy, stil="ikincil").pack(side="left", padx=8)
        a.w.bind("<Return>", lambda e: yukle())

    def kurtarma_anahtari_diyalogu(self, ilk=False):
        try:
            anahtar = g.kurtarma_anahtari()
        except g.GuvenlikHatasi as e:
            messagebox.showerror("Anahtar okunamadı", str(e), parent=self.root)
            return
        d = Diyalog(self.root, "Kurtarma anahtarını saklayın" if ilk else "Veri kurtarma anahtarı", 560, 420,
                    ("Başvuru sahiplerinin ad, telefon ve e-posta bilgileri artık şifreli saklanıyor. " if ilk else "") +
                    "Bilgisayar değişirse, Windows yeniden kurulursa veya veri_anahtari.bin silinirse verileri açmanın "
                    "TEK yolu bu anahtardır. Bir USB belleğe ya da kasaya kaydedin; veritabanıyla aynı klasörde tutmayın.")
        kopyalanabilir_kutu(d.govde, anahtar, 11).pack(fill="x", ipady=12)
        tk.Label(d.govde, text="Bu anahtarı gören kişi, veritabanı dosyasını da ele geçirirse kişisel verileri okuyabilir.",
                 font=F(8), bg=R["kart"], fg=R["uyari_koyu"], wraplength=500, justify="left").pack(anchor="w", pady=(8, 0))

        def dosyaya():
            yol = filedialog.asksaveasfilename(parent=d, defaultextension=".txt", initialfile="AHIKA_kurtarma_anahtari.txt",
                                               filetypes=[("Metin dosyası", "*.txt")], title="Kurtarma anahtarını kaydet")
            if yol:
                with open(yol, "w", encoding="utf-8") as fh:
                    fh.write("AHİKA Karar Destek Sistemi — Veri kurtarma anahtarı\n"
                             f"Oluşturma: {datetime.now():%d.%m.%Y %H:%M}\n\n{anahtar}\n")
                self.bildirim("Kurtarma anahtarı dosyaya kaydedildi.")

        def tamam():
            g.ayar_yaz("kurtarma_anahtari_onaylandi", "1")
            g.log_ekle(self.kadi, "ANAHTAR", "Kurtarma anahtarının saklandığı onaylandı.")
            d.destroy()

        def kopyala():
            self.root.clipboard_clear()
            self.root.clipboard_append(anahtar)
            self.bildirim("Kurtarma anahtarı panoya kopyalandı.", "bilgi")

        Buton(d.alt_ic, "Kaydettim", tamam, stil="basari").pack(side="right")
        Buton(d.alt_ic, "Dosyaya kaydet", dosyaya, stil="ikincil").pack(side="right", padx=8)
        Buton(d.alt_ic, "Kopyala", kopyala, stil="ikincil").pack(side="right")
        d.sigdir()
        g.log_ekle(self.kadi, "ANAHTAR", "Kurtarma anahtarı görüntülendi.")

    # ---------- genel ----------
    def model_yukle(self, yol):
        if not os.path.exists(yol):
            yol = os.path.join(g.APP_DIR, "sentetik.xlsx")
        try:
            self.motor.egit(yol)
            self.model_hatasi = None
        except Exception as e:
            self.model_hatasi = str(e)

    def yetki(self, islem):
        return self.kullanici is not None and g.yetki_var(self.kullanici["rol"], islem)

    @property
    def kadi(self):
        return self.kullanici["kullanici_adi"] if self.kullanici else "-"

    def bildirim(self, mesaj, tur="basari"):
        renk = {"basari": (R["basari"], "✓"), "hata": (R["tehlike"], "✕"), "bilgi": (R["vurgu"], "i"),
                "uyari": (R["uyari"], "!")}[tur]
        f = tk.Frame(self.root, bg=R["sidebar"], highlightthickness=0)
        tk.Frame(f, bg=renk[0], width=4).pack(side="left", fill="y")
        tk.Label(f, text=renk[1], font=F(11, True), fg=renk[0], bg=R["sidebar"]).pack(side="left", padx=(12, 6), pady=12)
        tk.Label(f, text=mesaj, font=F(10), fg="white", bg=R["sidebar"], wraplength=340, justify="left").pack(
            side="left", padx=(0, 18), pady=12)
        self._bildirimler.append(f)
        self._bildirim_diz()
        f.lift()

        def kapat():
            if f in self._bildirimler:
                self._bildirimler.remove(f)
            f.destroy()
            self._bildirim_diz()
        self.root.after(3800, kapat)

    def _bildirim_diz(self):
        y = 20
        for b in self._bildirimler:
            b.place(relx=1.0, x=-24, y=y, anchor="ne")
            b.update_idletasks()
            y += b.winfo_reqheight() + 10

    def onayla(self, baslik, mesaj, tehlikeli=False):
        return messagebox.askyesno(baslik, mesaj, icon="warning" if tehlikeli else "question", parent=self.root)

    def _temizle(self):
        for w in self.root.winfo_children():
            if isinstance(w, tk.Toplevel):
                w.destroy()
        for w in self.root.winfo_children():
            w.destroy()
        self._bildirimler = []
        self.sayfalar, self.nav = {}, {}

    def cikis(self):
        if self.kullanici:
            g.log_ekle(self.kadi, "ÇIKIŞ", "Uygulama kapatıldı.")
        self.root.destroy()

    # ---------- oturum zaman aşımı ----------
    def _hareket(self, e=None):
        self.son_hareket = time.time()

    def _oturum_kontrol(self):
        if not self.kullanici:
            return
        if time.time() - self.son_hareket > g.OTURUM_ZAMAN_ASIMI_DK * 60:
            g.log_ekle(self.kadi, "OTURUM ZAMAN AŞIMI", f"{g.OTURUM_ZAMAN_ASIMI_DK} dk hareketsizlik sonrası oturum kapatıldı.")
            self.kullanici = None
            self.giris_ekrani(bilgi=f"Güvenliğiniz için {g.OTURUM_ZAMAN_ASIMI_DK} dakika işlem yapılmadığında oturum otomatik olarak kapatılır. Lütfen yeniden giriş yapın.")
            return
        self._oturum_after = self.root.after(20000, self._oturum_kontrol)

    def oturumu_kapat(self, sor=True):
        if sor and not self.onayla("Oturumu kapat", "Oturumu kapatmak istediğinize emin misiniz?"):
            return
        g.log_ekle(self.kadi, "ÇIKIŞ", "Kullanıcı oturumu kapattı.")
        self.kullanici = None
        if self._oturum_after:
            self.root.after_cancel(self._oturum_after)
        self.giris_ekrani()

    # =================================================================================
    # GİRİŞ EKRANI
    # =================================================================================
    def giris_ekrani(self, bilgi=None):
        self._temizle()
        self.root.state("normal")
        self.root.minsize(900, 560)
        self.root.resizable(False, False)
        ortala(self.root, 980, 600)

        sol = tk.Frame(self.root, bg=R["sidebar"], width=420)
        sol.pack(side="left", fill="y")
        sol.pack_propagate(False)
        ic = tk.Frame(sol, bg=R["sidebar"])
        ic.pack(fill="both", expand=True, padx=44, pady=44)

        logo = tk.Canvas(ic, width=52, height=52, bg=R["sidebar"], highlightthickness=0)
        logo.create_oval(2, 2, 50, 50, fill=R["vurgu"], outline="")
        logo.create_text(26, 27, text="A", fill="white", font=F(20, True))
        logo.pack(anchor="w")
        tk.Label(ic, text="AHİKA", font=F(24, True), fg="white", bg=R["sidebar"]).pack(anchor="w", pady=(18, 0))
        tk.Label(ic, text="Karar Destek Sistemi", font=F(13), fg=R["sidebar_metin"], bg=R["sidebar"]).pack(anchor="w")
        tk.Label(ic, text="Hibe başvurularını yapay zekâ ile değerlendirin, bölgesel etkiyi ölçün ve "
                          "kararlarınızı güvenle kayıt altına alın.",
                 font=F(10), fg=R["sidebar_soluk"], bg=R["sidebar"], wraplength=320, justify="left").pack(anchor="w", pady=(18, 26))
        for madde in ("Sektör bazlı başarı olasılığı tahmini", "Bütçe duyarlılık simülasyonu",
                      "İlçe bazlı bölgesel etki haritası", "Rol tabanlı erişim ve denetim kaydı"):
            satir = tk.Frame(ic, bg=R["sidebar"])
            satir.pack(anchor="w", pady=5)
            tk.Label(satir, text="✓", font=F(10, True), fg="#60a5fa", bg=R["sidebar"]).pack(side="left")
            tk.Label(satir, text=madde, font=F(10), fg=R["sidebar_metin"], bg=R["sidebar"]).pack(side="left", padx=10)
        tk.Label(ic, text=f"Ahiler Kalkınma Ajansı  •  Sürüm {SURUM}", font=F(8), fg=R["sidebar_soluk"],
                 bg=R["sidebar"]).pack(side="bottom", anchor="w")

        sag = tk.Frame(self.root, bg=R["kart"])
        sag.pack(side="left", fill="both", expand=True)
        form = tk.Frame(sag, bg=R["kart"], width=360)
        form.place(relx=0.5, rely=0.5, anchor="center")

        if g.ilk_kurulum_gerekli():
            self._kurulum_formu(form)
        else:
            self._giris_formu(form, bilgi)

    def _banner(self, parent):
        b = tk.Label(parent, text="", font=F(9), wraplength=340, justify="left", anchor="w", padx=12, pady=10)
        return b

    def _banner_goster(self, b, mesaj, tur="hata", once=None):
        renk = {"hata": (R["tehlike_acik"], R["tehlike_koyu"]), "bilgi": (R["vurgu_acik"], R["vurgu_koyu"]),
                "uyari": (R["uyari_acik"], R["uyari_koyu"])}[tur]
        b.config(text=mesaj, bg=renk[0], fg=renk[1])
        if once is not None:
            b.pack(fill="x", pady=(0, 14), before=once)
        else:
            b.pack(fill="x", pady=(0, 14))

    def _giris_formu(self, form, bilgi):
        tk.Label(form, text="Hoş geldiniz", font=F(20, True), bg=R["kart"], fg=R["metin"]).pack(anchor="w")
        tk.Label(form, text="Devam etmek için hesabınıza giriş yapın.", font=F(10), bg=R["kart"],
                 fg=R["soluk"]).pack(anchor="w", pady=(4, 22))
        govde = tk.Frame(form, bg=R["kart"])
        govde.pack(fill="x")
        banner = self._banner(form)
        if bilgi:
            self._banner_goster(banner, bilgi, "bilgi", once=govde)

        a_kadi = Alan(govde, "Kullanıcı adı", placeholder="Kullanıcı adınızı girin", width=34)
        a_kadi.pack(fill="x", pady=(0, 6))
        a_sifre = sifre_alani(govde, "Şifre")
        a_sifre.pack(fill="x")
        caps = tk.Label(govde, text="", font=F(8, True), fg=R["uyari_koyu"], bg=R["kart"], anchor="w")
        caps.pack(fill="x")

        def caps_kontrol(e):
            caps.config(text="⚠  Caps Lock açık" if (e.state & 0x2) else "")
        a_sifre.w.bind("<KeyRelease>", caps_kontrol, add="+")

        btn = Buton(govde, "Giriş Yap", stil="birincil", boyut=11, pady=11)
        btn.pack(fill="x", pady=(12, 0))

        def giris():
            kadi, sifre = a_kadi.get(), a_sifre.w.get()
            a_kadi.hata("")
            a_sifre.hata("")
            if not kadi:
                a_kadi.hata("Kullanıcı adı gerekli.")
                a_kadi.odak()
                return
            if not sifre:
                a_sifre.hata("Şifre gerekli.")
                a_sifre.odak()
                return
            btn.config(text="Doğrulanıyor…")
            btn.etkin(False)
            self.root.update_idletasks()
            try:
                self.kullanici = g.giris_yap(kadi, sifre)
            except g.GuvenlikHatasi as e:
                btn.config(text="Giriş Yap")
                btn.etkin(True)
                self._banner_goster(banner, str(e), "hata", once=govde)
                a_sifre.w.temizle()
                a_sifre.odak()
                return
            if self.kullanici["sifre_degismeli"]:
                self.sifre_degistir_diyalogu(zorunlu=True, eski_sifre=sifre)
            else:
                self.ana_kabuk()

        btn.komut(giris)
        a_kadi.w.bind("<Return>", lambda e: a_sifre.odak())
        a_sifre.w.bind("<Return>", lambda e: giris())

        alt = tk.Frame(form, bg=R["kart"])
        alt.pack(fill="x", pady=(22, 0))
        tk.Label(alt, text="Davet kodunuz var mı?", font=F(9), bg=R["kart"], fg=R["soluk"]).pack(side="left")
        Baglanti(alt, "Hesap oluşturun", self.kayit_diyalogu).pack(side="left", padx=4)
        a_kadi.after(50, a_kadi.odak)

    def _kurulum_formu(self, form):
        tk.Label(form, text="İlk kurulum", font=F(20, True), bg=R["kart"], fg=R["metin"]).pack(anchor="w")
        tk.Label(form, text="Sistemi yönetecek Ana Yönetici hesabını oluşturun.\nBu hesap kullanıcıları ve davet kodlarını yönetir.",
                 font=F(10), bg=R["kart"], fg=R["soluk"], justify="left").pack(anchor="w", pady=(4, 20))
        banner = self._banner(form)
        govde = tk.Frame(form, bg=R["kart"])
        govde.pack(fill="x")
        a_kadi = Alan(govde, "Kullanıcı adı", placeholder="ör. ahika.yonetici", width=34)
        a_kadi.pack(fill="x", pady=(0, 6))
        a_s1 = sifre_alani(govde, "Şifre", guc=True)
        a_s1.pack(fill="x", pady=(0, 6))
        a_s2 = sifre_alani(govde, "Şifre (tekrar)")
        a_s2.pack(fill="x")

        def olustur():
            if a_s1.w.get() != a_s2.w.get():
                a_s2.hata("Şifreler eşleşmiyor.")
                return
            try:
                g.ilk_yonetici_olustur(a_kadi.get(), a_s1.w.get())
                self.kullanici = g.giris_yap(a_kadi.get(), a_s1.w.get())
            except g.GuvenlikHatasi as e:
                self._banner_goster(banner, str(e), "hata", once=govde)
                return
            self.ana_kabuk()
            self.bildirim("Kurulum tamamlandı. Hoş geldiniz!")

        Buton(govde, "Yönetici hesabını oluştur", olustur, boyut=11, pady=11).pack(fill="x", pady=(14, 0))

    def kayit_diyalogu(self):
        d = Diyalog(self.root, "Hesap oluştur", 460, 640,
                    "Yöneticinizden aldığınız tek kullanımlık davet kodu ile hesabınızı oluşturun. "
                    "Rolünüz davet koduna göre otomatik atanır.")
        banner = self._banner(d.govde)
        f = tk.Frame(d.govde, bg=R["kart"])
        f.pack(fill="x")
        a_kod = Alan(f, "Davet kodu", placeholder="AHK-XXXX-XXXX-XXXX")
        a_kod.pack(fill="x", pady=(0, 6))
        a_kadi = Alan(f, "Kullanıcı adı", placeholder="3-32 karakter", ipucu="harf, rakam, . _ -")
        a_kadi.pack(fill="x", pady=(0, 6))
        a_s1 = sifre_alani(f, "Şifre", guc=True)
        a_s1.pack(fill="x", pady=(0, 6))
        a_s2 = sifre_alani(f, "Şifre (tekrar)")
        a_s2.pack(fill="x")

        def kaydet():
            if not a_kod.get():
                a_kod.hata("Davet kodu gerekli.")
                return
            if a_s1.w.get() != a_s2.w.get():
                a_s2.hata("Şifreler eşleşmiyor.")
                return
            try:
                rol = g.kayit_ol(a_kadi.get(), a_s1.w.get(), a_kod.get())
            except g.GuvenlikHatasi as e:
                self._banner_goster(banner, str(e), "hata", once=f)
                return
            d.destroy()
            self.bildirim(f"Hesabınız oluşturuldu ({rol}). Şimdi giriş yapabilirsiniz.")

        Buton(d.alt_ic, "Hesap oluştur", kaydet).pack(side="right")
        Buton(d.alt_ic, "Vazgeç", d.destroy, stil="ikincil").pack(side="right", padx=8)
        d.sigdir(460)
        a_kod.after(60, a_kod.odak)

    def sifre_degistir_diyalogu(self, zorunlu=False, eski_sifre=None):
        if zorunlu:
            alt = ("Güvenlik güncellemesi nedeniyle hesabınızın şifresini yenilemeniz gerekiyor. "
                   "Yeni şifreniz en az 8 karakter olmalı, harf ve rakam içermelidir.")
        else:
            alt = "Hesabınızın güvenliği için tahmin edilmesi zor bir şifre seçin."
        d = Diyalog(self.root, "Şifrenizi yenileyin" if zorunlu else "Şifre değiştir", 460, 580 if not zorunlu else 560, alt)
        banner = self._banner(d.govde)
        f = tk.Frame(d.govde, bg=R["kart"])
        f.pack(fill="x")
        a_eski = None
        if not zorunlu:
            a_eski = sifre_alani(f, "Mevcut şifre")
            a_eski.pack(fill="x", pady=(0, 6))
        a_s1 = sifre_alani(f, "Yeni şifre", guc=True)
        a_s1.pack(fill="x", pady=(0, 6))
        a_s2 = sifre_alani(f, "Yeni şifre (tekrar)")
        a_s2.pack(fill="x")

        def kaydet():
            if a_s1.w.get() != a_s2.w.get():
                a_s2.hata("Şifreler eşleşmiyor.")
                return
            eski = eski_sifre if zorunlu else a_eski.w.get()
            try:
                g.sifre_degistir(self.kadi, eski, a_s1.w.get())
            except g.GuvenlikHatasi as e:
                self._banner_goster(banner, str(e), "hata", once=f)
                return
            d.destroy()
            if zorunlu:
                self.kullanici["sifre_degismeli"] = False
                self.ana_kabuk()
            self.bildirim("Şifreniz başarıyla güncellendi.")

        def vazgec():
            d.destroy()
            if zorunlu:
                g.log_ekle(self.kadi, "ÇIKIŞ", "Zorunlu şifre yenileme ertelendi, oturum kapatıldı.")
                self.kullanici = None
                self.giris_ekrani()

        d.protocol("WM_DELETE_WINDOW", vazgec)
        d.bind("<Escape>", lambda e: vazgec())
        Buton(d.alt_ic, "Şifreyi güncelle", kaydet).pack(side="right")
        Buton(d.alt_ic, "Çıkış yap" if zorunlu else "Vazgeç", vazgec, stil="ikincil").pack(side="right", padx=8)
        d.sigdir(460)
        (a_eski or a_s1).after(60, (a_eski or a_s1).odak)

    # =================================================================================
    # ANA KABUK (yan menü + üst bar + içerik + durum çubuğu)
    # =================================================================================
    def ana_kabuk(self):
        self._temizle()
        self.root.resizable(True, True)
        self.root.minsize(1180, 700)
        ortala(self.root, 1400, 840)
        if WINDOWS:
            try:
                self.root.state("zoomed")
            except tk.TclError:
                pass
        self.son_hareket = time.time()
        self._oturum_after = self.root.after(20000, self._oturum_kontrol)

        # --- yan menü ---
        sb = tk.Frame(self.root, bg=R["sidebar"], width=250)
        sb.pack(side="left", fill="y")
        sb.pack_propagate(False)
        marka = tk.Frame(sb, bg=R["sidebar"])
        marka.pack(fill="x", padx=20, pady=(22, 18))
        logo = tk.Canvas(marka, width=36, height=36, bg=R["sidebar"], highlightthickness=0)
        logo.create_oval(1, 1, 35, 35, fill=R["vurgu"], outline="")
        logo.create_text(18, 19, text="A", fill="white", font=F(14, True))
        logo.pack(side="left")
        mt = tk.Frame(marka, bg=R["sidebar"])
        mt.pack(side="left", padx=10)
        tk.Label(mt, text="AHİKA", font=F(13, True), fg="white", bg=R["sidebar"]).pack(anchor="w")
        tk.Label(mt, text="Karar Destek Sistemi", font=F(8), fg=R["sidebar_soluk"], bg=R["sidebar"]).pack(anchor="w")

        menu = [
            ("GENEL", None, None, None),
            ("panel", "▦", "Gösterge Paneli", "analiz"),
            ("analiz", "✚", "Yeni Analiz", "analiz"),
            ("KAYITLAR", None, None, None),
            ("gecmis", "◷", "Analiz Geçmişi", "analiz"),
            ("Onaylandi", "✓", "Onaylanan Hibeler", "analiz"),
            ("Reddedildi", "✕", "Reddedilen Hibeler", "analiz"),
            ("bolge", "◉", "Bölge Analizi", "analiz"),
            ("YÖNETİM", None, None, "kullanici_yonet|log_gor"),
            ("kullanicilar", "☰", "Kullanıcılar", "kullanici_yonet"),
            ("denetim", "⚑", "Denetim Kayıtları", "log_gor"),
        ]
        for anahtar, ikon, metin, izin in menu:
            if izin and not any(self.yetki(i) for i in izin.split("|")):
                continue
            if ikon is None:
                tk.Label(sb, text=anahtar, font=F(8, True), fg=R["sidebar_soluk"], bg=R["sidebar"]).pack(
                    anchor="w", padx=24, pady=(16, 6))
                continue
            self._nav_ogesi(sb, anahtar, ikon, metin)

        # kullanıcı kartı
        alt = tk.Frame(sb, bg=R["sidebar"])
        alt.pack(side="bottom", fill="x", padx=16, pady=16)
        tk.Frame(sb, bg=R["sidebar_hover"], height=1).pack(side="bottom", fill="x")
        uk = tk.Frame(alt, bg=R["sidebar"])
        uk.pack(fill="x")
        av = tk.Canvas(uk, width=38, height=38, bg=R["sidebar"], highlightthickness=0)
        av.create_oval(1, 1, 37, 37, fill=ROL_RENK.get(self.kullanici["rol"], R["vurgu"]), outline="")
        av.create_text(19, 20, text=self.kadi[:1].upper(), fill="white", font=F(13, True))
        av.pack(side="left")
        ut = tk.Frame(uk, bg=R["sidebar"])
        ut.pack(side="left", padx=10)
        tk.Label(ut, text=self.kadi, font=F(10, True), fg="white", bg=R["sidebar"]).pack(anchor="w")
        tk.Label(ut, text=self.kullanici["rol"], font=F(8), fg=R["sidebar_soluk"], bg=R["sidebar"]).pack(anchor="w")
        bt = tk.Frame(alt, bg=R["sidebar"])
        bt.pack(fill="x", pady=(12, 0))
        for metin, cmd in (("Şifre değiştir", self.sifre_degistir_diyalogu), ("Çıkış yap", self.oturumu_kapat)):
            l = tk.Label(bt, text=metin, font=F(9), fg=R["sidebar_metin"], bg=R["sidebar_hover"], padx=10, pady=6,
                         cursor="hand2")
            l.pack(side="left", expand=True, fill="x", padx=(0, 6) if metin.startswith("Ş") else 0)
            l.bind("<Button-1>", lambda e, c=cmd: c())
            l.bind("<Enter>", lambda e, w=l: w.config(bg="#334155"))
            l.bind("<Leave>", lambda e, w=l: w.config(bg=R["sidebar_hover"]))

        # --- sağ taraf ---
        sag = tk.Frame(self.root, bg=R["zemin"])
        sag.pack(side="left", fill="both", expand=True)

        ust = tk.Frame(sag, bg=R["kart"], height=74)
        ust.pack(fill="x")
        ust.pack_propagate(False)
        tk.Frame(sag, bg=R["kenar"], height=1).pack(fill="x")
        bas = tk.Frame(ust, bg=R["kart"])
        bas.pack(side="left", padx=28)
        self.lbl_baslik = tk.Label(bas, text="", font=F(16, True), fg=R["metin"], bg=R["kart"])
        self.lbl_baslik.pack(anchor="w", pady=(14, 0))
        self.lbl_alt = tk.Label(bas, text="", font=F(9), fg=R["soluk"], bg=R["kart"])
        self.lbl_alt.pack(anchor="w")
        self.ust_aksiyon = tk.Frame(ust, bg=R["kart"])
        self.ust_aksiyon.pack(side="right", padx=24)
        if self.yetki("yedekle"):
            Buton(self.ust_aksiyon, "⤓  Yedek al", self.yedek_al, stil="ikincil", boyut=9).pack(side="right")

        # durum çubuğu
        durum = tk.Frame(sag, bg=R["kart"], height=30)
        durum.pack(side="bottom", fill="x")
        durum.pack_propagate(False)
        tk.Frame(sag, bg=R["kenar"], height=1).pack(side="bottom", fill="x")
        self.lbl_model = tk.Label(durum, text="", font=F(8), bg=R["kart"], fg=R["soluk"])
        self.lbl_model.pack(side="left", padx=20)
        tk.Label(durum, text=f"Oturum {g.OTURUM_ZAMAN_ASIMI_DK} dk hareketsizlikte kapanır   •   Sürüm {SURUM}",
                 font=F(8), bg=R["kart"], fg=R["cok_soluk"]).pack(side="right", padx=20)
        self.durum_guncelle()

        self.icerik = tk.Frame(sag, bg=R["zemin"])
        self.icerik.pack(fill="both", expand=True)

        self.root.bind("<Control-Return>", lambda e: self.aktif_sayfa == "analiz" and self.sayfalar["analiz"].analiz_et())
        self.sayfa_goster("panel")
        if self.kullanici.get("son_giris"):
            self.bildirim(f"Son girişiniz: {m.tarih_tr(self.kullanici['son_giris'])}", "bilgi")
        if self.kullanici["rol"] == "Ana Yönetici" and g.ayar_al("kurtarma_anahtari_onaylandi") == "0":
            self.root.after(600, lambda: self.kurtarma_anahtari_diyalogu(ilk=True))

    def durum_guncelle(self):
        if self.motor.hazir:
            ad = os.path.basename(self.motor.veri_yolu)
            self.lbl_model.config(text=f"●  Model hazır   •   Veri seti: {ad}   •   {self.motor.kayit_sayisi:,} kayıt".replace(",", ".")
                                  + f"   •   Test doğruluğu {m.yuzde(self.motor.dogruluk)}", fg=R["basari_koyu"])
        else:
            self.lbl_model.config(text="●  Model yüklenemedi — veri seti yükleyin", fg=R["tehlike"])

    def _nav_ogesi(self, sb, anahtar, ikon, metin):
        f = tk.Frame(sb, bg=R["sidebar"], cursor="hand2")
        f.pack(fill="x", padx=12, pady=1)
        gosterge = tk.Frame(f, bg=R["sidebar"], width=3)
        gosterge.pack(side="left", fill="y")
        li = tk.Label(f, text=ikon, font=F(11), fg=R["sidebar_soluk"], bg=R["sidebar"], width=2)
        li.pack(side="left", padx=(10, 6), pady=9)
        lm = tk.Label(f, text=metin, font=F(10), fg=R["sidebar_metin"], bg=R["sidebar"])
        lm.pack(side="left")
        parcalar = (f, li, lm)

        def boya(bg):
            for w in parcalar:
                w.config(bg=bg)

        for w in parcalar:
            w.bind("<Button-1>", lambda e: self.sayfa_goster(anahtar))
            w.bind("<Enter>", lambda e: self.aktif_sayfa != anahtar and boya(R["sidebar_hover"]))
            w.bind("<Leave>", lambda e: self.aktif_sayfa != anahtar and boya(R["sidebar"]))
        self.nav[anahtar] = (parcalar, gosterge, li, lm, boya)

    def sayfa_goster(self, anahtar):
        for k, (parcalar, gosterge, li, lm, boya) in self.nav.items():
            aktif = k == anahtar
            boya(R["sidebar_hover"] if aktif else R["sidebar"])
            gosterge.config(bg=R["vurgu"] if aktif else R["sidebar"])
            lm.config(fg="white" if aktif else R["sidebar_metin"], font=F(10, aktif))
            li.config(fg="#60a5fa" if aktif else R["sidebar_soluk"])
        if self.aktif_sayfa in self.sayfalar:
            self.sayfalar[self.aktif_sayfa].pack_forget()
        if anahtar not in self.sayfalar:
            sinif = {"panel": PanelSayfasi, "analiz": AnalizSayfasi, "gecmis": KayitSayfasi,
                     "Onaylandi": KayitSayfasi, "Reddedildi": KayitSayfasi, "bolge": BolgeSayfasi,
                     "kullanicilar": KullaniciSayfasi, "denetim": DenetimSayfasi}[anahtar]
            self.sayfalar[anahtar] = sinif(self, anahtar) if sinif is KayitSayfasi else sinif(self)
        s = self.sayfalar[anahtar]
        self.aktif_sayfa = anahtar
        self.lbl_baslik.config(text=s.baslik)
        self.lbl_alt.config(text=s.alt_baslik)
        s.pack(fill="both", expand=True, padx=28, pady=24)
        s.yenile()

    def yedek_al(self):
        yol = filedialog.asksaveasfilename(parent=self.root, defaultextension=".db",
                                           initialfile=f"ahika_yedek_{datetime.now():%Y%m%d_%H%M}.db",
                                           filetypes=[("SQLite veritabanı", "*.db")], title="Veritabanı yedeğini kaydet")
        if not yol:
            return
        try:
            g.manuel_yedek_al(yol, self.kadi)
            self.bildirim("Yedek oluşturuldu. Kişisel veriler şifreli; başka bilgisayarda açmak için kurtarma anahtarı gerekir.")
        except Exception as e:
            messagebox.showerror("Yedekleme hatası", str(e), parent=self.root)

    def excel_kaydet(self, kayitlar, sutunlar, ad, log_aciklama):
        if not kayitlar:
            self.bildirim("Dışa aktarılacak kayıt yok.", "uyari")
            return
        yol = filedialog.asksaveasfilename(parent=self.root, defaultextension=".xlsx",
                                           initialfile=f"{ad}_{datetime.now():%Y%m%d}.xlsx",
                                           filetypes=[("Excel dosyası", "*.xlsx")], title="Excel'e aktar")
        if not yol:
            return
        try:
            m.excel_yaz(kayitlar, sutunlar, yol)
            g.log_ekle(self.kadi, "DIŞA AKTARMA", f"{log_aciklama}: {len(kayitlar)} kayıt Excel'e aktarıldı.")
            self.bildirim(f"{len(kayitlar)} kayıt Excel'e aktarıldı.")
        except PermissionError:
            messagebox.showerror("Dosya açık", "Dosya başka bir programda açık olabilir. Kapatıp tekrar deneyin.", parent=self.root)
        except Exception as e:
            messagebox.showerror("Dışa aktarma hatası", str(e), parent=self.root)

    def calistir(self):
        self.root.mainloop()

# =====================================================================================
# SAYFA: GÖSTERGE PANELİ
# =====================================================================================

class PanelSayfasi(tk.Frame):
    baslik = "Gösterge Paneli"
    alt_baslik = "Başvuru ve karar süreçlerinin genel görünümü"

    def __init__(self, app):
        super().__init__(app.icerik, bg=R["zemin"])
        self.app = app
        ust = tk.Frame(self, bg=R["zemin"])
        ust.pack(fill="x", pady=(0, 18))
        sol = tk.Frame(ust, bg=R["zemin"])
        sol.pack(side="left")
        saat = datetime.now().hour
        selam = "Günaydın" if 5 <= saat < 12 else ("İyi günler" if saat < 18 else "İyi akşamlar")
        tk.Label(sol, text=f"{selam}, {app.kadi}", font=F(15, True), bg=R["zemin"], fg=R["metin"]).pack(anchor="w")
        bugun = datetime.now()
        tk.Label(sol, text=f"{bugun.day} {AYLAR[bugun.month - 1]} {bugun.year}, {GUNLER[bugun.weekday()]}",
                 font=F(9), bg=R["zemin"], fg=R["soluk"]).pack(anchor="w")
        Buton(ust, "✚  Yeni analiz başlat", lambda: app.sayfa_goster("analiz")).pack(side="right")

        self.kpi = tk.Frame(self, bg=R["zemin"])
        self.kpi.pack(fill="x")
        for i in range(4):
            self.kpi.grid_columnconfigure(i, weight=1, uniform="kpi")

        alt = tk.Frame(self, bg=R["zemin"])
        alt.pack(fill="both", expand=True, pady=(18, 0))
        alt.grid_columnconfigure(0, weight=3, uniform="a")
        alt.grid_columnconfigure(1, weight=2, uniform="a")
        alt.grid_rowconfigure(0, weight=1)

        k1, ic1 = kart(alt, padding=20)
        k1.grid(row=0, column=0, sticky="nsew", padx=(0, 9))
        sag = kart_basligi(ic1, "Son analizler", "En son yapılan 8 değerlendirme")
        Baglanti(sag, "Tümünü gör →", lambda: app.sayfa_goster("gecmis")).pack()
        self.tablo = Tablo(ic1, [("tarih", "Tarih", 140, "w"), ("ad", "Başvuru sahibi", 150, "w"),
                                 ("sektor", "Sektör", 120, "w"), ("hibe", "Hibe", 110, "e"),
                                 ("oran", "Başarı", 70, "e")], secim="browse", height=8,
                           bos_metin="Henüz analiz yapılmadı.")
        self.tablo.pack(fill="both", expand=True, pady=(14, 0))

        k2, self.ic2 = kart(alt, padding=20)
        k2.grid(row=0, column=1, sticky="nsew", padx=(9, 0))
        kart_basligi(self.ic2, "Sektörlere göre başvurular", "Başvuru sayısı ve ortalama başarı")
        self.sektor_alan = tk.Frame(self.ic2, bg=R["kart"])
        self.sektor_alan.pack(fill="both", expand=True, pady=(14, 0))

    def _kpi_kart(self, sutun, baslik, deger, alt, renk):
        k = kart(self.kpi)
        k.grid(row=0, column=sutun, sticky="nsew", padx=(0 if sutun == 0 else 9, 0 if sutun == 3 else 9))
        tk.Frame(k, bg=renk, height=3).pack(fill="x")
        ic = tk.Frame(k, bg=R["kart"])
        ic.pack(fill="both", padx=20, pady=(14, 18))
        tk.Label(ic, text=baslik.upper(), font=F(8, True), bg=R["kart"], fg=R["soluk"]).pack(anchor="w")
        tk.Label(ic, text=deger, font=F(22, True), bg=R["kart"], fg=R["metin"]).pack(anchor="w", pady=(6, 2))
        tk.Label(ic, text=alt, font=F(9), bg=R["kart"], fg=R["soluk"]).pack(anchor="w")

    def yenile(self):
        o = m.ozet_istatistikler()
        for w in self.kpi.winfo_children():
            w.destroy()
        karar = o['onay_adet'] + o['red_adet']
        onay_orani = (o['onay_adet'] / karar * 100) if karar else 0
        self._kpi_kart(0, "Toplam analiz", f"{o['toplam_analiz']:,}".replace(",", "."), f"Bugün: {o['bugun']}", R["vurgu"])
        self._kpi_kart(1, "Onaylanan hibe", str(o['onay_adet']), f"Toplam {m.tl(o['onay_tutar'], 0)}", R["basari"])
        self._kpi_kart(2, "Reddedilen", str(o['red_adet']), f"Onay oranı {m.yuzde(onay_orani)}", R["tehlike"])
        dog = f"Model doğruluğu {m.yuzde(self.app.motor.dogruluk)}" if self.app.motor.hazir else "Model yüklenmedi"
        self._kpi_kart(3, "Ortalama başarı", m.yuzde(o['ortalama_basari']), dog, R["uyari"])

        satirlar = []
        for r in m.gecmis_kayitlari()[:8]:
            satirlar.append((r['id'], (m.tarih_tr(r['tarih']), r['ad_soyad'], r['sektor'], m.tl(r['hibe_tutari'], 0),
                                       m.yuzde(r['basari_orani'])),
                             (r['tarih'], r['ad_soyad'], r['sektor'], r['hibe_tutari'], r['basari_orani'])))
        self.tablo.doldur(satirlar)

        for w in self.sektor_alan.winfo_children():
            w.destroy()
        if not o['sektorler']:
            tk.Label(self.sektor_alan, text="Henüz veri yok.", font=F(10), bg=R["kart"], fg=R["cok_soluk"]).pack(pady=30)
            return
        en_cok = max(n for _, n, _ in o['sektorler'])
        for ad, n, ort in o['sektorler'][:8]:
            satir = tk.Frame(self.sektor_alan, bg=R["kart"])
            satir.pack(fill="x", pady=5)
            ust = tk.Frame(satir, bg=R["kart"])
            ust.pack(fill="x")
            tk.Label(ust, text=ad, font=F(9, True), bg=R["kart"], fg=R["metin2"]).pack(side="left")
            tk.Label(ust, text=f"{n} başvuru  •  ort. {m.yuzde(ort)}", font=F(8), bg=R["kart"], fg=R["soluk"]).pack(side="right")
            CubukGrafik(satir, n / en_cok * 100, R["vurgu"], 7).pack(fill="x", pady=(4, 0))

# =====================================================================================
# SAYFA: YENİ ANALİZ
# =====================================================================================

class AnalizSayfasi(tk.Frame):
    baslik = "Yeni Analiz"
    alt_baslik = "Başvuru bilgilerini girin; model tüm sektörler için başarı olasılığını hesaplasın"

    def __init__(self, app):
        super().__init__(app.icerik, bg=R["zemin"])
        self.app = app
        self.girdi = None
        self.sonuc = None
        self.karar_verildi = None

        self.grid_columnconfigure(0, weight=0, minsize=470)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # --- form ---
        fk = kart(self)
        fk.grid(row=0, column=0, sticky="nsew", padx=(0, 18))
        form_alt = tk.Frame(fk, bg=R["zebra"])
        form_alt.pack(side="bottom", fill="x")
        tk.Frame(fk, bg=R["kenar"], height=1).pack(side="bottom", fill="x")
        fs = KaydirilabilirAlan(fk, R["kart"])
        fs.pack(fill="both", expand=True)
        f = tk.Frame(fs.ic, bg=R["kart"])
        f.pack(fill="both", expand=True, padx=24, pady=22)
        sag = kart_basligi(f, "Başvuru formu", "Tüm alanlar zorunludur")
        if app.yetki("veri_yukle"):
            Baglanti(sag, "Veri seti yükle", self.veri_yukle, boyut=8).pack(anchor="e")
        Baglanti(sag, "Kriter ağırlıkları", self.kriter_grafigi, boyut=8).pack(anchor="e", pady=(2, 0))

        izgara = tk.Frame(f, bg=R["kart"])
        izgara.pack(fill="x", pady=(14, 0))
        izgara.grid_columnconfigure(0, weight=1, uniform="f")
        izgara.grid_columnconfigure(1, weight=1, uniform="f")

        def bolum(satir, metin):
            tk.Label(izgara, text=metin, font=F(8, True), bg=R["kart"], fg=R["vurgu"]).grid(
                row=satir, column=0, columnspan=2, sticky="w", pady=(8 if satir else 0, 6))

        def yerlestir(alan, satir, sutun, span=1):
            alan.grid(row=satir, column=sutun, columnspan=span, sticky="ew",
                      padx=(0, 8) if sutun == 0 and span == 1 else ((8, 0) if sutun == 1 else 0), pady=(0, 10))
            return alan

        mt = app.motor
        bolum(0, "BAŞVURU SAHİBİ")
        self.a_ad = yerlestir(Alan(izgara, "Ad soyad", placeholder="ör. Ayşe Yılmaz"), 1, 0, 2)
        self.a_tel = yerlestir(Alan(izgara, "Telefon", placeholder="05XXXXXXXXX", width=14), 2, 0)
        self.a_ep = yerlestir(Alan(izgara, "E-posta", placeholder="ornek@mail.com", width=14), 2, 1)
        bolum(3, "PROJE BİLGİLERİ")
        self.a_sektor = yerlestir(Alan(izgara, "Sektör", "combo", values=mt.sektorler, width=14), 4, 0)
        self.a_tur = yerlestir(Alan(izgara, "Hibe türü", "combo", values=mt.hibe_turleri, width=14), 4, 1)
        self.a_lok = yerlestir(Alan(izgara, "Lokasyon (ilçe)", "combo", values=mt.lokasyonlar, width=14), 5, 0)
        self.a_buyume = yerlestir(Alan(izgara, "Sektörel büyüme", placeholder="1,0 – 10,0", ipucu="1–10", width=14), 5, 1)
        bolum(6, "FİNANSAL HEDEFLER")
        self.a_hibe = yerlestir(Alan(izgara, "Hibe tutarı (TL)", placeholder="ör. 500.000", width=14), 7, 0)
        self.a_ciro = yerlestir(Alan(izgara, "Ciro artış hedefi (TL)", placeholder="ör. 900.000", width=14), 7, 1)
        self.a_ist = yerlestir(Alan(izgara, "Hedef istihdam (kişi)", placeholder="ör. 5", width=14), 8, 0)

        self.alanlar = [self.a_ad, self.a_tel, self.a_ep, self.a_sektor, self.a_tur, self.a_lok, self.a_buyume,
                        self.a_hibe, self.a_ciro, self.a_ist]
        for a, b in zip(self.alanlar, self.alanlar[1:]):
            a.w.bind("<Return>", lambda e, s=b: s.odak())
        self.a_ist.w.bind("<Return>", lambda e: self.analiz_et())
        for a in (self.a_hibe, self.a_ciro):
            a.w.bind("<FocusOut>", lambda e, al=a: self._para_bicimle(al), add="+")

        tk.Label(f, text="İpucu: Enter ile sonraki alana geçin, Ctrl+Enter ile analiz edin.", font=F(8),
                 bg=R["kart"], fg=R["cok_soluk"]).pack(anchor="w", pady=(12, 0))
        bt = tk.Frame(form_alt, bg=R["zebra"])
        bt.pack(fill="x", padx=22, pady=14)
        self.btn_analiz = Buton(bt, "Analiz et", self.analiz_et, boyut=11)
        self.btn_analiz.pack(side="left", fill="x", expand=True)
        Buton(bt, "Temizle", self.temizle, stil="ikincil").pack(side="left", padx=(10, 0))

        # --- sonuç ---
        sk = kart(self)
        sk.grid(row=0, column=1, sticky="nsew")
        self.aksiyon = tk.Frame(sk, bg=R["zebra"])
        self.aksiyon.pack(side="bottom", fill="x")
        tk.Frame(sk, bg=R["kenar"], height=1).pack(side="bottom", fill="x")
        ai = tk.Frame(self.aksiyon, bg=R["zebra"])
        ai.pack(fill="x", padx=22, pady=14)
        self.btn_onay = Buton(ai, "✓  Onayla", lambda: self.karar("Onaylandi"), stil="basari", padx=12)
        self.btn_red = Buton(ai, "✕  Reddet", lambda: self.karar("Reddedildi"), stil="tehlike", padx=12)
        if app.yetki("karar"):
            self.btn_onay.pack(side="left")
            self.btn_red.pack(side="left", padx=(8, 0))
        else:
            tk.Label(ai, text="Karar verme yetkiniz bulunmuyor.", font=F(9), bg=R["zebra"], fg=R["soluk"]).pack(side="left")
        self.btn_pdf = Buton(ai, "PDF rapor", self.pdf, stil="ikincil", padx=12)
        self.btn_pdf.pack(side="right")
        self.btn_duy = Buton(ai, "Duyarlılık", self.duyarlilik, stil="ikincil", padx=12)
        self.btn_duy.pack(side="right", padx=(0, 8))

        self.sonuc_alan = KaydirilabilirAlan(sk, R["kart"])
        self.sonuc_alan.pack(fill="both", expand=True)
        self._aksiyonlar_etkin(False)
        self._bos_durum()

    def yenile(self):
        mt = self.app.motor
        self.a_sektor.w.config(values=mt.sektorler)
        self.a_tur.w.config(values=mt.hibe_turleri)
        self.a_lok.w.config(values=mt.lokasyonlar)
        self.btn_analiz.etkin(mt.hazir)
        if not mt.hazir and self.app.model_hatasi:
            self.app.bildirim("Model yüklenemedi: " + self.app.model_hatasi.split("\n")[0], "hata")

    def _aksiyonlar_etkin(self, durum):
        for b in (self.btn_pdf, self.btn_duy):
            b.etkin(durum)
        for b in (self.btn_onay, self.btn_red):
            b.etkin(durum and not self.karar_verildi)

    def _para_bicimle(self, alan):
        try:
            v = m.sayi_coz(alan.get())
            if v > 0:
                alan.set(m.ondalik(v, 0) if v == int(v) else m.ondalik(v))
        except ValueError:
            pass

    def _sonuc_temizle(self):
        for w in self.sonuc_alan.ic.winfo_children():
            w.destroy()
        self.sonuc_alan.basa()

    def _bos_durum(self):
        self._sonuc_temizle()
        f = tk.Frame(self.sonuc_alan.ic, bg=R["kart"])
        f.pack(fill="both", expand=True, padx=40, pady=70)
        c = tk.Canvas(f, width=72, height=72, bg=R["kart"], highlightthickness=0)
        c.create_oval(2, 2, 70, 70, fill=R["vurgu_acik"], outline="")
        c.create_text(36, 37, text="◎", fill=R["vurgu"], font=F(26))
        c.pack()
        tk.Label(f, text="Analiz sonucu burada görünecek", font=F(13, True), bg=R["kart"], fg=R["metin"]).pack(pady=(16, 4))
        tk.Label(f, text="Soldaki formu doldurun ve “Analiz et” düğmesine basın.", font=F(10), bg=R["kart"],
                 fg=R["soluk"]).pack()
        adimlar = tk.Frame(f, bg=R["kart"])
        adimlar.pack(pady=(26, 0))
        for i, (b, a) in enumerate([("Başvuru bilgileri", "Kişi ve iletişim bilgileri"),
                                     ("Proje parametreleri", "Sektör, ilçe, hibe türü ve hedefler"),
                                     ("Sonuç ve karar", "Başarı olasılığı, riskler, rapor ve karar")], 1):
            s = tk.Frame(adimlar, bg=R["kart"])
            s.pack(anchor="w", pady=6)
            n = tk.Canvas(s, width=28, height=28, bg=R["kart"], highlightthickness=0)
            n.create_oval(1, 1, 27, 27, fill=R["vurgu"], outline="")
            n.create_text(14, 15, text=str(i), fill="white", font=F(9, True))
            n.pack(side="left")
            t = tk.Frame(s, bg=R["kart"])
            t.pack(side="left", padx=12)
            tk.Label(t, text=b, font=F(10, True), bg=R["kart"], fg=R["metin"]).pack(anchor="w")
            tk.Label(t, text=a, font=F(9), bg=R["kart"], fg=R["soluk"]).pack(anchor="w")

    def temizle(self):
        for a in self.alanlar:
            a.temizle()
        self.girdi = self.sonuc = None
        self.karar_verildi = None
        self._karar_etiketleri()
        self._aksiyonlar_etkin(False)
        self._bos_durum()
        self.a_ad.odak()

    # ---------- doğrulama ----------
    def _topla(self):
        hatalar = []

        def h(alan, mesaj):
            alan.hata(mesaj)
            hatalar.append(alan)

        for a in self.alanlar:
            a.hata("")
        ad = self.a_ad.get()
        if len(ad) < 3:
            h(self.a_ad, "Ad soyad en az 3 karakter olmalı.")
        tel = self.a_tel.get().replace(" ", "")
        if not (tel.isdigit() and len(tel) == 11 and tel.startswith("0")):
            h(self.a_tel, "11 hane, 0 ile başlamalı.")
        ep = self.a_ep.get()
        if not re.match(r"^[^@\s]+@[^@\s]+\.[A-Za-z]{2,}$", ep):
            h(self.a_ep, "Geçerli bir e-posta girin.")
        for a, ad_ in ((self.a_sektor, "Sektör"), (self.a_tur, "Hibe türü"), (self.a_lok, "Lokasyon")):
            if not a.get():
                h(a, f"{ad_} seçin.")
        sayilar = {}
        for a, anahtar, kural in ((self.a_buyume, "buyume", "1-10"), (self.a_hibe, "hibe", "+"),
                                  (self.a_ciro, "ciro", "0+"), (self.a_ist, "istihdam", "int")):
            try:
                v = m.sayi_coz(a.get())
            except ValueError:
                h(a, "Sayısal bir değer girin.")
                continue
            if kural == "1-10" and not (1 <= v <= 10):
                h(a, "1,0 ile 10,0 arasında olmalı.")
            elif kural == "+" and v <= 0:
                h(a, "Sıfırdan büyük olmalı.")
            elif kural == "0+" and v < 0:
                h(a, "Negatif olamaz.")
            elif kural == "int" and (v < 0 or v != int(v)):
                h(a, "Tam sayı olmalı.")
            sayilar[anahtar] = int(v) if kural == "int" else v
        if hatalar:
            hatalar[0].odak()
            return None
        return dict(ad_soyad=ad, telefon=tel, eposta=ep, sektor=self.a_sektor.get(), turu=self.a_tur.get(),
                    lokasyon=self.a_lok.get(), **sayilar)

    def analiz_et(self):
        if not self.app.motor.hazir:
            messagebox.showerror("Model hazır değil", "Analiz için önce geçerli bir veri seti yükleyin.", parent=self.app.root)
            return
        girdi = self._topla()
        if girdi is None:
            self.app.bildirim("Lütfen işaretli alanları düzeltin.", "uyari")
            return
        try:
            sonuc = self.app.motor.analiz(girdi)
            m.analiz_kaydet(girdi, sonuc, self.app.kadi)
        except Exception as e:
            messagebox.showerror("Analiz hatası", str(e), parent=self.app.root)
            return
        self.girdi, self.sonuc, self.karar_verildi = girdi, sonuc, None
        self._karar_etiketleri()
        self._sonuc_ciz()
        self._aksiyonlar_etkin(True)
        self.app.bildirim("Analiz tamamlandı ve geçmişe kaydedildi.")

    # ---------- sonuç ----------
    def _bolum(self, parent, baslik):
        tk.Label(parent, text=baslik, font=F(11, True), bg=R["kart"], fg=R["metin"]).pack(anchor="w", pady=(22, 10))

    def _sonuc_ciz(self):
        self._sonuc_temizle()
        g_, s = self.girdi, self.sonuc
        p = tk.Frame(self.sonuc_alan.ic, bg=R["kart"])
        p.pack(fill="both", expand=True, padx=26, pady=22)
        renk, acik = seviye_renk(s['oran'])

        # Özet
        oz = tk.Frame(p, bg=acik)
        oz.pack(fill="x")
        oi = tk.Frame(oz, bg=acik)
        oi.pack(fill="x", padx=22, pady=18)
        sg = tk.Frame(oi, bg=acik)
        sg.pack(side="right", padx=(16, 0))
        sol = tk.Frame(oi, bg=acik)
        sol.pack(side="left", fill="x", expand=True)
        tk.Label(sol, text="SEÇİLEN SEKTÖR İÇİN BAŞARI OLASILIĞI", font=F(8, True), bg=acik, fg=renk).pack(anchor="w")
        tk.Label(sol, text=s['sektor'], font=F(16, True), bg=acik, fg=R["metin"]).pack(anchor="w", pady=(4, 2))
        tk.Label(sol, text=g_['ad_soyad'], font=F(9, True), bg=acik, fg=R["metin2"], anchor="w").pack(anchor="w")
        tk.Label(sol, text=f"{s['lokasyon']}  •  {s['turu']}  •  {m.tl(g_['hibe'], 0)}", font=F(9), bg=acik,
                 fg=R["metin2"], anchor="w", justify="left", wraplength=320).pack(anchor="w")
        tk.Label(sg, text=m.yuzde(s['oran']), font=F(30, True), bg=acik, fg=renk).pack(anchor="e")
        cip(sg, s['karar_onerisi'], "white", renk, 9).pack(anchor="e")
        CubukGrafik(oz, s['oran'], renk, 6, zemin="white").pack(fill="x", padx=22, pady=(0, 18))

        # Sektör karşılaştırması
        self._bolum(p, "Sektör karşılaştırması")
        tk.Label(p, text="Aynı parametrelerle tüm sektörler için modelin öngördüğü başarı olasılığı",
                 font=F(9), bg=R["kart"], fg=R["soluk"]).pack(anchor="w", pady=(0, 8))
        for ad, (oran, tahmin) in sorted(s['sonuclar'].items(), key=lambda kv: -kv[1][0]):
            satir = tk.Frame(p, bg=R["vurgu_acik"] if ad == s['sektor'] else R["kart"])
            satir.pack(fill="x", pady=1)
            bg = satir.cget("bg")
            ic = tk.Frame(satir, bg=bg)
            ic.pack(fill="x", padx=10, pady=6)
            ic.grid_columnconfigure(1, weight=1)
            ad_f = tk.Frame(ic, bg=bg, width=210, height=22)
            ad_f.grid(row=0, column=0, sticky="w")
            ad_f.pack_propagate(False)
            tk.Label(ad_f, text=ad, font=F(10, ad == s['sektor']), bg=bg, fg=R["metin"]).pack(side="left")
            if ad == s['en_iyi']:
                cip(ad_f, "EN YÜKSEK", R["basari_koyu"], R["basari_acik"], 7).pack(side="left", padx=6)
            CubukGrafik(ic, oran, seviye_renk(oran)[0], 8, zemin=R["kenar"] if bg == R["kart"] else "white").grid(
                row=0, column=1, sticky="ew", padx=12)
            tk.Label(ic, text=m.yuzde(oran), font=F(10, True), bg=bg, fg=R["metin"], width=7, anchor="e").grid(row=0, column=2)

        # Göstergeler
        self._bolum(p, "Finansal ve sosyal göstergeler")
        gs = tk.Frame(p, bg=R["kart"])
        gs.pack(fill="x")
        for i in range(3):
            gs.grid_columnconfigure(i, weight=1, uniform="g")
        for i, (baslik, deger, iyi, aciklama) in enumerate([
            ("Ciro / hibe katsayısı", m.ondalik(s['verimlilik']), s['verimlilik'] >= 1.5, "Hedef ≥ 1,50"),
            ("İstihdam etkisi", m.ondalik(s['istihdam_kapasitesi']), s['istihdam_kapasitesi'] >= 1.0, "100.000 TL başına kişi (hedef ≥ 1)"),
            ("Sektörel büyüme", m.ondalik(g_['buyume'], 1), g_['buyume'] >= 4.0, "Bölge eşiği 4,0"),
        ]):
            k = tk.Frame(gs, bg=R["zebra"], highlightthickness=1, highlightbackground=R["kenar"])
            k.grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else 6, 0 if i == 2 else 6))
            ki = tk.Frame(k, bg=R["zebra"])
            ki.pack(fill="both", padx=14, pady=12)
            tk.Label(ki, text=baslik, font=F(8, True), bg=R["zebra"], fg=R["soluk"]).pack(anchor="w")
            sat = tk.Frame(ki, bg=R["zebra"])
            sat.pack(anchor="w", pady=(4, 2))
            tk.Label(sat, text=deger, font=F(16, True), bg=R["zebra"], fg=R["metin"]).pack(side="left")
            tk.Label(sat, text="  ✓" if iyi else "  ⚠", font=F(11, True), bg=R["zebra"],
                     fg=R["basari"] if iyi else R["uyari"]).pack(side="left")
            tk.Label(ki, text=aciklama, font=F(8), bg=R["zebra"], fg=R["cok_soluk"]).pack(anchor="w")

        # Riskler
        self._bolum(p, "Risk uyarıları")
        if s['riskler']:
            for r in s['riskler']:
                rf = tk.Frame(p, bg=R["uyari_acik"])
                rf.pack(fill="x", pady=2)
                tk.Label(rf, text="⚠", font=F(10, True), bg=R["uyari_acik"], fg=R["uyari_koyu"]).pack(side="left", padx=(12, 8), pady=8)
                tk.Label(rf, text=r, font=F(9), bg=R["uyari_acik"], fg=R["uyari_koyu"], justify="left", anchor="w").pack(side="left", fill="x")
        else:
            rf = tk.Frame(p, bg=R["basari_acik"])
            rf.pack(fill="x")
            tk.Label(rf, text="✓  Belirgin bir risk tespit edilmedi.", font=F(9, True), bg=R["basari_acik"],
                     fg=R["basari_koyu"]).pack(anchor="w", padx=12, pady=8)

        # Değerlendirme
        self._bolum(p, "Değerlendirme")
        yorum = tk.Label(p, text=s['yorum'], font=F(10), bg=R["kart"], fg=R["metin2"], justify="left", anchor="w")
        yorum.pack(fill="x")
        p.bind("<Configure>", lambda e: yorum.config(wraplength=max(300, e.width - 10)))
        tk.Label(p, text="Bu sonuç karar destek amaçlıdır; nihai karar yetkili personel tarafından verilir.",
                 font=F(8), bg=R["kart"], fg=R["cok_soluk"]).pack(anchor="w", pady=(16, 0))

    # ---------- aksiyonlar ----------
    def karar(self, durum):
        if not self.app.yetki("karar") or not self.sonuc or self.karar_verildi:
            return
        metin = "onaylamak" if durum == "Onaylandi" else "reddetmek"
        if not self.app.onayla("Kararı kaydet", f"{self.girdi['ad_soyad']} başvurusunu {metin} istediğinize emin misiniz?\n\n"
                                                 f"Sektör: {self.sonuc['sektor']}\nHibe: {m.tl(self.girdi['hibe'])}"):
            return
        try:
            m.karar_kaydet(self.girdi, self.sonuc, durum, self.app.kadi)
        except Exception as e:
            messagebox.showerror("Kayıt hatası", str(e), parent=self.app.root)
            return
        self.karar_verildi = durum
        self._aksiyonlar_etkin(True)
        self._karar_etiketleri()
        if durum == "Onaylandi":
            self.app.bildirim("Başvuru onaylandı ve listeye eklendi.")
        else:
            self.app.bildirim("Başvuru reddedildi ve listeye eklendi.", "bilgi")

    def _karar_etiketleri(self):
        self.btn_onay.config(text="✓  Onaylandı" if self.karar_verildi == "Onaylandi" else "✓  Onayla")
        self.btn_red.config(text="✕  Reddedildi" if self.karar_verildi == "Reddedildi" else "✕  Reddet")

    def pdf(self):
        if not self.sonuc:
            return
        ad = m.tr_ascii(self.girdi['ad_soyad']).replace(" ", "_")
        yol = filedialog.asksaveasfilename(parent=self.app.root, defaultextension=".pdf",
                                           initialfile=f"AHIKA_Rapor_{ad}_{datetime.now():%Y%m%d}.pdf",
                                           filetypes=[("PDF dosyası", "*.pdf")], title="Raporu kaydet")
        if not yol:
            return
        durum = {"Onaylandi": "Onaylandı", "Reddedildi": "Reddedildi"}.get(self.karar_verildi, "Henüz karar verilmedi")
        try:
            m.pdf_analiz_raporu(yol, self.girdi, self.sonuc, durum, self.app.kadi,
                                kisisel_goster=self.app.yetki("kisisel_veri_gor"))
            g.log_ekle(self.app.kadi, "PDF RAPOR", f"Analiz raporu oluşturuldu ({self.sonuc['sektor']}, {self.sonuc['lokasyon']}).")
            self.app.bildirim("PDF raporu kaydedildi.")
            self._dosya_ac_sor(yol)
        except PermissionError:
            messagebox.showerror("Dosya açık", "PDF başka bir programda açık olabilir. Kapatıp tekrar deneyin.", parent=self.app.root)
        except Exception as e:
            messagebox.showerror("PDF hatası", str(e), parent=self.app.root)

    def _dosya_ac_sor(self, yol):
        if WINDOWS and messagebox.askyesno("Rapor hazır", "Raporu şimdi açmak ister misiniz?", parent=self.app.root):
            try:
                os.startfile(yol)
            except Exception:
                pass

    def veri_yukle(self):
        yol = filedialog.askopenfilename(parent=self.app.root, filetypes=[("Excel dosyası", "*.xlsx")],
                                         title="Eğitim veri setini seçin")
        if not yol:
            return
        try:
            self.app.root.config(cursor="watch")
            self.app.root.update()
            self.app.motor.egit(yol)
        except Exception as e:
            messagebox.showerror("Veri seti yüklenemedi", str(e), parent=self.app.root)
            return
        finally:
            self.app.root.config(cursor="")
        g.ayar_yaz("veri_seti_yolu", yol)
        g.log_ekle(self.app.kadi, "VERİ YÜKLEME", f"Yeni veri seti: {os.path.basename(yol)} "
                                                 f"({self.app.motor.kayit_sayisi} kayıt, doğruluk {m.yuzde(self.app.motor.dogruluk)})")
        self.app.model_hatasi = None
        self.app.durum_guncelle()
        self.yenile()
        self.temizle()
        self.app.bildirim(f"Model yeniden eğitildi • {self.app.motor.kayit_sayisi} kayıt • doğruluk {m.yuzde(self.app.motor.dogruluk)}")

    def kriter_grafigi(self):
        if not self.app.motor.hazir:
            return
        from matplotlib.figure import Figure
        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
        d = Diyalog(self.app.root, "Kriter ağırlıkları", 760, 560,
                    "Modelin hibe başarısını tahmin ederken her kritere verdiği önem (Random Forest önem dereceleri).",
                    boyutlanabilir=True)
        veri = self.app.motor.onem_dereceleri()
        fig = Figure(figsize=(7, 3.8), dpi=100, facecolor="white")
        ax = fig.add_subplot(111)
        adlar = [a for a, _ in veri]
        degerler = [v * 100 for _, v in veri]
        bars = ax.barh(adlar, degerler, color=R["vurgu"], height=0.6)
        for b, v in zip(bars, degerler):
            ax.text(b.get_width() + 0.4, b.get_y() + b.get_height() / 2, m.yuzde(v), va="center", fontsize=9, color=R["metin2"])
        ax.set_xlabel("Önem derecesi (%)", fontsize=9, color=R["soluk"])
        for k in ("top", "right"):
            ax.spines[k].set_visible(False)
        for k in ("left", "bottom"):
            ax.spines[k].set_color(R["kenar_koyu"])
        ax.tick_params(colors=R["metin2"], labelsize=9)
        ax.set_xlim(0, max(degerler) * 1.18)
        fig.tight_layout()
        FigureCanvasTkAgg(fig, master=d.govde).get_tk_widget().pack(fill="both", expand=True)
        Buton(d.alt_ic, "Kapat", d.destroy, stil="ikincil").pack(side="right")
        g.log_ekle(self.app.kadi, "GRAFİK", "Kriter ağırlık grafiği görüntülendi.")

    def duyarlilik(self):
        if not self.sonuc:
            return
        from matplotlib.figure import Figure
        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
        senaryolar, en_iyi = self.app.motor.duyarlilik(self.girdi, self.sonuc)
        hedef = self.sonuc['sektor']
        d = Diyalog(self.app.root, "Bütçe duyarlılık analizi", 1180, 800,
                    f"Hibe bütçesi %5 adımlarla değiştirildiğinde başarı olasılığının nasıl değiştiğini gösterir. "
                    f"Başvuru: {self.girdi['ad_soyad']} • Talep: {m.tl(self.girdi['hibe'])} • Hedef sektör: {hedef}",
                    boyutlanabilir=True)

        oneri = tk.Frame(d.govde, bg=R["basari_acik"])
        oneri.pack(fill="x")
        tk.Label(oneri, text=f"Önerilen ideal bütçe: {m.tl(en_iyi['tutar'])}   •   {hedef} için başarı "
                             f"{m.yuzde(en_iyi['oranlar'][hedef])}   ({en_iyi['isim']})",
                 font=F(10, True), bg=R["basari_acik"], fg=R["basari_koyu"]).pack(anchor="w", padx=14, pady=10)

        fig = Figure(figsize=(10, 2.3), dpi=100, facecolor="white")
        ax = fig.add_subplot(111)
        x = [s['tutar'] / 1000 for s in senaryolar]
        ax.plot(x, [s['oranlar'][hedef] for s in senaryolar], color=R["vurgu"], linewidth=2)
        ax.axvline(self.girdi['hibe'] / 1000, color=R["soluk"], linestyle="--", linewidth=1)
        ax.scatter([en_iyi['tutar'] / 1000], [en_iyi['oranlar'][hedef]], color=R["basari"], zorder=5, s=50)
        ax.set_xlabel("Hibe tutarı (bin TL)", fontsize=8, color=R["soluk"])
        ax.set_ylabel("Başarı (%)", fontsize=8, color=R["soluk"])
        for k in ("top", "right"):
            ax.spines[k].set_visible(False)
        ax.tick_params(labelsize=8, colors=R["metin2"])
        ax.grid(axis="y", color=R["kenar"], linewidth=0.6)
        fig.tight_layout()
        FigureCanvasTkAgg(fig, master=d.govde).get_tk_widget().pack(fill="x", pady=(10, 6))

        sektorler = list(self.app.motor.le_sektor.classes_)
        sut = [("isim", "Senaryo", 170, "w"), ("tutar", "Hibe tutarı", 120, "e")] + \
              [(f"s{i}", s, 95, "center") for i, s in enumerate(sektorler)]
        tablo = Tablo(d.govde, sut, secim="browse", height=9)
        tablo.pack(fill="both", expand=True)
        satirlar = []
        for idx, s in enumerate(senaryolar):
            etiket = "iyi" if s is en_iyi else ("vurgu" if s['degisim'] == 0 else None)
            satirlar.append((idx, (s['isim'], m.tl(s['tutar'], 0)) + tuple(m.yuzde(s['oranlar'][k]) for k in sektorler),
                             (s['degisim'], s['tutar']) + tuple(s['oranlar'][k] for k in sektorler), etiket))
        tablo.doldur(satirlar)
        tablo.tree.see(str(senaryolar.index(en_iyi)))
        tablo.tree.selection_set(str(senaryolar.index(en_iyi)))

        tk.Label(d.alt_ic, text="■ Mavi: talep edilen   ■ Yeşil: önerilen", font=F(8), bg=R["zebra"], fg=R["soluk"]).pack(side="left")

        def uygula():
            sec = tablo.secili()
            if not sec:
                self.app.bildirim("Tablodan bir senaryo seçin.", "uyari")
                return
            s = senaryolar[int(sec[0])]
            if s['tutar'] <= 0:
                self.app.bildirim("Sıfır tutarlı senaryo uygulanamaz.", "uyari")
                return
            d.destroy()
            self.a_hibe.set(m.ondalik(s['tutar']))
            g.log_ekle(self.app.kadi, "BÜTÇE GÜNCELLEME", f"Duyarlılık analizinden yeni bütçe seçildi: {m.tl(s['tutar'])}")
            self.analiz_et()

        Buton(d.alt_ic, "Seçili tutarla yeniden analiz et", uygula, stil="basari").pack(side="right")
        Buton(d.alt_ic, "Kapat", d.destroy, stil="ikincil").pack(side="right", padx=8)
        tablo.tree.bind("<Double-1>", lambda e: uygula())
        g.log_ekle(self.app.kadi, "DUYARLILIK ANALİZİ", f"İdeal bütçe: {m.tl(en_iyi['tutar'])}")

# =====================================================================================
# SAYFA: KAYIT LİSTELERİ (geçmiş / onaylanan / reddedilen)
# =====================================================================================

class KayitSayfasi(tk.Frame):
    def __init__(self, app, tur):
        super().__init__(app.icerik, bg=R["zemin"])
        self.app, self.tur = app, tur
        self.kayitlar = []
        self.baslik, self.alt_baslik = {
            "gecmis": ("Analiz Geçmişi", "Yapılan tüm yapay zekâ değerlendirmeleri"),
            "Onaylandi": ("Onaylanan Hibeler", "Onay kararı verilen başvurular"),
            "Reddedildi": ("Reddedilen Hibeler", "Ret kararı verilen başvurular"),
        }[tur]
        self.maskele = not app.yetki("kisisel_veri_gor")

        # filtre çubuğu
        fk, fi = kart(self, padding=16)
        fk.pack(fill="x", pady=(0, 14))
        self.ara = Girdi(fi, placeholder="Ara: ad, sektör, ilçe veya hibe türü…", width=40)
        self.ara.pack(side="left", ipady=7)
        tk.Label(fi, text="Tarih", font=F(9, True), bg=R["kart"], fg=R["metin2"]).pack(side="left", padx=(20, 8))
        self.bas = Girdi(fi, placeholder="GG.AA.YYYY", width=12)
        self.bas.pack(side="left", ipady=7)
        tk.Label(fi, text="–", font=F(10), bg=R["kart"], fg=R["soluk"]).pack(side="left", padx=6)
        self.bit = Girdi(fi, placeholder="GG.AA.YYYY", width=12)
        self.bit.pack(side="left", ipady=7)
        Baglanti(fi, "Filtreleri temizle", self.filtre_temizle, renk=R["soluk"]).pack(side="left", padx=16)
        self.lbl_sayi = tk.Label(fi, text="", font=F(9, True), bg=R["kart"], fg=R["soluk"])
        self.lbl_sayi.pack(side="right")
        for w in (self.ara, self.bas, self.bit):
            w.bind("<KeyRelease>", lambda e: self.filtrele(), add="+")

        # tablo
        tk_, ti = kart(self, padding=2)
        tk_.pack(fill="both", expand=True)
        ortak = [("tarih", "Tarih", 140, "w"), ("ad", "Başvuru sahibi", 160, "w"), ("tel", "Telefon", 110, "w"),
                 ("sektor", "Sektör", 130, "w"), ("tur", "Hibe türü", 130, "w"), ("hibe", "Hibe tutarı", 120, "e"),
                 ("ist", "İstihdam", 80, "center"), ("ciro", "Ciro artışı", 120, "e"), ("lok", "İlçe", 115, "w")]
        if tur == "gecmis":
            ortak += [("buyume", "Büyüme", 70, "center"), ("oran", "Başarı", 75, "e"), ("kim", "Analiz eden", 100, "w")]
        else:
            ortak += [("oran", "Başarı", 75, "e"), ("kim", "Karar veren", 100, "w")]
        self.tablo = Tablo(ti, ortak, bos_metin="Kayıt bulunamadı.")
        self.tablo.pack(fill="both", expand=True)
        self.tablo.tree.bind("<Double-1>", lambda e: self.detay())
        self.tablo.tree.bind("<<TreeviewSelect>>", lambda e: self._secim_degisti())
        self.tablo.tree.bind("<Delete>", lambda e: self.secili_sil())

        # alt çubuk
        alt = tk.Frame(self, bg=R["zemin"])
        alt.pack(fill="x", pady=(14, 0))
        Buton(alt, "Karşılaştır", self.karsilastir, stil="ikincil", boyut=9).pack(side="left")
        if app.yetki("kayit_sil"):
            Buton(alt, "Seçilenleri sil", self.secili_sil, stil="tehlike_cizgi", boyut=9).pack(side="left", padx=8)
        if app.yetki("liste_temizle"):
            Buton(alt, "Tüm listeyi temizle", self.tumunu_sil, stil="hayalet", boyut=9).pack(side="left")
        self.lbl_secim = tk.Label(alt, text="İpucu: Ayrıntı için çift tıklayın, karşılaştırma için iki kayıt seçin (Ctrl+tık).",
                                  font=F(8), bg=R["zemin"], fg=R["cok_soluk"])
        self.lbl_secim.pack(side="left", padx=14)
        if app.yetki("disa_aktar"):
            Buton(alt, "⤓  Excel'e aktar", self.excel, stil="ikincil", boyut=9).pack(side="right")

    def yenile(self):
        self.kayitlar = m.gecmis_kayitlari() if self.tur == "gecmis" else m.karar_kayitlari(self.tur)
        self.filtrele()

    def filtre_temizle(self):
        for w in (self.ara, self.bas, self.bit):
            w.temizle()
            w.hata(False)
        self.filtrele()

    def _tarih(self, girdi):
        t = girdi.get().strip()
        if not t:
            girdi.hata(False)
            return None
        try:
            v = datetime.strptime(t, "%d.%m.%Y").strftime("%Y-%m-%d")
            girdi.hata(False)
            return v
        except ValueError:
            girdi.hata(len(t) >= 10)
            return None

    def _goster(self, r):
        tel = g.telefon_maskele(r['telefon']) if self.maskele else (r['telefon'] or "")
        ad = r['ad_soyad'] or ""
        ortak = (m.tarih_tr(r['tarih']), ad, tel, r['sektor'] or "", r['hibe_turu'] or "", m.tl(r['hibe_tutari'], 0),
                 r['istihdam'] if r['istihdam'] is not None else "", m.tl(r['ciro_artisi'], 0), r['lokasyon'] or "")
        ham = (r['tarih'], ad, tel, r['sektor'], r['hibe_turu'], r['hibe_tutari'] or 0, r['istihdam'] or 0,
               r['ciro_artisi'] or 0, r['lokasyon'])
        if self.tur == "gecmis":
            return (ortak + (m.ondalik(r['buyume_endeksi'] or 0, 1), m.yuzde(r['basari_orani']), r['analiz_eden'] or "—"),
                    ham + (r['buyume_endeksi'] or 0, r['basari_orani'] or 0, r['analiz_eden']))
        return (ortak + (m.yuzde(r['basari_orani']), r['karar_veren'] or "—"),
                ham + (r['basari_orani'] or 0, r['karar_veren']))

    def filtrele(self):
        q = m.tr_kucuk(self.ara.get())
        b, e = self._tarih(self.bas), self._tarih(self.bit)
        satirlar = []
        for r in self.kayitlar:
            metin = " ".join(m.tr_kucuk(r.get(k)) for k in ("ad_soyad", "sektor", "hibe_turu", "lokasyon"))
            if q and q not in metin:
                continue
            gun = str(r['tarih'])[:10]
            if (b and gun < b) or (e and gun > e):
                continue
            goster, ham = self._goster(r)
            satirlar.append((r['id'], goster, ham))
        self.tablo.doldur(satirlar)
        self.lbl_sayi.config(text=f"{len(satirlar)} / {len(self.kayitlar)} kayıt")

    def _secim_degisti(self):
        n = len(self.tablo.secili())
        self.lbl_secim.config(text=f"{n} kayıt seçili" if n else
                              "İpucu: Ayrıntı için çift tıklayın, karşılaştırma için iki kayıt seçin (Ctrl+tık).")

    def _kayit(self, iid):
        return next((r for r in self.kayitlar if str(r['id']) == str(iid)), None)

    def detay(self):
        sec = self.tablo.secili()
        if not sec:
            return
        r = self._kayit(sec[0])
        d = Diyalog(self.app.root, "Kayıt ayrıntısı", 520, 640, f"Kayıt #{r['id']} • {m.tarih_tr(r['tarih'])}")
        alanlar = [("Başvuru sahibi", r['ad_soyad']),
                   ("Telefon", g.telefon_maskele(r['telefon']) if self.maskele else r['telefon']),
                   ("E-posta", g.eposta_maskele(r['eposta']) if self.maskele else r['eposta']),
                   ("Sektör", r['sektor']), ("Hibe türü", r['hibe_turu']), ("İlçe", r['lokasyon']),
                   ("Hibe tutarı", m.tl(r['hibe_tutari'])), ("Hedef istihdam", f"{r['istihdam']} kişi"),
                   ("Ciro artış hedefi", m.tl(r['ciro_artisi'])), ("Başarı olasılığı", m.yuzde(r['basari_orani']))]
        if self.tur == "gecmis":
            alanlar += [("Sektörel büyüme", m.ondalik(r['buyume_endeksi'] or 0, 1)), ("Analiz eden", r['analiz_eden'] or "—")]
        else:
            alanlar += [("Karar", "Onaylandı" if self.tur == "Onaylandi" else "Reddedildi"), ("Karar veren", r['karar_veren'] or "—")]
        for i, (k, v) in enumerate(alanlar):
            s = tk.Frame(d.govde, bg=R["zebra"] if i % 2 else R["kart"])
            s.pack(fill="x")
            tk.Label(s, text=k, font=F(9), bg=s.cget("bg"), fg=R["soluk"], width=18, anchor="w").pack(side="left", padx=10, pady=7)
            tk.Label(s, text=v if v not in (None, "") else "—", font=F(10, True), bg=s.cget("bg"), fg=R["metin"],
                     anchor="w").pack(side="left")
        Buton(d.alt_ic, "Kapat", d.destroy, stil="ikincil").pack(side="right")
        d.sigdir(520)

    def karsilastir(self):
        sec = self.tablo.secili()
        if len(sec) != 2:
            self.app.bildirim("Karşılaştırma için tam olarak 2 kayıt seçin (Ctrl+tık).", "uyari")
            return
        a, b = self._kayit(sec[0]), self._kayit(sec[1])
        kriterler = [("Tarih", m.tarih_tr(a['tarih']), m.tarih_tr(b['tarih'])),
                     ("Sektör", a['sektor'], b['sektor']), ("Hibe türü", a['hibe_turu'], b['hibe_turu']),
                     ("İlçe", a['lokasyon'], b['lokasyon']), ("Hibe tutarı", m.tl(a['hibe_tutari']), m.tl(b['hibe_tutari'])),
                     ("Hedef istihdam", a['istihdam'], b['istihdam']),
                     ("Ciro artışı", m.tl(a['ciro_artisi']), m.tl(b['ciro_artisi'])),
                     ("Ciro / hibe", m.ondalik((a['ciro_artisi'] or 0) / (a['hibe_tutari'] or 1)),
                      m.ondalik((b['ciro_artisi'] or 0) / (b['hibe_tutari'] or 1))),
                     ("Başarı olasılığı", m.yuzde(a['basari_orani']), m.yuzde(b['basari_orani']))]
        oa, ob = a['basari_orani'] or 0, b['basari_orani'] or 0
        if abs(oa - ob) < 0.05:
            sonuc = "İki projenin başarı olasılığı eşit."
        else:
            kazanan = a if oa > ob else b
            sonuc = (f"{kazanan['ad_soyad']} projesi {m.yuzde(abs(oa - ob))} puan daha yüksek başarı olasılığıyla "
                     f"daha avantajlı görünmektedir.")
        d = Diyalog(self.app.root, "Proje karşılaştırması", 760, 620, "Seçilen iki başvurunun yan yana karşılaştırması")
        basliklar = [a['ad_soyad'] or "Proje 1", b['ad_soyad'] or "Proje 2"]
        t = Tablo(d.govde, [("k", "Kriter", 180, "w"), ("a", basliklar[0], 230, "center"), ("b", basliklar[1], 230, "center")],
                  secim="none", height=9)
        t.pack(fill="both", expand=True)
        t.doldur([(i, k, k) for i, k in enumerate(kriterler)])
        sf = tk.Frame(d.govde, bg=R["basari_acik"])
        sf.pack(fill="x", pady=(12, 0))
        tk.Label(sf, text=sonuc, font=F(10, True), bg=R["basari_acik"], fg=R["basari_koyu"], wraplength=680,
                 justify="left").pack(anchor="w", padx=14, pady=10)

        def pdf():
            yol = filedialog.asksaveasfilename(parent=d, defaultextension=".pdf", filetypes=[("PDF dosyası", "*.pdf")],
                                               initialfile=f"AHIKA_Karsilastirma_{datetime.now():%Y%m%d}.pdf")
            if not yol:
                return
            try:
                m.pdf_karsilastirma(yol, basliklar, kriterler, sonuc, self.app.kadi)
                g.log_ekle(self.app.kadi, "PDF RAPOR", f"Karşılaştırma raporu: kayıt #{a['id']} / #{b['id']}")
                self.app.bildirim("Karşılaştırma raporu kaydedildi.")
            except Exception as e:
                messagebox.showerror("PDF hatası", str(e), parent=d)

        Buton(d.alt_ic, "PDF olarak kaydet", pdf).pack(side="right")
        Buton(d.alt_ic, "Kapat", d.destroy, stil="ikincil").pack(side="right", padx=8)

    def secili_sil(self):
        if not self.app.yetki("kayit_sil"):
            return
        sec = self.tablo.secili()
        if not sec:
            self.app.bildirim("Silmek için önce kayıt seçin.", "uyari")
            return
        if not self.app.onayla("Kayıtları sil", f"Seçilen {len(sec)} kayıt kalıcı olarak silinecek. Bu işlem geri alınamaz.\n\nDevam edilsin mi?", True):
            return
        m.kayitlari_sil("gecmis" if self.tur == "gecmis" else "karar", sec, self.app.kadi)
        self.yenile()
        self.app.bildirim(f"{len(sec)} kayıt silindi.")

    def tumunu_sil(self):
        if not self.app.yetki("liste_temizle"):
            return
        if not self.kayitlar:
            return
        if not self.app.onayla("Listeyi temizle", f"“{self.baslik}” listesindeki {len(self.kayitlar)} kaydın tamamı kalıcı olarak silinecek.\n\n"
                                                    "Bu işlemden önce yedek almanız önerilir. Devam edilsin mi?", True):
            return
        n = m.liste_temizle(self.tur, self.app.kadi)
        self.yenile()
        self.app.bildirim(f"{n} kayıt silindi.")

    def excel(self):
        sut = [("tarih", "Tarih"), ("ad_soyad", "Başvuru Sahibi"), ("telefon", "Telefon"), ("eposta", "E-posta"),
               ("sektor", "Sektör"), ("hibe_turu", "Hibe Türü"), ("hibe_tutari", "Hibe Tutarı (TL)"),
               ("istihdam", "Hedef İstihdam"), ("ciro_artisi", "Ciro Artışı (TL)"), ("lokasyon", "İlçe")]
        if self.tur == "gecmis":
            sut += [("buyume_endeksi", "Sektörel Büyüme"), ("basari_orani", "Başarı Olasılığı (%)"), ("analiz_eden", "Analiz Eden")]
        else:
            sut += [("basari_orani", "Başarı Olasılığı (%)"), ("karar_veren", "Karar Veren")]
        gorunen = {i for i in self.tablo.tree.get_children()}
        kayitlar = [r for r in self.kayitlar if str(r['id']) in gorunen]
        self.app.excel_kaydet(kayitlar, sut, self.baslik.replace(" ", "_"), self.baslik)

# =====================================================================================
# SAYFA: BÖLGE ANALİZİ
# =====================================================================================

HARITA_DB = os.path.join(g.APP_DIR, "harita_onbellek.db")


def harita_onbellegi_var():
    if not os.path.exists(HARITA_DB):
        return False
    try:
        c = sqlite3.connect(HARITA_DB)
        n = c.execute("SELECT COUNT(*) FROM tiles").fetchone()[0]
        c.close()
        return n > 50
    except sqlite3.Error:
        return False


def basari_rengi(o):
    if o is None:
        return R["cok_soluk"]
    return R["basari"] if o >= 60 else (R["uyari"] if o >= 40 else R["tehlike"])


class SemaHarita(tk.Canvas):
    """İnternet gerektirmeyen şematik ilçe haritası (gerçek koordinatlara göre ölçekli)."""

    def __init__(self, parent, tikla):
        super().__init__(parent, bg="#f8fafc", highlightthickness=1, highlightbackground=R["kenar"])
        self.tikla = tikla
        self.veri = {}
        self.bind("<Configure>", lambda e: self.ciz())

    def guncelle(self, veri):
        self.veri = veri
        self.ciz()

    # Yönleri koruyan, okunabilirlik için ölçeksiz yerleşim (metro haritası mantığı)
    KONUM = {"Kozaklı": (0.52, 0.10), "Hacıbektaş": (0.24, 0.24), "Gülşehir": (0.30, 0.47),
             "Avanos": (0.76, 0.38), "Nevşehir Merkez": (0.52, 0.55), "Ürgüp": (0.80, 0.70),
             "Acıgöl": (0.24, 0.74), "Derinkuyu": (0.52, 0.88)}

    def ciz(self):
        self.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        if w < 50 or h < 50:
            return
        pad_x, pad_y = 70, 36

        def xy(ilce):
            if ilce in self.KONUM:
                fx, fy = self.KONUM[ilce]
            else:  # listede olmayan ilçe gelirse gerçek koordinattan yaklaşık konum
                lat, lon = m.NEVSEHIR_ILCE_KOORDINAT[ilce]
                fx, fy = (lon - 34.5) / 0.5, (39.3 - lat) / 1.0
            return pad_x + fx * (w - 2 * pad_x), pad_y + fy * (h - 2 * pad_y - 30)

        for i in range(1, 8):
            self.create_line(w * i / 8, 0, w * i / 8, h, fill="#eef2f7")
            self.create_line(0, h * i / 8, w, h * i / 8, fill="#eef2f7")
        merkez = xy("Nevşehir Merkez")
        for ilce in m.NEVSEHIR_ILCE_KOORDINAT:
            if ilce != "Nevşehir Merkez":
                self.create_line(*merkez, *xy(ilce), fill="#dbe3ec", width=2, dash=(4, 4))
        self.create_text(w - 16, 14, text="K ↑", anchor="ne", fill=R["soluk"], font=F(9, True))
        self.create_text(14, h - 12, text="Şematik görünüm — yönler korunur, ölçek yoktur", anchor="sw",
                         fill=R["cok_soluk"], font=F(8))

        for ilce in m.NEVSEHIR_ILCE_KOORDINAT:
            n, o = self.veri.get(m.tr_kucuk(ilce), (0, None))
            x, y = xy(ilce)
            r = 11 + min(14, 3 * math.sqrt(n))
            etiket = f"ilce_{abs(hash(ilce))}"
            self.create_oval(x - r - 4, y - r - 4, x + r + 4, y + r + 4, fill="white", outline="", tags=etiket)
            self.create_oval(x - r, y - r, x + r, y + r, fill=basari_rengi(o), outline="white", width=2, tags=etiket)
            if n:
                self.create_text(x, y, text=str(n), fill="white", font=F(9, True), tags=etiket)
            self.create_text(x, y + r + 12, text=ilce, fill=R["metin"], font=F(9, True), tags=etiket)
            self.create_text(x, y + r + 27, text=(f"{n} proje • {m.yuzde(o)}" if o is not None else "veri yok"),
                             fill=R["soluk"], font=F(8), tags=etiket)
            self.tag_bind(etiket, "<Button-1>", lambda e, i=ilce: self.tikla(i))
            self.tag_bind(etiket, "<Enter>", lambda e: self.config(cursor="hand2"))
            self.tag_bind(etiket, "<Leave>", lambda e: self.config(cursor=""))


class BolgeSayfasi(tk.Frame):
    baslik = "Bölge Analizi"
    alt_baslik = "İlçe bazında proje yoğunluğu ve ortalama başarı"

    def __init__(self, app):
        super().__init__(app.icerik, bg=R["zemin"])
        self.app = app
        self.markerlar = []
        self.grid_columnconfigure(0, weight=2, uniform="b")
        self.grid_columnconfigure(1, weight=3, uniform="b")
        self.grid_rowconfigure(0, weight=1)

        k1, i1 = kart(self, padding=20)
        k1.grid(row=0, column=0, sticky="nsew", padx=(0, 9))
        sag = kart_basligi(i1, "İlçe özeti", "Ayrıntılı dağılım için satıra çift tıklayın")
        self.tablo = Tablo(i1, [("ilce", "İlçe", 150, "w"), ("adet", "Proje", 70, "center"), ("ort", "Ort. başarı", 100, "e")],
                           secim="browse", bos_metin="Henüz bölge verisi yok.")
        self.tablo.pack(fill="both", expand=True, pady=(14, 0))
        self.tablo.tree.bind("<Double-1>", lambda e: self.tablo.secili() and self.ilce_detay(self.tablo.secili()[0]))
        bt = tk.Frame(i1, bg=R["kart"])
        bt.pack(fill="x", pady=(14, 0))
        if app.yetki("disa_aktar"):
            Buton(bt, "⤓  Excel'e aktar", self.excel, stil="ikincil", boyut=9).pack(side="left")
        if app.yetki("bolge_sil"):
            Buton(bt, "Tüm bölge verilerini sil", self.sil, stil="tehlike_cizgi", boyut=9).pack(side="right")

        k2, i2 = kart(self, padding=20)
        k2.grid(row=0, column=1, sticky="nsew", padx=(9, 0))
        sag2 = kart_basligi(i2, "Nevşehir ilçe haritası", "İlçeye tıklayarak ayrıntıyı açın")
        self.gorunum = tk.Frame(sag2, bg=R["kart"])
        self.gorunum.pack(anchor="e")
        self.sekme = {}
        for anahtar, metin in (("harita", "Harita"), ("sema", "Şema")):
            l = tk.Label(self.gorunum, text=metin, font=F(9, True), padx=12, pady=4, cursor="hand2",
                         highlightthickness=1, highlightbackground=R["kenar_koyu"])
            l.pack(side="left")
            l.bind("<Button-1>", lambda e, k=anahtar: self.gorunum_sec(k))
            self.sekme[anahtar] = l
        lej = tk.Frame(i2, bg=R["kart"])
        lej.pack(fill="x", pady=(10, 0))
        for metin, renk in (("≥ %60", R["basari"]), ("%40–60", R["uyari"]), ("< %40", R["tehlike"]), ("Veri yok", R["cok_soluk"])):
            tk.Label(lej, text="●", fg=renk, bg=R["kart"], font=F(10)).pack(side="left")
            tk.Label(lej, text=metin, fg=R["soluk"], bg=R["kart"], font=F(8)).pack(side="left", padx=(2, 10))
        self.lbl_harita = tk.Label(lej, text="", fg=R["cok_soluk"], bg=R["kart"], font=F(8))
        self.lbl_harita.pack(side="right")

        self.harita_kap = tk.Frame(i2, bg=R["kart"])
        self.harita_kap.pack(fill="both", expand=True, pady=(8, 0))
        self.sema = SemaHarita(self.harita_kap, self.ilce_detay)
        self.harita = None
        if tkintermapview:
            self.harita = tkintermapview.TkinterMapView(self.harita_kap, corner_radius=0, database_path=HARITA_DB)
            self.harita.set_position(38.76, 34.76)
            self.harita.set_zoom(9)
        self._indirme = {"durum": None}
        cevrimdisi = harita_onbellegi_var()
        if self.harita and cevrimdisi:
            self.lbl_harita.config(text="Çevrimdışı harita hazır")
            self.gorunum_sec("harita")
        else:
            self.gorunum_sec("sema")
            if self.harita:
                self.lbl_harita.config(text="İnternet kontrol ediliyor…")
                threading.Thread(target=self._onbellek_indir, daemon=True).start()
                self.after(1000, self._indirme_izle)
            else:
                self.lbl_harita.config(text="Şematik görünüm (çevrimdışı)")

    def gorunum_sec(self, k):
        if k == "harita" and not self.harita:
            return
        self.aktif_gorunum = k
        for ad, l in self.sekme.items():
            sec = ad == k
            l.config(bg=R["vurgu"] if sec else "white", fg="white" if sec else R["metin2"],
                     highlightbackground=R["vurgu"] if sec else R["kenar_koyu"])
        self.sema.pack_forget()
        if self.harita:
            self.harita.pack_forget()
        (self.harita if k == "harita" else self.sema).pack(fill="both", expand=True)

    def _onbellek_indir(self):
        """Arka planda: internet varsa Nevşehir karolarını bir kez indirip saklar (yaklaşık 2-3 MB)."""
        try:
            socket.create_connection(("a.tile.openstreetmap.org", 443), timeout=3).close()
        except OSError:
            self._indirme["durum"] = "yok"
            return
        self._indirme["durum"] = "iniyor"
        try:
            yukleyici = tkintermapview.OfflineLoader(path=HARITA_DB)
            yukleyici.number_of_threads = 4
            yukleyici.save_offline_tiles((39.40, 34.30), (38.20, 35.20), 8, 12)
            self._indirme["durum"] = "tamam" if harita_onbellegi_var() else "hata"
        except Exception:
            self._indirme["durum"] = "hata"

    def _indirme_izle(self):
        d = self._indirme["durum"]
        if d in (None, "iniyor"):
            if d == "iniyor":
                self.lbl_harita.config(text="Harita çevrimdışı kullanım için indiriliyor…")
            self.after(1000, self._indirme_izle)
        elif d == "tamam":
            self.lbl_harita.config(text="Çevrimdışı harita hazır")
            self.gorunum_sec("harita")
        elif d == "yok":
            self.lbl_harita.config(text="İnternet yok — şematik görünüm")
        else:
            self.lbl_harita.config(text="Harita indirilemedi — şematik görünüm")

    def yenile(self):
        self.veri = m.bolge_istatistikleri()
        self.tablo.doldur([(ad, (ad, n, m.yuzde(o)), (ad, n, o)) for ad, n, o in self.veri])
        sozluk = {m.tr_kucuk(ad): (n, o) for ad, n, o in self.veri}
        self.sema.guncelle(sozluk)
        if not self.harita:
            return
        for mk in self.markerlar:
            mk.delete()
        self.markerlar = []
        for ilce, (lat, lon) in m.NEVSEHIR_ILCE_KOORDINAT.items():
            n, o = sozluk.get(m.tr_kucuk(ilce), (0, None))
            renk = basari_rengi(o)
            metin = ilce if o is None else f"{ilce} {m.yuzde(o)}"
            try:
                mk = self.harita.set_marker(lat, lon, text=metin, marker_color_circle="white",
                                            marker_color_outside=renk, text_color=R["metin"], font=F(8, True),
                                            command=lambda _m, i=ilce: self.ilce_detay(i))
            except TypeError:
                mk = self.harita.set_marker(lat, lon, text=metin)
            self.markerlar.append(mk)

    def ilce_detay(self, ilce):
        adet, ort, sektor, tur = m.ilce_detay(ilce)
        d = Diyalog(self.app.root, f"{ilce} bölge ayrıntısı", 820, 620,
                    f"Toplam {adet} proje  •  Ortalama başarı {m.yuzde(ort)}", boyutlanabilir=True)
        if not adet:
            tk.Label(d.govde, text="Bu ilçede henüz kayıtlı analiz bulunmuyor.", font=F(10), bg=R["kart"],
                     fg=R["soluk"]).pack(pady=40)
        else:
            f = tk.Frame(d.govde, bg=R["kart"])
            f.pack(fill="both", expand=True)
            f.grid_columnconfigure(0, weight=1, uniform="d")
            f.grid_columnconfigure(1, weight=1, uniform="d")
            f.grid_rowconfigure(1, weight=1)
            for i, (bas, veri) in enumerate((("Sektör dağılımı", sektor), ("Hibe türü dağılımı", tur))):
                tk.Label(f, text=bas, font=F(10, True), bg=R["kart"], fg=R["metin"]).grid(row=0, column=i, sticky="w", pady=(0, 8))
                t = Tablo(f, [("ad", bas.split()[0], 150, "w"), ("n", "Proje", 60, "center"), ("o", "Ort. başarı", 90, "e")],
                          secim="browse", height=10)
                t.grid(row=1, column=i, sticky="nsew", padx=(0, 8) if i == 0 else (8, 0))
                t.doldur([(j, (a, n, m.yuzde(o)), (a, n, o)) for j, (a, n, o) in enumerate(veri)])
        Buton(d.alt_ic, "Kapat", d.destroy, stil="ikincil").pack(side="right")
        g.log_ekle(self.app.kadi, "BÖLGE DETAY", f"{ilce} bölge ayrıntısı görüntülendi.")

    def excel(self):
        kayit = [{"ilce": a, "adet": n, "ort": round(o, 2)} for a, n, o in self.veri]
        self.app.excel_kaydet(kayit, [("ilce", "İlçe"), ("adet", "Proje Sayısı"), ("ort", "Ortalama Başarı (%)")],
                              "Bolge_Analizi", "Bölge analizi")

    def sil(self):
        if not self.app.onayla("Bölge verilerini sil", "Tüm analiz geçmişi ile onay/ret kayıtları kalıcı olarak silinecek.\n\n"
                                                         "Bu işlem geri alınamaz. Devam etmeden önce yedek almanız önerilir. Emin misiniz?", True):
            return
        g.otomatik_yedek_al()
        m.bolge_verilerini_sil(self.app.kadi)
        self.yenile()
        self.app.bildirim("Bölge verileri silindi (silmeden önce otomatik yedek alındı).")

# =====================================================================================
# SAYFA: KULLANICILAR VE DAVETLER
# =====================================================================================

class KullaniciSayfasi(tk.Frame):
    baslik = "Kullanıcılar ve Davetler"
    alt_baslik = "Hesapları, rolleri ve tek kullanımlık davet kodlarını yönetin"

    def __init__(self, app):
        super().__init__(app.icerik, bg=R["zemin"])
        self.app = app
        self.grid_columnconfigure(0, weight=3, uniform="k")
        self.grid_columnconfigure(1, weight=2, uniform="k")
        self.grid_rowconfigure(0, weight=1)

        k1, i1 = kart(self, padding=20)
        k1.grid(row=0, column=0, sticky="nsew", padx=(0, 9))
        kart_basligi(i1, "Kullanıcı hesapları", "Yalnızca kendi seviyenizin altındaki hesapları yönetebilirsiniz")
        self.tablo = Tablo(i1, [("kadi", "Kullanıcı adı", 140, "w"), ("rol", "Rol", 110, "w"), ("durum", "Durum", 150, "w"),
                                ("son", "Son giriş", 130, "w")], secim="browse")
        self.tablo.pack(fill="both", expand=True, pady=(14, 0))
        bt = tk.Frame(i1, bg=R["kart"])
        bt.pack(fill="x", pady=(14, 0))
        Buton(bt, "Rol değiştir", self.rol, stil="ikincil", boyut=9).pack(side="left")
        self.btn_aktif = Buton(bt, "Devre dışı bırak", self.aktiflik, stil="ikincil", boyut=9)
        self.btn_aktif.pack(side="left", padx=6)
        Buton(bt, "Kilidi aç", self.kilit, stil="ikincil", boyut=9).pack(side="left")
        Buton(bt, "Şifre sıfırla", self.sifirla, stil="ikincil", boyut=9).pack(side="left", padx=6)
        Buton(bt, "Sil", self.sil, stil="tehlike_cizgi", boyut=9).pack(side="right")
        self.tablo.tree.bind("<<TreeviewSelect>>", lambda e: self._secim())

        k2, i2 = kart(self, padding=20)
        k2.grid(row=0, column=1, sticky="nsew", padx=(9, 0))
        kart_basligi(i2, "Davet kodları", f"Tek kullanımlıktır; varsayılan geçerlilik {g.DAVET_GECERLILIK_GUN} gün")
        form = tk.Frame(i2, bg=R["kart"])
        form.pack(fill="x", pady=(14, 0))
        form.grid_columnconfigure(0, weight=1)
        form.grid_columnconfigure(1, weight=1)
        self.a_rol = Alan(form, "Rol", "combo", values=g.atanabilir_roller(app.kullanici["rol"]), width=12)
        self.a_rol.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        roller = g.atanabilir_roller(app.kullanici["rol"])
        self.a_rol.set("Uzman" if "Uzman" in roller else (roller[-1] if roller else ""))
        self.a_gun = Alan(form, "Geçerlilik", "combo", values=["1 gün", "3 gün", "7 gün", "14 gün", "30 gün"], width=10)
        self.a_gun.grid(row=0, column=1, sticky="ew", padx=(6, 0))
        self.a_gun.set(f"{g.DAVET_GECERLILIK_GUN} gün")
        Buton(i2, "✚  Davet kodu oluştur", self.kod_olustur).pack(fill="x", pady=(4, 14))
        self.kodlar = Tablo(i2, [("kod", "Kod", 80, "w"), ("rol", "Rol", 90, "w"), ("son", "Son geçerlilik", 120, "w"),
                                 ("durum", "Durum", 130, "w")], secim="browse", bos_metin="Henüz davet kodu yok.")
        self.kodlar.pack(fill="both", expand=True)
        alt2 = tk.Frame(i2, bg=R["kart"])
        alt2.pack(fill="x", pady=(12, 0))
        Buton(alt2, "Seçili kodu iptal et", self.kod_iptal, stil="tehlike_cizgi", boyut=9).pack(side="right")
        if app.kullanici["rol"] == "Ana Yönetici":
            Buton(alt2, "Veri kurtarma anahtarı", app.kurtarma_anahtari_diyalogu, stil="ikincil", boyut=9).pack(side="left")

    def yenile(self):
        self.kullanicilar = g.kullanicilari_listele()
        simd = g.simdi()
        satirlar = []
        for u in self.kullanicilar:
            if not u['aktif']:
                durum, et = "Devre dışı", None
            elif u['kilit_bitis'] and u['kilit_bitis'] > simd:
                durum, et = "Kilitli", None
            elif u['sifre_degismeli']:
                durum, et = "Şifre yenileme bekliyor", None
            else:
                durum, et = "Aktif", None
            ad = u['kullanici_adi'] + ("  (siz)" if u['kullanici_adi'] == self.app.kadi else "")
            satirlar.append((u['id'], (ad, u['rol'], durum, m.tarih_tr(u['son_giris']) if u['son_giris'] else "—"),
                             (u['kullanici_adi'], g.ROL_SEVIYE.get(u['rol'], 0), durum, u['son_giris'] or "")))
        self.tablo.doldur(satirlar)
        kod_satir = []
        for k in g.davet_kodlarini_listele()[:100]:
            kod_satir.append((k['id'], ("…" + (k['kod_ipucu'] or ""), k['rol'], m.tarih_tr(k['son_kullanma']), k['durum']),
                              (k['kod_ipucu'], k['rol'], k['son_kullanma'], k['durum'])))
        self.kodlar.doldur(kod_satir)

    def _secili(self):
        s = self.tablo.secili()
        if not s:
            self.app.bildirim("Önce bir kullanıcı seçin.", "uyari")
            return None
        return next(u for u in self.kullanicilar if str(u['id']) == s[0])

    def _secim(self):
        s = self.tablo.secili()
        if s:
            u = next(u for u in self.kullanicilar if str(u['id']) == s[0])
            self.btn_aktif.config(text="Devre dışı bırak" if u['aktif'] else "Etkinleştir")

    def _calistir(self, fonk, basari):
        try:
            sonuc = fonk()
        except g.GuvenlikHatasi as e:
            self.app.bildirim(str(e), "hata")
            return None
        self.yenile()
        if basari:
            self.app.bildirim(basari)
        return sonuc if sonuc is not None else True

    def rol(self):
        u = self._secili()
        if not u:
            return
        d = Diyalog(self.app.root, "Rol değiştir", 420, 320, f"{u['kullanici_adi']} kullanıcısının yeni rolünü seçin.")
        a = Alan(d.govde, "Yeni rol", "combo", values=g.atanabilir_roller(self.app.kullanici["rol"]))
        a.pack(fill="x")
        a.set(u['rol'])

        def kaydet():
            if self._calistir(lambda: g.rol_degistir(self.app.kadi, u['id'], a.get()), f"{u['kullanici_adi']} rolü güncellendi."):
                d.destroy()
        Buton(d.alt_ic, "Kaydet", kaydet).pack(side="right")
        Buton(d.alt_ic, "Vazgeç", d.destroy, stil="ikincil").pack(side="right", padx=8)
        d.sigdir(420)

    def aktiflik(self):
        u = self._secili()
        if not u:
            return
        yeni = not u['aktif']
        if not yeni and not self.app.onayla("Hesabı devre dışı bırak", f"{u['kullanici_adi']} artık giriş yapamayacak. Devam edilsin mi?"):
            return
        self._calistir(lambda: g.aktiflik_degistir(self.app.kadi, u['id'], yeni),
                       f"{u['kullanici_adi']} {'etkinleştirildi' if yeni else 'devre dışı bırakıldı'}.")

    def kilit(self):
        u = self._secili()
        if u:
            self._calistir(lambda: g.kilidi_ac(self.app.kadi, u['id']), f"{u['kullanici_adi']} hesabının kilidi açıldı.")

    def sifirla(self):
        u = self._secili()
        if not u or not self.app.onayla("Şifre sıfırla", f"{u['kullanici_adi']} için geçici şifre oluşturulacak. "
                                                         "Kullanıcı ilk girişte şifresini değiştirmek zorunda kalacak. Devam edilsin mi?"):
            return
        gecici = self._calistir(lambda: g.sifre_sifirla(self.app.kadi, u['id']), None)
        if not gecici:
            return
        d = Diyalog(self.app.root, "Geçici şifre", 460, 360,
                    f"{u['kullanici_adi']} için oluşturulan geçici şifre aşağıdadır. Bu şifre yalnızca şimdi gösterilir; "
                    "kullanıcıya güvenli bir kanaldan iletin.")
        kopyalanabilir_kutu(d.govde, gecici).pack(fill="x", ipady=12)
        self._kopyala_butonu(d, gecici)

    def sil(self):
        u = self._secili()
        if not u or not self.app.onayla("Kullanıcıyı sil", f"{u['kullanici_adi']} hesabı kalıcı olarak silinecek. "
                                                           "Geçmiş kayıtlardaki adı korunur. Emin misiniz?", True):
            return
        self._calistir(lambda: g.kullanici_sil(self.app.kadi, u['id']), f"{u['kullanici_adi']} silindi.")

    def _kopyala_butonu(self, d, metin):
        def kopyala():
            self.app.root.clipboard_clear()
            self.app.root.clipboard_append(metin)
            btn.config(text="✓ Kopyalandı")
        btn = Buton(d.alt_ic, "Kopyala", kopyala)
        btn.pack(side="right")
        Buton(d.alt_ic, "Kapat", d.destroy, stil="ikincil").pack(side="right", padx=8)
        d.sigdir()

    def kod_olustur(self):
        rol = self.a_rol.get()
        if not rol:
            self.a_rol.hata("Rol seçin.")
            return
        gun = int(self.a_gun.get().split()[0])
        try:
            kod, son = g.davet_kodu_olustur(self.app.kadi, self.app.kullanici["rol"], rol, gun)
        except g.GuvenlikHatasi as e:
            self.app.bildirim(str(e), "hata")
            return
        self.yenile()
        d = Diyalog(self.app.root, "Davet kodu oluşturuldu", 500, 400,
                    f"Rol: {rol}  •  Son geçerlilik: {m.tarih_tr(son)}\n\nKod yalnızca bir kez kullanılabilir ve "
                    "güvenlik nedeniyle yalnızca şimdi gösterilir. Kopyalayıp kişiye iletin.")
        kopyalanabilir_kutu(d.govde, kod).pack(fill="x", ipady=12)
        self._kopyala_butonu(d, kod)

    def kod_iptal(self):
        s = self.kodlar.secili()
        if not s:
            self.app.bildirim("Önce bir davet kodu seçin.", "uyari")
            return
        g.davet_kodu_iptal(self.app.kadi, int(s[0]))
        self.yenile()
        self.app.bildirim("Davet kodu iptal edildi.")

# =====================================================================================
# SAYFA: DENETİM KAYITLARI
# =====================================================================================

class DenetimSayfasi(tk.Frame):
    baslik = "Denetim Kayıtları"
    alt_baslik = "Tüm kullanıcı işlemlerinin değiştirilemez kaydı"

    def __init__(self, app):
        super().__init__(app.icerik, bg=R["zemin"])
        self.app = app
        self.kayitlar = []
        fk, fi = kart(self, padding=16)
        fk.pack(fill="x", pady=(0, 14))
        self.ara = Girdi(fi, placeholder="Ara: kullanıcı, işlem veya ayrıntı…", width=40)
        self.ara.pack(side="left", ipady=7)
        self.ara.bind("<KeyRelease>", lambda e: self.filtrele(), add="+")
        tk.Label(fi, text="İşlem", font=F(9, True), bg=R["kart"], fg=R["metin2"]).pack(side="left", padx=(20, 8))
        self.islem = ttk.Combobox(fi, state="readonly", width=24, font=F(10))
        self.islem.pack(side="left")
        self.islem.bind("<<ComboboxSelected>>", lambda e: self.filtrele())
        self.lbl_sayi = tk.Label(fi, text="", font=F(9, True), bg=R["kart"], fg=R["soluk"])
        self.lbl_sayi.pack(side="right")

        self.banner = tk.Label(self, text="", font=F(9, True), anchor="w", padx=14, pady=10)

        tk_, ti = kart(self, padding=2)
        tk_.pack(fill="both", expand=True)
        self.tk_ = tk_
        self.tablo = Tablo(ti, [("t", "Zaman", 140, "w"), ("k", "Kullanıcı", 120, "w"), ("i", "İşlem", 170, "w"),
                                ("d", "Ayrıntı", 620, "w")], secim="browse")
        self.tablo.pack(fill="both", expand=True)

        alt = tk.Frame(self, bg=R["zemin"])
        alt.pack(fill="x", pady=(14, 0))
        Buton(alt, "Bütünlüğü doğrula", self.dogrula, stil="ikincil", boyut=9).pack(side="left")
        if app.yetki("log_sil"):
            Buton(alt, "Kayıtları temizle", self.temizle, stil="tehlike_cizgi", boyut=9).pack(side="left", padx=8)
        Buton(alt, "⤓  Excel'e aktar", self.excel, stil="ikincil", boyut=9).pack(side="right")

    def yenile(self):
        self.kayitlar = m.log_kayitlari()
        islemler = sorted({r['islem_tipi'] or "" for r in self.kayitlar})
        self.islem.config(values=["Tümü"] + islemler)
        if not self.islem.get():
            self.islem.set("Tümü")
        self.filtrele()

    def filtrele(self):
        q = m.tr_kucuk(self.ara.get())
        isl = self.islem.get()
        satirlar = []
        for r in self.kayitlar:
            if isl and isl != "Tümü" and r['islem_tipi'] != isl:
                continue
            if q and q not in m.tr_kucuk(f"{r['kullanici']} {r['islem_tipi']} {r['detay']}"):
                continue
            satirlar.append((r['id'], (m.tarih_tr(r['tarih']), r['kullanici'], r['islem_tipi'], r['detay']),
                             (r['tarih'], r['kullanici'], r['islem_tipi'], r['detay'])))
        self.tablo.doldur(satirlar)
        self.lbl_sayi.config(text=f"{len(satirlar)} / {len(self.kayitlar)} kayıt")

    def dogrula(self):
        ok, bozuk = g.log_zinciri_dogrula()
        if ok:
            self.banner.config(text=f"✓  Denetim kayıtlarının bütünlüğü doğrulandı — {len(self.kayitlar)} kayıt, değişiklik tespit edilmedi.",
                               bg=R["basari_acik"], fg=R["basari_koyu"])
        else:
            self.banner.config(text=f"✕  UYARI: #{bozuk} numaralı kayıt ve sonrası uygulama dışında değiştirilmiş veya silinmiş olabilir.",
                               bg=R["tehlike_acik"], fg=R["tehlike_koyu"])
        self.banner.pack(fill="x", pady=(0, 14), before=self.tk_)
        g.log_ekle(self.app.kadi, "BÜTÜNLÜK KONTROLÜ", "Başarılı" if ok else f"Bozulma tespit edildi (#{bozuk})")

    def temizle(self):
        if not self.app.onayla("Denetim kayıtlarını temizle", "Tüm denetim kayıtları silinecek. Silme işleminin kendisi yeni bir kayıt "
                                                               "olarak tutulacaktır.\n\nÖnce Excel'e aktarmanız önerilir. Devam edilsin mi?", True):
            return
        n = g.loglari_temizle(self.app.kadi)
        self.yenile()
        self.app.bildirim(f"{n} denetim kaydı silindi.")

    def excel(self):
        self.app.excel_kaydet(self.kayitlar, [("tarih", "Zaman"), ("kullanici", "Kullanıcı"), ("islem_tipi", "İşlem"),
                                              ("detay", "Ayrıntı")], "Denetim_Kayitlari", "Denetim kayıtları")


if __name__ == "__main__":
    Uygulama().calistir()
