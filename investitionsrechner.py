#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Investitionsrechner – Wirtschaftlichkeits- und Investitionsrechnung
====================================================================

Grafisches Programm (tkinter + matplotlib) zum Vergleich von bis zu drei
Investitionsalternativen mit allen gängigen Verfahren der Investitionsrechnung:

  Statische Verfahren
    * Kostenvergleichsrechnung (inkl. kritischer Menge)
    * Gewinnvergleichsrechnung (inkl. Break-even-Menge)
    * Rentabilitätsrechnung (ROI)
    * Statische Amortisationsrechnung (Durchschnitts- und Kumulationsmethode)

  Dynamische Verfahren
    * Kapitalwertmethode (NPV, Kapitalwertrate, Kapitalendwert)
    * Interne-Zinsfuß-Methode (IRR, MIRR, Fisher-Zins der Differenzinvestition)
    * Annuitätenmethode
    * Dynamische Amortisationsrechnung
    * Vermögensendwertmethode (gespaltener Soll-/Habenzins, Kontenausgleich)

  Wertorientiert / Risiko
    * Economic Value Added (EVA) und Market Value Added (MVA)
    * Sensitivitätsanalyse mit kritischen Werten (Tornado- und Spinnendiagramm)

Export: PDF (alle Texte und Diagramme), TXT (Bericht), CSV (Kennzahlen und
Zahlungsreihen, Excel-kompatibel mit Semikolon und Dezimalkomma).

Voraussetzungen: Python 3.8+ und matplotlib  ->  pip install matplotlib
Start:           python investitionsrechner.py
"""
from __future__ import annotations

import csv
import datetime
import math
import re
import sys
import textwrap
import warnings
from dataclasses import dataclass, field
from itertools import combinations

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from matplotlib.figure import Figure
from matplotlib.ticker import FuncFormatter
from matplotlib.backends.backend_pdf import PdfPages

APP_TITEL = "Investitionsrechner – Wirtschaftlichkeits- und Investitionsrechnung"
VERSION = "1.1"
FARBEN = ["#1f77b4", "#ff7f0e", "#2ca02c"]
BREITE = 96          # Zeichenbreite des Textberichts
MAX_JAHRE = 60


# =============================================================================
# Zahlen lesen und formatieren (deutsches Format)
# =============================================================================
def parse_num(text, name="Wert"):
    """Liest '1.234,56', '1234.56', '1234,5', '120.000', '6 %', '5 €'."""
    s = str(text).strip()
    for z in ("€", "%", " ", "\u00a0", "'"):
        s = s.replace(z, "")
    if s == "":
        raise ValueError(f"„{name}“ ist leer.")
    if "," in s:                                   # deutsches Format
        s = s.replace(".", "").replace(",", ".")
    elif re.fullmatch(r"-?[1-9]\d{0,2}(\.\d{3})+", s):   # 120.000 = Tausenderpunkt
        s = s.replace(".", "")
    try:
        return float(s)
    except ValueError:
        raise ValueError(f"„{name}“: '{text}' ist keine gültige Zahl.") from None


def ungueltig(v):
    if v is None:
        return True
    try:
        return math.isnan(v) or math.isinf(v)
    except TypeError:
        return True


def de(v, d=2):
    if ungueltig(v):
        return "–"
    s = f"{v:,.{d}f}".replace(",", "§").replace(".", ",").replace("§", ".")
    if s.startswith("-") and set(s[1:]) <= set("0,."):
        s = s[1:]
    return s


def eur(v, d=2):
    return "–" if ungueltig(v) else de(v, d) + " €"


def pct(v, d=2):
    return "–" if ungueltig(v) else de(v * 100, d) + " %"


def jahre(v, d=2):
    return "–" if ungueltig(v) else de(v, d) + " J."


def eingabe_fmt(v):
    """Zahl für ein Eingabefeld (ohne Tausenderpunkte, Dezimalkomma)."""
    if v is None:
        return ""
    s = f"{float(v):.6f}".rstrip("0").rstrip(".")
    return s.replace(".", ",")


def csvn(v, d=4):
    return "" if ungueltig(v) else f"{v:.{d}f}".replace(".", ",")


def mittel(xs):
    xs = list(xs)
    return sum(xs) / len(xs) if xs else 0.0


# =============================================================================
# Datenmodell
# =============================================================================
@dataclass
class Parameter:
    projekt: str = "Investitionsprojekt"
    i: float = 0.06          # Kalkulationszins
    r_re: float = 0.06       # Reinvestitionszins (MIRR)
    h: float = 0.02          # Habenzins (Vermögensendwert)
    s_soll: float = 0.08     # Sollzins (Vermögensendwert)
    ek: float = 0.0          # verfügbares Eigenkapital (Vermögensendwert)
    t_max: float = 5.0       # maximal akzeptierte Amortisationsdauer
    steuer: float = 0.30     # Ertragsteuersatz (EVA)
    wacc_modus: str = "kalk"  # "kalk" = WACC entspricht Kalkulationszins, sonst berechnet
    ek_quote: float = 0.40
    r_ek: float = 0.10
    r_fk: float = 0.05
    sens: float = 0.20       # Variation in der Sensitivitätsanalyse

    @property
    def wacc(self):
        if self.wacc_modus == "kalk":
            return self.i
        return self.ek_quote * self.r_ek + (1 - self.ek_quote) * self.r_fk * (1 - self.steuer)


@dataclass
class Alternative:
    name: str = "Alternative"
    A0: float = 0.0           # Anschaffungsauszahlung
    foerderung: float = 0.0   # Zuschuss / Förderung in t = 0
    n: int = 10               # Nutzungsdauer (Jahre)
    L: float = 0.0            # Restwert / Liquidationserlös am Ende
    WC: float = 0.0           # gebundenes Umlaufvermögen (Working Capital)
    einheit: str = "Stück"
    x: float = 0.0            # Menge / Leistung pro Jahr
    p: float = 0.0            # Preis / Erlös / Einsparung je Einheit
    kv: float = 0.0           # variable Kosten je Einheit
    kfix: float = 0.0         # fixe Betriebskosten pro Jahr (ohne AfA, ohne Zinsen)
    g_e: float = 0.0          # Steigerung Einzahlungen p.a.
    g_a: float = 0.0          # Steigerung Auszahlungen p.a.
    modus: str = "abgeleitet"  # "abgeleitet" oder "manuell"
    man_E: list = field(default_factory=list)
    man_A: list = field(default_factory=list)

    @property
    def I0(self):
        return self.A0 - self.foerderung


def pruefe(a: Alternative):
    f = []
    if not (1 <= a.n <= MAX_JAHRE):
        f.append(f"Nutzungsdauer muss zwischen 1 und {MAX_JAHRE} Jahren liegen.")
    if a.A0 < 0:
        f.append("Anschaffungsauszahlung darf nicht negativ sein.")
    if a.foerderung < 0 or a.foerderung > a.A0:
        f.append("Förderung muss zwischen 0 und der Anschaffungsauszahlung liegen.")
    if a.L < 0 or a.WC < 0:
        f.append("Restwert und Working Capital dürfen nicht negativ sein.")
    if a.x < 0:
        f.append("Menge darf nicht negativ sein.")
    if a.g_e <= -1 or a.g_a <= -1:
        f.append("Steigerungsraten müssen größer als −100 % sein.")
    if a.modus == "manuell" and (len(a.man_E) != a.n or len(a.man_A) != a.n):
        f.append("Die manuelle Zahlungsreihe passt nicht zur Nutzungsdauer – "
                 "bitte Tabelle neu erzeugen.")
    return [f"{a.name}: {t}" for t in f]


# =============================================================================
# Finanzmathematik
# =============================================================================
def npv(cf, i):
    return sum(c * (1 + i) ** -t for t, c in enumerate(cf))


def anf(i, n):
    """Annuitäten- bzw. Kapitalwiedergewinnungsfaktor."""
    if n <= 0:
        return float("nan")
    if abs(i) < 1e-12:
        return 1.0 / n
    q = (1 + i) ** n
    return i * q / (q - 1)


def irr_alle(cf):
    """Alle internen Zinsfüße im Bereich −99 % … 1000 % (Rasterung + Bisektion)."""
    if not cf or not (any(c > 0 for c in cf) and any(c < 0 for c in cf)):
        return []
    raster = [-0.99 + k * 0.0025 for k in range(int(1.99 / 0.0025) + 1)]
    raster += [1.0 + k * 0.05 for k in range(1, 181)]
    wurzeln = []
    r0, v0 = raster[0], npv(cf, raster[0])
    for r1 in raster[1:]:
        v1 = npv(cf, r1)
        if v0 == 0:
            wurzeln.append(r0)
        elif v0 * v1 < 0:
            lo, hi, flo = r0, r1, v0
            for _ in range(200):
                mid = (lo + hi) / 2
                fm = npv(cf, mid)
                if flo * fm <= 0:
                    hi = mid
                else:
                    lo, flo = mid, fm
                if hi - lo < 1e-13:
                    break
            wurzeln.append((lo + hi) / 2)
        r0, v0 = r1, v1
    if v0 == 0:
        wurzeln.append(r0)
    erg = []
    for w in sorted(wurzeln):
        if not erg or abs(w - erg[-1]) > 1e-7:
            erg.append(w)
    return erg


def mirr(cf, finanz, reinvest):
    n = len(cf) - 1
    if n < 1:
        return None
    pv_neg = sum(c * (1 + finanz) ** -t for t, c in enumerate(cf) if c < 0)
    fv_pos = sum(c * (1 + reinvest) ** (n - t) for t, c in enumerate(cf) if c > 0)
    if pv_neg >= 0 or fv_pos <= 0:
        return None
    return (fv_pos / -pv_neg) ** (1 / n) - 1


def amortisation(werte):
    """Erster Zeitpunkt, an dem die kumulierte Reihe ≥ 0 wird (lineare Interpolation)."""
    kum, s = [], 0.0
    for v in werte:
        s += v
        kum.append(s)
    if kum and kum[0] >= 0:
        return 0.0, kum
    for t in range(1, len(kum)):
        if kum[t - 1] < 0 <= kum[t]:
            return t - 1 + (-kum[t - 1]) / (kum[t] - kum[t - 1]), kum
    return None, kum


# =============================================================================
# Zahlungsreihe und Analyse
# =============================================================================
def zahlungsreihe(a: Alternative, f=None):
    """Erzeugt die Zahlungsreihe. f: Faktoren für die Sensitivitätsanalyse."""
    f = f or {}
    n = a.n
    I0 = a.I0 * f.get("inv", 1.0)
    L = a.L * f.get("rest", 1.0)
    E = [0.0] * (n + 1)
    fix = [0.0] * (n + 1)
    var = [0.0] * (n + 1)
    Aus = [0.0] * (n + 1)
    R = [0.0] * (n + 1)
    if a.modus == "abgeleitet":
        x = a.x * f.get("menge", 1.0)
        for t in range(1, n + 1):
            ge = (1 + a.g_e) ** (t - 1)
            ga = (1 + a.g_a) ** (t - 1)
            E[t] = a.p * f.get("preis", 1.0) * x * ge
            var[t] = a.kv * f.get("kv", 1.0) * x * ga
            fix[t] = a.kfix * f.get("kfix", 1.0) * ga
            Aus[t] = var[t] + fix[t]
    else:
        for t in range(1, n + 1):
            E[t] = a.man_E[t - 1] * f.get("E", 1.0)
            Aus[t] = a.man_A[t - 1] * f.get("A", 1.0)
    R[n] = L + a.WC
    CF = [-(I0 + a.WC)] + [E[t] - Aus[t] + R[t] for t in range(1, n + 1)]
    return dict(n=n, I0=I0, L=L, E=E, Aus=Aus, fix=fix, var=var, R=R, CF=CF)


def analysiere(a: Alternative, g: Parameter) -> dict:
    z = zahlungsreihe(a)
    n, i = a.n, g.i
    I0, L, WC = z["I0"], z["L"], a.WC
    E, Aus, CF = z["E"], z["Aus"], z["CF"]
    r = dict(alt=a, z=z, n=n, I0=I0, L=L, WC=WC, Ikap=I0 + WC)
    r["nur_kosten"] = sum(E[1:]) == 0

    # ---------------- statische Verfahren (Ø-Periode) ----------------
    Em, KB = mittel(E[1:]), mittel(Aus[1:])
    D = (I0 - L) / n
    Kap = (I0 + L) / 2 + WC
    Z = Kap * i
    K = KB + D + Z
    G = Em - K
    r.update(Em=Em, KB=KB, D=D, Kap=Kap, Z=Z, K=K, G=G)
    r["ROI"] = (G + Z) / Kap if Kap > 0 else None
    r["k_stueck"] = K / a.x if a.x > 0 else None
    r["abgeleitet"] = a.modus == "abgeleitet"
    if r["abgeleitet"]:
        fe = mittel((1 + a.g_e) ** (t - 1) for t in range(1, n + 1))
        fa = mittel((1 + a.g_a) ** (t - 1) for t in range(1, n + 1))
        r["fixm"] = a.kfix * fa
        r["varm"] = a.kv * a.x * fa
        r["kv_m"] = a.kv * fa
        r["p_m"] = a.p * fe
        r["Kf"] = r["fixm"] + D + Z
        r["db"] = r["p_m"] - r["kv_m"]
        r["BE"] = r["Kf"] / r["db"] if r["db"] > 0 else None
    else:
        for k in ("fixm", "varm", "kv_m", "p_m", "Kf", "db", "BE"):
            r[k] = None
    r["rueck"] = Em - KB
    r["t_avg"] = (I0 + WC) / r["rueck"] if r["rueck"] > 0 else None
    r["t_kum"], r["kum"] = amortisation(CF)

    # ---------------- dynamische Verfahren ----------------
    df = [(1 + i) ** -t for t in range(n + 1)]
    pv = [c * d for c, d in zip(CF, df)]
    r["df"], r["pv"] = df, pv
    r["NPV"] = sum(pv)
    r["t_dyn"], r["kum_pv"] = amortisation(pv)
    r["KWR"] = r["NPV"] / r["Ikap"] if r["Ikap"] > 0 else None
    r["KEW"] = r["NPV"] * (1 + i) ** n
    r["irrs"] = irr_alle(CF)
    r["IRR"] = r["irrs"][0] if len(r["irrs"]) == 1 else None
    r["MIRR"] = mirr(CF, i, g.r_re)
    f = anf(i, n)
    r["ANF"] = f
    r["PV_E"] = sum(E[t] * df[t] for t in range(1, n + 1))
    r["PV_A"] = sum(Aus[t] * df[t] for t in range(1, n + 1))
    r["PV_R"] = z["R"][n] * df[n]
    r["AN"] = r["NPV"] * f
    r["AN_E"] = r["PV_E"] * f
    r["AN_A"] = r["PV_A"] * f
    r["KD"] = (r["Ikap"] - r["PV_R"]) * f

    # Vermögensendwert mit Kontenausgleichsgebot
    V = [g.ek + CF[0]]
    Vz = [0.0]
    for t in range(1, n + 1):
        vor = V[-1]
        zs = g.h if vor >= 0 else g.s_soll
        Vz.append(vor * zs)
        V.append(vor * (1 + zs) + CF[t])
    r["V"], r["Vz"] = V, Vz
    r["VEW"] = V[-1]
    r["VEW_ref"] = g.ek * (1 + g.h) ** n
    r["VEW_diff"] = r["VEW"] - r["VEW_ref"]

    # Economic Value Added
    w, s = g.wacc, g.steuer
    cap, ebit, nopat, kk, eva = [], [], [], [], []
    for t in range(1, n + 1):
        C = I0 - (t - 1) * D + WC
        eb = E[t] - Aus[t] - D
        no = eb * (1 - s)
        cap.append(C)
        ebit.append(eb)
        nopat.append(no)
        kk.append(w * C)
        eva.append(no - w * C)
    r.update(eva_cap=cap, eva_ebit=ebit, eva_nopat=nopat, eva_kk=kk, eva=eva, wacc=w)
    r["MVA"] = sum(eva[t - 1] * (1 + w) ** -t for t in range(1, n + 1))
    cf_st = [-(I0 + WC)] + [(E[t] - Aus[t]) * (1 - s) + s * D + z["R"][t] for t in range(1, n + 1)]
    r["cf_st"] = cf_st
    r["NPV_st"] = npv(cf_st, w)
    return r


def sens_parameter(a: Alternative, g: Parameter):
    if a.modus == "abgeleitet":
        p = [("inv", "Investition (netto)", a.I0, "€"),
             ("menge", "Menge / Leistung", a.x, a.einheit),
             ("preis", "Preis/Erlös je Einheit", a.p, "€"),
             ("kv", "Variable Kosten je Einh.", a.kv, "€"),
             ("kfix", "Fixkosten p.a.", a.kfix, "€"),
             ("rest", "Restwert", a.L, "€")]
    else:
        p = [("inv", "Investition (netto)", a.I0, "€"),
             ("E", "Einzahlungen (alle J.)", 1.0, "×"),
             ("A", "Lfd. Auszahlungen (alle J.)", 1.0, "×"),
             ("rest", "Restwert", a.L, "€")]
    p.append(("zins", "Kalkulationszins", g.i, "%"))
    return [q for q in p if abs(q[2]) > 1e-12]


def npv_variiert(a, g, key, faktor):
    if key == "zins":
        return npv(zahlungsreihe(a)["CF"], g.i * faktor)
    return npv(zahlungsreihe(a, {key: faktor})["CF"], g.i)


def sensitivitaet(a: Alternative, g: Parameter, r: dict):
    d = g.sens
    zeilen = []
    stufen = [-d, -d / 2, 0.0, d / 2, d]
    for key, label, basis, einheit in sens_parameter(a, g):
        lo = npv_variiert(a, g, key, 1 - d)
        hi = npv_variiert(a, g, key, 1 + d)
        spinne = [npv_variiert(a, g, key, 1 + s) for s in stufen]
        if key == "zins":
            krit = r["IRR"]
            fk = krit / g.i if krit is not None and abs(g.i) > 1e-12 else None
        else:
            n0 = npv_variiert(a, g, key, 0.0)
            b = r["NPV"] - n0
            fk = -n0 / b if abs(b) > 1e-9 else None
            krit = basis * fk if fk is not None else None
        zeilen.append(dict(key=key, label=label, basis=basis, einheit=einheit,
                           lo=lo, hi=hi, fk=fk, krit=krit, spinne=spinne))
    return dict(stufen=stufen, zeilen=zeilen)


class Ergebnis:
    """Alle Berechnungen für einen Satz an Eingaben."""

    def __init__(self, alts, g: Parameter):
        self.alts = alts
        self.g = g
        self.rs = [analysiere(a, g) for a in alts]
        self.namen = [a.name for a in alts]
        self.sens = [sensitivitaet(a, g, r) for a, r in zip(alts, self.rs)]
        self.paare = []
        for ia, ib in combinations(range(len(alts)), 2):
            ra, rb = self.rs[ia], self.rs[ib]
            d = dict(a=ia, b=ib, xkrit=None)
            if ra["abgeleitet"] and rb["abgeleitet"] and abs(ra["kv_m"] - rb["kv_m"]) > 1e-12:
                xk = (ra["Kf"] - rb["Kf"]) / (rb["kv_m"] - ra["kv_m"])
                d["xkrit"] = xk if xk > 0 else None
            m = max(ra["n"], rb["n"]) + 1
            ca = ra["z"]["CF"] + [0.0] * (m - len(ra["z"]["CF"]))
            cb = rb["z"]["CF"] + [0.0] * (m - len(rb["z"]["CF"]))
            diff = [y - x for x, y in zip(ca, cb)]
            d["diff"] = diff
            d["fisher"] = irr_alle(diff)
            d["npv_diff"] = npv(diff, g.i)
            self.paare.append(d)
        self.erstellt = datetime.datetime.now()


# =============================================================================
# Textbericht
# =============================================================================
def wrap(t, einzug=""):
    return textwrap.fill(t, BREITE, initial_indent=einzug, subsequent_indent=einzug)


def tabelle(zeilen, namen, w0=38, w=19):
    namen = [n if len(n) < w else n[:w - 2] + "…" for n in namen]
    out = [f"{'':<{w0}}" + "".join(f"{n:>{w}}" for n in namen),
           "-" * (w0 + w * len(namen))]
    for z in zeilen:
        if z is None:
            out.append("")
            continue
        label, *werte = z
        out.append(f"{label[:w0 - 1]:<{w0}}" + "".join(f"{v:>{w}}" for v in werte))
    return out


def kopf(e: Ergebnis, titel):
    return ["=" * BREITE, titel.upper(), "=" * BREITE,
            f"Projekt: {e.g.projekt}",
            f"Kalkulationszins: {pct(e.g.i)}    Alternativen: {', '.join(e.namen)}",
            f"Erstellt: {e.erstellt:%d.%m.%Y %H:%M}", ""]


def beste(e, key, maximal=True):
    kand = [(r[key], k) for k, r in enumerate(e.rs) if not ungueltig(r.get(key))]
    if not kand:
        return None
    return (max if maximal else min)(kand)[1]


def hinweis_nur_kosten(e):
    out = []
    if any(r["Ikap"] <= 0 for r in e.rs):
        out += ["", wrap("HINWEIS: Mindestens eine Alternative hat keine Anfangsauszahlung (Nettoinvestition 0 €). "
                         "Kapitalwert, Annuität und Vermögensendwert bleiben aussagekräftig; Rentabilität, "
                         "Kapitalwertrate, interner Zinsfuß und Amortisation sind ohne Kapitaleinsatz nicht definiert.")]
    return out + _hinweis_kosten(e)


def _hinweis_kosten(e):
    if any(r["nur_kosten"] for r in e.rs):
        return ["", wrap("HINWEIS: Für mindestens eine Alternative sind keine Einzahlungen erfasst "
                         "(reine Kostenbetrachtung). Gewinn, Rentabilität, Amortisation und "
                         "interner Zinsfuß sind dann nicht aussagekräftig. Verwenden Sie die "
                         "Kostenvergleichsrechnung bzw. Kapitalwert/Annuität (höchster = geringster "
                         "Kostenbarwert) oder erfassen Sie Einsparungen als Einzahlungen.")]
    return []


def eingaben_text(e: Ergebnis):
    g = e.g
    L = ["=" * BREITE, "EINGABEDATEN", "=" * BREITE,
         f"Projekt: {g.projekt}", f"Erstellt: {e.erstellt:%d.%m.%Y %H:%M}", "",
         "Allgemeine Parameter",
         f"  Kalkulationszins i ............... {pct(g.i)}",
         f"  Reinvestitionszins (MIRR) ........ {pct(g.r_re)}",
         f"  Habenzins / Sollzins (VEW) ....... {pct(g.h)} / {pct(g.s_soll)}",
         f"  Eigenkapital (VEW) ............... {eur(g.ek)}",
         f"  Max. Amortisationsdauer .......... {de(g.t_max)} Jahre",
         f"  Ertragsteuersatz (EVA) ........... {pct(g.steuer)}",
         f"  WACC (EVA) ....................... {pct(g.wacc)}"
         + ("  (= Kalkulationszins)" if g.wacc_modus == "kalk" else
            f"  (EK-Quote {pct(g.ek_quote)}, r_EK {pct(g.r_ek)}, r_FK {pct(g.r_fk)})"),
         f"  Sensitivität ..................... ± {pct(g.sens, 0)}", ""]
    zeilen = []
    for lab, fn in [
        ("Anschaffungsauszahlung", lambda a: eur(a.A0)),
        ("Förderung / Zuschuss", lambda a: eur(a.foerderung)),
        ("Nettoinvestition I₀", lambda a: eur(a.I0)),
        ("Nutzungsdauer n", lambda a: f"{a.n} Jahre"),
        ("Restwert L", lambda a: eur(a.L)),
        ("Working Capital", lambda a: eur(a.WC)),
        ("Zahlungsreihe", lambda a: "aus Mengen/Preisen" if a.modus == "abgeleitet" else "manuell"),
        ("Menge / Leistung p.a.", lambda a: f"{de(a.x, 0)} {a.einheit}"),
        ("Preis / Erlös je Einheit", lambda a: eur(a.p, 4) if a.modus == "abgeleitet" else "–"),
        ("Variable Kosten je Einheit", lambda a: eur(a.kv, 4) if a.modus == "abgeleitet" else "–"),
        ("Fixe Betriebskosten p.a.", lambda a: eur(a.kfix) if a.modus == "abgeleitet" else "–"),
        ("Steigerung Einzahlungen p.a.", lambda a: pct(a.g_e) if a.modus == "abgeleitet" else "–"),
        ("Steigerung Auszahlungen p.a.", lambda a: pct(a.g_a) if a.modus == "abgeleitet" else "–"),
    ]:
        zeilen.append((lab, *[fn(a) for a in e.alts]))
    L += ["Alternativen"] + tabelle(zeilen, e.namen)
    L += ["", "Zahlungsreihen (vor Steuern)"]
    for r in e.rs:
        L += ["", f"  {r['alt'].name}", _jahreskopf(["Einzahl.", "Auszahl.", "Rest+WC", "Cashflow", "kum. CF"])]
        z = r["z"]
        for t in range(r["n"] + 1):
            L.append(_jahreszeile(t, [z["E"][t], z["Aus"][t] if t else r["Ikap"], z["R"][t],
                                      z["CF"][t], r["kum"][t]]))
    return "\n".join(L)


def _jahreskopf(spalten, w=13):
    return f"{'Jahr':>5}" + "".join(f"{s:>{w}}" for s in spalten)


def _jahreszeile(t, werte, w=13, dez=None):
    dez = dez or [0] * len(werte)
    return f"{t:>5}" + "".join(f"{de(v, d):>{w}}" for v, d in zip(werte, dez))


def amort_str(v, n):
    return f"> {n} J. (nicht erreicht)" if v is None else jahre(v)


# ---------------------------------------------------------------- Methoden-Texte
def text_kosten(e: Ergebnis):
    L = kopf(e, "Kostenvergleichsrechnung (statisch)")
    L += [wrap(BESCHREIBUNG["kosten"]), "",
          "Formeln",
          "  Kalk. Abschreibung      D = (I₀ − L) / n",
          "  Ø gebundenes Kapital    Ø K = (I₀ + L) / 2 + Working Capital",
          "  Kalk. Zinsen            Z = Ø K × i",
          "  Gesamtkosten p.a.       K = Ø Betriebskosten (fix + variabel) + D + Z",
          "  Stückkosten             k = K / x",
          "  Kritische Menge         x* = (Kfix_A − Kfix_B) / (kv_B − kv_A),  Kfix = fixe Betriebsk. + D + Z", ""]
    zeilen = [
        ("Nettoinvestition I₀", *[eur(r["I0"]) for r in e.rs]),
        ("Restwert L", *[eur(r["L"]) for r in e.rs]),
        ("Nutzungsdauer n", *[f"{r['n']} J." for r in e.rs]),
        ("Menge / Leistung p.a.", *[f"{de(r['alt'].x, 0)} {r['alt'].einheit}" for r in e.rs]),
        None,
        ("Ø fixe Betriebskosten", *[eur(r["fixm"]) for r in e.rs]),
        ("Ø variable Kosten", *[eur(r["varm"]) for r in e.rs]),
        ("Ø Betriebskosten gesamt", *[eur(r["KB"]) for r in e.rs]),
        ("Kalk. Abschreibung D", *[eur(r["D"]) for r in e.rs]),
        ("Ø gebundenes Kapital", *[eur(r["Kap"]) for r in e.rs]),
        ("Kalk. Zinsen Z", *[eur(r["Z"]) for r in e.rs]),
        ("GESAMTKOSTEN p.a.", *[eur(r["K"]) for r in e.rs]),
        ("Stückkosten", *[eur(r["k_stueck"], 4) for r in e.rs]),
        ("Fixkosten inkl. D und Z", *[eur(r["Kf"]) for r in e.rs]),
        ("Ø variable Kosten je Einheit", *[eur(r["kv_m"], 4) for r in e.rs]),
    ]
    L += tabelle(zeilen, e.namen)
    L += ["", "Ergebnis"]
    b = beste(e, "K", maximal=False)
    if len(e.rs) > 1:
        L.append(f"  Geringste Gesamtkosten p.a.: {e.namen[b]} ({eur(e.rs[b]['K'])})")
        mengen = {round(r["alt"].x, 6) for r in e.rs}
        if len(mengen) > 1 and all(r["k_stueck"] is not None for r in e.rs):
            bs = beste(e, "k_stueck", maximal=False)
            L.append(f"  Geringste Stückkosten: {e.namen[bs]} ({eur(e.rs[bs]['k_stueck'], 4)})")
            L.append(wrap("Die Mengen sind unterschiedlich – maßgeblich ist der Vergleich der "
                          "Stückkosten (Kostenvergleich pro Leistungseinheit).", "  "))
        for p in e.paare:
            na, nb = e.namen[p["a"]], e.namen[p["b"]]
            if p["xkrit"] is not None:
                ra = e.rs[p["a"]]
                unter = na if ra["Kf"] < e.rs[p["b"]]["Kf"] else nb
                ueber = nb if unter == na else na
                L.append(wrap(f"Kritische Menge {na} / {nb}: {de(p['xkrit'], 0)} "
                              f"{ra['alt'].einheit} p.a. – darunter ist {unter} günstiger, "
                              f"darüber {ueber}.", "  "))
            elif e.rs[p["a"]]["abgeleitet"] and e.rs[p["b"]]["abgeleitet"]:
                L.append(wrap(f"{na} / {nb}: keine positive kritische Menge (eine Alternative ist "
                              f"bei jeder Menge günstiger).", "  "))
    else:
        L.append(f"  Gesamtkosten p.a.: {eur(e.rs[0]['K'])}. Ein Kostenvergleich benötigt "
                 f"mindestens zwei Alternativen.")
    L += ["", "Grenzen der Methode",
          wrap("Statisch (Ø-Periode, keine zeitliche Verteilung), Erlöse werden ignoriert – nur "
               "sinnvoll bei gleicher Leistung bzw. Ersatz-/Rationalisierungsentscheidungen. "
               "Aussage zur absoluten Vorteilhaftigkeit ist nicht möglich.", "  ")]
    return "\n".join(L + hinweis_nur_kosten(e))


def text_gewinn(e: Ergebnis):
    L = kopf(e, "Gewinnvergleichsrechnung (statisch)")
    L += [wrap(BESCHREIBUNG["gewinn"]), "",
          "Formeln",
          "  Ø Gewinn p.a.           G = Ø Erlöse − (Ø Betriebskosten + D + Z)",
          "  Deckungsbeitrag/Einheit db = p − kv",
          "  Break-even-Menge        x_BE = (fixe Betriebskosten + D + Z) / db", ""]
    zeilen = [
        ("Ø Erlöse / Einzahlungen p.a.", *[eur(r["Em"]) for r in e.rs]),
        ("Ø Gesamtkosten p.a.", *[eur(r["K"]) for r in e.rs]),
        ("Ø GEWINN p.a.", *[eur(r["G"]) for r in e.rs]),
        ("Gewinn je Einheit", *[eur(r["G"] / r["alt"].x, 4) if r["alt"].x > 0 else "–" for r in e.rs]),
        None,
        ("Ø Preis je Einheit", *[eur(r["p_m"], 4) for r in e.rs]),
        ("Ø var. Kosten je Einheit", *[eur(r["kv_m"], 4) for r in e.rs]),
        ("Deckungsbeitrag je Einheit", *[eur(r["db"], 4) for r in e.rs]),
        ("Fixkosten inkl. D und Z", *[eur(r["Kf"]) for r in e.rs]),
        ("Break-even-Menge p.a.", *[f"{de(r['BE'], 0)} {r['alt'].einheit}" if r["BE"] else "–" for r in e.rs]),
        ("Auslastung des Break-even", *[pct(r["BE"] / r["alt"].x) if r["BE"] and r["alt"].x > 0 else "–" for r in e.rs]),
        ("Sicherheitsabstand", *[pct(1 - r["BE"] / r["alt"].x) if r["BE"] and r["alt"].x > 0 else "–" for r in e.rs]),
    ]
    L += tabelle(zeilen, e.namen) + ["", "Ergebnis"]
    for r in e.rs:
        urteil = "vorteilhaft (Gewinn > 0)" if r["G"] > 0 else "nicht vorteilhaft (Gewinn ≤ 0)"
        L.append(f"  {r['alt'].name}: {urteil}")
    if len(e.rs) > 1:
        b = beste(e, "G")
        L.append(f"  Höchster Ø Gewinn: {e.namen[b]} ({eur(e.rs[b]['G'])})")
    L += ["", "Grenzen der Methode",
          wrap("Statisch, Kapitaleinsatz und Nutzungsdauer gehen nur über D und Z ein; "
               "unterschiedliche Kapitalbindung wird beim Vergleich nicht vollständig "
               "berücksichtigt (→ Rentabilitätsrechnung, Kapitalwertmethode). Die "
               "Break-even-Menge unterstellt lineare Kosten- und Erlösverläufe.", "  ")]
    return "\n".join(L + hinweis_nur_kosten(e))


def text_rentabilitaet(e: Ergebnis):
    L = kopf(e, "Rentabilitätsrechnung / Return on Investment (statisch)")
    L += [wrap(BESCHREIBUNG["rentabilitaet"]), "",
          "Formel",
          "  Rentabilität R = (Ø Gewinn + kalk. Zinsen) / Ø gebundenes Kapital",
          "  Ø gebundenes Kapital = (I₀ + L) / 2 + Working Capital",
          "  Vorteilhaft, wenn R ≥ Kalkulationszins i", ""]
    zeilen = [
        ("Ø Gewinn p.a.", *[eur(r["G"]) for r in e.rs]),
        ("+ Kalk. Zinsen", *[eur(r["Z"]) for r in e.rs]),
        ("= Gewinn vor Zinsen", *[eur(r["G"] + r["Z"]) for r in e.rs]),
        ("Ø gebundenes Kapital", *[eur(r["Kap"]) for r in e.rs]),
        ("RENTABILITÄT R", *[pct(r["ROI"]) for r in e.rs]),
        ("Mindestverzinsung i", *[pct(e.g.i) for _ in e.rs]),
        ("Überrendite R − i", *[pct(r["ROI"] - e.g.i) if r["ROI"] is not None else "–" for r in e.rs]),
    ]
    L += tabelle(zeilen, e.namen) + ["", "Ergebnis"]
    for r in e.rs:
        ok = r["ROI"] is not None and r["ROI"] >= e.g.i
        L.append(f"  {r['alt'].name}: {'vorteilhaft' if ok else 'nicht vorteilhaft'} "
                 f"(R = {pct(r['ROI'])} {'≥' if ok else '<'} i = {pct(e.g.i)})")
    if len(e.rs) > 1:
        b = beste(e, "ROI")
        L.append(f"  Höchste Rentabilität: {e.namen[b]}")
    L += ["", "Grenzen der Methode",
          wrap("Statisch; relative Kennzahl ohne Aussage über die absolute Höhe des Gewinns. "
               "Beim Vergleich unterschiedlich großer Investitionen wird implizit unterstellt, "
               "dass Differenzbeträge zur gleichen Rendite angelegt werden.", "  ")]
    return "\n".join(L + hinweis_nur_kosten(e))


def text_amort_stat(e: Ergebnis):
    L = kopf(e, "Statische Amortisationsrechnung (Pay-back-Methode)")
    L += [wrap(BESCHREIBUNG["amort_stat"]), "",
          "Formeln",
          "  Durchschnittsmethode:  t_A = Kapitaleinsatz (I₀ + WC) / Ø Rückfluss",
          "                         Ø Rückfluss = Ø Einzahlungen − Ø lfd. Auszahlungen (= G + D + Z)",
          "  Kumulationsmethode:    Jahr, in dem die Summe der (unverzinsten) Cashflows ≥ 0 wird",
          "                         (innerhalb des Jahres linear interpoliert; Restwert erst im Jahr n)",
          f"  Vorteilhaft, wenn t_A ≤ maximal akzeptierte Dauer ({de(e.g.t_max)} Jahre)", ""]
    zeilen = [
        ("Kapitaleinsatz I₀ + WC", *[eur(r["Ikap"]) for r in e.rs]),
        ("Ø Rückfluss p.a.", *[eur(r["rueck"]) for r in e.rs]),
        ("Amortisation Ø-Methode", *[amort_str(r["t_avg"], r["n"]) for r in e.rs]),
        ("Amortisation Kumulation", *[amort_str(r["t_kum"], r["n"]) for r in e.rs]),
        ("Nutzungsdauer", *[f"{r['n']} J." for r in e.rs]),
    ]
    L += tabelle(zeilen, e.namen) + ["", "Ergebnis"]
    for r in e.rs:
        t = r["t_kum"]
        ok = t is not None and t <= e.g.t_max
        L.append(f"  {r['alt'].name}: {'vorteilhaft' if ok else 'nicht vorteilhaft'} "
                 f"(Kumulation {amort_str(t, r['n'])}, Grenze {de(e.g.t_max)} J.)")
    if len(e.rs) > 1:
        b = beste(e, "t_kum", maximal=False)
        if b is not None:
            L.append(f"  Kürzeste Amortisationsdauer: {e.namen[b]}")
    for r in e.rs:
        L += ["", f"  Kumulierte Cashflows – {r['alt'].name}", _jahreskopf(["Cashflow", "kumuliert"])]
        for t in range(r["n"] + 1):
            L.append(_jahreszeile(t, [r["z"]["CF"][t], r["kum"][t]]))
    L += ["", "Grenzen der Methode",
          wrap("Risiko-/Liquiditätskennzahl, kein Rentabilitätsmaß: Zahlungen nach dem "
               "Amortisationszeitpunkt und der Zeitwert des Geldes bleiben unberücksichtigt "
               "(→ dynamische Amortisationsrechnung).", "  ")]
    return "\n".join(L + hinweis_nur_kosten(e))


def _dyn_tabelle(r):
    z = r["z"]
    out = [_jahreskopf(["Einzahl.", "Auszahl.", "Rest+WC", "Cashflow", "Abzinsf.", "Barwert", "kum. BW"], 13)]
    for t in range(r["n"] + 1):
        out.append(_jahreszeile(t, [z["E"][t], z["Aus"][t] if t else r["Ikap"], z["R"][t], z["CF"][t],
                                    r["df"][t], r["pv"][t], r["kum_pv"][t]], 13, [0, 0, 0, 0, 4, 0, 0]))
    return out


def _hinweis_laufzeit(e):
    if len({r["n"] for r in e.rs}) > 1:
        return ["", wrap("HINWEIS: Die Nutzungsdauern sind unterschiedlich. Der Kapitalwertvergleich "
                         "unterstellt dann Ergänzungsinvestitionen zum Kalkulationszins (Kapitalwert 0). "
                         "Bei identischer Wiederholung der Investitionen ist die Annuitätenmethode "
                         "das geeignete Vergleichsmaß.")]
    return []


def text_kapitalwert(e: Ergebnis):
    L = kopf(e, "Kapitalwertmethode (Net Present Value)")
    L += [wrap(BESCHREIBUNG["kapitalwert"]), "",
          "Formeln",
          "  C₀ = −(I₀ + WC) + Σ (Eₜ − Aₜ)·(1+i)^−t + (L + WC)·(1+i)^−n",
          "  Kapitalwertrate  KWR = C₀ / (I₀ + WC)       Kapitalendwert  Cₙ = C₀·(1+i)^n",
          "  Vorteilhaft, wenn C₀ ≥ 0; beim Vergleich ist der höchste Kapitalwert vorzuziehen.", ""]
    zeilen = [
        ("Barwert Einzahlungen", *[eur(r["PV_E"]) for r in e.rs]),
        ("Barwert lfd. Auszahlungen", *[eur(-r["PV_A"]) for r in e.rs]),
        ("Barwert Restwert + WC", *[eur(r["PV_R"]) for r in e.rs]),
        ("Investition (t = 0)", *[eur(-r["Ikap"]) for r in e.rs]),
        ("KAPITALWERT C₀", *[eur(r["NPV"]) for r in e.rs]),
        ("Kapitalwertrate", *[pct(r["KWR"]) for r in e.rs]),
        ("Kapitalendwert Cₙ", *[eur(r["KEW"]) for r in e.rs]),
    ]
    L += tabelle(zeilen, e.namen) + ["", "Ergebnis"]
    for r in e.rs:
        L.append(f"  {r['alt'].name}: {'vorteilhaft' if r['NPV'] >= 0 else 'nicht vorteilhaft'} "
                 f"(C₀ = {eur(r['NPV'])})")
    if len(e.rs) > 1:
        b = beste(e, "NPV")
        L.append(f"  Höchster Kapitalwert: {e.namen[b]}")
        bk = beste(e, "KWR")
        if bk != b:
            L.append(f"  Höchste Kapitalwertrate (bei Kapitalknappheit relevant): {e.namen[bk]}")
    for r in e.rs:
        L += ["", f"  Zahlungsreihe und Barwerte – {r['alt'].name}"] + _dyn_tabelle(r)
    L += _hinweis_laufzeit(e)
    L += ["", "Grenzen der Methode",
          wrap("Unterstellt einen vollkommenen Kapitalmarkt (Anlage und Kreditaufnahme zum "
               "Kalkulationszins). Ergebnis reagiert empfindlich auf den Zinssatz "
               "(→ Sensitivitätsanalyse). Werte vor Steuern.", "  ")]
    return "\n".join(L + hinweis_nur_kosten(e))


def text_izf(e: Ergebnis):
    L = kopf(e, "Interne-Zinsfuß-Methode (IRR) und modifizierter interner Zinsfuß (MIRR)")
    L += [wrap(BESCHREIBUNG["izf"]), "",
          "Formeln",
          "  IRR:  r mit  Σ CFₜ·(1+r)^−t = 0   (numerisch bestimmt, alle Lösungen werden gesucht)",
          "  MIRR: ( Σ CFₜ⁺·(1+r_re)^(n−t) / |Σ CFₜ⁻·(1+i)^−t| )^(1/n) − 1",
          f"        Reinvestitionszins r_re = {pct(e.g.r_re)}, Finanzierungszins = i = {pct(e.g.i)}",
          "  Vorteilhaft, wenn IRR bzw. MIRR ≥ Kalkulationszins.", ""]

    def irr_txt(r):
        if not r["irrs"]:
            return "keiner"
        if len(r["irrs"]) > 1:
            return "mehrdeutig"
        return pct(r["IRR"])

    zeilen = [
        ("Interner Zinsfuß (IRR)", *[irr_txt(r) for r in e.rs]),
        ("Modifizierter IZF (MIRR)", *[pct(r["MIRR"]) for r in e.rs]),
        ("Kalkulationszins i", *[pct(e.g.i) for _ in e.rs]),
        ("Kapitalwert bei i", *[eur(r["NPV"]) for r in e.rs]),
    ]
    L += tabelle(zeilen, e.namen) + ["", "Ergebnis"]
    for r in e.rs:
        if len(r["irrs"]) > 1:
            L.append(wrap(f"{r['alt'].name}: Zahlungsreihe mit mehreren Vorzeichenwechseln – "
                          f"mehrere interne Zinsfüße ({', '.join(pct(x) for x in r['irrs'])}). "
                          f"Der IRR ist hier nicht interpretierbar; maßgeblich sind MIRR bzw. Kapitalwert.", "  "))
        elif not r["irrs"]:
            L.append(f"  {r['alt'].name}: kein interner Zinsfuß (kein Vorzeichenwechsel der Zahlungsreihe).")
        else:
            ok = r["IRR"] >= e.g.i
            L.append(f"  {r['alt'].name}: IRR {pct(r['IRR'])} {'≥' if ok else '<'} i → "
                     f"{'vorteilhaft' if ok else 'nicht vorteilhaft'}")
    if e.paare:
        L += ["", "Differenzinvestition / Fisher-Zins (Schnittpunkt der Kapitalwertkurven)"]
        for p in e.paare:
            na, nb = e.namen[p["a"]], e.namen[p["b"]]
            besser = nb if p["npv_diff"] > 0 else na
            if p["fisher"]:
                fz = ", ".join(pct(x) for x in p["fisher"])
                L.append(wrap(f"{nb} − {na}: Fisher-Zins {fz}. Beim Kalkulationszins von "
                              f"{pct(e.g.i)} ist {besser} vorteilhafter (Kapitalwertdifferenz "
                              f"{eur(abs(p['npv_diff']))}).", "  "))
            else:
                L.append(wrap(f"{nb} − {na}: kein Schnittpunkt – {besser} ist bei jedem "
                              f"Zinssatz vorteilhafter.", "  "))
        L.append(wrap("Achtung: Die Rangfolge nach IRR kann von der nach Kapitalwert abweichen. "
                      "Für die Auswahl zwischen Alternativen gilt der Kapitalwert bzw. der IRR der "
                      "Differenzinvestition (Fisher-Zins).", "  "))
    L += ["", "Grenzen der Methode",
          wrap("Der IRR unterstellt Wiederanlage der Rückflüsse zum IRR selbst; der MIRR korrigiert "
               "dies. Bei nicht-konventionellen Zahlungsreihen kann es keinen oder mehrere IRR geben.", "  ")]
    return "\n".join(L + hinweis_nur_kosten(e))


def text_annuitaet(e: Ergebnis):
    L = kopf(e, "Annuitätenmethode")
    L += [wrap(BESCHREIBUNG["annuitaet"]), "",
          "Formeln",
          "  Annuitätenfaktor ANF = i·(1+i)^n / ((1+i)^n − 1)   (Kapitalwiedergewinnungsfaktor)",
          "  Annuität AN = C₀ · ANF",
          "  Zerlegung: AN = Einzahlungsannuität − Auszahlungsannuität − Kapitaldienst",
          "  Kapitaldienst KD = ((I₀ + WC) − (L + WC)·(1+i)^−n) · ANF", ""]
    zeilen = [
        ("Annuitätenfaktor ANF", *[de(r["ANF"], 6) for r in e.rs]),
        ("Einzahlungsannuität", *[eur(r["AN_E"]) for r in e.rs]),
        ("− Auszahlungsannuität", *[eur(r["AN_A"]) for r in e.rs]),
        ("− Kapitaldienst", *[eur(r["KD"]) for r in e.rs]),
        ("= ANNUITÄT p.a.", *[eur(r["AN"]) for r in e.rs]),
        ("Kontrolle C₀ · ANF", *[eur(r["NPV"] * r["ANF"]) for r in e.rs]),
    ]
    L += tabelle(zeilen, e.namen) + ["", "Ergebnis"]
    for r in e.rs:
        if r["nur_kosten"]:
            L.append(f"  {r['alt'].name}: Kostenannuität {eur(-r['AN'])} p.a. (äquivalente jährliche Kosten)")
        else:
            L.append(f"  {r['alt'].name}: {'vorteilhaft' if r['AN'] >= 0 else 'nicht vorteilhaft'} "
                     f"(AN = {eur(r['AN'])} p.a.)")
    if len(e.rs) > 1:
        b = beste(e, "AN")
        L.append(f"  Höchste Annuität: {e.namen[b]}")
        if len({r["n"] for r in e.rs}) > 1:
            L.append(wrap("Bei unterschiedlichen Nutzungsdauern ist der Annuitätenvergleich korrekt, "
                          "wenn identische Wiederholung (Investitionskette) unterstellt wird.", "  "))
    L += ["", "Grenzen der Methode",
          wrap("Gleiche Annahmen wie die Kapitalwertmethode; die Annuität ist der in jedem Jahr "
               "gleich hohe Überschuss, der dem Kapitalwert entspricht.", "  ")]
    return "\n".join(L + hinweis_nur_kosten(e))


def text_amort_dyn(e: Ergebnis):
    L = kopf(e, "Dynamische Amortisationsrechnung")
    L += [wrap(BESCHREIBUNG["amort_dyn"]), "",
          "Formel",
          "  t_dyn = erstes Jahr, in dem Σ CFₜ·(1+i)^−t ≥ 0  (innerhalb des Jahres linear interpoliert)",
          f"  Vorteilhaft, wenn t_dyn ≤ maximal akzeptierte Dauer ({de(e.g.t_max)} J.)", ""]
    zeilen = [
        ("Amortisation statisch", *[amort_str(r["t_kum"], r["n"]) for r in e.rs]),
        ("Amortisation dynamisch", *[amort_str(r["t_dyn"], r["n"]) for r in e.rs]),
        ("Nutzungsdauer", *[f"{r['n']} J." for r in e.rs]),
        ("Kapitalwert", *[eur(r["NPV"]) for r in e.rs]),
    ]
    L += tabelle(zeilen, e.namen) + ["", "Ergebnis"]
    for r in e.rs:
        t = r["t_dyn"]
        ok = t is not None and t <= e.g.t_max
        L.append(f"  {r['alt'].name}: {'vorteilhaft' if ok else 'nicht vorteilhaft'} "
                 f"(dynamisch {amort_str(t, r['n'])})")
    for r in e.rs:
        L += ["", f"  Kumulierte Barwerte – {r['alt'].name}", _jahreskopf(["Cashflow", "Barwert", "kum. BW", "kum. CF"])]
        for t in range(r["n"] + 1):
            L.append(_jahreszeile(t, [r["z"]["CF"][t], r["pv"][t], r["kum_pv"][t], r["kum"][t]]))
    L += ["", "Grenzen der Methode",
          wrap("Berücksichtigt den Zeitwert des Geldes, aber keine Zahlungen nach dem "
               "Amortisationszeitpunkt. Ein Projekt, das sich innerhalb der Nutzungsdauer "
               "dynamisch amortisiert, hat (bei konventioneller Zahlungsreihe) einen positiven Kapitalwert.", "  ")]
    return "\n".join(L + hinweis_nur_kosten(e))


def text_vew(e: Ergebnis):
    g = e.g
    L = kopf(e, "Vermögensendwertmethode (gespaltener Soll-/Habenzins)")
    L += [wrap(BESCHREIBUNG["vew"]), "",
          "Vorgehen (Kontenausgleichsgebot)",
          f"  Startkonto V₀ = Eigenkapital ({eur(g.ek)}) − (I₀ + WC)",
          f"  Vₜ = Vₜ₋₁ · (1 + z) + CFₜ   mit z = Habenzins {pct(g.h)} bei Guthaben, Sollzins {pct(g.s_soll)} bei Schulden",
          f"  Vergleich mit der Unterlassungsalternative: Eigenkapital·(1 + Habenzins)^n", ""]
    zeilen = [
        ("Vermögensendwert Vₙ", *[eur(r["VEW"]) for r in e.rs]),
        ("Unterlassungsalternative", *[eur(r["VEW_ref"]) for r in e.rs]),
        ("VORTEIL der Investition", *[eur(r["VEW_diff"]) for r in e.rs]),
        ("Nutzungsdauer", *[f"{r['n']} J." for r in e.rs]),
        ("Zum Vergleich: Kapitalendwert", *[eur(r["KEW"]) for r in e.rs]),
    ]
    L += tabelle(zeilen, e.namen) + ["", "Ergebnis"]
    for r in e.rs:
        L.append(f"  {r['alt'].name}: {'vorteilhaft' if r['VEW_diff'] >= 0 else 'nicht vorteilhaft'} "
                 f"(Mehrvermögen {eur(r['VEW_diff'])})")
    if len(e.rs) > 1:
        b = beste(e, "VEW_diff")
        L.append(f"  Höchster Vermögenszuwachs: {e.namen[b]}")
        if len({r["n"] for r in e.rs}) > 1:
            L.append(wrap("Achtung: Unterschiedliche Nutzungsdauern – Endwerte beziehen sich auf "
                          "verschiedene Zeitpunkte; ggf. auf den gleichen Planungshorizont auffüllen.", "  "))
    for r in e.rs:
        L += ["", f"  Vermögenskonto – {r['alt'].name}", _jahreskopf(["Stand Beginn", "Zinsen", "Cashflow", "Stand Ende"], 15)]
        V, Vz, CF = r["V"], r["Vz"], r["z"]["CF"]
        L.append(f"{0:>5}" + f"{de(g.ek, 0):>15}{'':>15}{de(CF[0], 0):>15}{de(V[0], 0):>15}")
        for t in range(1, r["n"] + 1):
            L.append(_jahreszeile(t, [V[t - 1], Vz[t], CF[t], V[t]], 15))
    L += ["", "Grenzen der Methode",
          wrap("Realistischer als die Kapitalwertmethode, da Soll- und Habenzins verschieden sein "
               "dürfen (Grundidee des vollständigen Finanzplans, VOFI). Steuern und "
               "Tilgungspläne werden hier nicht modelliert.", "  ")]
    return "\n".join(L + hinweis_nur_kosten(e))


def text_eva(e: Ergebnis):
    g = e.g
    L = kopf(e, "Economic Value Added (EVA) und Market Value Added (MVA)")
    if g.wacc_modus == "kalk":
        wtxt = f"WACC = Kalkulationszins = {pct(g.wacc)}"
    else:
        wtxt = (f"WACC = EK-Quote·r_EK + (1 − EK-Quote)·r_FK·(1 − s) = {pct(g.ek_quote)}·{pct(g.r_ek)} + "
                f"{pct(1 - g.ek_quote)}·{pct(g.r_fk)}·(1 − {pct(g.steuer)}) = {pct(g.wacc)}")
    L += [wrap(BESCHREIBUNG["eva"]), "",
          "Formeln",
          "  EBIT  = Einzahlungen − lfd. Auszahlungen − Abschreibung D",
          f"  NOPAT = EBIT · (1 − s),  Steuersatz s = {pct(g.steuer)}",
          "  Gebundenes Kapital Cₜ₋₁ = Restbuchwert zu Jahresbeginn + Working Capital",
          "  EVAₜ  = NOPATₜ − WACC · Cₜ₋₁          MVA = Σ EVAₜ·(1+WACC)^−t",
          wrap(wtxt, "  "),
          "  Lücke-Theorem: MVA = Kapitalwert der Cashflows nach Steuern (Kontrollrechnung unten)", ""]
    zeilen = [
        ("WACC", *[pct(r["wacc"]) for r in e.rs]),
        ("EVA im Jahr 1", *[eur(r["eva"][0]) for r in e.rs]),
        ("Ø EVA p.a.", *[eur(mittel(r["eva"])) for r in e.rs]),
        ("Summe EVA (nominal)", *[eur(sum(r["eva"])) for r in e.rs]),
        ("MVA (Barwert der EVA)", *[eur(r["MVA"]) for r in e.rs]),
        ("Kontrolle: NPV nach Steuern", *[eur(r["NPV_st"]) for r in e.rs]),
        ("Zum Vergleich: NPV vor Steuern", *[eur(r["NPV"]) for r in e.rs]),
    ]
    L += tabelle(zeilen, e.namen) + ["", "Ergebnis"]
    for r in e.rs:
        L.append(f"  {r['alt'].name}: {'wertschaffend' if r['MVA'] >= 0 else 'wertvernichtend'} "
                 f"(MVA = {eur(r['MVA'])})")
        neg = [t + 1 for t, v in enumerate(r["eva"]) if v < 0]
        if neg and r["MVA"] >= 0:
            L.append(wrap(f"Hinweis: In den Jahren {', '.join(map(str, neg[:8]))}{' …' if len(neg) > 8 else ''} "
                          f"ist der EVA negativ, obwohl das Projekt insgesamt Wert schafft – typisch "
                          f"für die Anfangsjahre bei linearer Abschreibung.", "    "))
    if len(e.rs) > 1:
        b = beste(e, "MVA")
        L.append(f"  Höchster MVA: {e.namen[b]}")
    for r in e.rs:
        L += ["", f"  EVA-Verlauf – {r['alt'].name}",
              _jahreskopf(["Kapital", "EBIT", "NOPAT", "Kap.-kosten", "EVA", "Barwert EVA"], 14)]
        for t in range(1, r["n"] + 1):
            L.append(_jahreszeile(t, [r["eva_cap"][t - 1], r["eva_ebit"][t - 1], r["eva_nopat"][t - 1],
                                      r["eva_kk"][t - 1], r["eva"][t - 1],
                                      r["eva"][t - 1] * (1 + r["wacc"]) ** -t], 14))
    L += ["", "Grenzen der Methode",
          wrap("Periodenbezogene Kennzahl auf Basis von Buchwerten; einzelne Jahre können "
               "trotz insgesamt positiver Investition negativ sein. Negative EBIT werden mit "
               "Steuerersparnis gerechnet (Verlustverrechnung im Unternehmen unterstellt).", "  ")]
    return "\n".join(L)


def text_sens(e: Ergebnis):
    g = e.g
    L = kopf(e, "Sensitivitätsanalyse und kritische Werte (Basis: Kapitalwert)")
    L += [wrap(BESCHREIBUNG["sens"]), "",
          f"  Jeder Parameter wird einzeln um ± {pct(g.sens, 0)} verändert (ceteris paribus).",
          "  Kritischer Wert = Wert des Parameters, bei dem der Kapitalwert genau 0 wird.",
          "  Δ bis C₀ = 0  = nötige relative Änderung gegenüber dem Basiswert.", ""]
    for a, r, s in zip(e.alts, e.rs, e.sens):
        L += [f"  {a.name}  (Kapitalwert Basis: {eur(r['NPV'])})",
              f"{'Parameter':<28}{'Basiswert':>14}{'C₀ −' + pct(g.sens, 0):>14}{'C₀ +' + pct(g.sens, 0):>14}"
              f"{'krit. Wert':>14}{'Δ bis C₀=0':>12}",
              "-" * BREITE]
        for z in s["zeilen"]:
            if z["key"] == "zins":
                bw, kw = pct(z["basis"]), pct(z["krit"])
            elif z["einheit"] == "×":
                bw, kw = "100 %", pct(z["fk"]) if z["fk"] is not None else "–"
            elif z["einheit"] == "€":
                d = 4 if abs(z["basis"]) < 10 else 0
                bw, kw = de(z["basis"], d), de(z["krit"], d)
            else:
                bw, kw = de(z["basis"], 0), de(z["krit"], 0)
            delta = pct(z["fk"] - 1, 1) if z["fk"] is not None else "–"
            if z["fk"] is not None and z["fk"] < 0:
                delta += "*"
            L.append(f"{z['label']:<28}{bw:>14}{de(z['lo'], 0):>14}{de(z['hi'], 0):>14}{kw:>14}{delta:>12}")
        rang = sorted(s["zeilen"], key=lambda z: -abs(z["hi"] - z["lo"]))
        if rang:
            L.append(f"  Stärkster Einfluss: {rang[0]['label']}"
                     + (f", danach {rang[1]['label']}" if len(rang) > 1 else ""))
        L.append("")
    L.append(wrap("* Negativer kritischer Wert: Der Kapitalwert wird durch Änderung dieses Parameters "
                  "allein im sinnvollen Bereich nicht 0. Beim Zins entspricht der kritische Wert dem "
                  "internen Zinsfuß."))
    return "\n".join(L)


KENNZAHLEN = [
    ("Nettoinvestition inkl. WC", "€", lambda r: r["Ikap"], eur),
    ("Nutzungsdauer", "Jahre", lambda r: r["n"], lambda v: f"{v} J."),
    ("Ø Einzahlungen p.a.", "€", lambda r: r["Em"], eur),
    ("Ø lfd. Auszahlungen p.a.", "€", lambda r: r["KB"], eur),
    ("Ø Gesamtkosten p.a. (statisch)", "€", lambda r: r["K"], eur),
    ("Stückkosten", "€", lambda r: r["k_stueck"], lambda v: eur(v, 4)),
    ("Ø Gewinn p.a.", "€", lambda r: r["G"], eur),
    ("Break-even-Menge", "Einheiten", lambda r: r["BE"], lambda v: de(v, 0)),
    ("Rentabilität (ROI)", "%", lambda r: r["ROI"], pct),
    ("Amortisation statisch (Ø)", "Jahre", lambda r: r["t_avg"], jahre),
    ("Amortisation statisch (kum.)", "Jahre", lambda r: r["t_kum"], jahre),
    ("Kapitalwert C₀", "€", lambda r: r["NPV"], eur),
    ("Kapitalwertrate", "%", lambda r: r["KWR"], pct),
    ("Interner Zinsfuß (IRR)", "%", lambda r: r["IRR"], pct),
    ("Modifizierter IZF (MIRR)", "%", lambda r: r["MIRR"], pct),
    ("Annuität p.a.", "€", lambda r: r["AN"], eur),
    ("Amortisation dynamisch", "Jahre", lambda r: r["t_dyn"], jahre),
    ("Vermögensendwert", "€", lambda r: r["VEW"], eur),
    ("Vorteil ggü. Unterlassung", "€", lambda r: r["VEW_diff"], eur),
    ("Ø EVA p.a.", "€", lambda r: mittel(r["eva"]), eur),
    ("MVA", "€", lambda r: r["MVA"], eur),
]


def text_uebersicht(e: Ergebnis):
    g = e.g
    L = kopf(e, "Gesamtübersicht aller Verfahren")
    L += [wrap(BESCHREIBUNG["uebersicht"]), ""]
    L += tabelle([(lab, *[fmt(fn(r)) for r in e.rs]) for lab, _, fn, fmt in KENNZAHLEN], e.namen)
    L += ["", "Beurteilung", f"{'Verfahren':<34}{'Kriterium':<26}" + "".join(f"{n[:11]:>12}" for n in e.namen),
          "-" * BREITE]

    def zeile(verf, krit, urteile):
        return f"{verf:<34}{krit:<26}" + "".join(f"{u:>12}" for u in urteile)

    def ja(b):
        return "✔" if b else "✘"

    ein = len(e.rs) > 1
    L.append(zeile("Kostenvergleich", "geringste Kosten",
                   ["★" if ein and k == beste(e, "K", False) else "·" for k in range(len(e.rs))]))
    L.append(zeile("Gewinnvergleich", "Gewinn > 0", [ja(r["G"] > 0) for r in e.rs]))
    L.append(zeile("Rentabilität", "R ≥ i", [ja(r["ROI"] is not None and r["ROI"] >= g.i) for r in e.rs]))
    L.append(zeile("Amortisation statisch", f"≤ {de(g.t_max, 1)} J.",
                   [ja(r["t_kum"] is not None and r["t_kum"] <= g.t_max) for r in e.rs]))
    L.append(zeile("Kapitalwert", "C₀ ≥ 0", [ja(r["NPV"] >= 0) for r in e.rs]))
    L.append(zeile("Interner Zinsfuß", "IRR ≥ i",
                   [ja(r["IRR"] >= g.i) if r["IRR"] is not None else "?" for r in e.rs]))
    L.append(zeile("Annuität", "AN ≥ 0", [ja(r["AN"] >= 0) for r in e.rs]))
    L.append(zeile("Amortisation dynamisch", f"≤ {de(g.t_max, 1)} J.",
                   [ja(r["t_dyn"] is not None and r["t_dyn"] <= g.t_max) for r in e.rs]))
    L.append(zeile("Vermögensendwert", "Vorteil ≥ 0", [ja(r["VEW_diff"] >= 0) for r in e.rs]))
    L.append(zeile("EVA / MVA", "MVA ≥ 0", [ja(r["MVA"] >= 0) for r in e.rs]))
    L.append("  ✔ erfüllt   ✘ nicht erfüllt   ? nicht eindeutig   ★ beste Alternative")
    if ein:
        L += ["", "Rangfolge"]
        for lab, key, mx in [("Kapitalwert", "NPV", True), ("Annuität", "AN", True),
                             ("Vermögensendwert-Vorteil", "VEW_diff", True), ("MVA", "MVA", True),
                             ("Ø Gewinn", "G", True), ("Rentabilität", "ROI", True),
                             ("Amortisation dyn.", "t_dyn", False)]:
            b = beste(e, key, mx)
            L.append(f"  {lab:<28} → {e.namen[b] if b is not None else '–'}")
        b = beste(e, "NPV")
        L += ["", wrap(f"Empfehlung (theoretisch maßgeblich: Kapitalwert bzw. bei unterschiedlichen "
                       f"Laufzeiten die Annuität): {e.namen[b]}. Statische Verfahren dienen als "
                       f"Plausibilitäts- und Risikoindikatoren.")]
    else:
        r = e.rs[0]
        L += ["", wrap(f"Gesamturteil: Die Investition ist nach der Kapitalwertmethode "
                       f"{'vorteilhaft' if r['NPV'] >= 0 else 'nicht vorteilhaft'} "
                       f"(C₀ = {eur(r['NPV'])}).")]
    L += _hinweis_laufzeit(e) + hinweis_nur_kosten(e)
    return "\n".join(L)


# =============================================================================
# Diagramme
# =============================================================================
def _euro(ax, achse="y"):
    fm = FuncFormatter(lambda v, _: de(v, 0))
    (ax.yaxis if achse == "y" else ax.xaxis).set_major_formatter(fm)


def _prozent(ax, achse="y"):
    fm = FuncFormatter(lambda v, _: de(v * 100, 1) + " %")
    (ax.yaxis if achse == "y" else ax.xaxis).set_major_formatter(fm)


def _gitter(ax):
    ax.grid(True, alpha=0.3)
    ax.set_axisbelow(True)


def _werte(ax, balken, fmt):
    for b in balken:
        h = b.get_height()
        ax.annotate(fmt(h), (b.get_x() + b.get_width() / 2, h), ha="center",
                    va="bottom" if h >= 0 else "top", fontsize=8,
                    xytext=(0, 2 if h >= 0 else -2), textcoords="offset points")


def _kurz(n, k=22):
    return n if len(n) <= k else n[:k - 1] + "…"


def _gruppen(ax, kategorien, serien, namen, fmt=None):
    m = len(serien)
    w = 0.8 / max(m, 1)
    for k, (vals, name) in enumerate(zip(serien, namen)):
        xs = [j + (k - (m - 1) / 2) * w for j in range(len(kategorien))]
        b = ax.bar(xs, [0 if ungueltig(v) else v for v in vals], w, label=_kurz(name), color=FARBEN[k % 3])
        if fmt:
            _werte(ax, b, fmt)
    ax.set_xticks(range(len(kategorien)))
    ax.set_xticklabels(kategorien, fontsize=8)


def plot_kosten(fig, e):
    ax1, ax2 = fig.subplots(1, 2)
    namen = [_kurz(n, 16) for n in e.namen]
    teile = [("Fixe Betriebskosten", lambda r: r["fixm"] or 0, "#9ecae1"),
             ("Variable Kosten", lambda r: r["varm"] or 0, "#4292c6"),
             ("Lfd. Kosten (nicht aufgeteilt)", lambda r: 0 if r["abgeleitet"] else r["KB"], "#6baed6"),
             ("Kalk. Abschreibung", lambda r: r["D"], "#fd8d3c"),
             ("Kalk. Zinsen", lambda r: r["Z"], "#e6550d")]
    unten = [0.0] * len(e.rs)
    for lab, fn, col in teile:
        v = [fn(r) for r in e.rs]
        if any(abs(x) > 0 for x in v):
            ax1.bar(namen, v, bottom=unten, label=lab, color=col)
            unten = [u + x for u, x in zip(unten, v)]
    for k, r in enumerate(e.rs):
        ax1.annotate(eur(r["K"], 0), (k, r["K"]), ha="center", va="bottom", fontsize=9, fontweight="bold")
    ax1.set_title("Ø Gesamtkosten pro Jahr")
    ax1.set_ylabel("€ / Jahr")
    _euro(ax1)
    _gitter(ax1)
    ax1.legend(fontsize=7)
    abg = [k for k, r in enumerate(e.rs) if r["abgeleitet"]]
    if abg:
        xmax = max(max(e.rs[k]["alt"].x for k in abg) * 2, 1)
        for p in e.paare:
            if p["xkrit"]:
                xmax = max(xmax, p["xkrit"] * 1.3)
        xs = [xmax * j / 100 for j in range(101)]
        for k in abg:
            r = e.rs[k]
            ax2.plot(xs, [r["Kf"] + r["kv_m"] * x for x in xs], color=FARBEN[k], label=_kurz(e.namen[k]))
            ax2.plot([r["alt"].x], [r["K"]], "o", color=FARBEN[k])
        for p in e.paare:
            if p["xkrit"]:
                ra = e.rs[p["a"]]
                y = ra["Kf"] + ra["kv_m"] * p["xkrit"]
                ax2.plot([p["xkrit"]], [y], "kD")
                ax2.annotate(f"x* = {de(p['xkrit'], 0)}", (p["xkrit"], y), xytext=(6, -14),
                             textcoords="offset points", fontsize=8)
        ax2.set_title("Kostenfunktionen K(x) und kritische Menge")
        ax2.set_xlabel(f"Menge pro Jahr [{e.rs[abg[0]]['alt'].einheit}]")
        ax2.set_ylabel("€ / Jahr")
        _euro(ax2)
        _euro(ax2, "x")
        ax2.legend(fontsize=8)
    else:
        v = [r["k_stueck"] or 0 for r in e.rs]
        b = ax2.bar(namen, v, color=FARBEN[:len(v)])
        _werte(ax2, b, lambda h: de(h, 2))
        ax2.set_title("Stückkosten")
    _gitter(ax2)


def plot_gewinn(fig, e):
    ax1, ax2 = fig.subplots(1, 2)
    kat = ["Ø Erlöse", "Ø Gesamtkosten", "Ø Gewinn"]
    _gruppen(ax1, kat, [[r["Em"], r["K"], r["G"]] for r in e.rs], e.namen, lambda h: de(h, 0))
    ax1.axhline(0, color="black", lw=0.8)
    ax1.set_title("Ø Erlöse, Kosten und Gewinn pro Jahr")
    _euro(ax1)
    _gitter(ax1)
    ax1.legend(fontsize=8)
    abg = [k for k, r in enumerate(e.rs) if r["abgeleitet"]]
    if abg:
        xmax = 1.0
        for k in abg:
            r = e.rs[k]
            xmax = max(xmax, r["alt"].x * 1.6, (r["BE"] or 0) * 1.3)
        xs = [xmax * j / 100 for j in range(101)]
        for k in abg:
            r = e.rs[k]
            ax2.plot(xs, [r["p_m"] * x for x in xs], color=FARBEN[k], label=f"Erlöse {_kurz(e.namen[k], 14)}")
            ax2.plot(xs, [r["Kf"] + r["kv_m"] * x for x in xs], "--", color=FARBEN[k],
                     label=f"Kosten {_kurz(e.namen[k], 14)}")
            if r["BE"]:
                ax2.plot([r["BE"]], [r["p_m"] * r["BE"]], "kD")
                ax2.annotate(f"BE {de(r['BE'], 0)}", (r["BE"], r["p_m"] * r["BE"]), xytext=(6, -12),
                             textcoords="offset points", fontsize=8)
            ax2.axvline(r["alt"].x, color=FARBEN[k], lw=0.8, alpha=0.5, ls=":")
        ax2.set_title("Break-even-Analyse (Ø-Jahr)")
        ax2.set_xlabel("Menge pro Jahr")
        _euro(ax2)
        _euro(ax2, "x")
        ax2.legend(fontsize=7)
    else:
        ax2.text(0.5, 0.5, "Break-even nur bei abgeleiteter\nZahlungsreihe (Menge × Preis)",
                 ha="center", va="center", transform=ax2.transAxes)
    _gitter(ax2)


def plot_rentabilitaet(fig, e):
    ax = fig.subplots()
    v = [r["ROI"] if r["ROI"] is not None else 0 for r in e.rs]
    b = ax.bar([_kurz(n, 20) for n in e.namen], v, color=FARBEN[:len(v)])
    _werte(ax, b, lambda h: de(h * 100, 2) + " %")
    ax.axhline(e.g.i, color="red", ls="--", label=f"Kalkulationszins {pct(e.g.i)}")
    ax.axhline(0, color="black", lw=0.8)
    ax.set_title("Rentabilität (Gewinn vor Zinsen / Ø gebundenes Kapital)")
    _prozent(ax)
    _gitter(ax)
    ax.legend()


def plot_amort_stat(fig, e):
    ax = fig.subplots()
    for k, r in enumerate(e.rs):
        ax.plot(range(r["n"] + 1), r["kum"], "o-", color=FARBEN[k], label=_kurz(e.namen[k], 30), ms=4)
        if r["t_kum"] is not None:
            ax.axvline(r["t_kum"], color=FARBEN[k], ls=":", lw=1)
            ax.annotate(f"{de(r['t_kum'], 1)} J.", (r["t_kum"], 0), xytext=(4, 8 + 12 * k),
                        textcoords="offset points", color=FARBEN[k], fontsize=9)
    ax.axhline(0, color="black", lw=0.8)
    ax.axvline(e.g.t_max, color="red", ls="--", lw=1, label=f"max. Dauer {de(e.g.t_max, 1)} J.")
    ax.set_title("Statische Amortisation – kumulierte Cashflows (unverzinst)")
    ax.set_xlabel("Jahr")
    ax.set_ylabel("€")
    _euro(ax)
    _gitter(ax)
    ax.legend(fontsize=8)


def plot_kapitalwert(fig, e):
    ax1, ax2 = fig.subplots(1, 2, gridspec_kw={"width_ratios": [2, 1]})
    m = len(e.rs)
    w = 0.8 / m
    for k, r in enumerate(e.rs):
        xs = [t + (k - (m - 1) / 2) * w for t in range(r["n"] + 1)]
        ax1.bar(xs, r["pv"], w, color=FARBEN[k], label=_kurz(e.namen[k]))
    ax1.axhline(0, color="black", lw=0.8)
    ax1.set_title("Barwerte der Zahlungen je Jahr")
    ax1.set_xlabel("Jahr")
    _euro(ax1)
    _gitter(ax1)
    ax1.legend(fontsize=8)
    v = [r["NPV"] for r in e.rs]
    b = ax2.bar([_kurz(n, 14) for n in e.namen], v,
                color=["#2ca02c" if x >= 0 else "#d62728" for x in v])
    _werte(ax2, b, lambda h: de(h, 0) + " €")
    ax2.axhline(0, color="black", lw=0.8)
    ax2.set_title(f"Kapitalwert C₀ (i = {pct(e.g.i)})")
    _euro(ax2)
    _gitter(ax2)
    ax2.tick_params(axis="x", labelsize=8)


def plot_izf(fig, e):
    ax = fig.subplots()
    oben = max([0.3, e.g.i * 2.5] + [min(x * 1.5, 2.0) for r in e.rs for x in r["irrs"] if x > 0])
    unten = min([0.0] + [x * 1.2 for r in e.rs for x in r["irrs"] if x < 0])
    zs = [unten + (oben - unten) * j / 200 for j in range(201)]
    for k, r in enumerate(e.rs):
        ax.plot(zs, [npv(r["z"]["CF"], z) for z in zs], color=FARBEN[k], label=_kurz(e.namen[k], 30))
        for x in r["irrs"]:
            if unten <= x <= oben:
                ax.plot([x], [0], "o", color=FARBEN[k])
                ax.annotate(f"IRR {pct(x, 1)}", (x, 0), xytext=(4, 8 + 12 * k), textcoords="offset points",
                            color=FARBEN[k], fontsize=9)
    for p in e.paare:
        for x in p["fisher"]:
            if unten <= x <= oben:
                y = npv(e.rs[p["a"]]["z"]["CF"], x)
                ax.plot([x], [y], "kD", ms=5)
                ax.annotate(f"Fisher {pct(x, 1)}", (x, y), xytext=(5, 5), textcoords="offset points", fontsize=8)
    ax.axhline(0, color="black", lw=0.8)
    ax.axvline(e.g.i, color="red", ls="--", lw=1, label=f"Kalkulationszins {pct(e.g.i)}")
    ax.set_title("Kapitalwert in Abhängigkeit vom Zinssatz")
    ax.set_xlabel("Zinssatz")
    ax.set_ylabel("Kapitalwert [€]")
    _prozent(ax, "x")
    _euro(ax)
    _gitter(ax)
    ax.legend(fontsize=8)


def plot_annuitaet(fig, e):
    ax = fig.subplots()
    kat = ["Einzahlungs-\nannuität", "Auszahlungs-\nannuität", "Kapital-\ndienst", "Annuität"]
    _gruppen(ax, kat, [[r["AN_E"], -r["AN_A"], -r["KD"], r["AN"]] for r in e.rs], e.namen, lambda h: de(h, 0))
    ax.axhline(0, color="black", lw=0.8)
    ax.set_title(f"Annuitätenmethode – jährliche Äquivalente (i = {pct(e.g.i)})")
    ax.set_ylabel("€ / Jahr")
    _euro(ax)
    _gitter(ax)
    ax.legend(fontsize=8)


def plot_amort_dyn(fig, e):
    ax = fig.subplots()
    for k, r in enumerate(e.rs):
        ax.plot(range(r["n"] + 1), r["kum_pv"], "o-", color=FARBEN[k], ms=4,
                label=f"{_kurz(e.namen[k], 24)} (diskontiert)")
        ax.plot(range(r["n"] + 1), r["kum"], "--", color=FARBEN[k], alpha=0.45,
                label=f"{_kurz(e.namen[k], 24)} (unverzinst)")
        if r["t_dyn"] is not None:
            ax.axvline(r["t_dyn"], color=FARBEN[k], ls=":", lw=1)
            ax.annotate(f"{de(r['t_dyn'], 1)} J.", (r["t_dyn"], 0), xytext=(4, -16 - 12 * k),
                        textcoords="offset points", color=FARBEN[k], fontsize=9)
    ax.axhline(0, color="black", lw=0.8)
    ax.axvline(e.g.t_max, color="red", ls="--", lw=1, label=f"max. Dauer {de(e.g.t_max, 1)} J.")
    ax.set_title("Dynamische Amortisation – kumulierte Barwerte")
    ax.set_xlabel("Jahr")
    ax.set_ylabel("€")
    _euro(ax)
    _gitter(ax)
    ax.legend(fontsize=7)


def plot_vew(fig, e):
    ax = fig.subplots()
    nmax = max(r["n"] for r in e.rs)
    for k, r in enumerate(e.rs):
        ax.plot(range(r["n"] + 1), r["V"], "o-", color=FARBEN[k], ms=4, label=_kurz(e.namen[k], 30))
    ax.plot(range(nmax + 1), [e.g.ek * (1 + e.g.h) ** t for t in range(nmax + 1)], color="grey", ls="--",
            label="Unterlassung (EK zum Habenzins)")
    ax.axhline(0, color="black", lw=0.8)
    ax.set_title(f"Vermögenskonto (Haben {pct(e.g.h)}, Soll {pct(e.g.s_soll)})")
    ax.set_xlabel("Jahr")
    ax.set_ylabel("Kontostand [€]")
    _euro(ax)
    _gitter(ax)
    ax.legend(fontsize=8)


def plot_eva(fig, e):
    ax1, ax2 = fig.subplots(1, 2, gridspec_kw={"width_ratios": [2, 1]})
    m = len(e.rs)
    w = 0.8 / m
    for k, r in enumerate(e.rs):
        xs = [t + (k - (m - 1) / 2) * w for t in range(1, r["n"] + 1)]
        ax1.bar(xs, r["eva"], w, color=FARBEN[k], label=_kurz(e.namen[k]))
    ax1.axhline(0, color="black", lw=0.8)
    ax1.set_title(f"Economic Value Added je Jahr (WACC {pct(e.g.wacc)})")
    ax1.set_xlabel("Jahr")
    _euro(ax1)
    _gitter(ax1)
    ax1.legend(fontsize=8)
    _gruppen(ax2, ["MVA", "NPV n. St.", "NPV v. St."], [[r["MVA"], r["NPV_st"], r["NPV"]] for r in e.rs],
             e.namen, lambda h: de(h, 0))
    ax2.axhline(0, color="black", lw=0.8)
    ax2.set_title("MVA = Kapitalwert nach Steuern")
    _euro(ax2)
    _gitter(ax2)


def plot_sens(fig, e):
    m = len(e.rs)
    axes = fig.subplots(m, 2, squeeze=False)
    d = e.g.sens
    for k, (s, r) in enumerate(zip(e.sens, e.rs)):
        ax, ax2 = axes[k]
        zeilen = sorted(s["zeilen"], key=lambda z: abs(z["hi"] - z["lo"]))
        base = r["NPV"]
        ys = range(len(zeilen))
        ax.barh(ys, [z["lo"] - base for z in zeilen], left=base, color="#d62728", label=f"Parameter −{pct(d, 0)}")
        ax.barh(ys, [z["hi"] - base for z in zeilen], left=base, color="#2ca02c", alpha=0.8,
                label=f"Parameter +{pct(d, 0)}")
        ax.set_yticks(list(ys))
        ax.set_yticklabels([z["label"] for z in zeilen], fontsize=7)
        ax.axvline(base, color="black", lw=1)
        ax.axvline(0, color="grey", lw=0.8, ls="--")
        ax.set_title(f"Tornado – {_kurz(e.namen[k], 30)}", fontsize=9)
        _euro(ax, "x")
        _gitter(ax)
        if k == 0:
            ax.legend(fontsize=7, loc="lower right")
        for z in s["zeilen"]:
            ax2.plot([x * 100 for x in s["stufen"]], z["spinne"], "o-", ms=3, label=z["label"])
        ax2.axhline(0, color="grey", lw=0.8, ls="--")
        ax2.set_title(f"Spinnendiagramm – {_kurz(e.namen[k], 30)}", fontsize=9)
        ax2.set_xlabel("Änderung des Parameters [%]", fontsize=8)
        _euro(ax2)
        _gitter(ax2)
        ax2.legend(fontsize=6)


def plot_uebersicht(fig, e):
    (a1, a2), (a3, a4) = fig.subplots(2, 2)
    namen = [_kurz(n, 14) for n in e.namen]
    v = [r["NPV"] for r in e.rs]
    b = a1.bar(namen, v, color=["#2ca02c" if x >= 0 else "#d62728" for x in v])
    _werte(a1, b, lambda h: de(h, 0))
    a1.axhline(0, color="black", lw=0.8)
    a1.set_title("Kapitalwert [€]", fontsize=10)
    _euro(a1)
    _gruppen(a2, ["IRR", "MIRR", "ROI (statisch)"], [[r["IRR"], r["MIRR"], r["ROI"]] for r in e.rs], e.namen,
             lambda h: de(h * 100, 1) + "%")
    a2.axhline(e.g.i, color="red", ls="--", lw=1, label=f"i = {pct(e.g.i)}")
    a2.set_title("Renditekennzahlen", fontsize=10)
    _prozent(a2)
    a2.legend(fontsize=7)
    _gruppen(a3, ["statisch (Ø)", "statisch (kum.)", "dynamisch"],
             [[r["t_avg"], r["t_kum"], r["t_dyn"]] for r in e.rs], e.namen, lambda h: de(h, 1))
    a3.axhline(e.g.t_max, color="red", ls="--", lw=1, label=f"max. {de(e.g.t_max, 1)} J.")
    a3.set_title("Amortisationsdauer [Jahre] (0 = nicht erreicht)", fontsize=10)
    a3.legend(fontsize=7)
    _gruppen(a4, ["Annuität", "Ø Gewinn", "Ø EVA"], [[r["AN"], r["G"], mittel(r["eva"])] for r in e.rs],
             e.namen, lambda h: de(h, 0))
    a4.axhline(0, color="black", lw=0.8)
    a4.set_title("Jährliche Größen [€/Jahr]", fontsize=10)
    _euro(a4)
    for a in (a1, a2, a3, a4):
        _gitter(a)
        a.tick_params(axis="x", labelsize=8)


def layout(fig):
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        try:
            fig.tight_layout(rect=[0, 0, 1, 0.95])
        except Exception:
            pass


# =============================================================================
# Methodenkatalog
# =============================================================================
BESCHREIBUNG = {
    "uebersicht": "Alle Verfahren auf einen Blick: zentrale Kennzahlen, Erfüllung der jeweiligen "
                  "Vorteilhaftigkeitskriterien und Rangfolge der Alternativen.",
    "kosten": "Vergleicht die durchschnittlichen jährlichen Gesamtkosten (Betriebskosten + kalkulatorische "
              "Abschreibung + kalkulatorische Zinsen) der Alternativen. Geeignet bei gleicher Leistung, "
              "z. B. Ersatz- und Rationalisierungsinvestitionen.",
    "gewinn": "Vergleicht den durchschnittlichen Jahresgewinn (Erlöse − Gesamtkosten) und ermittelt die "
              "Break-even-Menge. Geeignet, wenn sich die Alternativen auch in den Erlösen unterscheiden.",
    "rentabilitaet": "Setzt den Gewinn vor kalkulatorischen Zinsen ins Verhältnis zum durchschnittlich "
                     "gebundenen Kapital (Return on Investment) und vergleicht mit der Mindestverzinsung.",
    "amort_stat": "Ermittelt, nach wie vielen Jahren das eingesetzte Kapital über die Rückflüsse wieder "
                  "zurückgeflossen ist (Durchschnitts- und Kumulationsmethode, ohne Verzinsung).",
    "kapitalwert": "Diskontiert alle Ein- und Auszahlungen mit dem Kalkulationszins auf den Zeitpunkt der "
                   "Investition. Ein positiver Kapitalwert bedeutet Vermögenszuwachs über die "
                   "Mindestverzinsung hinaus.",
    "izf": "Ermittelt den Zinssatz, bei dem der Kapitalwert null ist (effektive Rendite), sowie den "
           "modifizierten internen Zinsfuß mit realistischer Wiederanlageprämisse und den Fisher-Zins "
           "zwischen Alternativen.",
    "annuitaet": "Verteilt den Kapitalwert gleichmäßig auf die Nutzungsdauer: konstanter jährlicher "
                 "Überschuss. Besonders geeignet für Alternativen mit unterschiedlicher Nutzungsdauer.",
    "amort_dyn": "Wie die statische Amortisationsrechnung, aber mit abgezinsten Zahlungen – berücksichtigt "
                 "den Zeitwert des Geldes.",
    "vew": "Ermittelt das Vermögen am Ende der Nutzungsdauer mit getrennten Soll- und Habenzinsen "
           "(Kontenausgleichsgebot) und vergleicht es mit der Anlage des Eigenkapitals ohne Investition.",
    "eva": "Wertorientierte Kennzahl: Geschäftsergebnis nach Steuern (NOPAT) abzüglich der Kosten des "
           "gebundenen Kapitals (WACC). Der Barwert aller EVA (MVA) entspricht dem Kapitalwert nach Steuern.",
    "sens": "Zeigt, wie stark der Kapitalwert auf Änderungen einzelner Eingangsgrößen reagiert, und "
            "ermittelt kritische Werte, bei denen die Investition unvorteilhaft wird.",
}

METHODEN = {
    "uebersicht": ("Gesamtübersicht (alle Verfahren)", text_uebersicht, plot_uebersicht),
    "kosten": ("Statisch │ Kostenvergleichsrechnung", text_kosten, plot_kosten),
    "gewinn": ("Statisch │ Gewinnvergleichsrechnung (inkl. Break-even)", text_gewinn, plot_gewinn),
    "rentabilitaet": ("Statisch │ Rentabilitätsrechnung (ROI)", text_rentabilitaet, plot_rentabilitaet),
    "amort_stat": ("Statisch │ Amortisationsrechnung (Pay-back)", text_amort_stat, plot_amort_stat),
    "kapitalwert": ("Dynamisch │ Kapitalwertmethode (NPV)", text_kapitalwert, plot_kapitalwert),
    "izf": ("Dynamisch │ Interner Zinsfuß (IRR / MIRR)", text_izf, plot_izf),
    "annuitaet": ("Dynamisch │ Annuitätenmethode", text_annuitaet, plot_annuitaet),
    "amort_dyn": ("Dynamisch │ Dynamische Amortisationsrechnung", text_amort_dyn, plot_amort_dyn),
    "vew": ("Dynamisch │ Vermögensendwertmethode (Soll-/Habenzins)", text_vew, plot_vew),
    "eva": ("Wertorientiert │ Economic Value Added (EVA / MVA)", text_eva, plot_eva),
    "sens": ("Risiko │ Sensitivitätsanalyse & kritische Werte", text_sens, plot_sens),
}
REIHENFOLGE = list(METHODEN)


# =============================================================================
# Export
# =============================================================================
def bericht_text(e: Ergebnis, methoden):
    teile = [f"{APP_TITEL} (v{VERSION})", "", eingaben_text(e)]
    for key in methoden:
        teile += ["", "", METHODEN[key][1](e)]
    teile += ["", "-" * BREITE,
              wrap("Hinweis: Alle Werte sind Modellrechnungen auf Basis der Eingaben (vor Steuern, "
                   "außer EVA). Sie ersetzen keine steuerliche oder betriebswirtschaftliche Beratung.")]
    return "\n".join(teile)


def export_txt(pfad, e, methoden):
    with open(pfad, "w", encoding="utf-8") as f:
        f.write(bericht_text(e, methoden))


def _text_seiten(pdf, text, zaehler, zps=86):
    zeilen = text.split("\n")
    for k in range(0, len(zeilen), zps):
        fig = Figure(figsize=(8.27, 11.69))
        fig.text(0.05, 0.975, "\n".join(zeilen[k:k + zps]), family="monospace", fontsize=7.1,
                 va="top", ha="left")
        zaehler[0] += 1
        fig.text(0.95, 0.015, f"Seite {zaehler[0]}", fontsize=7, ha="right", color="grey")
        fig.text(0.05, 0.015, APP_TITEL, fontsize=7, ha="left", color="grey")
        pdf.savefig(fig)


def export_pdf(pfad, e, methoden):
    zaehler = [0]
    with PdfPages(pfad) as pdf:
        _text_seiten(pdf, f"{APP_TITEL} (v{VERSION})\n\n" + eingaben_text(e), zaehler)
        for key in methoden:
            titel, tf, pf = METHODEN[key]
            _text_seiten(pdf, tf(e), zaehler)
            hoehe = 8.27 if key != "sens" else max(8.27, 3.6 * len(e.rs))
            fig = Figure(figsize=(11.69, hoehe))
            pf(fig, e)
            fig.suptitle(f"{titel}  –  {e.g.projekt}", fontsize=12, fontweight="bold")
            layout(fig)
            zaehler[0] += 1
            fig.text(0.98, 0.01, f"Seite {zaehler[0]}", fontsize=7, ha="right", color="grey")
            pdf.savefig(fig)
        info = pdf.infodict()
        info["Title"] = f"Investitionsrechnung – {e.g.projekt}"
        info["Author"] = "Investitionsrechner"
        info["CreationDate"] = datetime.datetime.now()


def export_csv(pfad, e):
    g = e.g
    with open(pfad, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["Investitionsrechnung", g.projekt, e.erstellt.strftime("%d.%m.%Y %H:%M")])
        w.writerow([])
        w.writerow(["ALLGEMEINE PARAMETER", "Wert", "Einheit"])
        for lab, v, u in [("Kalkulationszins", g.i * 100, "%"), ("Reinvestitionszins MIRR", g.r_re * 100, "%"),
                          ("Habenzins", g.h * 100, "%"), ("Sollzins", g.s_soll * 100, "%"),
                          ("Eigenkapital VEW", g.ek, "€"), ("Max. Amortisationsdauer", g.t_max, "Jahre"),
                          ("Ertragsteuersatz", g.steuer * 100, "%"), ("WACC", g.wacc * 100, "%"),
                          ("Sensitivität +/-", g.sens * 100, "%")]:
            w.writerow([lab, csvn(v), u])
        w.writerow([])
        w.writerow(["EINGABEN", "Einheit"] + e.namen)
        for lab, u, fn in [("Anschaffungsauszahlung", "€", lambda a: a.A0),
                           ("Förderung", "€", lambda a: a.foerderung),
                           ("Nutzungsdauer", "Jahre", lambda a: a.n), ("Restwert", "€", lambda a: a.L),
                           ("Working Capital", "€", lambda a: a.WC), ("Menge p.a.", "Einheiten", lambda a: a.x),
                           ("Preis je Einheit", "€", lambda a: a.p), ("Var. Kosten je Einheit", "€", lambda a: a.kv),
                           ("Fixkosten p.a.", "€", lambda a: a.kfix),
                           ("Steigerung Einzahlungen", "%", lambda a: a.g_e * 100),
                           ("Steigerung Auszahlungen", "%", lambda a: a.g_a * 100)]:
            w.writerow([lab, u] + [csvn(fn(a)) for a in e.alts])
        w.writerow(["Zahlungsreihe", ""] + [a.modus for a in e.alts])
        w.writerow([])
        w.writerow(["KENNZAHLEN", "Einheit"] + e.namen)
        for lab, u, fn, _ in KENNZAHLEN:
            werte = []
            for r in e.rs:
                v = fn(r)
                werte.append(csvn(v * 100 if (u == "%" and not ungueltig(v)) else v))
            w.writerow([lab, u] + werte)
        w.writerow([])
        w.writerow(["ZAHLUNGSREIHEN", "Jahr", "Einzahlungen", "lfd. Auszahlungen bzw. Investition",
                    "Restwert+WC", "Cashflow", "kum. Cashflow", "Abzinsfaktor", "Barwert", "kum. Barwert",
                    "Vermögenskonto", "EVA"])
        for r in e.rs:
            z = r["z"]
            for t in range(r["n"] + 1):
                w.writerow([r["alt"].name, t, csvn(z["E"][t]), csvn(z["Aus"][t] if t else r["Ikap"]),
                            csvn(z["R"][t]), csvn(z["CF"][t]), csvn(r["kum"][t]), csvn(r["df"][t], 8),
                            csvn(r["pv"][t]), csvn(r["kum_pv"][t]), csvn(r["V"][t]),
                            csvn(r["eva"][t - 1]) if t else ""])
        w.writerow([])
        w.writerow(["SENSITIVITÄT", "Parameter", "Basiswert", f"NPV -{csvn(g.sens * 100, 0)}%",
                    f"NPV +{csvn(g.sens * 100, 0)}%", "kritischer Wert", "Änderung bis NPV=0 [%]"])
        for a, s in zip(e.alts, e.sens):
            for z in s["zeilen"]:
                bw = z["basis"] * 100 if z["key"] == "zins" else z["basis"]
                kw = (z["krit"] * 100 if z["key"] == "zins" and z["krit"] is not None else z["krit"])
                w.writerow([a.name, z["label"], csvn(bw), csvn(z["lo"]), csvn(z["hi"]), csvn(kw),
                            csvn((z["fk"] - 1) * 100) if z["fk"] is not None else ""])


# =============================================================================
# Vorlagen / Richtwerte
# =============================================================================
ZINS_VORLAGEN = [
    ("– Richtwert wählen –", None),
    ("2 % – sehr risikoarm / öffentliche Hand", 2),
    ("4 % – konservativ, z. B. Energie-/Gebäudeinvestitionen", 4),
    ("6 % – typischer Mischzins Mittelstand", 6),
    ("8 % – Unternehmen mit erhöhtem Risiko", 8),
    ("12 % – risikoreiche / innovative Projekte", 12),
]

ND_VORLAGEN = [
    ("– Richtwert wählen –", None),
    ("Pkw: 6 Jahre", 6),
    ("Lkw: 9 Jahre", 9),
    ("Büromöbel: 13 Jahre", 13),
    ("Photovoltaikanlage: 20 Jahre", 20),
    ("Heizungsanlage (wirtschaftlich): 20 Jahre", 20),
    ("Betriebsgebäude (3 % AfA): 33 Jahre", 33),
]

_G_STD = dict(projekt="Investitionsprojekt", i=6, r_re=6, h=2, s_soll=8, ek=0, t_max=5, steuer=30,
              wacc_modus="kalk", ek_quote=40, r_ek=10, r_fk=5, sens=20)
_A_STD = dict(name="Alternative", A0=0, foerderung=0, n=10, L=0, WC=0, einheit="Stück", x=0, p=0, kv=0,
              kfix=0, g_e=0, g_a=0, modus="abgeleitet")


def _g(**kw):
    d = dict(_G_STD)
    d.update(kw)
    return d


def _a(**kw):
    d = dict(_A_STD)
    d.update(kw)
    return d


VORLAGEN = {
    "Beispiel: Maschinenvergleich (Industrie)": dict(
        glob=_g(projekt="Ersatz einer Fertigungsmaschine", i=6, r_re=6, ek=50000, t_max=5),
        alts=[_a(name="Maschine A (Standard)", A0=120000, n=8, L=10000, x=10000, p=12, kv=6.5, kfix=25000),
              _a(name="Maschine B (automatisiert)", A0=180000, n=8, L=20000, x=10000, p=12, kv=5.0, kfix=28000),
              None]),
    "Beispiel: Heizung – Gas-Brennwert vs. Wärmepumpe": dict(
        glob=_g(projekt="Heizungserneuerung (reine Kostenbetrachtung)", i=4, r_re=4, t_max=10, sens=20),
        alts=[_a(name="Gas-Brennwert", A0=12000, n=20, einheit="kWh Wärme", x=20000, p=0, kv=0.13, kfix=350,
                 g_a=3),
              _a(name="Wärmepumpe", A0=32000, foerderung=9600, n=20, einheit="kWh Wärme", x=20000, p=0, kv=0.09,
                 kfix=250, g_a=2),
              None]),
    "Beispiel: Photovoltaik – ohne vs. mit Speicher": dict(
        glob=_g(projekt="PV-Anlage 30 kWp (Eigenverbrauch + Einspeisung)", i=4, r_re=4, ek=20000, t_max=12),
        alts=[_a(name="PV 30 kWp", A0=39000, n=20, einheit="kWh", x=28500, p=0.17, kfix=600, g_e=1, g_a=2),
              _a(name="PV 30 kWp + Speicher", A0=55000, n=20, einheit="kWh", x=28500, p=0.21, kfix=800, g_e=1,
                 g_a=2),
              None]),
    "Beispiel: Einsparinvestition – LED-Umrüstung": dict(
        glob=_g(projekt="LED-Umrüstung Verwaltungsgebäude", i=3, r_re=3, t_max=5),
        alts=[_a(name="LED-Umrüstung", A0=25000, n=15, einheit="kWh eingespart", x=40000, p=0.28, g_e=2),
              None, None]),
    "Standardwerte (eigene Eingaben)": dict(
        glob=_g(), alts=[_a(name="Alternative 1", A0=10000, x=1000, p=10, kv=5, kfix=2000), None, None]),
}


# =============================================================================
# GUI
# =============================================================================
class Tooltip:
    def __init__(self, widget, text):
        self.widget, self.text, self.tip = widget, text, None
        widget.bind("<Enter>", self.zeigen, add="+")
        widget.bind("<Leave>", self.verbergen, add="+")

    def zeigen(self, _=None):
        if self.tip or not self.text:
            return
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
        self.tip = tk.Toplevel(self.widget)
        self.tip.wm_overrideredirect(True)
        self.tip.wm_geometry(f"+{x}+{y}")
        tk.Label(self.tip, text=self.text, justify="left", background="#ffffe0", relief="solid",
                 borderwidth=1, wraplength=380, font=("TkDefaultFont", 9)).pack(ipadx=4, ipady=2)

    def verbergen(self, _=None):
        if self.tip:
            self.tip.destroy()
            self.tip = None


class ScrollFrame(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.canvas = tk.Canvas(self, highlightthickness=0)
        sb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.inner = ttk.Frame(self.canvas, padding=6)
        self.inner.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self._win = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfigure(self._win, width=e.width))
        self.canvas.configure(yscrollcommand=sb.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.bind("<Enter>", self._an)
        self.bind("<Leave>", self._aus)

    def _an(self, _):
        self.canvas.bind_all("<MouseWheel>", self._rad)
        self.canvas.bind_all("<Button-4>", lambda e: self.canvas.yview_scroll(-2, "units"))
        self.canvas.bind_all("<Button-5>", lambda e: self.canvas.yview_scroll(2, "units"))

    def _aus(self, _):
        try:
            w = self.winfo_containing(*self.winfo_pointerxy())
            if w is not None and str(w).startswith(str(self)):
                return  # Maus ist nur auf ein Kind-Widget gewechselt
        except (KeyError, tk.TclError):
            pass
        self.canvas.unbind_all("<MouseWheel>")
        self.canvas.unbind_all("<Button-4>")
        self.canvas.unbind_all("<Button-5>")

    def _rad(self, e):
        schritt = -e.delta if sys.platform == "darwin" else -int(e.delta / 120) * 2
        self.canvas.yview_scroll(schritt, "units")


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
        self._Canvas, self._Toolbar = FigureCanvasTkAgg, NavigationToolbar2Tk
        self.title(APP_TITEL)
        self.geometry("1560x940")
        self.minsize(1150, 720)
        try:
            ttk.Style(self).theme_use("clam")
        except tk.TclError:
            pass
        ttk.Style(self).configure("Titel.TLabel", font=("TkDefaultFont", 10, "bold"))
        ttk.Style(self).configure("Hinweis.TLabel", foreground="#555555", font=("TkDefaultFont", 8))
        ttk.Style(self).configure("Akzent.TButton", font=("TkDefaultFont", 10, "bold"))
        self.ergebnis = None
        self.gv = {}
        self.av = [{} for _ in range(3)]
        self.man_zeilen = [[] for _ in range(3)]
        self.man_rahmen = [None] * 3
        self.man_tabelle = [None] * 3
        self._baue_kopf()
        self._baue_hauptbereich()
        self.vorlage_var.set(list(VORLAGEN)[0])
        self.lade_vorlage()
        self.after(200, self._fokus_start)
        self.bind("<F5>", lambda e: self.berechnen())
        self.bind("<Return>", lambda e: self.berechnen())

    def _fokus_start(self):
        """Fenster nach vorn holen und Fokus auf die Verfahrensauswahl setzen."""
        try:
            self.lift()
            self.focus_force()
            self.methode_cb.focus_set()
        except tk.TclError:
            pass

    # ------------------------------------------------------------------ Aufbau
    def _baue_kopf(self):
        kopf = ttk.Frame(self, padding=(8, 6))
        kopf.pack(fill="x")
        ttk.Label(kopf, text="Verfahren:", style="Titel.TLabel").grid(row=0, column=0, sticky="w")
        self.methode_var = tk.StringVar(value=METHODEN["uebersicht"][0])
        cb = ttk.Combobox(kopf, textvariable=self.methode_var, state="readonly", width=52,
                          values=[METHODEN[k][0] for k in REIHENFOLGE])
        self.methode_cb = cb
        cb.grid(row=0, column=1, padx=(4, 16), sticky="w")
        cb.bind("<<ComboboxSelected>>", lambda e: self.berechnen())
        ttk.Label(kopf, text="Vorlage:").grid(row=0, column=2, sticky="w")
        self.vorlage_var = tk.StringVar()
        ttk.Combobox(kopf, textvariable=self.vorlage_var, state="readonly", width=46,
                     values=list(VORLAGEN)).grid(row=0, column=3, padx=4)
        ttk.Button(kopf, text="Vorlage laden", command=self.lade_vorlage).grid(row=0, column=4, padx=(0, 16))
        ttk.Button(kopf, text="▶ Berechnen (F5)", style="Akzent.TButton",
                   command=self.berechnen).grid(row=0, column=5, padx=4)
        ttk.Button(kopf, text="Export PDF", command=lambda: self.exportieren("pdf")).grid(row=0, column=6, padx=2)
        ttk.Button(kopf, text="Export TXT", command=lambda: self.exportieren("txt")).grid(row=0, column=7, padx=2)
        ttk.Button(kopf, text="Export CSV", command=lambda: self.exportieren("csv")).grid(row=0, column=8, padx=2)
        self.alle_var = tk.BooleanVar(value=True)
        c = ttk.Checkbutton(kopf, text="Export: alle Verfahren", variable=self.alle_var)
        c.grid(row=0, column=9, padx=(8, 0))
        Tooltip(c, "Aktiviert: PDF und TXT enthalten alle Verfahren mit allen Diagrammen.\n"
                   "Deaktiviert: nur das aktuell gewählte Verfahren. CSV enthält immer alle Kennzahlen.")
        self.beschr_var = tk.StringVar()
        ttk.Label(kopf, textvariable=self.beschr_var, wraplength=1450, foreground="#333333").grid(
            row=1, column=0, columnspan=10, sticky="w", pady=(6, 0))

    def _baue_hauptbereich(self):
        pw = ttk.PanedWindow(self, orient="horizontal")
        pw.pack(fill="both", expand=True, padx=6, pady=(0, 4))
        links = ttk.Notebook(pw, width=500)
        pw.add(links, weight=0)
        self._tab_allgemein(links)
        for k in range(3):
            self._tab_alternative(links, k)
        rechts = ttk.PanedWindow(pw, orient="vertical")
        pw.add(rechts, weight=1)
        tf = ttk.Frame(rechts)
        self.text = tk.Text(tf, wrap="none", font=("Courier New", 10) if sys.platform == "win32"
                            else "TkFixedFont", height=18, background="#fbfbfb")
        sy = ttk.Scrollbar(tf, orient="vertical", command=self.text.yview)
        sx = ttk.Scrollbar(tf, orient="horizontal", command=self.text.xview)
        self.text.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        sy.pack(side="right", fill="y")
        sx.pack(side="bottom", fill="x")
        self.text.pack(fill="both", expand=True)
        rechts.add(tf, weight=1)
        gf = ttk.Frame(rechts)
        self.fig = Figure(figsize=(9, 5), dpi=100)
        self.canvas = self._Canvas(self.fig, master=gf)
        # matplotlib setzt den Tastaturfokus beim Erzeugen und bei jedem Klick auf das Diagramm.
        # Unter Windows gibt Tk diesen Fokus beim Aktivieren des Fensters an das Diagramm zurück,
        # wodurch sich eine gerade geöffnete Combobox-Liste sofort wieder schließt
        # ("Dropdown öffnet sich erst beim zweiten Mal"). Daher: Diagramm nimmt keinen Fokus.
        tkc = self.canvas.get_tk_widget()
        tkc.focus_set = lambda *a, **k: None
        tkc.configure(takefocus=0)
        tb = self._Toolbar(self.canvas, gf, pack_toolbar=False)
        tb.update()
        tb.pack(side="bottom", fill="x")
        self.canvas.get_tk_widget().pack(fill="both", expand=True)
        rechts.add(gf, weight=2)
        self.status_var = tk.StringVar(value="Bereit.")
        ttk.Label(self, textvariable=self.status_var, relief="sunken", anchor="w",
                  padding=(6, 2)).pack(fill="x", side="bottom")

    def _feld(self, parent, row, label, var, hint="", einheit="", breite=14):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=2)
        ent = ttk.Entry(parent, textvariable=var, width=breite, justify="right")
        ent.grid(row=row, column=1, sticky="ew", padx=4, pady=2)
        ttk.Label(parent, text=einheit).grid(row=row, column=2, sticky="w")
        if hint:
            Tooltip(ent, hint)
            q = ttk.Label(parent, text="ⓘ", foreground="#1f77b4", cursor="question_arrow")
            q.grid(row=row, column=3, sticky="w", padx=(2, 0))
            Tooltip(q, hint)
        return ent

    def _tab_allgemein(self, nb):
        sf = ScrollFrame(nb)
        nb.add(sf, text="Allgemein")
        f = sf.inner
        f.columnconfigure(0, weight=1)
        v = self.gv
        for k in ("projekt", "i", "r_re", "h", "s_soll", "ek", "t_max", "steuer", "ek_quote", "r_ek", "r_fk",
                  "sens"):
            v[k] = tk.StringVar()
        v["wacc_modus"] = tk.StringVar(value="kalk")

        lf = ttk.LabelFrame(f, text="Projekt", padding=6)
        lf.grid(row=0, column=0, sticky="ew", pady=4)
        ttk.Label(lf, text="Bezeichnung").grid(row=0, column=0, sticky="w")
        ttk.Entry(lf, textvariable=v["projekt"], width=40).grid(row=0, column=1, columnspan=3, sticky="ew", padx=4)

        lf = ttk.LabelFrame(f, text="Zinssätze", padding=6)
        lf.grid(row=1, column=0, sticky="ew", pady=4)
        self._feld(lf, 0, "Kalkulationszins i", v["i"],
                   "Mindestverzinsung (Opportunitätskosten) für das eingesetzte Kapital. Grundlage für "
                   "kalkulatorische Zinsen, Kapitalwert, Annuität und dynamische Amortisation.", "%")
        zv = tk.StringVar(value=ZINS_VORLAGEN[0][0])
        cb = ttk.Combobox(lf, textvariable=zv, state="readonly", width=46, values=[z[0] for z in ZINS_VORLAGEN])
        cb.grid(row=1, column=0, columnspan=4, sticky="w", pady=(0, 4))

        def zins_gewaehlt(_):
            wert = dict(ZINS_VORLAGEN)[zv.get()]
            if wert is not None:
                v["i"].set(eingabe_fmt(wert))
                v["r_re"].set(eingabe_fmt(wert))
        cb.bind("<<ComboboxSelected>>", zins_gewaehlt)
        Tooltip(cb, "Grobe Richtwerte – kein Ersatz für unternehmensindividuelle Kapitalkosten.")
        self._feld(lf, 2, "Reinvestitionszins (MIRR)", v["r_re"],
                   "Zinssatz, zu dem Rückflüsse wieder angelegt werden können (für den modifizierten "
                   "internen Zinsfuß). Üblich: gleich dem Kalkulationszins.", "%")
        self._feld(lf, 3, "Habenzins (Vermögensendwert)", v["h"],
                   "Zinssatz für Guthaben auf dem Vermögenskonto.", "%")
        self._feld(lf, 4, "Sollzins (Vermögensendwert)", v["s_soll"],
                   "Zinssatz für Kredite/Schulden auf dem Vermögenskonto.", "%")
        self._feld(lf, 5, "Eigenkapital (Vermögensendwert)", v["ek"],
                   "Verfügbare Eigenmittel zu Beginn. Der Rest der Investition wird zum Sollzins finanziert. "
                   "Vergleichsmaßstab: Anlage dieses Eigenkapitals zum Habenzins.", "€")

        lf = ttk.LabelFrame(f, text="Amortisation", padding=6)
        lf.grid(row=2, column=0, sticky="ew", pady=4)
        self._feld(lf, 0, "Max. akzeptierte Dauer", v["t_max"],
                   "Vorgegebene Soll-Amortisationszeit für die Beurteilung der statischen und dynamischen "
                   "Amortisationsrechnung.", "Jahre")

        lf = ttk.LabelFrame(f, text="Economic Value Added (EVA)", padding=6)
        lf.grid(row=3, column=0, sticky="ew", pady=4)
        self._feld(lf, 0, "Ertragsteuersatz", v["steuer"],
                   "Kapitalgesellschaften in Deutschland: ca. 30 % (KSt + SolZ + GewSt, je nach Hebesatz). "
                   "Für öffentliche Hand / steuerbefreite Einrichtungen: 0 %.", "%")
        ttk.Radiobutton(lf, text="WACC = Kalkulationszins", variable=v["wacc_modus"], value="kalk",
                        command=self._wacc_anzeigen).grid(row=1, column=0, columnspan=3, sticky="w")
        ttk.Radiobutton(lf, text="WACC berechnen aus Kapitalstruktur:", variable=v["wacc_modus"], value="calc",
                        command=self._wacc_anzeigen).grid(row=2, column=0, columnspan=3, sticky="w")
        self._feld(lf, 3, "   Eigenkapitalquote", v["ek_quote"], "Anteil Eigenkapital am Gesamtkapital.", "%")
        self._feld(lf, 4, "   Eigenkapitalkosten r_EK", v["r_ek"],
                   "Renditeforderung der Eigenkapitalgeber (z. B. nach CAPM).", "%")
        self._feld(lf, 5, "   Fremdkapitalzins r_FK", v["r_fk"], "Durchschnittlicher Kreditzins (vor Steuern).",
                   "%")
        self.wacc_lbl = ttk.Label(lf, text="", style="Titel.TLabel")
        self.wacc_lbl.grid(row=6, column=0, columnspan=4, sticky="w", pady=(4, 0))
        for k in ("i", "steuer", "ek_quote", "r_ek", "r_fk"):
            v[k].trace_add("write", lambda *a: self._wacc_anzeigen())

        lf = ttk.LabelFrame(f, text="Sensitivitätsanalyse", padding=6)
        lf.grid(row=4, column=0, sticky="ew", pady=4)
        self._feld(lf, 0, "Variation der Parameter ±", v["sens"],
                   "Um wie viel Prozent jeder Parameter einzeln nach oben und unten verändert wird.", "%")

        ttk.Label(f, style="Hinweis.TLabel", wraplength=440, justify="left", text=(
            "Eingaben: Dezimalkomma oder -punkt; Tausenderpunkte erlaubt (z. B. 120.000). "
            "Prozentwerte als Zahl eingeben (6 = 6 %). Alle Beträge netto in €. "
            "Berechnung vor Steuern (außer EVA). Die Felder mit ⓘ erklären sich per Mauszeiger.")
        ).grid(row=5, column=0, sticky="w", pady=8)

    def _wacc_anzeigen(self):
        try:
            g = self.lese_parameter()
            self.wacc_lbl.configure(text=f"→ WACC = {pct(g.wacc)}")
        except Exception:
            self.wacc_lbl.configure(text="→ WACC = –")

    def _tab_alternative(self, nb, k):
        sf = ScrollFrame(nb)
        nb.add(sf, text=f"Alternative {k + 1}")
        f = sf.inner
        f.columnconfigure(0, weight=1)
        v = self.av[k]
        for key in ("name", "A0", "foerderung", "n", "L", "WC", "einheit", "x", "p", "kv", "kfix", "g_e", "g_a"):
            v[key] = tk.StringVar()
        v["modus"] = tk.StringVar(value="abgeleitet")
        v["aktiv"] = tk.BooleanVar(value=(k == 0))

        oben = ttk.Frame(f)
        oben.grid(row=0, column=0, sticky="ew")
        cb = ttk.Checkbutton(oben, text="Alternative in Berechnung einbeziehen", variable=v["aktiv"])
        cb.pack(side="left")
        if k == 0:
            cb.state(["disabled"])
        else:
            ttk.Button(oben, text="Werte von Alternative 1 kopieren",
                       command=lambda: self._kopiere(0, k)).pack(side="right")

        lf = ttk.LabelFrame(f, text="Investition", padding=6)
        lf.grid(row=1, column=0, sticky="ew", pady=4)
        ttk.Label(lf, text="Bezeichnung").grid(row=0, column=0, sticky="w")
        ttk.Entry(lf, textvariable=v["name"], width=30).grid(row=0, column=1, columnspan=3, sticky="ew", padx=4)
        self._feld(lf, 1, "Anschaffungsauszahlung", v["A0"],
                   "Anschaffungs-/Herstellungskosten inkl. Nebenkosten (Transport, Montage, Planung) "
                   "in t = 0, netto.", "€")
        self._feld(lf, 2, "Förderung / Zuschuss", v["foerderung"],
                   "Investitionszuschüsse in t = 0 (z. B. BEG/BAFA, KfW-Tilgungszuschuss). Mindert die "
                   "Nettoinvestition und damit auch Abschreibung und kalkulatorische Zinsen.", "€")
        self._feld(lf, 3, "Nutzungsdauer n", v["n"],
                   "Wirtschaftliche Nutzungsdauer in ganzen Jahren (1–60).", "Jahre")
        ndv = tk.StringVar(value=ND_VORLAGEN[0][0])
        ndcb = ttk.Combobox(lf, textvariable=ndv, state="readonly", width=40, values=[z[0] for z in ND_VORLAGEN])
        ndcb.grid(row=4, column=0, columnspan=4, sticky="w", pady=(0, 4))

        def nd_gewaehlt(_):
            w = dict(ND_VORLAGEN)[ndv.get()]
            if w:
                v["n"].set(str(w))
        ndcb.bind("<<ComboboxSelected>>", nd_gewaehlt)
        Tooltip(ndcb, "Richtwerte (überwiegend aus den amtlichen AfA-Tabellen). Die wirtschaftliche "
                      "Nutzungsdauer kann davon abweichen.")
        self._feld(lf, 5, "Restwert / Liquidationserlös", v["L"],
                   "Erlös aus dem Verkauf am Ende der Nutzungsdauer (abzüglich Abbruch-/Entsorgungskosten).", "€")
        self._feld(lf, 6, "Working Capital", v["WC"],
                   "Zusätzlich gebundenes Umlaufvermögen (Lager, Forderungen) in t = 0, wird am Ende "
                   "wieder freigesetzt. Meist 0.", "€")

        lf = ttk.LabelFrame(f, text="Laufende Zahlungen pro Jahr", padding=6)
        lf.grid(row=2, column=0, sticky="ew", pady=4)
        ttk.Radiobutton(lf, text="aus Menge, Preis und Kosten ableiten", variable=v["modus"], value="abgeleitet",
                        command=lambda: self._modus(k)).grid(row=0, column=0, columnspan=4, sticky="w")
        ttk.Radiobutton(lf, text="Zahlungsreihe je Jahr manuell eingeben", variable=v["modus"], value="manuell",
                        command=lambda: self._modus(k)).grid(row=1, column=0, columnspan=4, sticky="w")
        ttk.Label(lf, text="Einheit der Menge").grid(row=2, column=0, sticky="w")
        ttk.Entry(lf, textvariable=v["einheit"], width=14).grid(row=2, column=1, sticky="ew", padx=4)
        self._feld(lf, 3, "Menge / Leistung p.a.", v["x"],
                   "Produzierte/abgesetzte Menge oder Leistung pro Jahr (Stück, kWh, km, Fälle …). "
                   "Bei Einsparinvestitionen: eingesparte Menge.", "")
        self._feld(lf, 4, "Preis / Erlös je Einheit", v["p"],
                   "Verkaufspreis bzw. Wert/Einsparung je Einheit. 0 bei reiner Kostenbetrachtung.", "€")
        self._feld(lf, 5, "Variable Kosten je Einheit", v["kv"],
                   "Material, Energie, Verbrauchsstoffe, Fertigungslöhne je Einheit.", "€")
        self._feld(lf, 6, "Fixe Betriebskosten p.a.", v["kfix"],
                   "Personal, Wartung, Versicherung, Raumkosten u. Ä. – OHNE Abschreibung und Zinsen "
                   "(diese berechnet das Programm selbst).", "€")
        self._feld(lf, 7, "Steigerung Einzahlungen p.a.", v["g_e"],
                   "Jährliche Preis-/Mengenänderung der Einzahlungen (z. B. Strompreissteigerung). "
                   "Negativ z. B. bei Leistungsdegradation.", "%")
        self._feld(lf, 8, "Steigerung Auszahlungen p.a.", v["g_a"],
                   "Jährliche Steigerung der laufenden Kosten (Inflation, Energiepreise).", "%")

        rf = ttk.LabelFrame(f, text="Manuelle Zahlungsreihe (ohne Investition, Restwert und WC)", padding=6)
        self.man_rahmen[k] = rf
        rf.grid(row=3, column=0, sticky="ew", pady=4)
        ttk.Button(rf, text="Tabelle erzeugen (mit Werten aus 'ableiten' vorbefüllen)",
                   command=lambda: self._tabelle_erzeugen(k)).pack(anchor="w")
        tab = ttk.Frame(rf)
        tab.pack(fill="x", pady=4)
        self.man_tabelle[k] = tab
        rf.grid_remove()

    def _modus(self, k):
        if self.av[k]["modus"].get() == "manuell":
            self.man_rahmen[k].grid()
            if not self.man_zeilen[k]:
                self._tabelle_erzeugen(k)
        else:
            self.man_rahmen[k].grid_remove()

    def _tabelle_erzeugen(self, k, werte=None):
        tab = self.man_tabelle[k]
        for w in tab.winfo_children():
            w.destroy()
        self.man_zeilen[k] = []
        if werte is None:
            try:
                a = self._lese_alt(k, modus="abgeleitet")
                z = zahlungsreihe(a)
                werte = [(z["E"][t], z["Aus"][t]) for t in range(1, a.n + 1)]
            except ValueError as ex:
                messagebox.showerror("Eingabefehler", str(ex))
                return
        ttk.Label(tab, text="Jahr", style="Titel.TLabel").grid(row=0, column=0, padx=4)
        ttk.Label(tab, text="Einzahlungen €", style="Titel.TLabel").grid(row=0, column=1, padx=4)
        ttk.Label(tab, text="lfd. Auszahlungen €", style="Titel.TLabel").grid(row=0, column=2, padx=4)
        for t, (e_, a_) in enumerate(werte, start=1):
            ve, va = tk.StringVar(value=eingabe_fmt(round(e_, 2))), tk.StringVar(value=eingabe_fmt(round(a_, 2)))
            ttk.Label(tab, text=str(t)).grid(row=t, column=0)
            ttk.Entry(tab, textvariable=ve, width=16, justify="right").grid(row=t, column=1, padx=4, pady=1)
            ttk.Entry(tab, textvariable=va, width=16, justify="right").grid(row=t, column=2, padx=4, pady=1)
            self.man_zeilen[k].append((ve, va))

    def _kopiere(self, quelle, ziel):
        for key, var in self.av[quelle].items():
            if key == "aktiv":
                continue
            self.av[ziel][key].set(var.get())
        self.av[ziel]["name"].set(self.av[quelle]["name"].get() + " (Kopie)")
        self.av[ziel]["aktiv"].set(True)
        if self.man_zeilen[quelle]:
            self._tabelle_erzeugen(ziel, [(parse_num(e.get()), parse_num(a.get()))
                                          for e, a in self.man_zeilen[quelle]])
        self._modus(ziel)

    # ------------------------------------------------------------------ Vorlagen
    def lade_vorlage(self):
        vl = VORLAGEN.get(self.vorlage_var.get())
        if not vl:
            return
        for key, wert in vl["glob"].items():
            if key in ("projekt", "wacc_modus"):
                self.gv[key].set(wert)
            else:
                self.gv[key].set(eingabe_fmt(wert))
        for k in range(3):
            daten = vl["alts"][k]
            v = self.av[k]
            v["aktiv"].set(daten is not None)
            daten = daten or _a(name=f"Alternative {k + 1}")
            for key, wert in daten.items():
                if key in ("name", "einheit", "modus"):
                    v[key].set(wert)
                elif key == "n":
                    v[key].set(str(int(wert)))
                else:
                    v[key].set(eingabe_fmt(wert))
            self.man_zeilen[k] = []
            for w in self.man_tabelle[k].winfo_children():
                w.destroy()
            self._modus(k)
        self._wacc_anzeigen()
        self.berechnen()

    # ------------------------------------------------------------------ Einlesen
    def lese_parameter(self) -> Parameter:
        v = self.gv
        P = lambda k, name: parse_num(v[k].get(), name) / 100  # noqa: E731
        g = Parameter(
            projekt=v["projekt"].get().strip() or "Investitionsprojekt",
            i=P("i", "Kalkulationszins"), r_re=P("r_re", "Reinvestitionszins"),
            h=P("h", "Habenzins"), s_soll=P("s_soll", "Sollzins"),
            ek=parse_num(v["ek"].get(), "Eigenkapital"), t_max=parse_num(v["t_max"].get(), "Max. Dauer"),
            steuer=P("steuer", "Steuersatz"), wacc_modus=v["wacc_modus"].get(),
            ek_quote=P("ek_quote", "EK-Quote"), r_ek=P("r_ek", "EK-Kosten"), r_fk=P("r_fk", "FK-Zins"),
            sens=P("sens", "Sensitivität"))
        fehler = []
        for name, wert in [("Kalkulationszins", g.i), ("Reinvestitionszins", g.r_re), ("Habenzins", g.h),
                           ("Sollzins", g.s_soll)]:
            if wert <= -1:
                fehler.append(f"{name} muss größer als −100 % sein.")
        if not (0 <= g.steuer < 1):
            fehler.append("Steuersatz muss zwischen 0 und 100 % liegen.")
        if not (0 <= g.ek_quote <= 1):
            fehler.append("Eigenkapitalquote muss zwischen 0 und 100 % liegen.")
        if not (0 < g.sens < 1):
            fehler.append("Sensitivitäts-Variation muss zwischen 0 und 100 % liegen.")
        if g.wacc <= -1:
            fehler.append("WACC muss größer als −100 % sein.")
        if g.ek < 0:
            fehler.append("Eigenkapital darf nicht negativ sein.")
        if fehler:
            raise ValueError("\n".join(fehler))
        return g

    def _lese_alt(self, k, modus=None) -> Alternative:
        v = self.av[k]
        name = v["name"].get().strip() or f"Alternative {k + 1}"
        N = lambda key, lab: parse_num(v[key].get(), f"{name}: {lab}")  # noqa: E731
        n_wert = N("n", "Nutzungsdauer")
        if abs(n_wert - round(n_wert)) > 1e-9:
            raise ValueError(f"{name}: Nutzungsdauer muss eine ganze Zahl sein.")
        a = Alternative(
            name=name, A0=N("A0", "Anschaffungsauszahlung"), foerderung=N("foerderung", "Förderung"),
            n=int(round(n_wert)), L=N("L", "Restwert"), WC=N("WC", "Working Capital"),
            einheit=v["einheit"].get().strip() or "Einheiten", x=N("x", "Menge"), p=N("p", "Preis"),
            kv=N("kv", "variable Kosten"), kfix=N("kfix", "Fixkosten"),
            g_e=N("g_e", "Steigerung Einzahlungen") / 100, g_a=N("g_a", "Steigerung Auszahlungen") / 100,
            modus=modus or v["modus"].get())
        if a.modus == "manuell":
            a.man_E = [parse_num(e.get(), f"{name}: Einzahlungen Jahr {t + 1}")
                       for t, (e, _) in enumerate(self.man_zeilen[k])]
            a.man_A = [parse_num(x.get(), f"{name}: Auszahlungen Jahr {t + 1}")
                       for t, (_, x) in enumerate(self.man_zeilen[k])]
        return a

    def lese_alternativen(self):
        alts, fehler = [], []
        for k in range(3):
            if not self.av[k]["aktiv"].get():
                continue
            a = self._lese_alt(k)
            fehler += pruefe(a)
            alts.append(a)
        namen = [a.name for a in alts]
        if len(set(namen)) != len(namen):
            fehler.append("Die Alternativen benötigen unterschiedliche Bezeichnungen.")
        if fehler:
            raise ValueError("\n".join(fehler))
        return alts

    # ------------------------------------------------------------------ Rechnen
    def _methode_key(self):
        titel = self.methode_var.get()
        return next(k for k in REIHENFOLGE if METHODEN[k][0] == titel)

    def berechnen(self):
        key = self._methode_key()
        self.beschr_var.set("ℹ " + BESCHREIBUNG[key])
        try:
            g = self.lese_parameter()
            alts = self.lese_alternativen()
            e = Ergebnis(alts, g)
        except ValueError as ex:
            messagebox.showerror("Eingabefehler", str(ex))
            self.status_var.set("Fehler in den Eingaben – bitte korrigieren.")
            return None
        self.ergebnis = e
        titel, tf, pf = METHODEN[key]
        self.text.configure(state="normal")
        self.text.delete("1.0", "end")
        self.text.insert("1.0", tf(e))
        self._hervorheben()
        self.text.configure(state="disabled")
        self.fig.clear()
        try:
            pf(self.fig, e)
            self.fig.suptitle(titel, fontsize=11, fontweight="bold")
            layout(self.fig)
        except Exception as ex:  # Diagrammfehler sollen die Rechnung nicht blockieren
            self.fig.clear()
            self.fig.text(0.5, 0.5, f"Diagramm konnte nicht erstellt werden:\n{ex}", ha="center")
        self.canvas.draw_idle()
        r0 = e.rs[0]
        self.status_var.set(f"Berechnet: {len(alts)} Alternative(n) · Kapitalwert {e.namen[0]}: {eur(r0['NPV'])}"
                            f" · i = {pct(g.i)}")
        return e

    def _hervorheben(self):
        if not hasattr(self, "_fett"):
            import tkinter.font as tkfont
            self._fett = tkfont.Font(font=self.text.cget("font"))
            self._fett.configure(weight="bold")
        self.text.tag_configure("titel", font=self._fett, foreground="#1f4e79")
        self.text.tag_configure("pos", foreground="#1a7f37")
        self.text.tag_configure("neg", foreground="#c62828")
        inhalt = self.text.get("1.0", "end").split("\n")
        for nr, zeile in enumerate(inhalt, start=1):
            if nr == 2 or (zeile.isupper() and len(zeile) > 3 and not zeile.startswith(" ")):
                self.text.tag_add("titel", f"{nr}.0", f"{nr}.end")
            for wort, tag in (("nicht vorteilhaft", "neg"), ("wertvernichtend", "neg"), ("vorteilhaft", "pos"),
                              ("wertschaffend", "pos")):
                start = zeile.find(wort)
                if start >= 0:
                    self.text.tag_add(tag, f"{nr}.{start}", f"{nr}.{start + len(wort)}")
                    break

    def exportieren(self, art):
        e = self.berechnen()
        if e is None:
            return
        sicher = re.sub(r"[^\w\-]+", "_", e.g.projekt).strip("_")[:40] or "Investition"
        name = f"Investitionsrechnung_{sicher}_{datetime.date.today():%Y-%m-%d}.{art}"
        typen = {"pdf": [("PDF-Dokument", "*.pdf")], "txt": [("Textdatei", "*.txt")],
                 "csv": [("CSV (Excel, Semikolon)", "*.csv")]}[art]
        pfad = filedialog.asksaveasfilename(defaultextension=f".{art}", initialfile=name, filetypes=typen)
        if not pfad:
            return
        methoden = REIHENFOLGE if self.alle_var.get() else [self._methode_key()]
        try:
            self.config(cursor="watch")
            self.update_idletasks()
            if art == "pdf":
                export_pdf(pfad, e, methoden)
            elif art == "txt":
                export_txt(pfad, e, methoden)
            else:
                export_csv(pfad, e)
        except Exception as ex:
            messagebox.showerror("Export fehlgeschlagen", str(ex))
            return
        finally:
            self.config(cursor="")
        self.status_var.set(f"Exportiert: {pfad}")
        messagebox.showinfo("Export", f"Datei gespeichert:\n{pfad}")


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
