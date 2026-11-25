# -*- coding: utf-8 -*-
"""Mehrlinien-Stromlaufplan (allpolig) Stockwerksverteiler + Wohnungsverteiler, A3 quer.

Aufbau in Anlehnung an EPLAN-Pläne (Deckblatt, Inhalt, Übersicht, Ansicht, Anspeisung,
Abgänge, Kabelübersicht, Stückliste).  Aufruf:  python3 plan.py <ausgabe.pdf>
"""
import math
import sys

from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

import data as D

FD = "/usr/share/fonts/truetype/liberation/"
pdfmetrics.registerFont(TTFont("R", FD + "LiberationSans-Regular.ttf"))
pdfmetrics.registerFont(TTFont("B", FD + "LiberationSans-Bold.ttf"))
pdfmetrics.registerFont(TTFont("I", FD + "LiberationSans-Italic.ttf"))

PW, PH = 420.0, 297.0
BK = (0, 0, 0)
GR = (0.42, 0.42, 0.42)
LG = (0.88, 0.88, 0.88)
HD = (0.80, 0.84, 0.90)
BL = (0.10, 0.28, 0.70)
GN = (0.00, 0.48, 0.20)
RD = (0.75, 0.10, 0.10)
YE = (1.00, 0.96, 0.70)
WH = (1, 1, 1)
X0, Y0, CW = 20.0, 15.0, 39.0      # Zeichenbereich: x 20..410, y 15..252 (von oben)


# =========================================================================== Zeichenhilfen
class Sheet:
    def __init__(self, c):
        self.c = c

    def L(self, x1, y1, x2, y2, w=0.25, dash=None, rgb=BK):
        c = self.c
        c.saveState()
        c.setLineWidth(w * mm)
        c.setStrokeColorRGB(*rgb)
        if dash:
            c.setDash([d * mm for d in dash], 0)
        c.line(x1 * mm, (PH - y1) * mm, x2 * mm, (PH - y2) * mm)
        c.restoreState()

    def PL(self, pts, w=0.25, dash=None, rgb=BK, close=False, fill=None):
        c = self.c
        c.saveState()
        c.setLineWidth(w * mm)
        c.setStrokeColorRGB(*rgb)
        if fill:
            c.setFillColorRGB(*fill)
        if dash:
            c.setDash([d * mm for d in dash], 0)
        p = c.beginPath()
        p.moveTo(pts[0][0] * mm, (PH - pts[0][1]) * mm)
        for x, y in pts[1:]:
            p.lineTo(x * mm, (PH - y) * mm)
        if close:
            p.close()
        c.drawPath(p, stroke=1, fill=1 if fill else 0)
        c.restoreState()

    def R(self, x, y, w, h, lw=0.25, fill=None, dash=None, rgb=BK):
        c = self.c
        c.saveState()
        c.setLineWidth(max(lw, 0.01) * mm)
        c.setStrokeColorRGB(*rgb)
        if fill:
            c.setFillColorRGB(*fill)
        if dash:
            c.setDash([d * mm for d in dash], 0)
        c.rect(x * mm, (PH - y - h) * mm, w * mm, h * mm, stroke=1 if lw > 0 else 0, fill=1 if fill else 0)
        c.restoreState()

    def CI(self, x, y, r, lw=0.25, fill=None, rgb=BK):
        c = self.c
        c.saveState()
        c.setLineWidth(lw * mm)
        c.setStrokeColorRGB(*rgb)
        if fill:
            c.setFillColorRGB(*fill)
        c.circle(x * mm, (PH - y) * mm, r * mm, stroke=1, fill=1 if fill else 0)
        c.restoreState()

    def EL(self, x1, y1, x2, y2, lw=0.25, dash=None):
        c = self.c
        c.saveState()
        c.setLineWidth(lw * mm)
        if dash:
            c.setDash([d * mm for d in dash], 0)
        c.ellipse(x1 * mm, (PH - y1) * mm, x2 * mm, (PH - y2) * mm, stroke=1, fill=0)
        c.restoreState()

    def T(self, x, y, txt, sz=6.5, f="R", a="l", rot=0, rgb=BK):
        c = self.c
        c.saveState()
        c.setFillColorRGB(*rgb)
        c.setFont(f, sz)
        X, Y = x * mm, (PH - y) * mm
        if rot:
            c.translate(X, Y)
            c.rotate(rot)
            X, Y = 0, 0
        if a == "l":
            c.drawString(X, Y, txt)
        elif a == "c":
            c.drawCentredString(X, Y, txt)
        else:
            c.drawRightString(X, Y, txt)
        c.restoreState()

    @staticmethod
    def W(txt, sz, f="R"):
        return pdfmetrics.stringWidth(txt, f, sz) / mm

    def TF(self, x, y, txt, wmax, sz=6.5, f="R", a="l", rgb=BK, smin=4.2):
        """Text, der bei Bedarf verkleinert wird, um in wmax zu passen."""
        while sz > smin and self.W(txt, sz, f) > wmax:
            sz -= 0.2
        self.T(x, y, txt, sz, f, a, rgb=rgb)

    def dot(self, x, y, r=0.6):
        self.CI(x, y, r, lw=0.1, fill=BK)

    def term(self, x, y, r=0.8):
        self.CI(x, y, r, lw=0.22, fill=WH)


def wrap(txt, wmax, sz, f="R"):
    words = str(txt).split(" ")
    lines, cur = [], ""
    for w in words:
        t = (cur + " " + w).strip()
        if Sheet.W(t, sz, f) <= wmax or not cur:
            cur = t
        else:
            lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return lines or [""]


def colof(x):
    return max(0, min(9, int((x - X0) // CW)))


def fnum(v, d=1):
    return f"{v:.{d}f}".replace(".", ",")


# =========================================================================== Seitenverwaltung
PAGES = []
PN = {}


def add(key, title, anl, fn, *args):
    PAGES.append(dict(key=key, title=title, anl=anl, fn=fn, args=args))


def xr(key, x_or_col):
    col = x_or_col if isinstance(x_or_col, int) else colof(x_or_col)
    return f"{PN[key]}.{col}"


ANL_TXT = {"ALL": "Gesamtanlage", "SV": "=SV  Stockwerksverteiler (Stiegenhaus)"}
for _t in D.TOPS:
    ANL_TXT[f"WV{_t}"] = f"=WV{_t}  Wohnungsverteiler Top {_t}"


# =========================================================================== Rahmen + Schriftfeld
def frame(s, pno, total, title, anl):
    s.R(5, 5, 410, 287, lw=0.5)
    s.R(X0, 10, 390, 277, lw=0.35)
    for i in range(10):
        x = X0 + CW * i
        if i:
            s.L(x, 10, x, 15, 0.2)
            s.L(x, 252, x, 257, 0.2)
        s.T(x + CW / 2, 13.7, str(i), 6.5, "R", "c")
        s.T(x + CW / 2, 255.7, str(i), 6.5, "R", "c")
    s.L(X0, 15, 410, 15, 0.25)
    s.L(X0, 252, 410, 252, 0.25)
    s.L(X0, 257, 410, 257, 0.35)
    # Schutzvermerk links
    s.T(13, 250, "Schutzvermerk nach ISO 16016 beachten · Plan mit Schrack-Komponenten, Stand " + D.DATE,
        5, "R", "l", rot=90, rgb=GR)
    # Schriftfeld y 257..287
    y = 257
    # A Änderungen
    s.L(92, y, 92, 287, 0.3)
    s.R(X0, y, 72, 5, lw=0.2, fill=LG)
    for xx, t in ((X0 + 1, "Änd."), (29, "Datum"), (46, "Name"), (59, "Änderung")):
        s.T(xx, y + 3.6, t, 5.5, "B")
    for k in range(1, 4):
        s.L(X0, y + 5 + 7.33 * k, 92, y + 5 + 7.33 * k, 0.15)
    for xx in (28, 45, 58):
        s.L(xx, y, xx, 287, 0.15)
    s.T(X0 + 2, y + 10, "0", 6)
    s.T(29, y + 10, D.DATE, 5.5)
    s.T(46.5, y + 10, D.BEARB, 5.5)
    s.T(59, y + 10, "Erstausgabe", 5.5)
    # B Datum / Bearbeiter
    s.L(140, y, 140, 287, 0.3)
    for k, (a, b) in enumerate((("Datum", D.DATE), ("Bearb.", D.BEARB), ("Gepr.", "—"), ("Norm", "OVE E 8101"))):
        yy = y + 7.5 * k
        if k:
            s.L(92, yy, 140, yy, 0.15)
        s.T(93.5, yy + 4.8, a, 5.5, "B")
        s.T(107, yy + 4.8, b, 6)
    # C Firma
    s.L(210, y, 210, 287, 0.3)
    s.T(175, y + 10.5, D.COMPANY, 13, "B", "c")
    s.T(175, y + 16, "Elektroplanung", 6.5, "R", "c")
    s.T(175, y + 24, "Plan-Nr. " + D.PLAN_NO, 6, "R", "c")
    # D Projekt
    s.L(300, y, 300, 287, 0.3)
    s.T(211.5, y + 4, "Projekt", 5, "R", rgb=GR)
    s.TF(211.5, y + 9.5, D.PROJECT, 86, 7.5, "B")
    s.T(211.5, y + 15, "Netzbetreiber: " + D.NETZ, 6)
    s.T(211.5, y + 20, "Bauherr / Kunde: lt. Projekt", 6)
    s.T(211.5, y + 25.5, "Auftr.-Nr.: —", 6)
    # E Benennung + Anlage
    s.L(375, y, 375, 287, 0.3)
    s.T(301.5, y + 4, "Benennung", 5, "R", rgb=GR)
    tl = wrap(title, 71, 8.5, "B")
    for k, t in enumerate(tl[:2]):
        s.T(301.5, y + 10 + 4.2 * k, t, 8.5, "B")
    s.L(300, y + 19, 375, y + 19, 0.15)
    s.T(301.5, y + 22.3, "Anlage", 5, "R", rgb=GR)
    s.TF(301.5, y + 27, ANL_TXT.get(anl, anl), 72, 6.5, "B")
    # F Blatt
    s.T(376.5, y + 4, "Blatt", 5, "R", rgb=GR)
    s.T(392.5, y + 14, str(pno), 14, "B", "c")
    s.L(375, y + 19, 410, y + 19, 0.15)
    s.T(376.5, y + 22.3, "Blätter", 5, "R", rgb=GR)
    s.T(392.5, y + 27, str(total), 8, "B", "c")


XE = 392.0


# =========================================================================== Symbole
def pole(s, x, y0, h=16, kind="sw", nt=None, nb=None):
    """Ein Schaltkontakt (offen) zwischen y0 und y0+h. Rückgabe: Mitte Schaltstück."""
    yt = y0 + h * 0.30
    yb = y0 + h * 0.72
    s.L(x, y0, x, yt)
    s.L(x, yb, x, y0 + h)
    s.L(x, yb, x - 3.3, yt + 0.2)
    if kind == "ls":
        d = 0.85
        s.L(x - d, yt - d, x + d, yt + d, 0.25)
        s.L(x - d, yt + d, x + d, yt - d, 0.25)
    elif kind == "last":
        s.L(x - 1.1, yt, x + 1.1, yt, 0.3)
        s.CI(x, yt - 1.25, 0.55, 0.2)
    elif kind == "tr":
        s.L(x - 1.1, yt, x + 1.1, yt, 0.3)
    if nt:
        s.T(x + 0.9, y0 + 2.3, nt, 4.3)
    if nb:
        s.T(x + 0.9, y0 + h - 0.7, nb, 4.3)
    return (x - 1.65, (yt + yb) / 2)


def linkage(s, mids, handle=True, extra_left=6):
    ym = mids[0][1]
    xa = mids[0][0] - (extra_left if handle else 0)
    s.L(xa, ym, mids[-1][0], ym, 0.22, dash=(0.9, 0.7))
    if handle:
        s.L(xa, ym - 1.3, xa, ym + 1.3, 0.3)
        s.L(xa - 1.2, ym - 1.3, xa, ym - 1.3, 0.3)


def fuse_pole(s, x, y0, h=20, nt=None, nb=None, fitted=True):
    """D0-Sicherungslasttrennschalter: Lasttrennkontakt, Schaltstück als Sicherung."""
    yt = y0 + 5.0
    yb = y0 + 14.0
    s.L(x, y0, x, yt)
    s.L(x - 1.1, yt, x + 1.1, yt, 0.3)
    s.CI(x, yt - 1.25, 0.55, 0.2)
    s.L(x, yb, x, y0 + h)
    ex, ey = x - 3.6, yt + 0.4
    s.L(x, yb, ex, ey)
    # Sicherungskörper auf Schaltstück
    mx, my = (x + ex) / 2, (yb + ey) / 2
    ux, uy = ex - x, ey - yb
    ln = math.hypot(ux, uy)
    ux, uy = ux / ln, uy / ln
    vx, vy = -uy, ux
    a, b = 3.0, 1.15
    pts = [(mx + ux * a + vx * b, my + uy * a + vy * b), (mx - ux * a + vx * b, my - uy * a + vy * b),
           (mx - ux * a - vx * b, my - uy * a - vy * b), (mx + ux * a - vx * b, my + uy * a - vy * b)]
    s.PL(pts, 0.25, close=True, fill=WH if fitted else None, dash=None if fitted else (0.6, 0.5))
    s.L(mx + ux * a, my + uy * a, mx - ux * a, my - uy * a, 0.2)
    if nt:
        s.T(x + 0.9, y0 + 2.3, nt, 4.3)
    if nb:
        s.T(x + 0.9, y0 + h - 0.7, nb, 4.3)
    return (mx - 1.0, my)


def label(s, x, y, lines, sz=5.8, lh=2.9):
    """Gerätebeschriftung: erste Zeile fett (BMK)."""
    for k, t in enumerate(lines):
        if not t:
            continue
        s.T(x, y + k * lh, t, sz + (0.9 if k == 0 else 0), "B" if k == 0 else "R",
            rgb=BK if k == 0 else (0.15, 0.15, 0.15))


def mcb(s, xs, y0, h=16, nt=("1", "3", "5"), nb=("2", "4", "6")):
    mids = [pole(s, x, y0, h, "ls", nt[i] if i < len(nt) else None, nb[i] if i < len(nb) else None)
            for i, x in enumerate(xs)]
    bx = xs[0] - 9.5
    ym = mids[0][1]
    s.R(bx, ym - 2.2, 4.2, 4.4, 0.22, fill=WH)
    s.T(bx + 2.1, ym + 0.9, "I>", 4.2, "R", "c")
    s.L(bx + 4.2, ym, mids[-1][0], ym, 0.22, dash=(0.9, 0.7))
    return ym


def lts(s, xs, y0, h=16):
    mids = [pole(s, x, y0, h, "last", ("1", "3", "5")[i], ("2", "4", "6")[i]) for i, x in enumerate(xs)]
    linkage(s, mids)


def rcd(s, xs, y0, h=16, nt=("1", "3", "5", "N"), nb=("2", "4", "6", "N"), ls_first=False):
    mids = []
    for i, x in enumerate(xs):
        mids.append(pole(s, x, y0, h, "ls" if (ls_first and i == 0) else "sw", nt[i], nb[i]))
    yc = y0 + h + 3.6
    for x in xs:
        s.L(x, y0 + h, x, y0 + h + 7.2)
    xa, xb = xs[0] - 2.8, xs[-1] + 2.8
    s.EL(xa, yc - 1.4, xb, yc + 1.4, 0.25)
    bx = xs[0] - 11.5
    s.R(bx, yc - 2.6, 6.0, 5.2, 0.22, fill=WH)
    s.T(bx + 3.0, yc + 0.9, "IΔ" if not ls_first else "IΔ I>", 4.3 if ls_first else 5, "R", "c")
    s.L(bx + 6.0, yc, xa, yc, 0.22)
    ym = mids[0][1]
    s.L(bx + 3.0, yc - 2.6, bx + 3.0, ym, 0.22, dash=(0.9, 0.7))
    s.L(bx + 3.0, ym, mids[-1][0], ym, 0.22, dash=(0.9, 0.7))
    return y0 + h + 7.2


def meter(s, xs, y0, h=22, fitted=True):
    x1, x2 = xs[0] - 4.5, xs[-1] + 4.5
    s.R(x1, y0, x2 - x1, h, 0.35, dash=None if fitted else (1.2, 0.8))
    s.T((x1 + x2) / 2, y0 + h / 2 + 0.2, "Wh", 10, "B", "c", rgb=BK if fitted else GR)
    s.T((x1 + x2) / 2, y0 + h / 2 + 4.6, "Smart Meter" if fitted else "kein Zähler", 4.8, "R", "c",
        rgb=BK if fitted else GR)
    for x, ni, no in zip(xs, ("1", "4", "7", "10"), ("3", "6", "9", "12")):
        s.T(x + 0.8, y0 + 2.6, ni, 4.2)
        s.T(x + 0.8, y0 + h - 0.9, no, 4.2)


def spd(s, xs, y0):
    """Überspannungsableiter 3+1; xs = L1,L2,L3,N. Rückgabe: y PE-Anschluss unten (bei xs[3])."""
    x1, x2 = xs[0] - 3.2, xs[-1] + 3.2
    h = 17
    s.R(x1, y0, x2 - x1, h, 0.3)
    yb = y0 + 10.5
    for x in xs[:3]:
        s.L(x, y0, x, y0 + 2.5)
        s.R(x - 1.0, y0 + 2.5, 2.0, 5.2, 0.22, fill=WH)
        s.L(x - 1.6, y0 + 7.0, x + 1.6, y0 + 3.0, 0.2)
        s.L(x, y0 + 7.7, x, yb)
    s.L(xs[0], yb, xs[3], yb, 0.25)
    s.L(xs[3], y0, xs[3], yb)
    # Funkenstrecke N-PE
    xn = xs[3]
    s.L(xn, yb, xn, yb + 1.6)
    s.PL([(xn - 1.0, yb + 1.6), (xn + 1.0, yb + 1.6), (xn, yb + 2.8)], 0.2, close=True, fill=BK)
    s.PL([(xn - 1.0, yb + 4.6), (xn + 1.0, yb + 4.6), (xn, yb + 3.4)], 0.2, close=True, fill=BK)
    s.L(xn, yb + 4.6, xn, y0 + h)
    return y0 + h


def cablemark(s, xs, y, name, typ, extra=None, colors=None, tx=None):
    xa, xb = xs[0] - 2.4, xs[-1] + 2.4
    s.L(xa, y + 1.6, xb, y - 1.6, 0.3)
    if colors:
        for x, col in zip(xs, colors):
            s.T(x - 0.5, y + 7.0, col, 3.7, "R", "l", rot=90, rgb=GR)
    t = tx if tx is not None else xb + 1.5
    s.T(t, y - 0.3, name, 6.3, "B")
    s.T(t, y + 2.6, typ, 5.4)
    if extra:
        s.T(t, y + 5.3, extra, 5.0, rgb=GR)


def xref_down(s, x, y, txt):
    s.L(x, y, x, y + 2.0)
    s.PL([(x - 0.9, y + 1.0), (x, y + 2.4), (x + 0.9, y + 1.0)], 0.22)
    s.T(x, y + 5.4, txt, 4.6, "R", "c", rgb=BL)


def bus(s, y, xa, xb, name, lref=None, rref=None, w=0.3):
    s.L(xa, y, xb, y, w)
    if lref:
        s.T(xa - 0.8, y + 0.9, f"{name} / {lref}", 4.8, "R", "r", rgb=BL)
    if rref:
        s.T(xb + 0.8, y + 0.9, f"{name} / {rref}", 4.8, "R", "l", rgb=BL)


def consumer(s, x, y, kind, txt):
    if kind == "licht":
        s.CI(x, y + 3.2, 3.0, 0.3)
        d = 2.1
        s.L(x - d, y + 3.2 - d, x + d, y + 3.2 + d, 0.25)
        s.L(x - d, y + 3.2 + d, x + d, y + 3.2 - d, 0.25)
    elif kind == "sd":
        c = s.c
        c.saveState()
        c.setLineWidth(0.3 * mm)
        p = c.beginPath()
        p.arc((x - 3.2) * mm, (PH - y - 6.4) * mm, (x + 3.2) * mm, (PH - y) * mm, 0, 180)
        c.drawPath(p, stroke=1, fill=0)
        c.restoreState()
        s.L(x - 3.2, y + 3.2, x + 3.2, y + 3.2, 0.3)
        s.L(x, y + 3.2, x, y + 6.4, 0.3)
    elif kind == "res":
        s.R(x - 4, y, 8, 6, 0.25, dash=(0.8, 0.6))
    else:
        s.R(x - 4, y, 8, 6.4, 0.3)
        s.T(x, y + 4.3, "M" if kind == "motor" else "≈", 6, "B", "c")


# =========================================================================== Seiten
def pg_cover(s):
    s.T(215, 36, "STROMLAUFPLAN  (allpolig / Mehrlinien)", 20, "B", "c")
    s.T(215, 46, D.SUBTITLE, 13, "B", "c")
    s.T(215, 54, D.PROJECT + "  ·  Netzbetreiber " + D.NETZ + "  ·  Komponenten Schrack Technik",
        9, "R", "c", rgb=GR)

    def box(x, y, w, h, title):
        s.R(x, y, w, h, 0.3)
        s.R(x, y, w, 6, 0.3, fill=HD)
        s.T(x + 2, y + 4.3, title, 7, "B")

    def kv(x, y, rows, wk=46, sz=6.2, lh=4.3, wmax=None):
        for k, (a, b) in enumerate(rows):
            s.T(x, y + k * lh, a, sz, "B")
            s.TF(x + wk, y + k * lh, b, wmax or 70, sz)

    box(25, 64, 128, 104, "Technische Daten  =SV  Stockwerksverteiler")
    kv(28, 75, [
        ("Gehäuse", "Schrack Modul 2000 UP 3U-33G/WIEN (IL166333WS)"),
        ("Abmessungen", "H 1605 × B 810 × T 250 mm, Putznische 1575 × 800"),
        ("Bestückung", "6 Zählerplätze (2 Reihen à 3), 3 Nachzählerreihen 99 TE"),
        ("Normen Gehäuse", "EN 62208, EN 60529; Zählertafeln OVE E 8640:2022"),
        ("Netzform", "TN-C-S (Nullung), PEN-Aufteilung im HV"),
        ("Bemessungsspannung Ue", "3/N/PE AC 400/230 V, 50 Hz"),
        ("Bemessungsisolationsspg. Ui", "1000 V (Gehäuse lt. Datenblatt)"),
        ("Bemessungsstrom InA", "80 A (Vorsicherung NH00 gG 80 A im HV)"),
        ("Bed. Kurzschlussstrom Icc", "25 kA (Gehäuse); Ik am SV ≤ 10 kA (Annahme)"),
        ("Schutzart / -klasse", "IP20C / I, Farbe RAL 9016"),
        ("Vorzählersicherung", "NEOZED D02 gG 3×35 A (Top), 3×25 A (Allg.)"),
        ("Nachzähler je Zählerplatz", "Hauptschalter 40 A 3-pol. + N-Klemme"),
        ("Zähler", "Smart Meter direkt (Beistellung Wiener Netze)"),
        ("Zähleranschluss", "Zählersteckklemme 3p+2N 50 A, OVE R 21:2019"),
        ("Umgebung", "Stiegenhaus (Fluchtweg), −5…+40 °C, Kabel halogenfrei"),
        ("Verlustleistung", f"≈ {fnum(D.pv_sv()[1], 0)} W (RDF 0,6) ≤ 455 W zulässig"),
        ("Überspannungsschutz", "kein Ableiter im SV (Typ 1 im HV, Typ 2 je WV)"),
    ], wk=44, wmax=78)

    box(158, 64, 128, 104, "Technische Daten  =WV1 … =WV4  Wohnungsverteiler")
    kv(161, 75, [
        ("Gehäuse", "Schrack BK0850042 + BK0851042F, UP Massivwand"),
        ("Abmessungen", "H 717 × B 346 × T 94 mm, 4 Reihen à 14 TE (56 TE)"),
        ("Norm", "ÖVE/ÖNORM EN 61439-3 (DBO)"),
        ("Bemessungsspannung Ue", "3/N/PE AC 400/230 V, 50 Hz"),
        ("Bemessungsstrom InA", "35 A (Vorsicherung D02 35 A im SV)"),
        ("Bed. Kurzschlussstrom", "Inc FI 6 kA; Ik am WV ≈ 3,4 kA (10 kA am SV)"),
        ("Schutzart / -klasse", "IP40 / I (Stahlblechtür), RAL 9003"),
        ("Zuleitung", "NHXMH-J 5×10 mm², ≤ 25 m, ΔU ≈ 0,8 %"),
        ("Schutzmaßnahme", "Nullung + Zusatzschutz FI 30 mA (alle Endstromkreise)"),
        ("FI-Gruppen", "2 × FI 40/4/0,03 Typ A, G + FI/LS B16 Kühlgerät"),
        ("Leitungsschutz", "LS AMPARO 10 kA, B-Charakteristik"),
        ("Überspannungsschutz", "Typ 2 (+T3) 3+1, 275 V (IS111314) vor den FI"),
        ("Endstromkreise", "18 + 2 Reserve (3-Zimmer-Wohnung ~75 m²)"),
        ("Heizung / WW", "Fernwärme-Wohnungsstation (kein E-Boiler)"),
        ("Verlustleistung", f"≈ {fnum(D.pv_wv()[1], 0)} W (RDF 0,5) ≤ 54 W zulässig"),
        ("Hauptschalter", "im SV je Top (-nQ1); WV ohne eigenen Hauptschalter"),
    ], wk=44, wmax=78)

    box(291, 64, 114, 104, "Verdrahtung / Kennzeichnung")
    rows = [("Hauptstrom L1, L2, L3", "schwarz (H07V-K)"), ("Neutralleiter N", "hellblau"),
            ("Schutzleiter PE", "grün-gelb"), ("Kabeladern (HD 308)", "L1 BN · L2 BK · L3 GY · N BU"),
            ("Querschnitt 25/35/40 A", "10 mm² Cu (einheitlich)"),
            ("Querschnitt B16", "2,5 mm² Cu"), ("Querschnitt B6/B10", "1,5 mm² Cu"),
            ("Ableiteranschluss", "6 mm², Gesamtlänge ≤ 0,5 m"),
            ("Vorzählerbereich", "kurz- und erdschlusssicher, plombiert"),
            ("Klemmen -nX2", "Abgang Zählerplatz n (L1-L3, N, PE)"),
            ("Klemmen -XN1/-XN2", "N je FI-Gruppe (getrennt!)"),
            ("Klemme -XPE", "PE-Klemmleiste Wohnungsverteiler"),
            ("BMK-Schema", "-nF1 Vorzähler · -nP1 Zähler · -nQ1 HS"),
            ("Querverweis", "Blatt.Spalte, z. B. 7.3")]
    kv(294, 75, rows, wk=40, wmax=68)

    box(25, 174, 380, 74, "Planungsgrundlagen (Vorgaben Auftraggeber, Multiple-Choice-Abfrage)")
    left = [
        "• Umfang: Stockwerksverteiler inkl. Zähler + 4 Wohnungsverteiler (allpolig)",
        "• Zusatzplätze SV: Allgemeinzähler + Reserve-Zählerplatz (6 Zählerplätze gesamt)",
        "• Vorzählersicherung: NEOZED D02 3×35 A (Sicherungslasttrennschalter, plombierbar)",
        "• Einspeisung: Stich aus HV, NHXHX-J 5×25 mm², NH00 gG 80 A",
        "• Nachzählerteil: Hauptschalter 40 A 3-polig + N-Klemme je Zählerplatz",
        "• Überspannungsschutz: Typ 2 je Wohnungsverteiler (Typ 1 im HV vorausgesetzt)",
        "• Zuleitung Top: NHXMH-J 5×10 mm², ≤ 25 m (halogenfrei, Fluchtweg)",
        "• Allgemein: Stiegenhauslicht, Sprechanlage, RWA/Sicherheitsbeleuchtung",
        "• SV Unterputz in Wandnische; WV: 3 Zi. ~75 m², Fernwärme, 2 FI-Gruppen + FI/LS Kühlgerät",
    ]
    right = [
        "Normen / Regelwerke:",
        "• OVE E 8101 (Errichtung NS-Anlagen), TAEV 2020 + Ausführungsbest. Wiener Netze",
        "• ÖVE/ÖNORM EN 61439-1/-3, EN 62208; OVE R 21:2019 (Zählersteckklemmen)",
        "• EN 60898-1 (LS), EN 61008/61009 (FI, FI/LS), EN 61643-11 (SPD)",
        "Hinweise:",
        "• Leitungslängen, Ik und Netzimpedanz sind Annahmen → vor Ausführung prüfen",
        "• Zählerplatz-/Plombierungsvorgaben mit Wiener Netze abstimmen",
        "• Brandschutz Stiegenhaus (OIB-RL 2) lt. Brandschutzkonzept beachten",
        "• Ausführung nur durch konzessioniertes Elektrounternehmen",
    ]
    for k, t in enumerate(left):
        s.T(29, 184 + k * 6.6, t, 6.6)
    for k, t in enumerate(right):
        s.T(216, 184 + k * 6.6, t, 6.6, "B" if t.endswith(":") else "R")


def pg_toc(s):
    s.T(25, 25, "Inhaltsverzeichnis", 12, "B")
    widths = [16, 150, 120, 26, 18]
    hdr = ["Blatt", "Benennung", "Anlage", "Datum", "Bearb."]
    x, y = 25, 32
    s.R(x, y, sum(widths), 6, 0.25, fill=HD)
    xx = x
    for w, h in zip(widths, hdr):
        s.T(xx + 1.5, y + 4.2, h, 6.5, "B")
        xx += w
    y += 6
    rh = 5.35
    for p in PAGES:
        xx = x
        vals = [str(PN[p["key"]]), p["title"], ANL_TXT.get(p["anl"], p["anl"]), D.DATE, D.BEARB]
        for w, v in zip(widths, vals):
            s.TF(xx + 1.5, y + 3.8, v, w - 3, 6.3)
            xx += w
        s.L(x, y + rh, x + sum(widths), y + rh, 0.12, rgb=GR)
        y += rh
    s.R(x, 32, sum(widths), y - 32, 0.25)


def pg_overview(s):
    """Übersichtsschaltplan einpolig (Energiefluss)."""
    s.T(25, 22, "Übersichtsschaltplan (einpolig) – Energiefluss HV → Stockwerksverteiler → Wohnungsverteiler",
        8.5, "B")
    # HV
    s.R(30, 30, 58, 22, 0.3, dash=(1.2, 0.8))
    s.T(33, 36, "Hauptverteiler (HV) / Bestand", 6.3, "B")
    s.T(33, 41, "Abgang NH00 gG 80 A", 6)
    s.T(33, 45.5, "TN-C-S, PEN-Aufteilung im HV", 5.5, rgb=GR)
    s.T(33, 49.5, "Ik am SV ≤ 10 kA (Annahme)", 5.5, rgb=GR)
    xm = 59
    s.L(xm, 52, xm, 74, 0.45)
    for k in range(3):
        s.L(xm - 1.6, 61 + k * 1.3, xm + 1.6, 59 + k * 1.3, 0.25)
    s.T(xm + 3, 61, "-W0  NHXHX-J 5×25 mm², ≈30 m", 6, "B")
    s.T(xm + 3, 64.5, f"Ib ≈ {fnum(D.calc()['Ib'])} A, ΔU ≈ {fnum(D.calc()['du_w0'][1], 2)} %", 5.5, rgb=GR)
    # SV Sammelschiene
    s.R(25, 74, 380, 92, 0.3, dash=(2, 1))
    s.T(66, 79, "=SV  Stockwerksverteiler UP 3U-33G/WIEN (Stiegenhaus)", 7, "B")
    s.T(66, 83.5, f"→ Blatt {PN['sv_feed']}–{PN['sv_allg']}", 5.5, rgb=BL)
    s.L(40, 90, 395, 90, 0.7)
    s.T(66, 88.5, "-X0  Hauptabzweigklemme (ungezählt, plombiert)", 5.5, rgb=GR)
    s.L(xm, 74, xm, 90, 0.45)
    s.dot(xm, 90, 0.8)
    xs = [65, 120, 175, 230, 285, 340]
    for z, x in zip(D.ZP, xs):
        s.dot(x, 90, 0.8)
        y = 90
        s.L(x, y, x, 96)
        # Sicherung
        s.R(x - 1.3, 96, 2.6, 8, 0.3, fill=WH if z["fitted"] else None, dash=None if z["fitted"] else (0.6, 0.5))
        s.L(x, 96, x, 104)
        s.T(x + 3, 99.5, f"-{z['n']}F1", 6, "B")
        s.T(x + 3, 102.8, f"D02 3×{z['fuse']} A", 5.3)
        s.L(x, 104, x, 110)
        # Zähler
        s.R(x - 4.5, 110, 9, 10, 0.35, dash=None if z["fitted"] else (1, 0.7))
        s.T(x, 116.3, "Wh", 6.5, "B", "c")
        s.T(x + 6, 114, f"-{z['n']}P1", 6, "B")
        s.T(x + 6, 117.3, "Smart Meter" if z["fitted"] else "Reserve", 5.2)
        s.L(x, 120, x, 127)
        # Hauptschalter
        pole(s, x, 127, 12, "last")
        s.T(x + 3, 132, f"-{z['n']}Q1", 6, "B")
        s.T(x + 3, 135.3, "HS 40 A 3p", 5.3)
        s.L(x, 139, x, 152)
        for k in range(3):
            s.L(x - 1.5, 146 + k * 1.3, x + 1.5, 144 + k * 1.3, 0.25)
        s.T(x - 2, 81 + 14, z["name"], 7, "B", "r")
        if z["kind"] == "top":
            s.T(x + 3, 147, f"-W{z['n']} NHXMH-J 5×10", 5.4, "B")
        elif z["kind"] == "allg":
            s.T(x + 3, 147, "-W5.1…-W5.3", 5.4, "B")
        else:
            s.T(x + 3, 147, "-W6 (vorbereitet)", 5.4, "B", rgb=GR)
        s.T(x + 3, 158, f"→ Blatt {PN['sv_zp1'] if z['n'] <= 3 else PN['sv_zp2']}", 5.2, rgb=BL)
    # WV Boxen
    for z, x in zip(D.ZP, xs):
        if z["kind"] == "top":
            t = z["top"]
            s.L(x, 166, x, 176, 0.45)
            s.R(x - 25, 176, 50, 66, 0.3)
            s.R(x - 25, 176, 50, 6, 0.3, fill=HD)
            s.T(x, 180.3, f"=WV{t}  Wohnungsverteiler Top {t}", 6, "B", "c")
            lines = ["-X0 Einspeisung, -F01 Ableiter T2", "-F1 FI 40/4/0,03 A,G → 8 Kreise",
                     "-F2 FI 40/4/0,03 A,G → 11 Kreise", "-F3 FI/LS B16 Kühlgerät",
                     "Licht 3× B10 · SD/Geräte B16", "Kochfeld B16 3-pol.", "FW-Station B10",
                     f"→ Blatt {PN[f'wv{t}_view']}–{PN[f'wv{t}_g2b']}"]
            for k, l in enumerate(lines):
                s.T(x - 23, 188 + k * 6.2, l, 5.4, rgb=BL if l.startswith("→") else BK)
        elif z["kind"] == "allg":
            s.L(x, 166, x, 176, 0.45)
            s.R(x - 25, 176, 50, 66, 0.3)
            s.R(x - 25, 176, 50, 6, 0.3, fill=HD)
            s.T(x, 180.3, "Allgemeinbereich", 6, "B", "c")
            lines = ["-5F2 FI 40/4/0,03 A,G", "-5F3 B10 Stiegenhaus + SiBel", "-5K1 Treppenlichtzeitschalter",
                     "-5F4 B6 Sprechanlage", "-5F5 B16 Reserve", "-5F6 B10 RWA (ohne FI)",
                     f"→ Blatt {PN['sv_allg']}"]
            for k, l in enumerate(lines):
                s.T(x - 23, 188 + k * 6.2, l, 5.4, rgb=BL if l.startswith("→") else BK)
        else:
            s.L(x, 166, x, 176, 0.45, dash=(1, 0.8))
            s.R(x - 25, 176, 50, 30, 0.3, dash=(1.2, 0.8))
            s.T(x, 182, "Reserve", 6.5, "B", "c")
            s.T(x, 188, "Wallbox / Wärmepumpe / PV", 5.4, "R", "c")
            s.T(x, 193, "Zählerplatz + Hauptschalter", 5.4, "R", "c")
            s.T(x, 198, "vorbereitet", 5.4, "R", "c")


def pg_proof(s):
    c = D.calc()
    s.T(25, 22, "Auslegung und Nachweise (Kurzfassung – Details in Excel „Auslegung“)", 9, "B")

    def tab(x, y, widths, hdr, rows, sz=6.2, rh=5.4, title=None):
        if title:
            s.T(x, y - 2, title, 7.2, "B")
        s.R(x, y, sum(widths), rh, 0.25, fill=HD)
        xx = x
        for w, h in zip(widths, hdr):
            s.TF(xx + 1.2, y + rh - 1.5, h, w - 2, sz, "B")
            xx += w
        for k, r in enumerate(rows):
            yy = y + rh * (k + 1)
            xx = x
            for w, v in zip(widths, r):
                ok = isinstance(v, str) and (v.startswith("OK") or v.startswith("!"))
                s.TF(xx + 1.2, yy + rh - 1.5, v, w - 2, sz, "B" if ok else "R",
                     rgb=GN if v.startswith("OK") else (RD if v.startswith("!") else BK))
                xx += w
            s.L(x, yy + rh, x + sum(widths), yy + rh, 0.1, rgb=GR)
        s.R(x, y, sum(widths), rh * (len(rows) + 1), 0.25)
        return y + rh * (len(rows) + 1)

    i = D.INPUT
    rows = [["Wohnungen ohne el. WW (Richtwert je WE)", f"{fnum(i['P_WE'])} kW", "Annahme, Anhalt DIN 18015-1"],
            ["Gleichzeitigkeit 4 WE", fnum(i["g_WE"], 2), "Annahme"],
            ["Allgemein (Licht, Sprechanlage, RWA)", f"{fnum(i['P_ALLG'])} kW", "Annahme"],
            ["Reserve (z. B. Wallbox 11 kW)", f"{fnum(i['P_RES'])} kW", "g = 1,0"],
            ["Bemessungsleistung P", f"{fnum(c['P'])} kW", "4·P·g + Allg + Res"],
            ["Betriebsstrom Ib (cos φ 0,95)", f"{fnum(c['Ib'])} A", "OK ≤ In 80 A (NH00)"]]
    tab(25, 32, [80, 30, 70], ["Leistungsbedarf Steigleitung", "Wert", "Bemerkung"], rows,
        title="1  Leistungsbedarf")

    rows = [["-W0 HV→SV", "NHXHX-J 5×25", "NH00 gG 80", "C: 96 A", "128 ≤ 139", "OK bei C; B2 → 5×35 mm²"],
            ["-W1…-W4 SV→WV", "NHXMH-J 5×10", "D02 gG 35", "B2: 46 A", "56 ≤ 67", "OK"],
            ["Endstromkr. Licht", "NYM-J 3×1,5", "B10", "A1: 14,5 A", "14,5 ≤ 21", "OK"],
            ["Endstromkr. SD/Geräte", "NYM-J 3×2,5", "B16", "A1: 18,5 A", "23,2 ≤ 26,8", "OK"],
            ["Kochfeld 3~", "NYM-J 5×2,5", "B16 3p", "B2: 20 A", "23,2 ≤ 29", "OK"],
            ["Allgemein", "NHXMH-J 3/5×1,5", "B6/B10", "B2: 16,5 A", "14,5 ≤ 23,9", "OK"]]
    tab(25, 82, [38, 32, 24, 26, 26, 34], ["Leitung", "Typ", "Schutz In", "Iz (70 °C)", "I2 ≤ 1,45·Iz",
                                            "Ergebnis"], rows, title="2  Überlast-/Leiterschutz (OVE E 8101-4-43)")

    rows = [["-W0 (HV→SV)", f"{fnum(i['L_W0'], 0)} m / 25 mm²", f"{fnum(c['Ib'])} A",
             f"{fnum(c['du_w0'][1], 2)} %", f"max. {fnum(c['L_w0_05'], 0)} m für 0,5 %"],
            ["-W1…-W4 (SV→WV)", f"{fnum(i['L_WTOP'], 0)} m / 10 mm²", "35 A", f"{fnum(c['du_wtop'][1], 2)} %", "OK"],
            ["Endstromkreis B16", "20 m / 2,5 mm²", "16 A", f"{fnum(c['du_final'][1], 2)} %", "1-phasig"],
            ["Summe Zähler → Verbraucher", "", "", f"{fnum(c['du_wtop'][1] + c['du_final'][1], 2)} %",
             "OK ≤ 4 % (Empf. OVE E 8101)"]]
    tab(25, 132, [46, 34, 22, 20, 58], ["Abschnitt", "Länge / Querschnitt", "I", "ΔU", "Bemerkung"], rows,
        title="3  Spannungsfall (κ = 1/0,0225 m/Ωmm², x = 0,08 mΩ/m)")

    rows = [["NH00 gG 80 A → D02 gG 35 A", "2,29 : 1", "OK selektiv (≥ 1,6 : 1)"],
            ["NH00 gG 80 A → D02 gG 25 A", "3,2 : 1", "OK selektiv"],
            ["D02 gG 35 A → LS B16/B10 (WV)", "–", "teilselektiv (Herstellertabelle)"],
            ["FI Typ G im WV / FI Typ G Allgemein", "–", "OK, kein FI in Reihe"]]
    tab(215, 32, [70, 24, 86], ["Selektivität", "Verhältnis", "Ergebnis"], rows, title="4  Selektivität")

    pv_sv, pv_sv_r = D.pv_sv()
    pv_wv, pv_wv_r = D.pv_wv()
    rows = [["=SV Stockwerksverteiler", f"{fnum(pv_sv, 0)} W", "0,6", f"{fnum(pv_sv_r, 0)} W", "455 W",
             "OK"],
            ["=WVn Wohnungsverteiler", f"{fnum(pv_wv, 0)} W", "0,5", f"{fnum(pv_wv_r, 0)} W", "54 W", "OK"]]
    tab(215, 72, [52, 26, 16, 26, 26, 34], ["Verteiler", "Pv Nennstrom", "RDF", "Pv betrieb", "Pv,max",
                                             "Ergebnis"], rows,
        title="5  Erwärmung (EN 61439-3, Tab. 101: RDF; Pv ~ RDF²)")

    rows = [["=SV Vorzählerreihe", "6 × TYTAN 71 mm + Hauptklemmen", "≈ 30 TE", "33 TE"],
            ["=SV Nachzählerreihen", "6 HS (18 TE) + Allg. (8 TE) + Klemmen", "≈ 47 TE", "99 TE"],
            ["=WVn Reihe 1–3", "X0, F01, F3 / FI1 + 10 TE / FI2 + 11 TE", "35 TE", "56 TE"]]
    tab(215, 102, [42, 84, 26, 28], ["Platz", "Belegung", "belegt", "verfügbar"], rows,
        title="6  Platzbedarf (TE = 17,8 mm)")

    s.R(215, 132, 180, 52, 0.25, fill=(0.97, 0.97, 0.97))
    notes = ["Hinweise zu den Nachweisen",
             "• W0 nur mit Verlegeart C (offen auf Wand/Kabelrinne, ohne Häufung) zulässig;",
             "   bei Verlegung im Rohr (B2) oder Häufung → NHXHX-J 5×35 mm² verwenden.",
             f"• ΔU W0 bei 30 m ≈ {fnum(c['du_w0'][1], 2)} %; falls Wiener Netze für die Hauptleitung 0,5 %",
             f"   verlangt: Länge ≤ {fnum(c['L_w0_05'], 0)} m oder 5×35 mm² (TAEV Abschn. 2.4 prüfen).",
             "• Abschaltbedingung (Nullung) bei IBN durch Schleifenimpedanzmessung nachweisen;",
             "   B16: Zs ≤ 2,88 Ω, B10: Zs ≤ 4,60 Ω (0,4 s, U0 = 230 V).",
             "• RWA-Kreis -5F6 bewusst ohne FI (Betriebssicherheit), N vor FI abgegriffen."]
    for k, t in enumerate(notes):
        s.T(218, 138 + k * 5.6, t, 6.2, "B" if k == 0 else "R")

    rows = [[n, f"{fnum(zs, 2)} Ω", f"{fnum(la, 0)} m", f"{fnum(ld, 0)} m"] for n, zs, la, ld in c["lim"]]
    tab(25, 170, [62, 32, 42, 44], ["Endstromkreis", "Zs,max (0,4 s)", "Lmax Abschaltung", "Lmax ΔU ≤ 3 %"],
        rows, title=f"7  Abschaltbedingung / max. Leitungslängen (Zs am WV ≈ {fnum(i['Z_WV'], 2)} Ω, Annahme)")
    s.T(25, 196, "Maßgebend ist der Spannungsfall: Endstromkreise ≤ 23 m (1,5/2,5 mm²) bzw. ≤ 48 m (Kochfeld) "
                 "planen, sonst Querschnitt erhöhen.", 5.8, rgb=GR)

    rows = [["Ik3 am Stockwerksverteiler", f"≤ {fnum(i['IK_SV'], 0)} kA", "Annahme – bei Wiener Netze anfragen"],
            ["Ik3 am Wohnungsverteiler", f"≈ {fnum(c['ik_wv'], 1)} kA", "über 25 m NHXMH-J 5×10 mm²"],
            ["Schaltvermögen Geräte SV", "D02 50 kA · LS 10 kA", "OK ≥ Ik am SV"],
            ["Schaltvermögen Geräte WV", "FI 6 kA · LS 10 kA · FI/LS 10 kA", "OK ≥ Ik am WV"]]
    tab(215, 196, [58, 52, 70], ["Kurzschluss", "Wert", "Bemerkung"], rows, title="8  Kurzschlussfestigkeit")


# ------------------------------------------------------------------ SV Ansicht
def pg_sv_view(s):
    sc = 1 / 8.0
    x0, y0 = 40, 30
    W, H = 810 * sc, 1605 * sc
    s.T(x0, 24, "Frontansicht Stockwerksverteiler (Tür offen) – M 1:8, Einteilung schematisch", 7.5, "B")
    s.R(x0, y0, W, H, 0.5)
    s.R(x0 + 4, y0 + 4, W - 8, H - 8, 0.25)
    # Felder (Modulhöhen 46 mm)
    mh = 46 * sc
    ix, iw = x0 + 5.5, W - 11
    yy = y0 + 5.5
    fields = [("VZ", 3, "Vorzählerfeld – plombierbare Abdeckung"), ("ZR1", 9, "Zählerreihe 1"),
              ("ZR2", 9, "Zählerreihe 2"), ("FREE", 1, ""), ("NZ1", 3, "Nachzählerreihe 1"),
              ("NZ2", 3, "Nachzählerreihe 2"), ("NZ3", 3, "Nachzählerreihe 3"), ("CAB", 2, "Kabelraum")]
    te = 17.8 * sc
    pos = {}
    for key, n, txt in fields:
        h = n * mh
        if key in ("VZ",):
            s.R(ix, yy, iw, h, 0.25, fill=(1.0, 0.95, 0.85))
        elif key.startswith("ZR"):
            s.R(ix, yy, iw, h, 0.25)
        elif key.startswith("NZ"):
            s.R(ix, yy, iw, h, 0.25, fill=(0.93, 0.96, 1.0))
        else:
            s.R(ix, yy, iw, h, 0.15, dash=(0.8, 0.6))
        pos[key] = (yy, h)
        yy += h
    # Vorzählerreihe: 6 TYTAN + Hauptklemme
    yv, hv = pos["VZ"]
    xx = ix + 2
    s.R(xx, yv + 4, 4 * te, hv - 8, 0.2, fill=WH)
    s.T(xx + 2 * te, yv + hv / 2 + 0.9, "-X0", 4.6, "B", "c")
    xx += 4 * te + 1
    for z in D.ZP:
        w = 71 * sc
        s.R(xx, yv + 3, w, hv - 6, 0.2, fill=WH)
        s.T(xx + w / 2, yv + hv / 2 + 0.9, f"-{z['n']}F1", 4.4, "B", "c")
        xx += w + 0.6
    s.T(ix + iw - 1, yv + 3.2, "plombiert", 4.4, "I", "r", rgb=RD)
    # Zählerreihen
    for r, key in enumerate(("ZR1", "ZR2")):
        yz, hz = pos[key]
        for k in range(3):
            n = r * 3 + k + 1
            z = D.ZP[n - 1]
            cx = ix + iw / 6 * (2 * k + 1)
            tw, th = 210 * sc, 370 * sc
            s.R(cx - tw / 2, yz + 1.2, tw, th, 0.25, fill=(0.96, 0.96, 0.96), dash=None if z["fitted"] else (1, .7))
            s.R(cx - 8.5, yz + 4, 17, 21, 0.3, fill=WH if z["fitted"] else None,
                dash=None if z["fitted"] else (1, .7))
            s.T(cx, yz + 12, "Wh", 7, "B", "c", rgb=BK if z["fitted"] else GR)
            s.T(cx, yz + 16.5, f"-{n}P1", 5, "B", "c")
            s.R(cx - 9.3, yz + 28.5, 18.6, 3.6, 0.2, fill=(1, 1, 0.9))
            s.T(cx, yz + 31.4, f"-{n}X1", 4.2, "R", "c")
            s.T(cx, yz + 39, z["name"], 5.6, "B", "c")
            s.T(cx, yz + 42.5, "Reserve" if not z["fitted"] else "", 4.5, "R", "c", rgb=GR)
    # Nachzählerreihen
    yn, hn = pos["NZ1"]
    xx = ix + 2
    for z in D.ZP:
        w = 3 * te
        s.R(xx, yn + 3, w, hn - 6, 0.2, fill=WH)
        s.T(xx + w / 2, yn + hn / 2 + 0.9, f"-{z['n']}Q1", 4.4, "B", "c")
        xx += w + 0.3
    yn2, hn2 = pos["NZ2"]
    xx = ix + 2
    for b, wte in (("-5F2", 3), ("-5F3", 1), ("-5K1", 1), ("-5F4", 1), ("-5F5", 1), ("-5F6", 1)):
        w = wte * te
        s.R(xx, yn2 + 3, w, hn2 - 6, 0.2, fill=WH)
        s.T(xx + w / 2 + (0 if wte > 1 else 1.2), yn2 + hn2 / 2 + 0.9, b, 3.6, "B", "c", rot=0 if wte > 1 else 90)
        xx += w + 0.3
    yn3, hn3 = pos["NZ3"]
    xx = ix + 2
    for n in (1, 2, 3, 4, 6):
        w = 5 * 12.2 * sc
        s.R(xx, yn3 + 4, w, hn3 - 8, 0.2, fill=(0.92, 0.92, 0.92))
        s.T(xx + w / 2, yn3 + hn3 / 2 + 0.9, f"-{n}X2", 3.8, "B", "c")
        xx += w + 0.8
    w = 14 * 6.2 * sc
    s.R(xx, yn3 + 4, w, hn3 - 8, 0.2, fill=(0.92, 0.92, 0.92))
    s.T(xx + w / 2, yn3 + hn3 / 2 + 0.9, "-5X2", 3.8, "B", "c")
    # Maße
    s.L(x0, y0 + H + 6, x0 + W, y0 + H + 6, 0.2)
    s.T(x0 + W / 2, y0 + H + 5, "810", 5.5, "R", "c")
    s.L(x0 - 6, y0, x0 - 6, y0 + H, 0.2)
    s.T(x0 - 7, y0 + H / 2, "1605", 5.5, "R", "c", rot=90)
    # Legende rechts
    lx = 160
    s.T(lx, 32, "Bestückung / Belegung", 8, "B")
    rows = [("Vorzählerfeld (ungezählt, plombiert)", ""),
            ("-X0", "Hauptabzweigklemme 4p 125 A IKB04035 + PE SI022270"),
            ("-1F1 … -6F1", "TYTAN T D02 3p 63 A IS503031, Einsätze 35 A / 25 A (Allg.)"),
            ("Zählerreihen 1 + 2 (Zählertafeln OVE E 8640)", ""),
            ("-1P1 … -5P1", "Smart Meter Wiener Netze (Beistellung), Platz 6 Reserve"),
            ("-1X1 … -6X1", "Zählersteckklemme 3p+2N 50 A IL900189 + Abdeckung IL900191"),
            ("Nachzählerreihen (gezählt)", ""),
            ("-1Q1 … -6Q1", "Hauptschalter 40 A 3p BZ900243 (je 3 TE)"),
            ("-5F2 … -5F6, -5K1", "Allgemein: FI 40/4/0,03 A,G, LS B10/B6/B16, Treppenlicht"),
            ("-1X2 … -6X2", "Abgangsklemmen TB 16 I / TB 16 I BU / TB 16-PE I"),
            ("-5X2", "Allgemein-Klemmen TB 4 I / TB 4 I BU / TB 10-PE I"),
            ("Gehäuse", "IL166333WS: H1605 × B810 × T250, Putznische 1575 × 800"),
            ("", "Tür: Drehriegel mit Doppelbart (Option Zylinder IL900200-F)"),
            ("", "Kabeleinführung oben (W0) / unten (W1–W6) lt. Baustelle")]
    for k, (a, b) in enumerate(rows):
        yy = 40 + k * 6.3
        if not b and a:
            s.R(lx, yy - 4.2, 240, 5.4, 0.1, fill=LG)
            s.T(lx + 1.5, yy, a, 6.3, "B")
        else:
            s.T(lx + 1.5, yy, a, 6.2, "B")
            s.T(lx + 42, yy, b, 6.2)
    s.R(lx, 140, 240, 52, 0.25, fill=(0.97, 0.97, 0.97))
    tx = ["Montage- und Ausführungshinweise",
          "• Unterputz in Wandnische Stiegenhaus; Zählermitte 0,8 … 1,8 m über FFB (Ablesehöhe).",
          "• Vorzählerbereich: Verdrahtung H07V-K 10 mm² kurz- und erdschlusssicher, getrennt geführt.",
          "• Zählersteckklemmen nach OVE R 21:2019, Plombierung durch Wiener Netze.",
          "• Beschriftung je Zählerplatz: Top-Nr. auf Zählertafel, Lasttrenner und Hauptschalter.",
          "• Freie TE mit Leerplatzabdeckung IL900251 schließen (IP20C).",
          "• Brandschutzanforderung Stiegenhaus (OIB-RL 2) mit Brandschutzplaner klären.",
          "• Feldeinteilung = Konfiguration IL166333WS; exakte Lage lt. Schrack-Fertigung."]
    for k, t in enumerate(tx):
        s.T(lx + 2, 146 + k * 5.7, t, 6.2, "B" if k == 0 else "R")


# ------------------------------------------------------------------ SV Anspeisung
def pg_sv_feed(s):
    s.T(25, 22, "Anspeisung Stockwerksverteiler vom Hauptverteiler (HV)", 8, "B")
    s.T(25, 27, "3/N/PE AC 400/230 V, 50 Hz · TN-C-S (Nullung) · PEN-Aufteilung im HV", 6.2, rgb=GR)
    xs = [50, 56, 62, 68, 74]
    s.R(40, 32, 60, 15, 0.3, dash=(1.2, 0.8))
    s.T(42, 37, "HV – Abgang NH00 gG 80 A", 6.2, "B")
    s.T(42, 41.5, "(Plan Hauptverteiler / Bestand)", 5.5, rgb=GR)
    for x in xs:
        s.L(x, 47, x, 120)
    cablemark(s, xs, 72, "-W0", "NHXHX-J 5×25 mm² (halogenfrei)",
              f"L ≈ {fnum(D.INPUT['L_W0'], 0)} m · Verlegeart C · Ib ≈ {fnum(D.calc()['Ib'])} A",
              ["BN", "BK", "GY", "BU", "GNYE"])
    for x, n in zip(xs, ("L1", "L2", "L3", "N", "PE")):
        s.T(x - 0.8, 54, n, 5.2, "R", "r")
    # X0
    s.R(46, 118, 26, 6, 0.3, fill=WH)
    s.T(59, 122.3, "IKB04035 4p 125 A", 4.4, "R", "c")
    s.R(72.5, 118, 4, 6, 0.3, fill=(0.85, 0.95, 0.85))
    for x in xs:
        s.term(x, 118)
    label(s, 80, 115, ["-X0", "Hauptabzweigklemme 4-polig 125 A", "Eingang 1×35 mm², Abgänge 10×16 mm²",
                       "PE: Hauptklemme 1×6×16 mm² SI022270"])
    # Sammelschienen nach rechts
    ys = [140, 146, 152, 158, 164]
    for x, y in zip(xs, ys):
        s.L(x, 124, x, y)
        s.dot(x, y)
    for x, y, n in zip(xs, ys, ("L1", "L2", "L3", "N", "PE")):
        bus(s, y, x, XE, n, rref=xr("sv_zp1", 0))
    s.T(90, 137, "Verteilung zu den Vorzählersicherungen -1F1 … -6F1, H07V-K 10 mm², kurz- und erdschlusssicher",
        5.6, rgb=GR)
    # ungezählter Bereich
    s.R(36, 108, 372, 62, 0.3, dash=(2.0, 1.2), rgb=RD)
    s.T(404, 112.5, "ungezählter Bereich – plombiert (Wiener Netze)", 5.6, "I", "r", rgb=RD)
    # Gehäuse-PE
    s.L(74, 164, 74, 190)
    s.dot(74, 164)
    s.L(70, 190, 78, 190, 0.35)
    s.L(71.5, 191.5, 76.5, 191.5, 0.3)
    s.L(73, 193, 75, 193, 0.3)
    label(s, 80, 186, ["PE-Gehäuse", "Schutzleiterbolzen M6 Mauerwanne / Tür", "(Potentialausgleich lt. HV)"])
    # Hinweisbox
    c = D.calc()
    s.R(160, 182, 245, 64, 0.25, fill=(0.97, 0.97, 0.97))
    tx = ["Auslegung Anspeisung",
          f"• Leistungsbedarf P ≈ {fnum(c['P'])} kW → Ib ≈ {fnum(c['Ib'])} A ≤ In = 80 A (NH00 gG im HV)",
          "• Leiterschutz -W0: Iz = 96 A (Verlegeart C, 70 °C) ≥ 80 A; I2 = 128 A ≤ 1,45·Iz = 139 A OK",
          "  bei Verlegeart B2 (Rohr) oder Häufung → 5×35 mm² erforderlich",
          f"• Spannungsfall -W0 (30 m): ΔU ≈ {fnum(c['du_w0'][1], 2)} % – Grenzwert Hauptleitung lt. TAEV/Wiener Netze prüfen",
          "• Selektivität NH00 gG 80 A → D02 gG 35 A: 2,29 : 1 OK (≥ 1,6 : 1)",
          "• Hauptabzweigklemme -X0 und Verdrahtung bis zur Vorzählersicherung im plombierten Bereich",
          "• Kein Überspannungsableiter im SV (Typ 1 im HV vorausgesetzt, Typ 2 je Wohnungsverteiler)"]
    for k, t in enumerate(tx):
        s.T(163, 189 + k * 6.6, t, 6.3, "B" if k == 0 else "R")


# ------------------------------------------------------------------ SV Zählerplätze
def label_r(s, x, y, lines, sz=5.8, lh=2.9):
    for k, t in enumerate(lines):
        if t:
            s.T(x, y + k * lh, t, sz + (0.9 if k == 0 else 0), "B" if k == 0 else "R", "r",
                rgb=BK if k == 0 else (0.15, 0.15, 0.15))


def kterm(s, x, y, t, side="l"):
    if t.startswith("N"):
        s.R(x - 1.1, y - 1.1, 2.2, 2.2, 0.22, fill=(0.8, 0.88, 1.0))
    elif t.startswith("PE"):
        s.R(x - 1.1, y - 1.1, 2.2, 2.2, 0.22, fill=(0.8, 0.95, 0.8))
    else:
        s.term(x, y)
    if side == "l":
        s.T(x - 1.4, y - 1.7, t, 4.3, "R", "r")
    else:
        s.T(x + 1.5, y - 1.7, t, 4.3)


def pg_sv_zp(s, zps, lref, rref):
    ys = {"L1": 30, "L2": 35, "L3": 40, "N": 45, "PE": 50}
    xlast = 52 + 120 * (len(zps) - 1) + 24
    for n, y in ys.items():
        bus(s, y, 34, XE if rref else xlast, n, lref=lref, rref=rref)
    for k, z in enumerate(zps):
        xb = 52 + 120 * k
        xl = [xb, xb + 6, xb + 12]
        xn, xp = xb + 18, xb + 24
        n = z["n"]
        s.T(xb - 6, 19.5, f"Zählerplatz {n} – {z['name']}", 7.5, "B")
        s.T(xb - 6, 24.5, z["desc"], 5.8, rgb=GR)
        for x, b in zip(xl + [xn, xp], ("L1", "L2", "L3", "N", "PE")):
            s.dot(x, ys[b])
        s.T(xb + 30, 56.5, "H07V-K 10 mm², kurz- und erdschlusssicher", 5.0, rgb=GR)
        for x, b in zip(xl, ("L1", "L2", "L3")):
            s.L(x, ys[b], x, 60)
        mids = [fuse_pole(s, x, 60, 20, ("1", "3", "5")[i], ("2", "4", "6")[i], z["fitted"])
                for i, x in enumerate(xl)]
        linkage(s, mids)
        label(s, xb + 30, 63, [f"-{n}F1", "Vorzählersicherung (plombierbar)",
                               f"NEOZED D02 3×{z['fuse']} A gG" + ("" if z["fitted"] else " – nicht bestückt"),
                               "TYTAN T 63 A · IS503031",
                               "Passhülse " + ("IS504136" if z["fuse"] == 35 else "IS504135")
                               + (" · Einsatz IS504037" if z["fuse"] == 35 else " · Einsatz IS504036")])
        for x in xl:
            s.L(x, 80, x, 98)
        s.L(xn, ys["N"], xn, 98)
        s.L(xp, ys["PE"], xp, 172 if z["kind"] == "allg" else 222)
        for x in xl + [xn]:
            s.term(x, 96)
        s.T(xb - 2, 97.2, f"-{n}X1", 5.6, "B", "r")
        label(s, xb + 30, 91, [f"-{n}X1", "Zählersteckklemme 3p+2N 50 A", "IL900189 + Abdeckung IL900191"])
        meter(s, xl + [xn], 98, 22, z["fitted"])
        label(s, xb + 30, 106, [f"-{n}P1", "Smart Meter 3~ direkt" if z["fitted"] else "Zählerplatz Reserve",
                                "Beistellung Wiener Netze" if z["fitted"] else "Zähler nicht montiert"])
        for x in xl + [xn]:
            s.L(x, 120, x, 128)
            s.term(x, 124)
        s.T(xb - 2, 125.2, f"-{n}X1", 5.6, "B", "r")
        s.L(xb - 10, 131, xb + 96, 131, 0.3, dash=(2.0, 1.2), rgb=RD)
        s.T(xb + 96, 129.6, "plombiert  /  Kundenanlage", 4.6, "I", "r", rgb=RD)
        for x in xl:
            s.L(x, 128, x, 134)
        lts(s, xl, 134, 16)
        label(s, xb + 30, 140, [f"-{n}Q1", "Hauptschalter 40 A, 3-polig", "BZ900243"])
        if z["kind"] == "allg":
            for x in xl:
                s.L(x, 150, x, 172)
            s.L(xn, 128, xn, 172)
            for x in xl + [xn]:
                xref_down(s, x, 172, "")
            xref_down(s, xp, 172, "")
            s.T(xb + 12, 184, f"→ Allgemein-Abgänge  Blatt {xr('sv_allg', 0)}", 6.0, "B", "c", rgb=BL)
            s.T(xb + 12, 189, "5Q1: L1, L2, L3, N und PE", 5.2, "R", "c", rgb=BL)
            continue
        for x in xl:
            s.L(x, 150, x, 222)
        s.L(xn, 128, xn, 222)
        yk = 170
        for i, x in enumerate(xl):
            kterm(s, x, yk, str(i + 1))
        kterm(s, xn, yk, "N")
        kterm(s, xp, yk, "PE", side="r")
        s.T(xb - 4, yk + 1.0, f"-{n}X2", 5.6, "B", "r")
        label(s, xb + 30, 166, [f"-{n}X2", "TB 16 I · TB 16 I BU (N trennbar)", "TB 16-PE I"])
        allx = xl + [xn, xp]
        if z["kind"] == "top":
            cablemark(s, allx, 188, f"-W{n}", "NHXMH-J 5×10 mm²",
                      f"≤ {fnum(D.INPUT['L_WTOP'], 0)} m · ΔU ≈ {fnum(D.calc()['du_wtop'][1], 1)} %",
                      ["BN", "BK", "GY", "BU", "GNYE"])
            key = "wv%d_feed" % z["top"]
            s.R(xb - 6, 222, 42, 14, 0.3)
            s.T(xb + 15, 227.5, f"Wohnungsverteiler Top {z['top']}", 6.0, "B", "c")
            s.T(xb + 15, 232.5, f"=WV{z['top']} -X0  → Blatt {xr(key, 0)}", 5.4, "R", "c", rgb=BL)
        else:
            cablemark(s, allx, 188, "-W6", "vorbereitet – nicht verlegt", "Reserve-Abgang", None)
            s.R(xb - 6, 222, 42, 14, 0.3, dash=(1.2, 0.8))
            s.T(xb + 15, 227.5, "Reserve", 6.0, "B", "c")
            s.T(xb + 15, 232.5, "Wallbox / WP / PV", 5.4, "R", "c", rgb=GR)


# ------------------------------------------------------------------ SV Allgemein
def pg_sv_allg(s):
    s.T(25, 21, "Zählerplatz 5 – Allgemein: Abgänge (gezählter Bereich nach -5Q1)", 7.5, "B")
    ys = {"L1": 30, "L2": 35, "L3": 40, "N": 45, "PE": 50}
    lref = xr("sv_zp2", 4)
    xf = [50, 56, 62, 68]
    bus(s, ys["L1"], 34, 50, "L1", lref=lref)
    bus(s, ys["L2"], 34, 56, "L2", lref=lref)
    bus(s, ys["L3"], 34, 340, "L3", lref=lref)
    bus(s, ys["N"], 34, 346, "N", lref=lref)
    bus(s, ys["PE"], 34, 352, "PE", lref=lref)
    for x, b in zip(xf, ("L1", "L2", "L3", "N")):
        s.dot(x, ys[b])
        s.L(x, ys[b], x, 62)
    yb = rcd(s, xf, 62, 16)
    label(s, 74, 64, ["-5F2", "FI-Schutzschalter 40 A, 4-polig", "30 mA, Typ A, G", "AR634103-A"])
    ysf = {"L1": 100, "L2": 105, "L3": 110, "N": 115}
    for x, b in zip(xf, ("L1", "L2", "L3", "N")):
        s.L(x, yb, x, ysf[b])
        s.dot(x, ysf[b])
        s.L(x, ysf[b], 300, ysf[b], 0.3)
        s.T(301, ysf[b] + 0.9, b + "'", 4.8, rgb=GR)
    s.T(76, 97.5, "nach FI (Zusatzschutz 30 mA)", 5.2, rgb=GR)
    yk = 196
    # ---------------- Stiegenhaus: Klemmen 2 (Dauer) 1 (Licht) 3 (Taster) N1 PE1
    xF = 110
    s.dot(xF, ysf["L1"])
    s.L(xF, ysf["L1"], xF, 122)
    mcb(s, [xF], 122, 16, ("1",), ("2",))
    label_r(s, 99, 124, ["-5F3", "LS B10, 1-polig", "AM018110"])
    s.L(xF, 138, xF, 150)
    s.dot(xF, 144)
    s.L(xF, 144, 100, 144)
    s.L(100, 144, 100, 236)
    s.R(104, 150, 22, 16, 0.3, fill=WH)
    s.T(115, 157.0, "-5K1", 6.2, "B", "c")
    s.T(115, 160.6, "Treppenlicht-", 4.6, "R", "c")
    s.T(115, 163.6, "zeitschalter", 4.6, "R", "c")
    s.T(110.6, 152.6, "L", 4.2)
    s.T(118.6, 152.6, "N", 4.2)
    s.T(110.6, 165.2, "LA", 4.2)
    s.T(120.6, 165.2, "T", 4.2)
    s.dot(118, ysf["N"])
    s.L(118, ysf["N"], 118, 150)
    s.L(110, 166, 110, 236)
    s.L(120, 166, 120, 236)
    s.dot(128, ysf["N"])
    s.L(128, ysf["N"], 128, 236)
    s.dot(134, ys["PE"])
    s.L(134, ys["PE"], 134, 236)
    label(s, 138, 152, ["-5K1", "Treppenlichtzeitschalter VOWA+", "1–12 min, 16 A, Vorwarnung", "BZ327360"])
    for x, t in ((100, "2"), (110, "1"), (120, "3"), (128, "N1"), (134, "PE1")):
        kterm(s, x, yk, t, side="r" if t.startswith("PE") else "l")
    s.T(94, yk + 1.0, "-5X2", 5.6, "B", "r")
    cablemark(s, [100, 110, 120, 128, 134], 212, "-W5.1", "NHXMH-J 5×1,5 mm²",
              "≈25 m · BN Dauer · BK Licht · GY Taster", ["BN", "BK", "GY", "BU", "GNYE"])
    s.T(117, 241, "Stiegenhauslicht (geschaltet) · SiBel + Taster (Dauerphase)", 5.0, "R", "c")
    # ---------------- Sprechanlage
    x = 196
    s.dot(x, ysf["L2"])
    s.L(x, ysf["L2"], x, 122)
    mcb(s, [x], 122, 16, ("1",), ("2",))
    label(s, x + 15, 124, ["-5F4", "LS B6, 1-polig", "AM018106"])
    s.L(x, 138, x, 236)
    s.dot(x + 6, ysf["N"])
    s.L(x + 6, ysf["N"], x + 6, 236)
    s.dot(x + 12, ys["PE"])
    s.L(x + 12, ys["PE"], x + 12, 236)
    for xx, t in ((x, "4"), (x + 6, "N2"), (x + 12, "PE2")):
        kterm(s, xx, yk, t, side="r" if t.startswith("PE") else "l")
    cablemark(s, [x, x + 6, x + 12], 212, "-W5.2", "NHXMH-J 3×1,5 mm²", "≈15 m", ["BN", "BU", "GNYE"])
    s.T(x + 6, 241, "Netzgerät Sprechanlage", 5.0, "R", "c")
    # ---------------- Reserve
    x = 262
    s.dot(x, ysf["L3"])
    s.L(x, ysf["L3"], x, 122)
    mcb(s, [x], 122, 16, ("1",), ("2",))
    label(s, x + 15, 124, ["-5F5", "LS B16, 1-polig", "AM018116"])
    s.L(x, 138, x, yk)
    s.dot(x + 6, ysf["N"])
    s.L(x + 6, ysf["N"], x + 6, yk)
    s.dot(x + 12, ys["PE"])
    s.L(x + 12, ys["PE"], x + 12, yk)
    for xx, t in ((x, "6"), (x + 6, "N4"), (x + 12, "PE4")):
        kterm(s, xx, yk, t, side="r" if t.startswith("PE") else "l")
    s.R(x - 4, 206, 20, 10, 0.25, dash=(1, .7))
    s.T(x + 6, 212.5, "Reserve", 5.4, "R", "c", rgb=GR)
    # ---------------- RWA (vor FI)
    x = 340
    s.dot(x, ys["L3"])
    s.L(x, ys["L3"], x, 122)
    mcb(s, [x], 122, 16, ("1",), ("2",))
    label(s, x + 15, 124, ["-5F6", "LS B10, 1-polig", "AM018110", "ohne FI – RWA", "nicht abschalten!"])
    s.L(x, 138, x, 236)
    s.dot(x + 6, ys["N"])
    s.L(x + 6, ys["N"], x + 6, 236)
    s.dot(x + 12, ys["PE"])
    s.L(x + 12, ys["PE"], x + 12, 236)
    for xx, t in ((x, "5"), (x + 6, "N3"), (x + 12, "PE3")):
        kterm(s, xx, yk, t, side="r" if t.startswith("PE") else "l")
    cablemark(s, [x, x + 6, x + 12], 212, "-W5.3", "NHXMH-J 3×1,5 mm²", "≈20 m", ["BN", "BU", "GNYE"])
    s.T(x + 6, 241, "RWA-Zentrale (N vor FI!)", 5.0, "R", "c")
    s.R(213, 221, 124, 17, 0.25, fill=(1.0, 0.95, 0.95))
    s.T(215, 226, "RWA-Stromkreis vor dem FI abgegriffen (L3 und N vor -5F2),", 5.6, rgb=RD)
    s.T(215, 230.5, "damit Fehler im Allgemeinbereich die RWA nicht abschalten.", 5.6, rgb=RD)
    s.T(215, 235, "Kennzeichnung „RWA – nicht abschalten“ an -5F6 anbringen.", 5.6, rgb=RD)


# ------------------------------------------------------------------ WV Ansicht
def pg_wv_view(s, t):
    sc = 1 / 3.5
    W, H = 346 * sc, 717 * sc
    x0, y0 = 40, 26
    s.T(x0, 22, f"Frontansicht Wohnungsverteiler Top {t} (Abdeckung offen) – M 1:3,5", 7.5, "B")
    s.R(x0, y0, W, H, 0.5)
    s.R(x0 + 3, y0 + 3, W - 6, H - 6, 0.25)
    te = 17.8 * sc
    rows_y = [y0 + 34, y0 + 74, y0 + 114, y0 + 154]
    rx = x0 + (W - 14 * te) / 2
    devs = [
        [("-X0", 4, (0.92, 0.92, 0.92)), ("-F01", 2, (1, 0.93, 0.8)), ("-F3", 2, (0.9, 0.95, 1))],
        [("-F1", 3, (0.9, 0.95, 1))] + [(c["bmk"], c["poles"], WH) for c in D.WV_G1],
        [("-F2", 3, (0.9, 0.95, 1))] + [(c["bmk"], c["poles"], WH) for c in D.WV_G2],
        [],
    ]
    for r, (yr, row) in enumerate(zip(rows_y, devs)):
        s.L(rx - 3, yr + 9, rx + 14 * te + 3, yr + 9, 0.6, rgb=GR)
        s.T(rx - 4, yr + 10, f"R{r + 1}", 5, "B", "r", rgb=GR)
        xx = rx
        for b, n, col in row:
            w = n * te
            s.R(xx, yr, w, 18, 0.22, fill=col)
            s.T(xx + w / 2, yr + 10.5, b.replace("-", ""), 3.6 if n == 1 else 4.4, "B", "c", rot=90 if n == 1 else 0)
            xx += w
        if not row:
            s.R(rx, yr, 14 * te, 18, 0.2, dash=(0.8, 0.6), fill=WH)
            s.T(rx + 7 * te, yr + 6.0, "Reserve 14 TE", 5, "R", "c", rgb=GR)
    s.R(x0 + 6, y0 + H - 22, 20, 8, 0.2, fill=(0.8, 0.88, 1.0))
    s.T(x0 + 16, y0 + H - 17, "N-Klemmen", 4.4, "R", "c")
    s.R(x0 + W - 26, y0 + H - 22, 20, 8, 0.2, fill=(0.8, 0.95, 0.8))
    s.T(x0 + W - 16, y0 + H - 17, "PE-Klemmen", 4.4, "R", "c")
    s.L(x0, y0 + H + 5, x0 + W, y0 + H + 5, 0.2)
    s.T(x0 + W / 2, y0 + H + 4, "346", 5.5, "R", "c")
    s.L(x0 - 5, y0, x0 - 5, y0 + H, 0.2)
    s.T(x0 - 6, y0 + H / 2, "717", 5.5, "R", "c", rot=90)
    # Tabelle
    lx = 160
    s.T(lx, 30, "Belegung (Phasenschiene L1-L2-L3 umlaufend)", 8, "B")
    hdr = ["Reihe", "BMK", "Gerät", "Art.-Nr.", "TE", "Phase"]
    widths = [14, 18, 110, 30, 10, 20]
    y = 35
    s.R(lx, y, sum(widths), 5.4, 0.25, fill=HD)
    xx = lx
    for w, h in zip(widths, hdr):
        s.T(xx + 1, y + 3.9, h, 6, "B")
        xx += w
    y += 5.4
    rows = [("R1", "-X0", "Hauptabzweigklemme 4p 125 A (L1-L3, N)", "IKB04035", "4", "–"),
            ("R1", "-F01", "Überspannungsableiter Typ 2 (+T3) 3+1, 275 V", "IS111314", "2", "L1-L3,N"),
            ("R1", "-F3", "FI/LS B16 1+N 30 mA Typ A,G – Kühlgerät", "BO218616", "2", "L2"),
            ("R2", "-F1", "FI 40 A 4p 30 mA Typ A,G – Gruppe 1", "AR634103-A", "3", "L1-L3,N")]
    for c in D.WV_G1:
        rows.append(("R2", c["bmk"], f"LS B{c['In']} {c['poles']}p – {c['short']}", c["art"].rstrip("-"),
                     str(c["poles"]), c["ph"]))
    rows.append(("R3", "-F2", "FI 40 A 4p 30 mA Typ A,G – Gruppe 2", "AR634103-A", "3", "L1-L3,N"))
    for c in D.WV_G2:
        rows.append(("R3", c["bmk"], f"LS B{c['In']} {c['poles']}p – {c['short']}", c["art"].rstrip("-"),
                     str(c["poles"]), c["ph"]))
    rows.append(("R4", "–", "Reserve (14 TE)", "–", "–", "–"))
    for r in rows:
        xx = lx
        for w, v in zip(widths, r):
            s.TF(xx + 1, y + 3.7, v, w - 2, 5.8)
            xx += w
        s.L(lx, y + 5.0, lx + sum(widths), y + 5.0, 0.1, rgb=GR)
        y += 5.0
    s.R(lx, 35, sum(widths), y - 35, 0.25)
    pv, pvr = D.pv_wv()
    s.T(lx, y + 7, f"Belegt 35 TE von 56 TE (Reserve 21 TE ≈ 37 %) · Verlustleistung ≈ {fnum(pvr, 0)} W "
                   f"(RDF 0,5) ≤ 54 W zulässig · Gehäuse BK0850042 + BK0851042F", 6.2, "B")
    s.T(lx, y + 12.5, "Kühlgerät über eigenen FI/LS (-F3), damit ein FI-Auslösen der Gruppen die Kühlung nicht "
                      "unterbricht.", 6.0)


# ------------------------------------------------------------------ WV Anspeisung
def pg_wv_feed(s, t):
    s.T(25, 21, f"Anspeisung Wohnungsverteiler Top {t} · Überspannungsschutz · Kühlgerät", 7.5, "B")
    xs = [44, 50, 56, 62, 68]
    s.T(56, 26.5, f"von =SV -{t}X2  (Blatt {xr('sv_zp1' if t <= 3 else 'sv_zp2', 52.0 + 120 * ((t - 1) % 3))})",
        5.6, "R", "c", rgb=BL)
    for x in xs:
        s.L(x, 30, x, 60)
    cablemark(s, xs, 42, f"-W{t} (=SV)", "NHXMH-J 5×10 mm²", "≤ 25 m · D02 35 A im SV",
              ["BN", "BK", "GY", "BU", "GNYE"])
    s.R(41, 60, 24, 6, 0.3, fill=WH)
    s.T(53, 64.3, "IKB04035", 4.4, "R", "c")
    for x in xs[:4]:
        s.term(x, 60)
    s.R(xs[4] - 1.1, 58.9, 2.2, 2.2, 0.22, fill=(0.8, 0.95, 0.8))
    s.T(xs[4] + 1.6, 59.5, "XPE:1", 4.4)
    label(s, 76, 58, ["-X0", "Hauptabzweigklemme 4p 125 A", "IKB04035"])
    ys = {"L1": 82, "L2": 88, "L3": 94, "N": 100, "PE": 106}
    for x, (b, y) in zip(xs, ys.items()):
        s.L(x, 66 if b != "PE" else 61, x, y)
        s.dot(x, y)
        bus(s, y, x, XE, b, rref=xr(f"wv{t}_g1", 0))
    # SPD
    xsp = [110, 116, 122, 128]
    for x, b in zip(xsp, ("L1", "L2", "L3", "N")):
        s.dot(x, ys[b])
        s.L(x, ys[b], x, 124)
    ype = spd(s, xsp, 124)
    s.L(128, ype, 128, ype + 6)
    s.L(128, ype + 6, 102, ype + 6)
    s.L(102, ype + 6, 102, ys["PE"])
    s.dot(102, ys["PE"])
    s.R(102 - 1.1, ype + 6 - 1.1, 2.2, 2.2, 0.22, fill=(0.8, 0.95, 0.8))
    s.T(100.5, ype + 5.0, "XPE:2", 4.4, "R", "r")
    s.T(104, ype + 10.0, "PE 6 mm², Anschlussleitung ≤ 0,5 m", 4.6, rgb=GR)
    label(s, 134, 126, ["-F01", "Überspannungsableiter-Set 3+1", "Typ 2 (+T3), 275 V, 20 kA", "IS111314",
                        "TN-S, vor den FI-Schutzschaltern"])
    # FI/LS Kühlgerät
    xL, xN, xP = 220, 226, 232
    s.dot(xL, ys["L2"])
    s.dot(xN, ys["N"])
    s.L(xL, ys["L2"], xL, 124)
    s.L(xN, ys["N"], xN, 124)
    yo = rcd(s, [xL, xN], 124, 16, ("1", "N"), ("2", "N"), ls_first=True)
    label(s, 238, 126, ["-F3", "FI/LS B16, 1+N, 30 mA", "Typ A, G · 10 kA", "BO218616"])
    s.L(xL, yo, xL, 210)
    s.L(xN, yo, xN, 210)
    s.dot(xP, ys["PE"])
    s.L(xP, ys["PE"], xP, 210)
    s.R(xP - 1.1, 168.9, 2.2, 2.2, 0.22, fill=(0.8, 0.95, 0.8))
    s.T(xP + 1.6, 169.6, "XPE:3", 4.4)
    cablemark(s, [xL, xN, xP], 186, "-W3", "NYM-J 3×2,5 mm²", "≈10 m", ["BN", "BU", "GNYE"])
    consumer(s, xN, 212, "geraet", "")
    s.T(xN, 225, "Kühl-/Gefriergerät", 5.6, "B", "c")
    # Hinweise
    s.R(268, 120, 137, 60, 0.25, fill=(0.97, 0.97, 0.97))
    tx = ["Hinweise",
          f"• Hauptschalter Top {t}: -{t}Q1 im Stockwerksverteiler",
          "• Vorsicherung: NEOZED D02 3×35 A gG (plombiert)",
          "• Zusatzschutz: alle Endstromkreise FI 30 mA",
          "• FI Typ G (kurzzeitverzögert) gegen Fehlauslösung",
          "• N der FI-Gruppen getrennt (-XN1 / -XN2)",
          "• Ableiter vor den FI, Anschlussleitung ≤ 0,5 m",
          "• PE-Klemmleiste -XPE im Gehäuse (Lieferumfang)"]
    for k, tt in enumerate(tx):
        s.T(271, 127 + k * 6.6, tt, 6.0, "B" if k == 0 else "R")


# ------------------------------------------------------------------ WV FI-Gruppe
def pg_wv_group(s, t, circuits, fi_bmk=None, bus_lref=None, bus_rref=None, pe_lref=None, pe_rref=None,
                fi_lref=None, fi_rref=None, xn_start=2, xpe_start=4):
    ysb = {"L1": 30, "L2": 34, "L3": 38, "N": 42, "PE": 46}
    ysf = {"L1": 98, "L2": 103, "L3": 108, "N": 113}
    xn_name = "XN1" if circuits[0]["bmk"].startswith("-F1.") else "XN2"
    # Lage der Stromkreise vorab bestimmen
    x = 100 if fi_bmk else 52
    pos = []
    for cc in circuits:
        pos.append(x)
        x += 40 if cc["poles"] == 3 else 32
    xend = pos[-1] + (20 if circuits[-1]["poles"] == 3 else 10) + 3
    if fi_bmk:
        xf = [44, 50, 56, 62]
        for (b, y), xx in zip(list(ysb.items())[:4], xf):
            bus(s, y, 34, XE if bus_rref else xx, b, lref=bus_lref, rref=bus_rref)
            s.dot(xx, y)
            s.L(xx, y, xx, 58)
        bus(s, ysb["PE"], 34, XE if (pe_rref or bus_rref) else xend, "PE", lref=bus_lref, rref=pe_rref or bus_rref)
        yb = rcd(s, xf, 58, 16)
        label(s, 68, 62, [fi_bmk, "FI 40 A, 4-polig, 30 mA", "Typ A, G (kurzzeitverz.)", "AR634103-A"])
        for xx, b in zip(xf, ("L1", "L2", "L3", "N")):
            s.L(xx, yb, xx, ysf[b])
            s.dot(xx, ysf[b])
        s.term(62, 90)
        s.T(63.5, 90.8, f"{xn_name}:1", 4.4)
        starts = dict(zip(("L1", "L2", "L3", "N"), xf))
    else:
        bus(s, ysb["PE"], 34, XE if pe_rref else xend, "PE", lref=pe_lref, rref=pe_rref)
        starts = dict.fromkeys(("L1", "L2", "L3", "N"), 34)
        for b, y in ysf.items():
            s.T(33.2, y + 0.9, f"{b}' / {fi_lref}", 4.8, "R", "r", rgb=BL)
    for b, y in ysf.items():
        s.L(starts[b], y, XE if fi_rref else xend, y, 0.3)
        if fi_rref:
            s.T(XE + 0.8, y + 0.9, f"{b}' / {fi_rref}", 4.8, rgb=BL)
    s.T((68 if fi_bmk else 36), 95.4, f"nach FI – Neutralleiter über -{xn_name}", 5.0, rgb=GR)
    xn_i, xpe_i = xn_start, xpe_start
    yk = 166
    for cc, x in zip(circuits, pos):
        three = cc["poles"] == 3
        if three:
            xl = [x, x + 5, x + 10]
            xnn, xpp = x + 15, x + 20
            phs = ["L1", "L2", "L3"]
        else:
            xl = [x]
            xnn, xpp = x + 5, x + 10
            phs = [cc["ph"]]
        for k, ln in enumerate(wrap(cc["f"], 30, 5.3)[:3]):
            s.T(x - 4, 19.0 + k * 2.9, ln, 5.3, "B" if k == 0 else "R")
        for xx, ph in zip(xl, phs):
            s.dot(xx, ysf[ph])
            s.L(xx, ysf[ph], xx, 124)
        mcb(s, xl, 124, 16, ("1", "3", "5"), ("2", "4", "6"))
        label(s, xpp + 2.5, 145, [cc["bmk"], f"B{cc['In']}{' 3-pol.' if three else ''}", cc["art"].rstrip("-")],
              5.2, 2.7)
        yend = 205 if cc["typ"] else yk + 6
        for xx in xl:
            s.L(xx, 140, xx, yend)
        s.dot(xnn, ysf["N"])
        s.L(xnn, ysf["N"], xnn, yend)
        s.dot(xpp, ysb["PE"])
        s.L(xpp, ysb["PE"], xpp, yend)
        kterm(s, xnn, yk, "N")
        s.T(xnn - 1.4, yk + 2.6, f"{xn_name}:{xn_i}", 3.9, "R", "r")
        kterm(s, xpp, yk + 4, "PE", side="r")
        s.T(xpp + 1.5, yk + 7.2, f"XPE:{xpe_i}", 3.9)
        xn_i += 1
        xpe_i += 1
        allx = xl + [xnn, xpp]
        if cc["typ"]:
            typ, cs = cc["typ"].split(" ")
            cols = ["BN", "BK", "GY", "BU", "GNYE"] if three else ["BN", "BU", "GNYE"]
            cablemark(s, allx, 186, cc["bmk"].replace("-F", "-W"), f"{typ} {cs.replace('x', '×')}",
                      f"≈{cc['L']} m", cols)
            cx = (allx[0] + allx[-1]) / 2
            consumer(s, cx, 208, cc["cons"], cc["short"])
            s.T(cx, 221, cc["short"], 5.0, "R", "c")
        else:
            cx = (allx[0] + allx[-1]) / 2
            consumer(s, cx, 178, "res", "")
            s.T(cx, 190, "Reserve", 5.0, "R", "c", rgb=GR)


def pg_wv_g1(s, t):
    pg_wv_group(s, t, D.WV_G1, fi_bmk="-F1", bus_lref=xr(f"wv{t}_feed", 9), bus_rref=xr(f"wv{t}_g2a", 0),
                xn_start=2, xpe_start=4)


def pg_wv_g2a(s, t):
    pg_wv_group(s, t, D.WV_G2[:7], fi_bmk="-F2", bus_lref=xr(f"wv{t}_g1", 9), pe_rref=xr(f"wv{t}_g2b", 0),
                fi_rref=xr(f"wv{t}_g2b", 0), xn_start=2, xpe_start=4 + len(D.WV_G1))


def pg_wv_g2b(s, t):
    pg_wv_group(s, t, D.WV_G2[7:], fi_bmk=None, pe_lref=xr(f"wv{t}_g2a", 9), fi_lref=xr(f"wv{t}_g2a", 9),
                xn_start=9, xpe_start=4 + len(D.WV_G1) + 7)


# ------------------------------------------------------------------ Tabellenseiten
TAB_Y0, TAB_YMAX = 36.0, 244.0


def cell_lines(widths, r, sz):
    return [wrap(str(v), w - 2.4, sz) for w, v in zip(widths, r)]


def row_h(widths, r, sz):
    n = max(len(c) for c in cell_lines(widths, r, sz))
    return 1.6 + n * sz * 0.3528 * 1.22


def table_pages(widths, rows, sz, reserve=0.0):
    pages, cur, h = [], [], 0.0
    for r in rows:
        rh = row_h(widths, r, sz)
        if cur and TAB_Y0 + h + rh > TAB_YMAX - reserve:
            pages.append(cur)
            cur, h = [], 0.0
        cur.append(r)
        h += rh
    pages.append(cur)
    return pages


def draw_table(s, title, widths, hdr, rows, sz=5.9, rh=4.6, foot=None, colorfn=None):
    s.T(25, 24, title, 10, "B")
    x, y = 25, 30
    s.R(x, y, sum(widths), 6, 0.25, fill=HD)
    xx = x
    for w, h in zip(widths, hdr):
        s.TF(xx + 1.2, y + 4.2, h, w - 2, 6.2, "B")
        xx += w
    y += 6
    lh = sz * 0.3528 * 1.22
    for r in rows:
        cl = cell_lines(widths, r, sz)
        h = 1.6 + max(len(c) for c in cl) * lh
        fill = colorfn(r) if colorfn else None
        if fill:
            s.R(x, y, sum(widths), h, 0, fill=fill)
        xx = x
        for w, lines in zip(widths, cl):
            for k, t in enumerate(lines):
                s.T(xx + 1.2, y + 0.9 + lh * (k + 1) - 0.55, t, sz)
            xx += w
        s.L(x, y + h, x + sum(widths), y + h, 0.1, rgb=GR)
        y += h
    s.R(x, 30, sum(widths), y - 30, 0.25)
    xx = x
    for w in widths[:-1]:
        xx += w
        s.L(xx, 30, xx, y, 0.1, rgb=GR)
    if foot:
        for k, f in enumerate(foot):
            s.T(x, y + 6 + k * 4.6, f, 5.8, rgb=GR)


ST_SHORT = {"V": "verifiziert", "K": "Schrack-Katalog", "O": "Option", "A": "Art.-Nr. anfragen",
            "B": "Beistellung WN"}

CAB_W = [18, 20, 40, 40, 22, 18, 16, 210]
CAB_H = ["Anlage", "Kabel", "von", "nach", "Typ", "Adern × mm²", "Länge m", "Funktion"]
BOM_W = [12, 14, 46, 13, 10, 160, 26, 56, 14, 31]
BOM_H = ["Pos", "Anlage", "BMK", "Menge", "Einh.", "Bezeichnung", "Art.-Nr.", "Hersteller", "Pv W",
         "Status"]


def cable_rows():
    out = []
    for a, n, f, t, typ, cs, L, fn in D.cables():
        out.append(("=" + a, n, f, t, typ, cs.replace("x", "×"), fnum(L, 0) if L else "—", fn))
    return out


def bom_rows():
    sv, wv = D.bom()
    out = []
    pos = 1
    for a, b, q, u, art, rem in sv:
        A = D.ART[art]
        out.append((str(pos), "=SV", b, fnum(q, 0), u, A["t"] + (f" – {rem}" if rem else ""),
                    art if A["st"] not in ("A", "B") else "—", A["m"], fnum(A["pv"], 2) if A["pv"] else "",
                    ST_SHORT[A["st"]]))
        pos += 1
    for a, b, q, u, art, rem in wv:
        A = D.ART[art]
        out.append((str(pos), "=WV1…4", b, f"4×{fnum(q, 0)}", u, A["t"] + (f" – {rem}" if rem else ""),
                    art if A["st"] not in ("A", "B") else "—", A["m"], fnum(A["pv"], 2) if A["pv"] else "",
                    ST_SHORT[A["st"]]))
        pos += 1
    return out


def bom_color(r):
    st = r[-1]
    if st.startswith("Option") or st.startswith("Art.-Nr."):
        return YE
    return None


# =========================================================================== Aufbau
def build(out):
    cab = cable_rows()
    bomr = bom_rows()
    cab_pages = table_pages(CAB_W, cab, 5.9, reserve=8)
    bom_pages = table_pages(BOM_W, bomr, 5.6, reserve=14)

    add("cover", "Deckblatt", "ALL", pg_cover)
    add("toc", "Inhaltsverzeichnis", "ALL", pg_toc)
    add("ov", "Übersichtsschaltplan (Energiefluss)", "ALL", pg_overview)
    add("proof", "Auslegung und Nachweise", "ALL", pg_proof)
    add("sv_view", "Ansicht Stockwerksverteiler", "SV", pg_sv_view)
    add("sv_feed", "Anspeisung / Hauptabzweig", "SV", pg_sv_feed)
    add("sv_zp1", "Zählerplätze Top 1 – Top 3", "SV", pg_sv_zp, D.ZP[0:3], "PREV_FEED", "NEXT_ZP2")
    add("sv_zp2", "Zählerplätze Top 4, Allgemein, Reserve", "SV", pg_sv_zp, D.ZP[3:6], "PREV_ZP1", None)
    add("sv_allg", "Allgemein – Abgänge", "SV", pg_sv_allg)
    for t in D.TOPS:
        add(f"wv{t}_view", f"Ansicht Wohnungsverteiler Top {t}", f"WV{t}", pg_wv_view, t)
        add(f"wv{t}_feed", f"Anspeisung, Ableiter, Kühlgerät Top {t}", f"WV{t}", pg_wv_feed, t)
        add(f"wv{t}_g1", f"FI-Gruppe 1 Top {t}", f"WV{t}", pg_wv_g1, t)
        add(f"wv{t}_g2a", f"FI-Gruppe 2 (1/2) Top {t}", f"WV{t}", pg_wv_g2a, t)
        add(f"wv{t}_g2b", f"FI-Gruppe 2 (2/2) Top {t}", f"WV{t}", pg_wv_g2b, t)
    for k, rows in enumerate(cab_pages):
        add(f"cab{k}", f"Kabelübersicht ({k + 1}/{len(cab_pages)})", "ALL", draw_table,
            f"Kabelübersicht ({k + 1}/{len(cab_pages)})", CAB_W, CAB_H, rows, 5.9, 4.6,
            ["Längen = Annahmen bzw. Höchstwerte; Ausführung lt. Installationsplan. Kabel im Stiegenhaus halogenfrei."]
            if k == len(cab_pages) - 1 else None)
    for k, rows in enumerate(bom_pages):
        add(f"bom{k}", f"Stückliste ({k + 1}/{len(bom_pages)})", "ALL", draw_table,
            f"Stückliste ({k + 1}/{len(bom_pages)}) – Schrack Technik und Unterlieferanten", BOM_W, BOM_H, rows,
            5.6, 4.6,
            ["Status: verifiziert = Artikel auf schrack.com (Datenblatt/Shopseite) geprüft; Schrack-Katalog = Nummer aus "
             "Schrack-Zubehörliste/Konfiguration; Beistellung WN = Wiener Netze; gelb = Option bzw. Art.-Nr. (Meterware) anfragen.",
             "Pv = Nennverlustleistung je Stück lt. Schrack (Richtwerte gekennzeichnet). Preise siehe Excel "
             "(Eingabefelder). Mengen Meterware = Annahmen."] if k == len(bom_pages) - 1 else None,
            bom_color)
    for i, p in enumerate(PAGES):
        PN[p["key"]] = i + 1
    # Platzhalter-Querverweise auflösen
    for p in PAGES:
        if p["fn"] is pg_sv_zp:
            zps, a, b = p["args"]
            a = xr("sv_feed", 9) if a == "PREV_FEED" else (xr("sv_zp1", 9) if a == "PREV_ZP1" else a)
            b = xr("sv_zp2", 0) if b == "NEXT_ZP2" else b
            p["args"] = (zps, a, b)

    c = canvas.Canvas(out, pagesize=(PW * mm, PH * mm))
    c.setTitle("Stromlaufplan Stockwerksverteiler + Wohnungsverteiler Top 1–4, Wien")
    c.setAuthor(D.COMPANY)
    c.setSubject(D.PROJECT)
    total = len(PAGES)
    for p in PAGES:
        s = Sheet(c)
        frame(s, PN[p["key"]], total, p["title"], p["anl"])
        p["fn"](s, *p["args"])
        c.showPage()
    c.save()
    return total


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else "plan.pdf"
    n = build(out)
    print("Seiten:", n, "->", out)
