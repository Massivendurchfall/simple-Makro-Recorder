import sys, os, json, time, threading, ctypes
from ctypes import wintypes
from PySide6.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QFileDialog, QComboBox, QDoubleSpinBox, QSystemTrayIcon, QMenu, QStyle
from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QIcon, QAction

if os.name != "nt":
    print("Nur unter Windows nutzbar.")
    sys.exit(1)

if not hasattr(wintypes, "ULONG_PTR"):
    if ctypes.sizeof(ctypes.c_void_p) == ctypes.sizeof(ctypes.c_ulonglong):
        wintypes.ULONG_PTR = ctypes.c_ulonglong
    else:
        wintypes.ULONG_PTR = ctypes.c_ulong
if not hasattr(wintypes, "LRESULT"):
    if ctypes.sizeof(ctypes.c_void_p) == ctypes.sizeof(ctypes.c_longlong):
        wintypes.LRESULT = ctypes.c_longlong
    else:
        wintypes.LRESULT = ctypes.c_long

user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

WH_MOUSE_LL = 14
WH_KEYBOARD_LL = 13

WM_MOUSEMOVE = 0x0200
WM_LBUTTONDOWN = 0x0201
WM_LBUTTONUP = 0x0202
WM_RBUTTONDOWN = 0x0204
WM_RBUTTONUP = 0x0205
WM_MBUTTONDOWN = 0x0207
WM_MBUTTONUP = 0x0208
WM_MOUSEWHEEL = 0x020A
WM_MOUSEHWHEEL = 0x020E
WM_XBUTTONDOWN = 0x020B
WM_XBUTTONUP = 0x020C

WM_KEYDOWN = 0x0100
WM_KEYUP = 0x0101
WM_SYSKEYDOWN = 0x0104
WM_SYSKEYUP = 0x0105

LLMHF_INJECTED = 0x00000001
LLKHF_INJECTED = 0x00000010

INPUT_MOUSE = 0
INPUT_KEYBOARD = 1

MOUSEEVENTF_LEFTDOWN = 0x0002
MOUSEEVENTF_LEFTUP = 0x0004
MOUSEEVENTF_RIGHTDOWN = 0x0008
MOUSEEVENTF_RIGHTUP = 0x0010
MOUSEEVENTF_MIDDLEDOWN = 0x0020
MOUSEEVENTF_MIDDLEUP = 0x0040
MOUSEEVENTF_WHEEL = 0x0800
MOUSEEVENTF_HWHEEL = 0x01000

KEYEVENTF_KEYUP = 0x0002

VK_F9 = 0x78
VK_F10 = 0x79
VK_F12 = 0x7B

WHEEL_DELTA = 120

class PUNKT(ctypes.Structure):
    _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]

class MSLLHOOKSTRUCT(ctypes.Structure):
    _fields_ = [("pt", PUNKT), ("mouseData", wintypes.DWORD), ("flags", wintypes.DWORD), ("time", wintypes.DWORD), ("dwExtraInfo", wintypes.ULONG_PTR)]

class KBDLLHOOKSTRUCT(ctypes.Structure):
    _fields_ = [("vkCode", wintypes.DWORD), ("scanCode", wintypes.DWORD), ("flags", wintypes.DWORD), ("time", wintypes.DWORD), ("dwExtraInfo", wintypes.ULONG_PTR)]

class MOUSEINPUT(ctypes.Structure):
    _fields_ = [("dx", wintypes.LONG), ("dy", wintypes.LONG), ("mouseData", wintypes.DWORD), ("dwFlags", wintypes.DWORD), ("time", wintypes.DWORD), ("dwExtraInfo", wintypes.ULONG_PTR)]

class KEYBDINPUT(ctypes.Structure):
    _fields_ = [("wVk", wintypes.WORD), ("wScan", wintypes.WORD), ("dwFlags", wintypes.DWORD), ("time", wintypes.DWORD), ("dwExtraInfo", wintypes.ULONG_PTR)]

class HARDWAREINPUT(ctypes.Structure):
    _fields_ = [("uMsg", wintypes.DWORD), ("wParamL", ctypes.c_short), ("wParamH", ctypes.c_short)]

class EINGABE_UNION(ctypes.Union):
    _fields_ = [("mi", MOUSEINPUT), ("ki", KEYBDINPUT), ("hi", HARDWAREINPUT)]

class EINGABE(ctypes.Structure):
    _fields_ = [("type", wintypes.DWORD), ("union", EINGABE_UNION)]

LowLevelMouseProc = ctypes.WINFUNCTYPE(wintypes.LRESULT, ctypes.c_int, wintypes.WPARAM, wintypes.LPARAM)
LowLevelKeyboardProc = ctypes.WINFUNCTYPE(wintypes.LRESULT, ctypes.c_int, wintypes.WPARAM, wintypes.LPARAM)

user32.SetWindowsHookExW.restype = wintypes.HHOOK
user32.SetWindowsHookExW.argtypes = [ctypes.c_int, ctypes.c_void_p, wintypes.HINSTANCE, wintypes.DWORD]
user32.CallNextHookEx.restype = wintypes.LRESULT
user32.CallNextHookEx.argtypes = [wintypes.HHOOK, ctypes.c_int, wintypes.WPARAM, wintypes.LPARAM]
user32.UnhookWindowsHookEx.restype = wintypes.BOOL
user32.UnhookWindowsHookEx.argtypes = [wintypes.HHOOK]
kernel32.GetModuleHandleW.restype = wintypes.HMODULE
kernel32.GetModuleHandleW.argtypes = [wintypes.LPCWSTR]
user32.SendInput.restype = wintypes.UINT
user32.SendInput.argtypes = [wintypes.UINT, ctypes.c_void_p, ctypes.c_int]
user32.SetCursorPos.restype = wintypes.BOOL
user32.SetCursorPos.argtypes = [ctypes.c_int, ctypes.c_int]

def setze_cursor_pos(x, y):
    user32.SetCursorPos(int(x), int(y))

def sende_maus_flags(flags, data=0):
    extra = wintypes.ULONG_PTR(0)
    inp = EINGABE()
    inp.type = INPUT_MOUSE
    inp.union.mi = MOUSEINPUT(0, 0, int(data), int(flags), 0, extra)
    user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(EINGABE))

def sende_taste(vk, runter):
    extra = wintypes.ULONG_PTR(0)
    inp = EINGABE()
    inp.type = INPUT_KEYBOARD
    inp.union.ki = KEYBDINPUT(int(vk), 0, 0 if runter else KEYEVENTF_KEYUP, 0, extra)
    user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(EINGABE))

class MakroSteuerung:
    def __init__(self):
        self.aufnahme_aktiv = False
        self.wiedergabe_aktiv = False
        self.schleife_aktiv = False
        self.geschwindigkeit = 1.0
        self.liste_ereignisse = []
        self.zeit_start = 0.0
        self.sperre = threading.Lock()
        self.hook_maus = None
        self.hook_tastatur = None
        self.cb_maus = None
        self.cb_tastatur = None
        self.letzte_bewegung_zeit = 0.0
        self.letzte_bewegung_pos = (None, None)
        self.stop_signal = threading.Event()
        self.hotkey_aufnehmen = VK_F9
        self.hotkey_abspielen = VK_F10
        self.hotkey_stop = VK_F12
        self.zaehler_maus = 0
        self.zaehler_tastatur = 0

    def starte_aufnahme(self):
        with self.sperre:
            if self.aufnahme_aktiv:
                return
            self.liste_ereignisse = []
            self.zeit_start = time.perf_counter()
            self.aufnahme_aktiv = True
            self.letzte_bewegung_zeit = 0.0
            self.letzte_bewegung_pos = (None, None)
            self.zaehler_maus = 0
            self.zaehler_tastatur = 0
            self.installiere_maus_hook()

    def stoppe_aufnahme(self):
        with self.sperre:
            if not self.aufnahme_aktiv:
                return
            self.aufnahme_aktiv = False
            self.deinstalliere_maus_hook()

    def starte_wiedergabe(self, schleife, geschwindigkeit):
        with self.sperre:
            if self.wiedergabe_aktiv or not self.liste_ereignisse:
                return
            self.schleife_aktiv = schleife
            self.geschwindigkeit = max(0.05, float(geschwindigkeit))
            self.wiedergabe_aktiv = True
            self.stop_signal.clear()
            t = threading.Thread(target=self._faden_wiedergabe, daemon=True)
            t.start()

    def stoppe_wiedergabe(self):
        with self.sperre:
            if not self.wiedergabe_aktiv:
                return
            self.stop_signal.set()
            self.wiedergabe_aktiv = False

    def installiere_keyboard_hook(self):
        if self.hook_tastatur:
            return
        self.cb_tastatur = LowLevelKeyboardProc(self._callback_tastatur)
        hinst = kernel32.GetModuleHandleW(None)
        self.hook_tastatur = user32.SetWindowsHookExW(WH_KEYBOARD_LL, self.cb_tastatur, hinst, 0)

    def installiere_maus_hook(self):
        if self.hook_maus:
            return
        self.cb_maus = LowLevelMouseProc(self._callback_maus)
        hinst = kernel32.GetModuleHandleW(None)
        self.hook_maus = user32.SetWindowsHookExW(WH_MOUSE_LL, self.cb_maus, hinst, 0)

    def deinstalliere_keyboard_hook(self):
        if self.hook_tastatur:
            user32.UnhookWindowsHookEx(self.hook_tastatur)
            self.hook_tastatur = None
        self.cb_tastatur = None

    def deinstalliere_maus_hook(self):
        if self.hook_maus:
            user32.UnhookWindowsHookEx(self.hook_maus)
            self.hook_maus = None
        self.cb_maus = None

    def _zeitstempel(self):
        return time.perf_counter() - self.zeit_start

    def _hinzufuegen(self, eintrag):
        with self.sperre:
            if self.aufnahme_aktiv:
                self.liste_ereignisse.append(eintrag)

    def _callback_maus(self, nCode, wParam, lParam):
        if nCode >= 0 and self.aufnahme_aktiv:
            info = ctypes.cast(lParam, ctypes.POINTER(MSLLHOOKSTRUCT)).contents
            if info.flags & LLMHF_INJECTED:
                return user32.CallNextHookEx(self.hook_maus, nCode, wParam, lParam)
            x = int(info.pt.x)
            y = int(info.pt.y)
            jetzt = time.perf_counter()
            if wParam == WM_MOUSEMOVE:
                if self.letzte_bewegung_pos[0] is None:
                    self.letzte_bewegung_pos = (x, y)
                    self.letzte_bewegung_zeit = jetzt
                    self._hinzufuegen({"typ": "maus_bewegung", "x": x, "y": y, "zeit": self._zeitstempel()})
                    self.zaehler_maus += 1
                else:
                    dx = abs(x - self.letzte_bewegung_pos[0])
                    dy = abs(y - self.letzte_bewegung_pos[1])
                    dt = jetzt - self.letzte_bewegung_zeit
                    if dx + dy >= 2 or dt >= 0.01:
                        self.letzte_bewegung_pos = (x, y)
                        self.letzte_bewegung_zeit = jetzt
                        self._hinzufuegen({"typ": "maus_bewegung", "x": x, "y": y, "zeit": self._zeitstempel()})
                        self.zaehler_maus += 1
            elif wParam == WM_LBUTTONDOWN:
                self._hinzufuegen({"typ": "maus_links_down", "x": x, "y": y, "zeit": self._zeitstempel()}); self.zaehler_maus += 1
            elif wParam == WM_LBUTTONUP:
                self._hinzufuegen({"typ": "maus_links_up", "x": x, "y": y, "zeit": self._zeitstempel()}); self.zaehler_maus += 1
            elif wParam == WM_RBUTTONDOWN:
                self._hinzufuegen({"typ": "maus_rechts_down", "x": x, "y": y, "zeit": self._zeitstempel()}); self.zaehler_maus += 1
            elif wParam == WM_RBUTTONUP:
                self._hinzufuegen({"typ": "maus_rechts_up", "x": x, "y": y, "zeit": self._zeitstempel()}); self.zaehler_maus += 1
            elif wParam == WM_MBUTTONDOWN:
                self._hinzufuegen({"typ": "maus_mitte_down", "x": x, "y": y, "zeit": self._zeitstempel()}); self.zaehler_maus += 1
            elif wParam == WM_MBUTTONUP:
                self._hinzufuegen({"typ": "maus_mitte_up", "x": x, "y": y, "zeit": self._zeitstempel()}); self.zaehler_maus += 1
            elif wParam == WM_MOUSEWHEEL:
                delta = ctypes.c_short((info.mouseData >> 16) & 0xFFFF).value
                self._hinzufuegen({"typ": "rad_vertikal", "x": x, "y": y, "delta": int(delta), "zeit": self._zeitstempel()}); self.zaehler_maus += 1
            elif wParam == WM_MOUSEHWHEEL:
                delta = ctypes.c_short((info.mouseData >> 16) & 0xFFFF).value
                self._hinzufuegen({"typ": "rad_horizontal", "x": x, "y": y, "delta": int(delta), "zeit": self._zeitstempel()}); self.zaehler_maus += 1
            elif wParam == WM_XBUTTONDOWN:
                taste = (info.mouseData >> 16) & 0xFFFF
                self._hinzufuegen({"typ": f"maus_x{taste}_down", "x": x, "y": y, "zeit": self._zeitstempel()}); self.zaehler_maus += 1
            elif wParam == WM_XBUTTONUP:
                taste = (info.mouseData >> 16) & 0xFFFF
                self._hinzufuegen({"typ": f"maus_x{taste}_up", "x": x, "y": y, "zeit": self._zeitstempel()}); self.zaehler_maus += 1
        return user32.CallNextHookEx(self.hook_maus, nCode, wParam, lParam)

    def _callback_tastatur(self, nCode, wParam, lParam):
        if nCode >= 0:
            info = ctypes.cast(lParam, ctypes.POINTER(KBDLLHOOKSTRUCT)).contents
            vk = int(info.vkCode)
            inj = bool(info.flags & LLKHF_INJECTED)
            if wParam in (WM_KEYDOWN, WM_SYSKEYDOWN):
                if vk == self.hotkey_aufnehmen:
                    if not self.wiedergabe_aktiv:
                        if self.aufnahme_aktiv:
                            self.stoppe_aufnahme()
                        else:
                            self.starte_aufnahme()
                    return 1
                if vk == self.hotkey_abspielen:
                    if not self.aufnahme_aktiv:
                        if self.wiedergabe_aktiv:
                            self.stoppe_wiedergabe()
                        else:
                            self.starte_wiedergabe(self.schleife_aktiv, self.geschwindigkeit)
                    return 1
                if vk == self.hotkey_stop:
                    if self.aufnahme_aktiv:
                        self.stoppe_aufnahme()
                    if self.wiedergabe_aktiv:
                        self.stoppe_wiedergabe()
                    return 1
                if self.aufnahme_aktiv and not inj:
                    self._hinzufuegen({"typ": "taste_down", "vk": vk, "zeit": self._zeitstempel()})
                    self.zaehler_tastatur += 1
            elif wParam in (WM_KEYUP, WM_SYSKEYUP):
                if self.aufnahme_aktiv and not inj and vk not in (self.hotkey_aufnehmen, self.hotkey_abspielen, self.hotkey_stop):
                    self._hinzufuegen({"typ": "taste_up", "vk": vk, "zeit": self._zeitstempel()})
                    self.zaehler_tastatur += 1
        return user32.CallNextHookEx(self.hook_tastatur, nCode, wParam, lParam)

    def _warten_bis(self, ziel):
        jetzt = time.perf_counter()
        rest = ziel - jetzt
        if rest > 0:
            time.sleep(rest)

    def _faden_wiedergabe(self):
        while True:
            if not self.liste_ereignisse:
                break
            start = time.perf_counter()
            erstes = self.liste_ereignisse[0]["zeit"]
            for e in self.liste_ereignisse:
                if self.stop_signal.is_set():
                    break
                ziel = start + (e["zeit"] - erstes) / max(0.0001, self.geschwindigkeit)
                self._warten_bis(ziel)
                if e["typ"] == "maus_bewegung":
                    setze_cursor_pos(e["x"], e["y"])
                elif e["typ"] == "maus_links_down":
                    setze_cursor_pos(e["x"], e["y"]); sende_maus_flags(MOUSEEVENTF_LEFTDOWN)
                elif e["typ"] == "maus_links_up":
                    setze_cursor_pos(e["x"], e["y"]); sende_maus_flags(MOUSEEVENTF_LEFTUP)
                elif e["typ"] == "maus_rechts_down":
                    setze_cursor_pos(e["x"], e["y"]); sende_maus_flags(MOUSEEVENTF_RIGHTDOWN)
                elif e["typ"] == "maus_rechts_up":
                    setze_cursor_pos(e["x"], e["y"]); sende_maus_flags(MOUSEEVENTF_RIGHTUP)
                elif e["typ"] == "maus_mitte_down":
                    setze_cursor_pos(e["x"], e["y"]); sende_maus_flags(MOUSEEVENTF_MIDDLEDOWN)
                elif e["typ"] == "maus_mitte_up":
                    setze_cursor_pos(e["x"], e["y"]); sende_maus_flags(MOUSEEVENTF_MIDDLEUP)
                elif e["typ"] == "rad_vertikal":
                    setze_cursor_pos(e["x"], e["y"]); sende_maus_flags(MOUSEEVENTF_WHEEL, int(e.get("delta", 0)))
                elif e["typ"] == "rad_horizontal":
                    setze_cursor_pos(e["x"], e["y"]); sende_maus_flags(MOUSEEVENTF_HWHEEL, int(e.get("delta", 0)))
                elif e["typ"] == "taste_down":
                    sende_taste(int(e["vk"]), True)
                elif e["typ"] == "taste_up":
                    sende_taste(int(e["vk"]), False)
            if self.stop_signal.is_set():
                break
            if not self.schleife_aktiv:
                break
        self.wiedergabe_aktiv = False

    def loesche(self):
        with self.sperre:
            self.liste_ereignisse = []

    def speichere_datei(self, pfad):
        try:
            with open(pfad, "w", encoding="utf-8") as f:
                json.dump(self.liste_ereignisse, f, ensure_ascii=False, indent=2)
            return True
        except:
            return False

    def lade_datei(self, pfad):
        try:
            with open(pfad, "r", encoding="utf-8") as f:
                daten = json.load(f)
            if isinstance(daten, list):
                self.liste_ereignisse = daten
                return True
        except:
            pass
        return False

class Fenster(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Makro Recorder")
        self.setMinimumSize(460, 280)
        self.steuerung = MakroSteuerung()
        self.steuerung.installiere_keyboard_hook()
        haupt = QWidget()
        layout = QVBoxLayout()
        zeile1 = QHBoxLayout()
        self.knopf_aufnehmen = QPushButton("Aufnahme starten (F9)")
        self.knopf_stop_aufn = QPushButton("Aufnahme stoppen")
        self.knopf_abspielen = QPushButton("Abspielen (F10)")
        self.knopf_stop_absp = QPushButton("Stop (F12)")
        zeile1.addWidget(self.knopf_aufnehmen)
        zeile1.addWidget(self.knopf_stop_aufn)
        zeile1.addWidget(self.knopf_abspielen)
        zeile1.addWidget(self.knopf_stop_absp)
        zeile2 = QHBoxLayout()
        self.auswahl_modus = QComboBox()
        self.auswahl_modus.addItems(["Einmal", "Schleife"])
        self.feld_geschw = QDoubleSpinBox()
        self.feld_geschw.setRange(0.1, 5.0)
        self.feld_geschw.setSingleStep(0.1)
        self.feld_geschw.setValue(1.0)
        self.lbl_modus = QLabel("Modus")
        self.lbl_geschw = QLabel("Geschwindigkeit")
        zeile2.addWidget(self.lbl_modus)
        zeile2.addWidget(self.auswahl_modus)
        zeile2.addWidget(self.lbl_geschw)
        zeile2.addWidget(self.feld_geschw)
        zeile3 = QHBoxLayout()
        self.knopf_speichern = QPushButton("Speichern")
        self.knopf_laden = QPushButton("Laden")
        self.knopf_loeschen = QPushButton("Löschen")
        zeile3.addWidget(self.knopf_speichern)
        zeile3.addWidget(self.knopf_laden)
        zeile3.addWidget(self.knopf_loeschen)
        zeile4 = QHBoxLayout()
        self.lbl_status = QLabel("Status: Bereit")
        self.lbl_anzahl = QLabel("Ereignisse: 0")
        self.lbl_hooks = QLabel("Hooks: Maus 0 | Tastatur 0")
        zeile4.addWidget(self.lbl_status)
        zeile4.addWidget(self.lbl_anzahl)
        zeile4.addWidget(self.lbl_hooks)
        layout.addLayout(zeile1)
        layout.addLayout(zeile2)
        layout.addLayout(zeile3)
        layout.addLayout(zeile4)
        haupt.setLayout(layout)
        self.setCentralWidget(haupt)
        self.knopf_aufnehmen.clicked.connect(self.klick_aufnahme_start)
        self.knopf_stop_aufn.clicked.connect(self.klick_aufnahme_stop)
        self.knopf_abspielen.clicked.connect(self.klick_abspielen)
        self.knopf_stop_absp.clicked.connect(self.klick_stop)
        self.knopf_speichern.clicked.connect(self.klick_speichern)
        self.knopf_laden.clicked.connect(self.klick_laden)
        self.knopf_loeschen.clicked.connect(self.klick_loeschen)
        self.tray = QSystemTrayIcon(self)
        self.tray.setIcon(self.style().standardIcon(QStyle.SP_ComputerIcon))
        menu = QMenu()
        ak_show = QAction("Öffnen", self)
        ak_quit = QAction("Beenden", self)
        menu.addAction(ak_show)
        menu.addSeparator()
        menu.addAction(ak_quit)
        ak_show.triggered.connect(self.zeige)
        ak_quit.triggered.connect(self.schliessen)
        self.tray.setContextMenu(menu)
        self.tray.show()
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.tick)
        self.timer.start(200)

    def tick(self):
        self.lbl_anzahl.setText(f"Ereignisse: {len(self.steuerung.liste_ereignisse)}")
        self.lbl_hooks.setText(f"Hooks: Maus {self.steuerung.zaehler_maus} | Tastatur {self.steuerung.zaehler_tastatur}")
        modus = self.auswahl_modus.currentText() == "Schleife"
        self.steuerung.schleife_aktiv = modus
        self.steuerung.geschwindigkeit = float(self.feld_geschw.value())
        if self.steuerung.wiedergabe_aktiv:
            self.lbl_status.setText("Status: Abspielen")
        elif self.steuerung.aufnahme_aktiv:
            self.lbl_status.setText("Status: Aufnahme")
        else:
            self.lbl_status.setText("Status: Bereit")

    def klick_aufnahme_start(self):
        if self.steuerung.wiedergabe_aktiv:
            return
        self.steuerung.starte_aufnahme()

    def klick_aufnahme_stop(self):
        self.steuerung.stoppe_aufnahme()

    def klick_abspielen(self):
        if self.steuerung.aufnahme_aktiv:
            return
        self.steuerung.starte_wiedergabe(self.auswahl_modus.currentText() == "Schleife", float(self.feld_geschw.value()))

    def klick_stop(self):
        self.steuerung.stoppe_aufnahme()
        self.steuerung.stoppe_wiedergabe()

    def klick_speichern(self):
        pfad, _ = QFileDialog.getSaveFileName(self, "Makro speichern", os.path.expanduser("~"), "JSON (*.json)")
        if pfad:
            ok = self.steuerung.speichere_datei(pfad)
            self.lbl_status.setText("Status: Gespeichert" if ok else "Status: Fehler beim Speichern")

    def klick_laden(self):
        pfad, _ = QFileDialog.getOpenFileName(self, "Makro laden", os.path.expanduser("~"), "JSON (*.json)")
        if pfad:
            ok = self.steuerung.lade_datei(pfad)
            self.lbl_status.setText("Status: Geladen" if ok else "Status: Fehler beim Laden")

    def klick_loeschen(self):
        self.steuerung.loesche()
        self.lbl_status.setText("Status: Gelöscht")

    def zeige(self):
        self.showNormal()
        self.activateWindow()

    def schliessen(self):
        self.close()

    def closeEvent(self, e):
        try:
            self.steuerung.stoppe_aufnahme()
            self.steuerung.deinstalliere_maus_hook()
            self.steuerung.deinstalliere_keyboard_hook()
        except:
            pass
        e.accept()

def main():
    app = QApplication(sys.argv)
    fenster = Fenster()
    fenster.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
