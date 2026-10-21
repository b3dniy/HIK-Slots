# -*- coding: utf-8 -*-
"""
HIK-Slots — камеры регистраторов Hikvision по слотам (каналам Dxx) и их
проверка. Работает по ISAPI, всё в терминале, без веб-интерфейса.

Камеры:
  • проверка всех камер: связь, запись (расписание и реальный архив),
    детекция движения, звук → Excel;
  • подробный аудит одной камеры и исправление найденного, перезагрузка.
Перестановка:
  • перенос камеры в другой слот (свободный или занятый — замена / обмен);
  • перенос со сдвигом остальных (D115 → D50) — до первого свободного слота;
  • обмен двух камер; отмена любой из последних перестановок.
Имя, расписание записи, связки событий и графики охраны переезжают вместе
с камерой. Перед изменениями — снимок и журнал шагов: прерванную операцию
можно продолжить или отменить.

Регистраторы задаются в меню настроек («0») и хранятся внутри exe (NTFS ADS).
Пароли нигде не сохраняются.
"""

import sys
import os
import re
import time
import json
import copy
import glob
import colorsys
import socket
import getpass
import threading
import subprocess
import datetime as dt
import xml.etree.ElementTree as ET
import uuid
from concurrent.futures import ThreadPoolExecutor, as_completed


# ------------------------------------------------------------------ #
#  Автоустановка зависимостей
# ------------------------------------------------------------------ #
def _ensure(pkg, import_name=None):
    import_name = import_name or pkg
    try:
        __import__(import_name)
    except ImportError:
        print(f"  Устанавливаю зависимость: {pkg} ...")
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "--quiet", pkg]
        )


# В собранном .exe зависимости уже внутри — ставить ничего не нужно.
if not getattr(sys, "frozen", False):
    for _p, _i in (("requests", "requests"),
                   ("openpyxl", "openpyxl"),
                   ("colorama", "colorama")):
        _ensure(_p, _i)

import requests
from requests.auth import HTTPDigestAuth, HTTPBasicAuth
from colorama import init as colorama_init, Fore, Style
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Версия программы. Обновление при запуске сравнивает её с последним
# Release на GitHub и с этой же строкой в ветке main — берётся более новая.
VERSION = "1.1.0"
REPO = "b3dniy/HIK-Slots"

def _enable_vt():
    """Включает ANSI/VT в консоли Windows (для 24-битного цвета). True если ок."""
    if os.name != "nt":
        return True
    try:
        import ctypes
        k = ctypes.windll.kernel32
        ok = False
        for handle_id in (-11, -12):          # STDOUT, STDERR
            h = k.GetStdHandle(handle_id)
            mode = ctypes.c_uint32()
            if k.GetConsoleMode(h, ctypes.byref(mode)):
                k.SetConsoleMode(h, mode.value | 0x0004)  # VT processing
                ok = True
        return ok
    except Exception:
        return False


TRUECOLOR = _enable_vt()
if TRUECOLOR:
    colorama_init(strip=False, convert=False)  # пропускаем ANSI как есть (24-бит)
else:
    colorama_init()
requests.packages.urllib3.disable_warnings()  # на случай https с самоподписанным


# ------------------------------------------------------------------ #
#  Консоль в UTF-8 (для русского текста и псевдографики)
# ------------------------------------------------------------------ #
try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stdin.reconfigure(encoding="utf-8")
except Exception:
    pass


def output_dir():
    """Куда сохранять отчёт: рядом с программой, иначе на Рабочий стол."""
    if getattr(sys, "frozen", False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    try:
        t = os.path.join(base, ".write_test")
        open(t, "w").close()
        os.remove(t)
        return base
    except Exception:
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        return desktop if os.path.isdir(desktop) else os.path.expanduser("~")



# ------------------------------------------------------------------ #
#  Конфигурация регистраторов
# ------------------------------------------------------------------ #
HTTP_TIMEOUT = 8
WRITE_TIMEOUT = 25      # запись конфигурации канала идёт дольше чтения

# По умолчанию список пуст: пользователь добавляет свои регистраторы
# через меню настроек («0»). Настройки сохраняются внутри самого exe (ADS).
DEFAULT_RECORDERS = []


def _cfg_stream():
    """Путь к хранилищу настроек — альтернативный поток NTFS внутри самого exe.
    Отдельного видимого файла нет; настройки «прилипают» к exe."""
    base = sys.executable if getattr(sys, "frozen", False) \
        else os.path.abspath(__file__)
    return base + ":hikslots"


def load_recorders():
    """Читает регистраторы из ADS exe; если пусто/ошибка — значения по умолчанию."""
    try:
        with open(_cfg_stream(), "r", encoding="utf-8") as f:
            data = json.load(f)
        out = []
        for r in (data.get("recorders") or []):
            name = str(r.get("name", "")).strip() or str(r.get("lan_ip", ""))
            out.append({
                "name": name,
                "lan_ip": str(r.get("lan_ip", "")).strip(),
                "lan_port": int(r.get("lan_port", 80) or 80),
                "rem_host": str(r.get("rem_host", "")).strip(),
                "rem_port": int(r.get("rem_port", 0) or 0),
                # IP камер, выключенных намеренно (проверка не считает их ошибкой)
                "disabled": [str(x) for x in (r.get("disabled") or [])],
            })
        if out:
            return out
    except Exception:
        pass
    return [dict(r) for r in DEFAULT_RECORDERS]


def save_recorders(recs):
    """Сохраняет регистраторы в ADS exe. Возвращает (ok, сообщение об ошибке)."""
    try:
        with open(_cfg_stream(), "w", encoding="utf-8") as f:
            json.dump({"recorders": recs}, f, ensure_ascii=False, indent=2)
        return True, ""
    except Exception as e:
        return False, str(e)

# ------------------------------------------------------------------ #
#  Вспомогательные функции вывода
# ------------------------------------------------------------------ #
def cls():
    try:
        CON.clear_badges()
    except Exception:
        pass
    os.system("cls" if os.name == "nt" else "clear")


def banner(text):
    line = "═" * (len(text) + 2)
    print(Fore.CYAN + "╔" + line + "╗")
    print(Fore.CYAN + "║ " + Style.BRIGHT + text + Style.NORMAL + Fore.CYAN + " ║")
    print(Fore.CYAN + "╚" + line + "╝" + Style.RESET_ALL)



_START_HUE = 0.78   # фиолетовый — стартовый цвет радуги


def _rainbow_rgb(h):
    """Полная радуга из HSV. h — любое число, берётся дробная часть."""
    r, g, b = colorsys.hsv_to_rgb(h % 1.0, 1.0, 1.0)
    return int(r * 255), int(g * 255), int(b * 255)


_BAR_W = 34


class _Console:
    """Прокси stdout + один фоновый поток анимации.

    • снизу рисуется радужный прогресс-бар;
    • выше — бейджи SUCCESS/ERROR, которые ПРОДОЛЖАЮТ мигать на следующих
      шагах, пока их строка видна на экране (ушла в прокрутку — застывает).
    Процесс не задерживается: всё рисует фоновый поток. Позиция курсора
    сохраняется/восстанавливается (ESC 7 / ESC 8), поэтому даже ввод с
    клавиатуры не портится."""

    def __init__(self, real):
        self.real = real
        self.lock = threading.RLock()
        self.badges = []        # [{off, on, off_str}]
        self.bar = None         # {prefix, frac, phase} | None
        self.frame = 0
        self._stop = threading.Event()
        self._thr = None

    # ---- интерфейс потока вывода (sys.stdout) ----
    def write(self, s):
        if not s:
            return 0
        with self.lock:
            n = s.count("\n")
            if n and self.badges:
                for b in self.badges:
                    b["off"] += n
            self.real.write(s)
            self.real.flush()
            if n:
                self._drop_scrolled()
        return len(s)

    def flush(self):
        try:
            self.real.flush()
        except Exception:
            pass

    def isatty(self):
        try:
            return self.real.isatty()
        except Exception:
            return False

    def fileno(self):
        return self.real.fileno()

    def __getattr__(self, name):
        return getattr(self.real, name)

    # ---- служебное ----
    def _term_lines(self):
        try:
            return os.get_terminal_size().lines
        except Exception:
            return 25

    def _drop_scrolled(self):
        if self.badges:
            h = self._term_lines()
            self.badges = [b for b in self.badges if b["off"] < h - 1]

    def start(self):
        if TRUECOLOR and self._thr is None:
            self._thr = threading.Thread(target=self._run, daemon=True)
            self._thr.start()
        return self

    def stop(self):
        self._stop.set()
        if self._thr:
            self._thr.join(timeout=0.5)
        try:
            self.real.write("\x1b[0m")
            self.real.flush()
        except Exception:
            pass

    def _run(self):
        while not self._stop.is_set():
            self._tick()
            time.sleep(0.035)

    def _bar_str(self):
        bar = self.bar
        if bar is None:
            return None
        bar["phase"] = (bar["phase"] + 0.015) % 1.0
        frac, prefix, phase = bar["frac"], bar["prefix"], bar["phase"]
        filled = int(round(_BAR_W * frac))
        cells = []
        for i in range(filled):
            hL = _START_HUE - i / _BAR_W + phase
            hR = _START_HUE - (i + 0.5) / _BAR_W + phase
            rL, gL, bL = _rainbow_rgb(hL)
            rR, gR, bR = _rainbow_rgb(hR)
            cells.append(f"\x1b[38;2;{rL};{gL};{bL};48;2;{rR};{gR};{bR}m▌")
        tail = "\x1b[0m"
        if filled < _BAR_W:
            tail += "\x1b[38;2;55;58;64m" + "░" * (_BAR_W - filled) + "\x1b[0m"
        pref = f"\x1b[92m{prefix:<26}\x1b[0m"
        return f"\r{pref} |{''.join(cells)}{tail}| {int(frac * 100):3d}%"

    def _tick(self):
        with self.lock:
            self.frame += 1
            blink_on = (self.frame // 12) % 2 == 0     # смена ~каждые 0.42 c
            h = self._term_lines()
            out = []
            for b in self.badges:
                d = b["off"]
                if d <= 0 or d >= h - 1:
                    continue
                line = b["on"] if blink_on else b["off_str"]
                out.append(f"\x1b7\x1b[{d}A\r{line}\x1b[K\x1b8")
            bar = self._bar_str()
            if bar is not None:
                out.append(bar)
            if out:
                try:
                    self.real.write("".join(out))
                    self.real.flush()
                except Exception:
                    pass

    # ---- API ----
    def bar_start(self, prefix):
        with self.lock:
            self.bar = {"prefix": prefix, "frac": 0.0, "phase": 0.0}

    def bar_set(self, frac, prefix=None):
        with self.lock:
            if self.bar is None:
                return
            self.bar["frac"] = 0.0 if frac < 0 else (1.0 if frac > 1 else frac)
            if prefix is not None:
                self.bar["prefix"] = prefix

    def _push_badge(self, on, off_str, plain):
        if TRUECOLOR:
            self.real.write("\r\x1b[K" + on + "\n")
            self.real.flush()
            for bd in self.badges:
                bd["off"] += 1
            self.badges.append({"off": 1, "on": on, "off_str": off_str})
            self._drop_scrolled()
        else:
            self.real.write("\r" + plain + "\n")
            self.real.flush()

    def bar_finish(self, ok=True, prefix=None):
        with self.lock:
            b = self.bar
            self.bar = None
            prefix = prefix or (b["prefix"] if b else "")
            word = "✓  SUCCESS" if ok else "✗  ERROR"
            field = word.center(_BAR_W)
            cmain = "\x1b[1;92m" if ok else "\x1b[1;91m"
            pref = f"\x1b[92m{prefix:<26}\x1b[0m"
            on = f"{pref} |{cmain}{field}\x1b[0m|"
            off = f"{pref} |\x1b[1;97m{field}\x1b[0m|"
            self._push_badge(on, off, f"{prefix:<26} |{field}|")

    def add_badge(self, text, ok=True):
        with self.lock:
            cmain = "\x1b[1;92m" if ok else "\x1b[1;91m"
            on = f"{cmain}{text}\x1b[0m"
            off = f"\x1b[1;97m{text}\x1b[0m"
            self._push_badge(on, off, text)

    def clear_badges(self):
        """Сброс отслеживаемых бейджей/бара (перед очисткой экрана)."""
        with self.lock:
            self.badges = []
            self.bar = None


class ProgressBar:
    """Тонкая обёртка: сам бар рисует фоновый поток _Console (без задержек)."""

    def __init__(self, prefix="", width=_BAR_W):
        self.prefix = prefix

    def start(self):
        CON.bar_start(self.prefix)
        if not TRUECOLOR:
            sys.stdout.write(f"{self.prefix:<26} |{'░' * _BAR_W}|   0%")
            sys.stdout.flush()
        return self

    def set(self, frac, prefix=None):
        CON.bar_set(frac, prefix)
        if not TRUECOLOR:
            p = prefix or self.prefix
            f = 0.0 if frac < 0 else (1.0 if frac > 1 else frac)
            filled = int(round(_BAR_W * f))
            sys.stdout.write(
                f"\r{p:<26} |{'█' * filled}{'░' * (_BAR_W - filled)}| "
                f"{int(f * 100):3d}%")
            sys.stdout.flush()

    def finish(self, ok=True, prefix=None):
        CON.bar_finish(ok, prefix)


CON = _Console(sys.stdout).start()
sys.stdout = CON


def getch_choice(valid):
    """Ждём одиночное нажатие одной из клавиш valid (без Enter). Иное — игнор."""
    try:
        import msvcrt
    except ImportError:
        # не Windows — читаем строкой
        while True:
            s = input().strip()
            if s in valid:
                return s
    while True:
        ch = msvcrt.getch()
        if ch in (b"\x00", b"\xe0"):   # спец-клавиши — съедаем второй байт
            msvcrt.getch()
            continue
        try:
            c = ch.decode("ascii")
        except UnicodeDecodeError:
            continue
        if c == "\x03":                # Ctrl+C
            raise KeyboardInterrupt
        if c in valid:
            return c
        # любая другая клавиша — ничего не делаем, ждём дальше


def read_line_visible(prompt):
    """Видимый построчный ввод. Подтверждение — Enter ИЛИ Tab.
    getwch — сразу Unicode, поэтому кириллица вводится в любой раскладке."""
    try:
        import msvcrt
    except ImportError:
        return input(prompt)
    sys.stdout.write(prompt)
    sys.stdout.flush()
    buf = []
    while True:
        c = msvcrt.getwch()
        if c in ("\x00", "\xe0"):          # спец-клавиши
            msvcrt.getwch()
            continue
        if c in ("\r", "\n", "\t"):        # Enter или Tab — подтвердить
            sys.stdout.write("\n")
            sys.stdout.flush()
            return "".join(buf)
        if c == "\x08":                    # Backspace
            if buf:
                buf.pop()
                sys.stdout.write("\b \b")
                sys.stdout.flush()
            continue
        if c == "\x03":                    # Ctrl+C
            raise KeyboardInterrupt
        if c == "\x1b" or not c.isprintable():   # Esc и прочее — игнор
            continue
        buf.append(c)
        sys.stdout.write(c)
        sys.stdout.flush()


def read_password(prompt="Пароль: "):
    """Скрытый ввод (эхо «*») через общий поток вывода. Подтверждение — Enter."""
    try:
        import msvcrt
    except ImportError:
        return getpass.getpass(prompt)
    sys.stdout.write(prompt)
    sys.stdout.flush()
    buf = []
    while True:
        c = msvcrt.getwch()
        if c in ("\x00", "\xe0"):
            msvcrt.getwch()
            continue
        if c in ("\r", "\n"):
            sys.stdout.write("\n")
            sys.stdout.flush()
            return "".join(buf)
        if c == "\x08":
            if buf:
                buf.pop()
                sys.stdout.write("\b \b")
                sys.stdout.flush()
            continue
        if c == "\x03":
            raise KeyboardInterrupt
        if not c.isprintable():
            continue
        buf.append(c)
        sys.stdout.write("*")
        sys.stdout.flush()


def pause_exit(ok=False, label=None):
    """Показывает мигающий ERROR-бейдж и ждёт любую клавишу."""
    text = label or ("✗  ERROR — нажмите любую клавишу для выхода"
                     if not ok else "Нажмите любую клавишу для выхода")
    CON.add_badge(text, ok=ok)
    try:
        import msvcrt
    except ImportError:
        try:
            input()
        except Exception:
            pass
        return
    try:
        ch = msvcrt.getch()
        if ch in (b"\x00", b"\xe0"):
            msvcrt.getch()
    except Exception:
        pass



# ------------------------------------------------------------------ #
#  Доступность / подключение
# ------------------------------------------------------------------ #
def tcp_ok(host, port, timeout=2.0):
    try:
        with socket.create_connection((host, port), timeout=timeout):
            return True
    except OSError:
        return False


def resolve_endpoint(rec):
    """Возвращает (base_url, via, host, port) или None. Приоритет — локальная сеть."""
    h, p = rec.get("lan_ip", ""), rec.get("lan_port", 80)
    if h and tcp_ok(h, p):
        return (f"http://{h}:{p}", "локальная сеть", h, p)
    h, p = rec.get("rem_host", ""), rec.get("rem_port", 0)
    if h and p and tcp_ok(h, p):
        return (f"http://{h}:{p}", "удалённый порт", h, p)
    return None


def check_all(recs):
    print(Fore.YELLOW + "\nПроверка доступности регистраторов...\n" + Style.RESET_ALL)
    result = [None] * len(recs)
    n = max(len(recs), 1)
    bar = ProgressBar("Проверяю регистраторы").start()
    try:
        for i, rec in enumerate(recs):
            bar.set(i / n, f"Проверяю {rec['name']}")
            result[i] = resolve_endpoint(rec)
    finally:
        bar.finish(ok=any(v is not None for v in result), prefix="Доступность")
    return result


# ------------------------------------------------------------------ #
#  XML помощники (без привязки к namespace)
# ------------------------------------------------------------------ #
def _local(tag):
    return tag.split("}")[-1] if "}" in tag else tag


def findall_local(root, name):
    return [e for e in root.iter() if _local(e.tag) == name]


def first_text(elem, name, default=""):
    for e in elem.iter():
        if _local(e.tag) == name:
            return (e.text or "").strip()
    return default


# ------------------------------------------------------------------ #
#  Авторизация
# ------------------------------------------------------------------ #
def make_session(base, user, pw):
    """Пробуем Digest, затем Basic. Возвращаем (session, None) или (None, сообщение)."""
    last = ""
    for auth in (HTTPDigestAuth(user, pw), HTTPBasicAuth(user, pw)):
        s = requests.Session()
        s.auth = auth
        s.verify = False
        try:
            r = s.get(base + "/ISAPI/System/deviceInfo", timeout=HTTP_TIMEOUT)
        except requests.RequestException as e:
            last = f"сеть: {e}"
            continue
        if r.status_code == 200:
            return s, None
        if r.status_code == 401:
            last = "неверный логин или пароль"
        else:
            last = f"HTTP {r.status_code}"
    return None, last


# ------------------------------------------------------------------ #
#  XML: правка и запись (без потери namespace Hikvision)
# ------------------------------------------------------------------ #
def _ns(elem):
    t = elem.tag
    return t[1:].split("}")[0] if t.startswith("{") else ""


def _q(ns, name):
    return f"{{{ns}}}{name}" if ns else name


def child(elem, name):
    """Прямой потомок по локальному имени (или None)."""
    if elem is None:
        return None
    for c in list(elem):
        if _local(c.tag) == name:
            return c
    return None


def ctext(elem, name):
    c = child(elem, name)
    return (c.text or "").strip() if c is not None else ""


def set_text(elem, name, value, after=None):
    """Ставит текст прямого потомка name; если его нет — создаёт
    (сразу после потомка after, иначе в конце)."""
    c = child(elem, name)
    if c is None:
        c = ET.Element(_q(_ns(elem), name))
        kids = list(elem)
        pos = len(kids)
        if after:
            for i, k in enumerate(kids):
                if _local(k.tag) == after:
                    pos = i + 1
                    break
        elem.insert(pos, c)
    c.text = str(value)
    return c


def xml_bytes(elem):
    """Сериализация без префиксов ns0: — регистратор их не понимает."""
    ns = _ns(elem)
    if ns:
        ET.register_namespace("", ns)
    return ET.tostring(elem, encoding="utf-8", xml_declaration=True)


def xml_str(elem):
    return xml_bytes(elem).decode("utf-8")


def parse_xml(data):
    if isinstance(data, str):
        data = data.encode("utf-8")
    return ET.fromstring(data)


def canon(e):
    """Содержимое элемента без namespace — чтобы сравнивать «то же самое ли»."""
    if e is None:
        return None
    return (_local(e.tag), (e.text or "").strip(), tuple(canon(c) for c in e))


def _renamespace(elem, ns):
    for e in elem.iter():
        e.tag = _q(ns, _local(e.tag))


def transplant(dst, src, keys):
    """Переносит прямых потомков keys из src в dst (по локальному имени),
    сохраняя их место в dst и namespace dst."""
    ns = _ns(dst)
    for k in keys:
        s = child(src, k)
        if s is None:
            continue
        new = copy.deepcopy(s)
        _renamespace(new, ns)
        d = child(dst, k)
        if d is not None:
            idx = list(dst).index(d)
            dst.remove(d)
            dst.insert(idx, new)
        else:
            dst.append(new)


def _int(s, default=None):
    try:
        return int(str(s).strip())
    except (TypeError, ValueError):
        return default


# ------------------------------------------------------------------ #
#  ISAPI-клиент
# ------------------------------------------------------------------ #
BUSY_RETRIES = 5       # «503 Device Busy» — регистратор занят (камеры переподключаются)
BUSY_PAUSE = 2         # пауза растёт: 2, 4, 6, 8, 10 с


class Nvr:
    """Тонкая обёртка над ISAPI: GET → Element, запись → (ok, сообщение).
    На «503 Device Busy» запрос повторяется с нарастающей паузой."""

    def __init__(self, session, base):
        self.s = session
        self.base = base
        self.last_fail = None
        self.cache = {}

    def _request(self, method, path, **kw):
        # Чтение: 503 часто означает «не поддерживается» — один короткий
        # повтор, не больше (иначе проверка 143 камер растягивается на минуты).
        # Запись: 503 — «занят, камеры переподключаются» — ждём дольше.
        retries = 1 if method == "GET" else BUSY_RETRIES
        for attempt in range(retries + 1):
            r = self.s.request(method, self.base + path, **kw)
            if r.status_code != 503 or attempt == retries:
                break
            time.sleep(1 if method == "GET" else BUSY_PAUSE * (attempt + 1))
        return r

    def get(self, path, timeout=HTTP_TIMEOUT):
        return self._request("GET", path, timeout=timeout)

    def get_xml(self, path, timeout=HTTP_TIMEOUT):
        """Element или None (не 200 / сеть / битый XML)."""
        try:
            r = self.get(path, timeout)
        except requests.RequestException:
            return None
        if r.status_code != 200:
            return None
        try:
            return ET.fromstring(r.content)
        except ET.ParseError:
            return None

    def send(self, method, path, body=None, timeout=WRITE_TIMEOUT):
        """(ok, сообщение, Element ответа | None). Разбирает ResponseStatus."""
        if body is not None and not isinstance(body, (bytes, str)):
            body = xml_bytes(body)
        if isinstance(body, str):
            body = body.encode("utf-8")
        headers = {"Content-Type": "application/xml"} if body is not None else None
        try:
            r = self._request(method, path, data=body,
                              headers=headers, timeout=timeout)
        except requests.RequestException as e:
            return False, f"сеть: {e}", None
        root = None
        if r.content:
            try:
                root = ET.fromstring(r.content)
            except ET.ParseError:
                root = None
        code = first_text(root, "statusCode") if root is not None else ""
        sstr = first_text(root, "statusString") if root is not None else ""
        sub = first_text(root, "subStatusCode") if root is not None else ""
        if r.status_code == 200 and code in ("", "1", "7"):
            return True, ("OK" if code != "7" else "OK (нужна перезагрузка)"), root
        detail = "/".join(x for x in (sstr, sub) if x) or (r.reason or "")
        # ответ регистратора целиком — в журнал для разбора (тело запроса
        # не сохраняем: в нём пароль камеры)
        self.last_fail = {"method": method, "path": path, "http": r.status_code,
                          "response": r.text[:1500]}
        return False, f"HTTP {r.status_code}: {detail}", root

    def put(self, path, body):
        return self.send("PUT", path, body)

    def post(self, path, body):
        return self.send("POST", path, body)

    def delete(self, path):
        return self.send("DELETE", path)


# ------------------------------------------------------------------ #
#  Каналы (слоты Dxx)
# ------------------------------------------------------------------ #
CH_PATH = "/ISAPI/ContentMgmt/InputProxy/channels"
TRIG_PATH = "/ISAPI/Event/triggers"


def cam_key(c):
    """Камера = IP + номер канала источника (у видеосервера один IP на
    несколько каналов, это не конфликт)."""
    return f"{c.get('ip', '')}#{c.get('src_ch') or '1'}"


def _cam_from_xml(elem):
    sid = child(elem, "sourceInputPortDescriptor")
    src = sid if sid is not None else elem
    conn = (first_text(elem, "connMode") or "").lower()
    locked = ("plug" in conn or "poe" in conn
              or any(_local(e.tag).lower() in ("poeportno", "poeport")
                     for e in elem.iter()))
    return {
        "slot": _int(ctext(elem, "id")),
        "name": ctext(elem, "name"),
        "ip": (first_text(src, "ipAddress") or first_text(src, "hostName")
               or first_text(src, "domainName")),
        "port": first_text(src, "managePortNo"),
        "proto": first_text(src, "adminProtocol") or first_text(src, "proxyProtocol"),
        "user": first_text(src, "userName"),
        "src_ch": first_text(src, "srcInputPort"),
        "locked": locked,
        "xml": xml_str(elem),
    }


def read_channels(nvr):
    """{слот: камера}. Исключение — если список не читается вовсе."""
    root = nvr.get_xml(CH_PATH)
    if root is None:
        raise RuntimeError("регистратор не отдал список IP-каналов (InputProxy)")
    chans = {}
    for e in findall_local(root, "InputProxyChannel"):
        cam = _cam_from_xml(e)
        if cam["slot"]:
            chans[cam["slot"]] = cam
    return chans


def read_status(nvr):
    """{слот: {online, detect}}."""
    out = {}
    root = nvr.get_xml(CH_PATH + "/status")
    if root is None:
        return out
    for e in findall_local(root, "InputProxyChannelStatus"):
        sl = _int(ctext(e, "id"))
        if sl:
            out[sl] = {"online": first_text(e, "online").lower() == "true",
                       "detect": first_text(e, "chanDetectResult")}
    return out


def read_max_slot(nvr, chans):
    """Сколько IP-каналов у регистратора (из capabilities), иначе — по факту.
    Capabilities не меняются — спрашиваем регистратор один раз за сеанс."""
    if "max_slot" not in nvr.cache:
        cap = 0
        root = nvr.get_xml(CH_PATH + "/capabilities")
        if root is not None:
            for e in root.iter():
                if _local(e.tag) == "id" and e.get("max"):
                    cap = _int(e.get("max"), 0)
                    break
        nvr.cache["max_slot"] = cap
    return max(nvr.cache["max_slot"], max(chans) if chans else 0)


DETECT_RU = {
    "connect": "на связи",
    "overSysBandwidth": "превышена полоса регистратора",
    "domainError": "ошибка домена",
    "ipcStreamFail": "камера не отдаёт поток",
    "ipcReconnect": "переподключается",
    "netUnreachable": "нет сети до камеры",
    "notExist": "камера не найдена",
    "userNameOrPasswordError": "неверный логин/пароль камеры",
    "errorUserNameOrPasswd": "неверный логин/пароль камеры",
    "passwordError": "неверный пароль камеры",
    "userLocked": "камера заблокировала вход (много попыток)",
    "ipConflict": "конфликт IP",
    "notSupport": "не поддерживается",
    "unknownError": "неизвестная ошибка",
}


def detect_ru(st):
    if not st:
        return "нет данных"
    if st.get("online"):
        return "на связи"
    d = st.get("detect") or ""
    return DETECT_RU.get(d, d or "офлайн")


def is_pw_error(st):
    d = (st or {}).get("detect", "").lower()
    return bool(st) and not st.get("online") and ("passw" in d or "lock" in d)


def free_ranges(chans, max_slot):
    """'D7, D20–D31' — свободные слоты одной строкой."""
    free = [s for s in range(1, max_slot + 1) if s not in chans]
    parts, i = [], 0
    while i < len(free):
        j = i
        while j + 1 < len(free) and free[j + 1] == free[j] + 1:
            j += 1
        parts.append(f"D{free[i]}" if i == j else f"D{free[i]}–D{free[j]}")
        i = j + 1
    return ", ".join(parts)


def slot_occupant(nvr, slot):
    """cam_key камеры в слоте, '' — слот пуст, None — не удалось прочитать."""
    try:
        r = nvr.get(f"{CH_PATH}/{slot}")
    except requests.RequestException:
        return None
    if r.status_code == 200:
        try:
            root = ET.fromstring(r.content)
        except ET.ParseError:
            return None
        if _local(root.tag) == "InputProxyChannel":
            return cam_key(_cam_from_xml(root))
        return ""
    if r.status_code in (400, 403, 404):
        return ""
    return None


def channel_body(cam, slot, password):
    """XML канала из снимка, переставленный в слот slot, с паролем камеры
    (регистратор пароль не отдаёт, поэтому подставляем его сами)."""
    e = parse_xml(cam["xml"])
    set_text(e, "id", slot)
    for c in list(e):
        if _local(c.tag) == "devIndex":       # уникален для слота — не копируем
            e.remove(c)
    sid = child(e, "sourceInputPortDescriptor")
    if sid is not None and password is not None:
        set_text(sid, "password", password, after="userName")
    return e


# Временные адреса из диапазонов документации (RFC 5737) — в сетях их нет.
PARK_NET = "192.0.2."       # «заглушка» в занятом слоте вместо удаления камеры
FILL_NET = "198.51.100."    # заглушки, временно занимающие свободные слоты
MAX_FILLERS = 64


def is_placeholder(ip):
    return (ip or "").startswith((PARK_NET, FILL_NET))


def placeholder_body(template_xml, slot, ip, name, password):
    """Канал-заглушка на основе XML настоящего канала: тот же протокол и порт,
    но несуществующий IP — регистратор держит слот, камера никуда не ходит."""
    e = parse_xml(template_xml)
    set_text(e, "id", slot)
    set_text(e, "name", name)
    for c in list(e):
        if _local(c.tag) == "devIndex":
            e.remove(c)
    sid = child(e, "sourceInputPortDescriptor")
    if sid is not None:
        for c in list(sid):
            if _local(c.tag) in ("hostName", "domainName"):
                sid.remove(c)
        set_text(sid, "addressingFormatType", "ipaddress")
        set_text(sid, "ipAddress", ip, after="addressingFormatType")
        if password is not None:
            set_text(sid, "password", password, after="userName")
    return e


MIN_SRC_KEYS = ("proxyProtocol", "adminProtocol", "addressingFormatType", "ipAddress",
                "hostName", "managePortNo", "srcInputPort", "userName", "password",
                "streamType")


def minimal_channel(full):
    """Тот же канал, но только с обязательными полями — на случай, если
    прошивка отвечает badXmlContent на что-то из необязательного."""
    ns = _ns(full)
    e = ET.Element(_q(ns, "InputProxyChannel"))
    for k in ("id", "name"):
        c = child(full, k)
        if c is not None:
            e.append(copy.deepcopy(c))
    sid = child(full, "sourceInputPortDescriptor")
    out = ET.SubElement(e, _q(ns, "sourceInputPortDescriptor"))
    for k in MIN_SRC_KEYS:
        c = child(sid, k)
        if c is not None and (c.text or "").strip():
            out.append(copy.deepcopy(c))
    return e


def send_channel(nvr, method, path, body):
    """Запись канала; при badXmlContent — повтор с минимальным XML."""
    ok, msg, _ = nvr.send(method, path, body)
    if ok or "badxml" not in msg.lower():
        return ok, msg
    ok2, msg2, _ = nvr.send(method, path, minimal_channel(body))
    return ok2, (msg2 if ok2 else f"{msg}; упрощённый XML: {msg2}")


def _channel_xml(nvr, slot):
    """XML канала в слоте (str) или None, если слот пуст / не читается."""
    try:
        r = nvr.get(f"{CH_PATH}/{slot}")
    except requests.RequestException:
        return None
    if r.status_code != 200:
        return None
    try:
        root = ET.fromstring(r.content)
    except ET.ParseError:
        return None
    return xml_str(root) if _local(root.tag) == "InputProxyChannel" else None


def do_delete(nvr, slot):
    """Освобождает слот. Уже пустой слот — тоже успех (важно для продолжения)."""
    if slot_occupant(nvr, slot) == "":
        return True, "уже пусто"
    ok, msg, _ = nvr.delete(f"{CH_PATH}/{slot}")
    if ok:
        return True, msg
    if slot_occupant(nvr, slot) == "":
        return True, "удалено"
    return False, msg


def do_park(nvr, slot, idx, password):
    """Вместо удаления: в слот ставится заглушка с несуществующим IP.
    Слот остаётся занятым, поэтому камеру в него потом можно записать
    обычным изменением канала — это умеют все прошивки."""
    cur = _channel_xml(nvr, slot)
    if cur is None:
        if slot_occupant(nvr, slot) == "":
            return True, "слот пуст"
        return False, f"D{slot} не читается"
    ip = f"{PARK_NET}{2 + idx % 250}"
    return send_channel(nvr, "PUT", f"{CH_PATH}/{slot}",
                        placeholder_body(cur, slot, ip, f"HIK-Slots temp {slot}", password))


def sweep_fillers(nvr):
    """Убирает заглушки FILL_NET (в т.ч. оставшиеся после обрыва)."""
    try:
        chans = read_channels(nvr)
    except Exception:
        return
    for sl, c in chans.items():
        if (c["ip"] or "").startswith(FILL_NET):
            nvr.delete(f"{CH_PATH}/{sl}")


def create_at(nvr, slot, cam, password):
    """Добавить камеру именно в ПУСТОЙ слот. Многие прошивки при добавлении
    сами берут первый свободный слот, поэтому все свободные слоты ниже нужного
    временно занимаем заглушками, добавляем камеру и убираем заглушки."""
    key = cam_key(cam)
    try:
        chans = read_channels(nvr)
    except Exception as e:
        return False, f"список каналов не читается: {e}"
    below = [x for x in range(1, slot) if x not in chans]
    if len(below) > MAX_FILLERS:
        return False, (f"ниже D{slot} свободно {len(below)} слотов — слишком много "
                       "временных заглушек")
    try:
        for k, f in enumerate(below):
            ok, msg = send_channel(nvr, "POST", CH_PATH, placeholder_body(
                cam["xml"], f, f"{FILL_NET}{1 + k % 250}", f"HIK-Slots fill {f}", password))
            if not ok:
                return False, f"не удалось временно занять D{f}: {msg}"
        ok, msg = send_channel(nvr, "POST", CH_PATH, channel_body(cam, slot, password))
        if not ok:
            return False, f"добавление камеры: {msg}"
        try:
            chans = read_channels(nvr)
        except Exception as e:
            return False, f"после добавления список каналов не читается: {e}"
        where = [x for x, c in chans.items() if cam_key(c) == key]
        if slot in where:
            # при добавлении прошивка ставит своё имя («IPCamera 149») —
            # записываем наше отдельным изменением канала
            if chans[slot]["name"] != cam["name"]:
                send_channel(nvr, "PUT", f"{CH_PATH}/{slot}",
                             channel_body(cam, slot, password))
            return True, "OK (добавлена в пустой слот)"
        for x in where:
            nvr.delete(f"{CH_PATH}/{x}")
        if where:
            return False, (f"камера встала в D{where[0]} вместо D{slot} — "
                           "убрана оттуда")
        return False, "камера не появилась после добавления"
    finally:
        sweep_fillers(nvr)


def do_place(nvr, slot, cam, password):
    """Ставит камеру в слот slot. Занятый слот (камера или заглушка) —
    изменение канала (PUT); пустой — create_at с заглушками."""
    key = cam_key(cam)
    # Камеру могли вручную добавить заново, и она встала в другой слот.
    # Тогда её место там занимаем заглушкой (иначе будет дубль), а после
    # успешной установки этот слот освобождаем.
    elsewhere = []
    try:
        elsewhere = [s for s, c in read_channels(nvr).items()
                     if cam_key(c) == key and s != slot]
    except Exception:
        pass
    for k, s in enumerate(elsewhere):
        ok, msg = do_park(nvr, s, 200 + k, password)
        if not ok:
            return False, f"камера уже стоит в D{s}, убрать оттуда не вышло: {msg}"
    ok, msg = _place_here(nvr, slot, cam, key, password)
    for s in elsewhere:
        if ok:
            do_delete(nvr, s)
        else:                                   # не вышло — возвращаем как было
            send_channel(nvr, "PUT", f"{CH_PATH}/{s}", channel_body(cam, s, password))
    if ok and elsewhere:
        msg += " (камера была в " + ", ".join(f"D{s}" for s in elsewhere) + \
            " — перенесена оттуда)"
    return ok, msg


def _place_here(nvr, slot, cam, key, password):
    occ = slot_occupant(nvr, slot)
    if occ == "":
        return create_at(nvr, slot, cam, password)
    ok, msg = send_channel(nvr, "PUT", f"{CH_PATH}/{slot}", channel_body(cam, slot, password))
    if not ok:
        return False, f"изменение канала D{slot}: {msg}"
    occ = slot_occupant(nvr, slot)
    if occ == key or occ is None:
        return True, msg
    if occ == "":
        return False, f"регистратор ответил OK, но слот D{slot} пуст"
    return False, f"в слоте D{slot} оказалась другая камера"


# ------------------------------------------------------------------ #
#  Дорожки записи и события (привязаны к слоту, а не к камере)
# ------------------------------------------------------------------ #
TRACK_KEYS = ("Enable", "LoopEnable", "DefaultRecordingMode", "Duration",
              "TrackSchedule", "CustomExtensionList")

_SUFFIX = re.compile(r"^(.*?\D)(\d+)$")


def track_id(slot, off):
    return (slot + off) * 100 + 1          # основной поток: D5 → 501


def read_tracks(nvr):
    """{id дорожки: Element}."""
    out = {}
    root = nvr.get_xml("/ISAPI/ContentMgmt/record/tracks")
    if root is not None:
        for t in findall_local(root, "Track"):
            tid = _int(ctext(t, "id"))
            if tid:
                out[tid] = t
    return out


def detect_offset(chans, tracks):
    """Сдвиг нумерации IP-каналов в дорожках/событиях: обычно 0 (D1 → 101);
    на гибридных/старых моделях IP-каналы идут после аналоговых (D1 → 3301)."""
    if not chans or not tracks:
        return 0
    main = {t // 100 for t in tracks if t % 100 == 1}
    best, best_hits = 0, -1
    for off in (0, 32, 16, 64, 8, 4):
        hits = sum(1 for s in chans if s + off in main)
        if hits > best_hits:
            best, best_hits = off, hits
    return best


def track_info(t):
    """Расписание дорожки: включено ли, режимы записи, есть ли блоки."""
    if t is None:
        return {"enabled": None, "modes": [], "blocks": 0}
    en = first_text(t, "enableSchedule") or ctext(t, "Enable")
    modes = sorted({(e.text or "").strip() for e in t.iter()
                    if _local(e.tag) == "ActionRecordingMode" and (e.text or "").strip()})
    blocks = sum(1 for e in t.iter() if _local(e.tag) == "ScheduleBlock")
    return {"enabled": (en.lower() == "true") if en else None,
            "modes": modes, "blocks": blocks}


MODE_RU = {"CMR": "постоянная", "MOTION": "по движению", "ALARM": "по тревоге"}


def modes_ru(modes):
    if not modes:
        return "—"
    return ", ".join(dict.fromkeys(MODE_RU.get(m, "по событиям") for m in modes))


def read_triggers(nvr):
    root = nvr.get_xml(TRIG_PATH)
    if root is None:
        return []
    return [xml_str(t) for t in findall_local(root, "EventTrigger")]


def own_slot(elem, off):
    """Слот, к которому относится триггер/график (None — не канальный)."""
    dyn = _int(ctext(elem, "dynVideoInputChannelID"))
    if dyn:
        return dyn
    v = _int(ctext(elem, "videoInputChannelID"))
    return v - off if v else None


def renum(text, old, new, off):
    """'VMD-50' → 'VMD-51' (суффикс может быть со сдвигом нумерации off)."""
    mt = _SUFFIX.match(text or "")
    if not mt:
        return text
    n = int(mt.group(2))
    if n == old:
        return f"{mt.group(1)}{new}"
    if off and n == old + off:
        return f"{mt.group(1)}{new + off}"
    return text


def remap_event(elem, mapping, off):
    """Переписывает в триггере/графике номера каналов по карте
    {старый слот: новый}: свой канал, каналы в связках (запись других камер),
    и id вида 'VMD-50' / 'record-50'. Возвращает множество старых слотов,
    на которые элемент ссылался."""
    refs = set()

    def m(s):
        return mapping.get(s, s)

    for holder in list(elem.iter()):
        for c in list(holder):
            ln = _local(c.tag)
            if ln not in ("dynVideoInputChannelID", "dynVideoInputID",
                          "videoInputChannelID", "videoInputID"):
                continue
            v = _int(c.text)
            if v is None:
                continue
            dyn = ln.startswith("dyn")
            slot = v if dyn else v - off
            refs.add(slot)
            nv = m(slot) if dyn else m(slot) + off
            if nv != v:
                c.text = str(nv)
                idc = child(holder, "id")
                if idc is not None:
                    mt = _SUFFIX.match((idc.text or "").strip())
                    if mt and int(mt.group(2)) == v:
                        idc.text = f"{mt.group(1)}{nv}"
    return refs


# Графики охраны: раздел и форматы id. Реальная прошивка (2.249) хранит их
# как «VMD_video1», «Tamperdetection_video1»…; старые — по id триггера «VMD-1».
SCHED_KINDS = (
    ("motionDetections", ("VMD_video{}", "VMD-{}", "VMD_{}")),
    ("tamperDetections", ("Tamperdetection_video{}", "tamper-{}", "Tamper-{}")),
    ("lineDetections", ("linedetection_video{}", "linedetection-{}")),
    ("fieldDetections", ("fielddetection_video{}", "fielddetection-{}")),
)


def read_schedule(nvr, typ, sid, v):
    """График или None. Прошивка «угадывает» id по цифрам в конце и может
    вернуть чужой график — берём ответ, только если это ровно запрошенный."""
    root = nvr.get_xml(f"/ISAPI/Event/schedules/{typ}/{sid}")
    if root is None or ctext(root, "id") != sid:
        return None
    ch = _int(ctext(root, "videoInputChannelID"))
    return root if ch in (None, v) else None


def schedule_hours(root):
    """Сколько часов в неделю покрывает график охраны (TimeBlock'и)."""
    total = 0.0
    for tb in findall_local(root, "TimeBlock"):
        a = first_text(tb, "beginTime")
        b = first_text(tb, "endTime")
        try:
            ha, ma, sa = (int(x) for x in a.split(":"))
            hb, mb, sb = (int(x) for x in b.split(":"))
        except ValueError:
            continue
        total += max(0, (hb * 3600 + mb * 60 + sb) - (ha * 3600 + ma * 60 + sa)) / 3600
    return total


def read_schedules(nvr, slots, off):
    """Графики охраны для слотов (с камерами): [[тип, id, слот, xml]].
    Рабочий формат id запоминается; раздел без ответа больше не трогаем."""
    out, fmt_ok = [], {}
    for s in sorted(slots):
        v = s + off
        for typ, fmts in SCHED_KINDS:
            if typ in fmt_ok and fmt_ok[typ] is None:
                continue
            for f in ([fmt_ok[typ]] if typ in fmt_ok else fmts):
                root = read_schedule(nvr, typ, f.format(v), v)
                if root is not None:
                    fmt_ok[typ] = f
                    out.append([typ, f.format(v), s, xml_str(root)])
                    break
            else:
                fmt_ok.setdefault(typ, None)
    return out


# ------------------------------------------------------------------ #
#  Планирование перестановки (чистая логика, без сети)
# ------------------------------------------------------------------ #
def plan_move(chans, src, dst):
    """Камеру из src ставим в dst. Камеры между ними сдвигаются на один слот
    в сторону src — но только до первого пустого слота (дальше незачем).

    Возвращает (plan, None) или (None, ошибка). plan:
      pairs — [(старый слот, новый слот)] для всех переезжающих камер;
      steps — цепочка ("park", слот) / ("place", слот, исходный слот) /
              ("delete", слот): одна камера ни в какой момент не стоит
              в двух слотах, а каналы не удаляются (кроме освобождаемого);
      freed — слот, который останется пустым (или None)."""
    if src == dst:
        return None, "Слот-источник и слот-назначение совпадают."
    if src not in chans:
        return None, f"В слоте D{src} нет камеры."
    if chans[src].get("locked"):
        return None, f"D{src} — канал PoE-порта, его нельзя перемещать."
    step = 1 if src > dst else -1
    occ = set(chans) - {src}
    hole = dst
    while hole in occ:
        hole += step
    shifted = list(range(dst, hole, step))
    for s in shifted:
        if chans[s].get("locked"):
            return None, f"D{s} — канал PoE-порта, сдвинуть его нельзя."

    pairs = [(src, dst)] + [(s, s + step) for s in shifted]
    steps = [("park", src)]
    cur = hole
    while cur != dst:
        prev = cur - step
        steps.append(("park", prev))
        steps.append(("place", cur, prev))
        cur = prev
    steps.append(("place", dst, src))
    if hole != src:
        steps.append(("delete", src))
    return {"pairs": pairs, "steps": steps,
            "freed": src if hole != src else None,
            "title": f"D{src} → D{dst}"}, None


def plan_swap(chans, a, b):
    """Меняет местами камеры слотов a и b (пустой слот = просто перенос)."""
    if a == b:
        return None, "Нужны два разных слота."
    if a not in chans and b not in chans:
        return None, "Оба слота пусты."
    if b not in chans:
        return plan_move(chans, a, b)
    if a not in chans:
        return plan_move(chans, b, a)
    for s in (a, b):
        if chans[s].get("locked"):
            return None, f"D{s} — канал PoE-порта, его нельзя перемещать."
    return {"pairs": [(a, b), (b, a)],
            "steps": [("park", a), ("place", b, a), ("place", a, b)],
            "freed": None, "title": f"D{a} ⇄ D{b}"}, None


def plan_replace(chans, src, dst):
    """Камера из src встаёт на место камеры в dst; камера из dst удаляется
    с регистратора, src освобождается. Пустой dst — обычный перенос."""
    if src == dst:
        return None, "Слот-источник и слот-назначение совпадают."
    if src not in chans:
        return None, f"В слоте D{src} нет камеры."
    if dst not in chans:
        return plan_move(chans, src, dst)
    for s in (src, dst):
        if chans[s].get("locked"):
            return None, f"D{s} — канал PoE-порта, его нельзя трогать."
    return {"pairs": [(src, dst)],
            "steps": [("park", src), ("place", dst, src), ("delete", src)],
            "freed": src, "replaced": dst,
            "title": f"D{src} → D{dst} (замена)"}, None


def simulate_steps(steps, lay):
    """Раскладка после шагов: {слот: исходный слот камеры | "P" (заглушка)}."""
    lay = dict(lay)
    for st in steps:
        if st[0] == "delete":
            lay.pop(st[1], None)
        elif st[0] == "park":
            if st[1] in lay:
                lay[st[1]] = "P"
        else:
            lay[st[1]] = st[2]
    return lay


def rollback_steps(done, occupied, touched):
    """Шаги, возвращающие затронутые слоты к исходному состоянию после того,
    как выполнены шаги done (операция могла оборваться на любом шаге).
    occupied — слоты, где камеры стояли до операции.
    1) всё, что стоит не на своём месте, — в заглушку (слот не удаляется);
    2) каждую камеру — в её исходный слот;
    3) слоты, пустые до операции, — снова освобождаем."""
    occupied = set(occupied)
    lay = simulate_steps(done, {s: s for s in touched if s in occupied})
    out = []
    for s in sorted(touched):
        want = s if s in occupied else None
        if lay.get(s) not in (want, None, "P"):
            out.append(["park", s])
    for s in sorted(touched):
        if s in occupied and lay.get(s) != s:
            out.append(["place", s, s])
    for s in sorted(touched):
        if s not in occupied and s in lay:
            out.append(["delete", s])
    return out


# ------------------------------------------------------------------ #
#  Журнал операций (снимок + шаги) — рядом с программой
# ------------------------------------------------------------------ #
def backup_dir():
    d = os.path.join(output_dir(), "HIK-Slots_backup")
    os.makedirs(d, exist_ok=True)
    return d


def _safe(s):
    return "".join(ch if ch.isalnum() or ch in "-_" else "-" for ch in s)


def job_path(job):
    return os.path.join(backup_dir(), job["file"])


def save_job(job):
    """Атомарная запись: при обрыве питания журнал не превратится в мусор."""
    p = job_path(job)
    tmp = p + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(job, f, ensure_ascii=False, indent=1)
    os.replace(tmp, p)


def load_jobs(rec_name):
    """Журналы этого регистратора, новые первыми."""
    pat = os.path.join(backup_dir(), _safe(rec_name) + "_*.json")
    out = []
    for p in sorted(glob.glob(pat), reverse=True):
        try:
            with open(p, "r", encoding="utf-8") as f:
                j = json.load(f)
        except Exception:
            continue
        if j.get("recorder") == rec_name:
            out.append(j)
    return out


def new_job(rec_name, kind, title, plan_or_steps, snap, off, touched, pairs):
    ts = dt.datetime.now()
    return {
        "version": 1,
        # миллисекунды — чтобы журналы одной секунды сортировались по времени
        "file": f"{_safe(rec_name)}_{ts.strftime('%Y%m%d_%H%M%S_%f')[:-3]}_{kind}.json",
        "recorder": rec_name,
        "kind": kind,                  # move | swap | rollback
        "title": title,
        "created": ts.strftime("%d.%m.%Y %H:%M:%S"),
        "state": "chain",              # chain → settings → verify → done
        "steps": [list(s) for s in plan_or_steps],
        "done": 0,
        "error": "",
        "pairs": [list(p) for p in pairs],
        "touched": sorted(touched),
        "off": off,
        "snap": snap,
        "results": {},
        "finished": "",
        "rollback_of": "",
    }


def take_snapshot(nvr, chans, status, touched):
    """Всё, что нужно для переноса и отката: каналы, статусы,
    дорожки и графики затронутых слотов, все триггеры событий."""
    bar = ProgressBar("Снимок настроек").start()
    ok = False
    try:
        bar.set(0.05, "Дорожки записи")
        tracks_all = read_tracks(nvr)
        off = detect_offset(chans, tracks_all)
        tracks = {}
        for s in touched:
            t = tracks_all.get(track_id(s, off))
            if t is not None:
                tracks[str(s)] = xml_str(t)
        bar.set(0.35, "Связки событий")
        trig = read_triggers(nvr)
        bar.set(0.6, "Графики охраны")
        scheds = read_schedules(nvr, set(touched) & set(chans), off)
        ok = True
    finally:
        bar.finish(ok=ok, prefix="Снимок настроек")
    snap = {
        "chans": {str(s): c for s, c in chans.items()},
        "status": {str(s): v for s, v in status.items()},
        "tracks": tracks,
        "triggers": trig,
        "scheds": scheds,
    }
    return snap, off


# ------------------------------------------------------------------ #
#  Выполнение операции
# ------------------------------------------------------------------ #
class Ctx:
    """Подключение + пароли камер (только в памяти, никуда не пишутся)."""

    def __init__(self, nvr, rec, via, nvr_pw):
        self.nvr = nvr
        self.rec = rec
        self.via = via
        self.nvr_pw = nvr_pw
        self.pw_common = None
        self.pw_by_cam = {}

    def password_for(self, cam):
        if cam_key(cam) in self.pw_by_cam:
            return self.pw_by_cam[cam_key(cam)]
        return self.pw_common if self.pw_common is not None else self.nvr_pw


def _step_text(st, snap_chans):
    if st[0] == "delete":
        return f"освободить D{st[1]}"
    if st[0] == "park":
        return f"заглушка в D{st[1]}"
    cam = snap_chans.get(str(st[2]), {})
    return f"D{st[1]} ← {cam.get('name') or cam.get('ip', '?')}"


def run_chain(ctx, job):
    """Цепочка заглушка/поставить. Журнал пишется после каждого шага,
    поэтому прерванную операцию можно продолжить с того же места."""
    nvr = ctx.nvr
    snap = job["snap"]["chans"]
    steps = job["steps"]
    n = max(len(steps), 1)
    ok_all = True
    sweep_fillers(nvr)                 # хвосты прошлого обрыва
    bar = ProgressBar("Перестановка").start()
    try:
        while job["done"] < len(steps):
            i = job["done"]
            st = steps[i]
            bar.set(i / n, _step_text(st, snap)[:26])
            if st[0] == "delete":
                ok, msg = do_delete(nvr, st[1])
            elif st[0] == "park":
                ok, msg = do_park(nvr, st[1], i, ctx.pw_common or ctx.nvr_pw)
            else:
                cam = snap[str(st[2])]
                ok, msg = do_place(nvr, st[1], cam, ctx.password_for(cam))
            if not ok:
                job["error_detail"] = nvr.last_fail
                job["error"] = (f"шаг {i + 1} из {len(steps)} "
                                f"({_step_text(st, snap)}): {msg}")
                save_job(job)
                ok_all = False
                break
            job["done"] = i + 1
            job["error"] = ""
            save_job(job)
    finally:
        bar.finish(ok=ok_all, prefix="Перестановка")
    return ok_all


def run_settings(ctx, job):
    """Расписание записи, связки событий и графики охраны — за камерой.
    Пишем только то, что реально отличается (меньше запросов и риска)."""
    nvr = ctx.nvr
    snap, off = job["snap"], job["off"]
    mapping = {int(o): int(n) for o, n in job["pairs"]}
    touched = set(job["touched"])
    holes = touched - set(mapping)        # пустые до операции — их старое не берём
    res = {"tracks": [0, []], "triggers": [0, []], "scheds": [0, []]}

    trig_cur = {}
    root = nvr.get_xml(TRIG_PATH)
    if root is not None:
        for t in findall_local(root, "EventTrigger"):
            trig_cur[ctext(t, "id")] = t

    # Камера офлайн (выключена или неверный пароль) — её связки регистратор
    # не примет («503 Device Busy»), повторы только тянут время. Пропускаем
    # честно и помечаем отдельно; дорожки записи (на стороне регистратора) пишем.
    offline = {s for s, v in read_status(nvr).items() if not v.get("online")}
    skipped = set()

    work = []
    for o, n in sorted(mapping.items(), key=lambda p: p[1]):
        if str(o) in snap["tracks"]:
            work.append(("track", o, n))
    for tx in snap["triggers"]:
        e = parse_xml(tx)
        own = own_slot(e, off)
        if own in holes:
            continue
        refs = remap_event(e, mapping, off)
        if own in mapping or (refs & touched):
            new_own = mapping.get(own, own) if own is not None else None
            if new_own in offline:
                skipped.add(new_own)
                continue
            work.append(("trigger", own, e))
    for typ, tid, own, sx in snap["scheds"]:
        if own in mapping:
            if mapping[own] in offline:
                skipped.add(mapping[own])
                continue
            work.append(("sched", (typ, tid, own), sx))
    job["results"]["skipped"] = sorted(skipped)

    total = max(len(work), 1)
    bar = ProgressBar("Перенос настроек").start()
    try:
        for i, w in enumerate(work):
            kind = w[0]
            if kind == "track":
                _, o, n = w
                bar.set(i / total, f"Расписание D{n}")
                tid = track_id(n, off)
                cur = nvr.get_xml(f"/ISAPI/ContentMgmt/record/tracks/{tid}")
                if cur is None:
                    res["tracks"][1].append(f"D{n}: дорожка {tid} не читается")
                    continue
                want = copy.deepcopy(cur)
                transplant(want, parse_xml(snap["tracks"][str(o)]), TRACK_KEYS)
                if canon(want) == canon(cur):
                    res["tracks"][0] += 1
                    continue
                ok, msg, _ = nvr.put(f"/ISAPI/ContentMgmt/record/tracks/{tid}", want)
                if ok:
                    res["tracks"][0] += 1
                else:
                    res["tracks"][1].append(f"D{n}: {msg}")
            elif kind == "trigger":
                _, own, e = w
                tid = ctext(e, "id")
                bar.set(i / total, f"Событие {tid}"[:26])
                if canon(e) == canon(trig_cur.get(tid)):
                    res["triggers"][0] += 1
                    continue
                ok, msg, _ = nvr.put(f"{TRIG_PATH}/{tid}", e)
                if ok:
                    res["triggers"][0] += 1
                else:
                    res["triggers"][1].append(f"{tid}: {msg}")
            else:
                _, (typ, tid, own), sx = w
                e = parse_xml(sx)
                remap_event(e, mapping, off)
                new_id = renum(tid, own, mapping[own], off)
                if child(e, "id") is not None:
                    set_text(e, "id", new_id)
                bar.set(i / total, f"График {new_id}"[:26])
                path = f"/ISAPI/Event/schedules/{typ}/{new_id}"
                if canon(e) == canon(nvr.get_xml(path)):
                    res["scheds"][0] += 1
                    continue
                ok, msg, _ = nvr.put(path, e)
                if ok:
                    res["scheds"][0] += 1
                else:
                    res["scheds"][1].append(f"{new_id}: {msg}")
    finally:
        fails = sum(len(v[1]) for v in res.values())
        bar.finish(ok=(fails == 0), prefix="Перенос настроек")
    job["results"]["settings"] = res
    save_job(job)
    return res


VERIFY_WAIT = 90        # сколько ждать, пока камеры выйдут на связь
VERIFY_MIN = 15         # раньше этого не сдаёмся даже при ошибках пароля


def wait_online(nvr, slots, timeout, bar=None):
    """Ждём, пока slots выйдут на связь. Ошибки пароля сами не пройдут —
    если остались только они, выходим раньше."""
    start = time.time()
    st = {}
    while True:
        st = read_status(nvr)
        on = sum(1 for s in slots if st.get(s, {}).get("online"))
        if bar:
            bar.set(on / max(len(slots), 1), f"На связи {on} из {len(slots)}")
        if on == len(slots):
            break
        el = time.time() - start
        if el > timeout:
            break
        pending = [s for s in slots if not st.get(s, {}).get("online")]
        if el > VERIFY_MIN and all(is_pw_error(st.get(s)) for s in pending):
            break
        time.sleep(3)
    return st


def run_verify(ctx, job):
    """Камеры на связи? В каждом слоте та камера и то имя, что по плану?"""
    nvr = ctx.nvr
    snap = job["snap"]
    # ждём только камеры, которые были на связи до операции:
    # выключенная камера и после переноса не появится — незачем ждать её 90 с
    targets = sorted({int(n) for o, n in job["pairs"]
                      if (snap["status"].get(str(o)) or {}).get("online") is not False})
    bar = ProgressBar("Проверка связи").start()
    st = {}
    try:
        st = wait_online(nvr, targets, VERIFY_WAIT, bar)
    finally:
        bar.finish(ok=all(st.get(s, {}).get("online") for s in targets),
                   prefix="Проверка связи")
    if not st:
        st = read_status(nvr)
    try:
        live = read_channels(nvr)
    except Exception:
        live = {}
    res = {}
    for o, n in job["pairs"]:
        cam = snap["chans"][str(o)]
        lv = live.get(int(n))
        s = st.get(int(n), {})
        was = snap["status"].get(str(o)) or {}
        res[str(n)] = {
            "from": int(o),
            "placed": bool(lv) and cam_key(lv) == cam_key(cam),
            "name_ok": bool(lv) and lv["name"] == cam["name"],
            "live_name": lv["name"] if lv else "",
            "online": bool(s.get("online")),
            "detect": s.get("detect", ""),
            "was_online": was.get("online"),
            "pw_fixed": False,
        }
    job["results"]["verify"] = res
    fix_names(ctx, job)
    save_job(job)
    return res


def fix_names(ctx, job):
    """Имя не совпало (регистратор мог подставить своё, пока камера
    переподключалась) — записываем канал ещё раз, когда камера уже на связи."""
    res = job["results"].get("verify", {})
    snap = job["snap"]["chans"]
    for n, r in res.items():
        if not r.get("placed") or r.get("name_ok"):
            continue
        cam = snap[str(r["from"])]
        send_channel(ctx.nvr, "PUT", f"{CH_PATH}/{n}",
                     channel_body(cam, int(n), ctx.password_for(cam)))
        cur = _channel_xml(ctx.nvr, int(n))
        if cur is not None:
            r["live_name"] = ctext(parse_xml(cur), "name")
            r["name_ok"] = r["live_name"] == cam["name"]


def refresh_verify(ctx, job):
    """Свежие статусы камер операции (на связи / ошибка пароля)."""
    st = read_status(ctx.nvr)
    for n, r in job["results"].get("verify", {}).items():
        s = st.get(int(n))
        if s:
            r["online"], r["detect"] = bool(s.get("online")), s.get("detect", "")


def settings_fails(job):
    sett = job["results"].get("settings", {})
    return sum(len(v[1]) for v in sett.values())


def repair_settings(ctx, job):
    """Повторить перенос настроек, пока регистратор не перестанет быть занят
    (пишется только то, что ещё отличается)."""
    while settings_fails(job):
        print(Fore.YELLOW + f"\n  Не перенеслось настроек: {settings_fails(job)} "
              "(регистратор был занят)." + Style.RESET_ALL)
        print(Fore.YELLOW + "  1 — повторить через 15 сек   ESC — пропустить"
              + Style.RESET_ALL)
        if getch_choice({"1", "\x1b"}) == "\x1b":
            return
        print("  Жду 15 сек…")
        time.sleep(15)
        run_settings(ctx, job)


# ------------------------------------------------------------------ #
#  Проверка камер (связь, запись, звук)
# ------------------------------------------------------------------ #
def parse_local(s):
    """'2026-10-20T11:22:33+03:00' / '...Z' → наивное локальное время."""
    s = (s or "").strip()
    if not s:
        return None
    s = re.sub(r"(Z|[+-]\d{2}:?\d{2})$", "", s).replace(" ", "T").split(".")[0]
    try:
        return dt.datetime.strptime(s, "%Y-%m-%dT%H:%M:%S")
    except ValueError:
        return None


def nvr_now(nvr):
    """Текущее время регистратора и его смещение от UTC (минуты)."""
    root = nvr.get_xml("/ISAPI/System/time")
    if root is not None:
        raw = first_text(root, "localTime")
        d = parse_local(raw)
        if d:
            mt = re.search(r"([+-])(\d{2}):?(\d{2})$", raw)
            tz = 0
            if mt:
                tz = (int(mt.group(2)) * 60 + int(mt.group(3))) \
                    * (1 if mt.group(1) == "+" else -1)
            return d, tz
    return dt.datetime.now().replace(microsecond=0), 0


def read_storage(nvr):
    """[{id, status, cap_gb, free_gb}] — диски регистратора."""
    out = []
    root = nvr.get_xml("/ISAPI/ContentMgmt/Storage")
    if root is None:
        return out
    for h in findall_local(root, "hdd"):
        cap = _int(first_text(h, "capacity"), 0)
        free = _int(first_text(h, "freeSpace"), 0)
        out.append({"id": first_text(h, "id"), "status": first_text(h, "status"),
                    "cap_gb": round(cap / 1024), "free_gb": round(free / 1024)})
    return out


def search_records(nvr, track, t0, t1, max_pages=10):
    """Отрезки архива [(начало, конец)] за период. None — поиск не работает."""
    segs, pos = [], 0
    for _ in range(max_pages):
        sid = "{" + str(uuid.uuid4()).upper() + "}"
        body = (
            '<?xml version="1.0" encoding="utf-8"?>'
            "<CMSearchDescription>"
            f"<searchID>{sid}</searchID>"
            f"<trackList><trackID>{track}</trackID></trackList>"
            "<timeSpanList><timeSpan>"
            f"<startTime>{t0.strftime('%Y-%m-%dT%H:%M:%SZ')}</startTime>"
            f"<endTime>{t1.strftime('%Y-%m-%dT%H:%M:%SZ')}</endTime>"
            "</timeSpan></timeSpanList>"
            "<maxResults>100</maxResults>"
            f"<searchResultPostion>{pos}</searchResultPostion>"
            "<metadataList><metadataDescriptor>//recordType.meta.std-cgi.com"
            "</metadataDescriptor></metadataList>"
            "</CMSearchDescription>"
        )
        try:
            r = nvr.s.post(nvr.base + "/ISAPI/ContentMgmt/search",
                           data=body.encode("utf-8"),
                           headers={"Content-Type": "application/xml"},
                           timeout=HTTP_TIMEOUT)
            root = ET.fromstring(r.content) if r.status_code == 200 else None
        except (requests.RequestException, ET.ParseError):
            root = None
        if root is None:
            return segs if pos else None
        items = findall_local(root, "searchMatchItem")
        for it in items:
            a = parse_local(first_text(it, "startTime"))
            b = parse_local(first_text(it, "endTime"))
            if a and b:
                segs.append((a, b))
        pos += len(items)
        if first_text(root, "responseStatusStrg") != "MORE" or not items:
            break
    return segs


def coverage(segs, t0, t1):
    """Доля периода [t0, t1], покрытая архивом (0..1)."""
    total = (t1 - t0).total_seconds()
    if total <= 0 or not segs:
        return 0.0
    iv = sorted((max(a, t0), min(b, t1)) for a, b in segs if b > t0 and a < t1)
    covered, cur_a, cur_b = 0.0, None, None
    for a, b in iv:
        if cur_b is None or a > cur_b:
            if cur_b is not None:
                covered += (cur_b - cur_a).total_seconds()
            cur_a, cur_b = a, b
        else:
            cur_b = max(cur_b, b)
    if cur_b is not None:
        covered += (cur_b - cur_a).total_seconds()
    return min(covered / total, 1.0)


def stream_info(nvr, slot, off):
    """Параметры основного потока и звука (через регистратор)."""
    for path in (f"/ISAPI/ContentMgmt/StreamingProxy/channels/{slot}01",
                 f"/ISAPI/Streaming/channels/{slot + off}01"):
        root = nvr.get_xml(path)
        if root is None:
            continue
        v = child(root, "Video")
        a = child(root, "Audio")
        info = {"codec": "", "res": "", "fps": "", "kbps": "",
                "audio": None, "acodec": ""}
        if v is not None:
            info["codec"] = ctext(v, "videoCodecType")
            w, h = ctext(v, "videoResolutionWidth"), ctext(v, "videoResolutionHeight")
            info["res"] = f"{w}×{h}" if w and h else ""
            fr = _int(ctext(v, "maxFrameRate"))
            info["fps"] = str(fr // 100 if fr and fr >= 100 else fr or "")
            cbr = ctext(v, "videoQualityControlType").upper() == "CBR"
            info["kbps"] = ctext(v, "constantBitRate" if cbr else "vbrUpperCap")
        if a is not None:
            info["audio"] = ctext(a, "enabled").lower() == "true"
            info["acodec"] = ctext(a, "audioCompressionType")
        return info
    return None


# заводские имена Hikvision: «IPCamera 01», «IPdome 05», «Camera 01»
DEFAULT_NAME = re.compile(r"^(ip ?camera|ip ?dome|camera)[\s_-]*0*\d+$", re.I)
LAG_OK = dt.timedelta(minutes=10)


def check_one(nvr, slot, cam, st, track, off, now, tz, hours, vmd_rec=None):
    """Одна камера → строка отчёта с вердиктом 0 (ок) / 1 (внимание) / 2 (ошибка)."""
    t0 = now - dt.timedelta(hours=hours)
    ti = track_info(track)
    segs = search_records(nvr, track_id(slot, off), t0, now)
    if segs and tz:
        # часть прошивок отдаёт время архива в UTC — выравниваем
        last = max(b for _, b in segs)
        lag = now - last
        if lag > LAG_OK and abs(lag - dt.timedelta(minutes=tz)) < LAG_OK:
            sh = dt.timedelta(minutes=tz)
            segs = [(a + sh, b + sh) for a, b in segs]
    online = bool(st and st.get("online"))
    stream = stream_info(nvr, slot, off) if online else None

    notes, lvl = [], 0

    def add(msg, level):
        nonlocal lvl
        notes.append(msg)
        lvl = max(lvl, level)

    if not online:
        add(f"нет связи: {detect_ru(st)}", 2)
    continuous = "CMR" in ti["modes"]
    if ti["enabled"] is False:
        add("запись выключена", 2)
    elif track is not None and ti["blocks"] == 0:
        add("нет расписания записи", 2)
    elif track is None:
        add("дорожка записи не найдена", 1)

    last = max((b for _, b in segs), default=None) if segs else None
    cov = coverage(segs, t0, now) if segs is not None else None
    if segs is None:
        add("архив не читается (поиск не поддерживается?)", 1)
    elif continuous:
        if last is None or now - last > LAG_OK:
            add("постоянная запись, но архив не пополняется", 2)
        elif cov < 0.95:
            add(f"пропуски в архиве: покрыто {cov * 100:.0f}% за {hours} ч", 1)
    elif ti["modes"] and not segs:
        add(f"запись по событиям: за {hours} ч записей нет", 1)

    # запись по движению сработает, только если движение «включает запись»
    if "MOTION" in ti["modes"] and not continuous:
        if vmd_rec is False:
            add("запись по движению, но в связке движения нет записи канала",
                2 if not segs else 1)
        # Настройку детекции отдают не все прошивки: после 3 неудач подряд
        # (и ни одной удачи) больше не спрашиваем — не тратим время.
        md_state = nvr.cache.setdefault("md", {"ok": 0, "fail": 0})
        if online and (md_state["ok"] or md_state["fail"] < 3):
            md = nvr.get_xml(f"/ISAPI/System/Video/inputs/channels/{slot + off}/motionDetection")
            md_state["ok" if md is not None else "fail"] += 1
            if md is not None and first_text(md, "enabled").lower() == "false":
                add("запись по движению, но детекция движения в камере выключена", 2)
            elif md is not None:
                grid = first_text(md, "gridMap")
                if grid and not grid.strip("0"):
                    add("детекция движения: зона не нарисована — движения не будет", 2)
                if first_text(md, "sensitivityLevel") == "0":
                    add("детекция движения: чувствительность 0", 2)
        # график охраны движения (когда детекция вообще работает)
        vs_state = nvr.cache.setdefault("vmd_sched", {"ok": 0, "fail": 0})
        if vs_state["ok"] or vs_state["fail"] < 3:
            sc = read_schedule(nvr, "motionDetections", f"VMD_video{slot + off}", slot + off)
            vs_state["ok" if sc is not None else "fail"] += 1
            if sc is not None:
                h = schedule_hours(sc)
                if h == 0:
                    add("график охраны движения пуст — запись по движению не сработает", 2)
                elif h < 7 * 24:
                    notes.append(f"график охраны движения: {h:.0f} ч в неделю (не круглосуточно)")

    if online and stream is None:
        add("параметры потока не читаются", 1)
    elif stream is not None:
        if stream["audio"] is False:
            add("звук выключен", 1)
        elif stream["audio"] is None:
            notes.append("звук не поддерживается камерой")

    if not cam["name"] or DEFAULT_NAME.match(cam["name"]):
        add("имя не задано (стандартное)", 1)

    return {
        "slot": slot, "name": cam["name"], "ip": cam["ip"],
        "online": online, "detect": detect_ru(st),
        "rec_enabled": ti["enabled"], "modes": modes_ru(ti["modes"]),
        "last": last, "cov": cov, "stream": stream,
        "continuous": continuous, "events": len(segs) if segs is not None else None,
        "notes": notes, "level": lvl,
    }


def run_check(nvr, hours=24, workers=4):
    """Проверка всех камер регистратора. Возвращает (rows, meta)."""
    bar = ProgressBar("Чтение каналов").start()
    ok = False
    try:
        chans = read_channels(nvr)
        bar.set(0.3, "Статусы")
        status = read_status(nvr)
        bar.set(0.5, "Дорожки записи")
        tracks = read_tracks(nvr)
        off = detect_offset(chans, tracks)
        bar.set(0.8, "Диски и время")
        now, tz = nvr_now(nvr)
        hdds = read_storage(nvr)
        bar.set(0.9, "Связки движения")
        vmd = {}                       # слот -> есть ли «запись» в связке движения
        for tx in read_triggers(nvr):
            e = parse_xml(tx)
            if ctext(e, "eventType").upper() != "VMD":
                continue
            own = own_slot(e, off)
            if own:
                vmd[own] = any(first_text(n, "notificationMethod").lower() == "record"
                               for n in findall_local(e, "EventTriggerNotification"))
        ok = True
    finally:
        bar.finish(ok=ok, prefix="Чтение каналов")

    rows = []
    n = max(len(chans), 1)
    bar = ProgressBar("Проверка камер").start()
    ok = False
    try:
        with ThreadPoolExecutor(max_workers=workers) as ex:
            futs = {ex.submit(check_one, nvr, s, c, status.get(s),
                              tracks.get(track_id(s, off)), off, now, tz, hours,
                              vmd.get(s)): s
                    for s, c in chans.items()}
            for i, f in enumerate(as_completed(futs), 1):
                s = futs[f]
                try:
                    rows.append(f.result())
                except Exception as e:
                    c = chans[s]
                    rows.append({"slot": s, "name": c["name"], "ip": c["ip"],
                                 "online": False, "detect": "", "rec_enabled": None,
                                 "modes": "—", "last": None, "cov": None,
                                 "stream": None, "notes": [f"ошибка проверки: {e}"],
                                 "level": 2})
                bar.set(i / n, f"Проверено {i} из {len(chans)}")
        ok = True
    finally:
        bar.finish(ok=ok, prefix="Проверка камер")

    # один IP в двух слотах — почти всегда ошибка при добавлении
    by_key = {}
    for r in rows:
        by_key.setdefault(cam_key(chans[r["slot"]]), []).append(r["slot"])
    for r in rows:
        dup = [s for s in by_key[cam_key(chans[r["slot"]])] if s != r["slot"]]
        if dup:
            r["notes"].append("та же камера в " + ", ".join(f"D{s}" for s in dup))
            r["level"] = max(r["level"], 1)

    rows.sort(key=lambda r: r["slot"])
    meta = {"now": now, "hours": hours, "hdds": hdds, "off": off,
            "max_slot": read_max_slot(nvr, chans), "chans": chans}
    return rows, meta


# ------------------------------------------------------------------ #
#  Настройки регистраторов (меню по «0»)
# ------------------------------------------------------------------ #
def press_any():
    try:
        import msvcrt
        print(Fore.YELLOW + "  — нажмите любую клавишу —" + Style.RESET_ALL)
        ch = msvcrt.getch()
        if ch in (b"\x00", b"\xe0"):
            msvcrt.getch()
    except Exception:
        try:
            input()
        except Exception:
            pass


def _edit_recorder_dialog(rec):
    """Диалог создания/редактирования. Пустой ввод = оставить текущее."""
    cur = rec or {"name": "", "lan_ip": "", "lan_port": 80,
                  "rem_host": "", "rem_port": 0, "disabled": []}

    def ask(label, val):
        hint = f" [{val}]" if val not in (None, "", 0) else ""
        return read_line_visible(f"  {label}{hint}: ").strip()

    print()
    print(Fore.CYAN + ("  Новый регистратор" if rec is None
                       else f"  Редактирование: {cur['name']}") + Style.RESET_ALL)
    name = ask("Название", cur["name"]) or cur["name"]
    lan_ip = ask("Внутренний IP (локальная сеть)", cur["lan_ip"]) or cur["lan_ip"]
    lp = ask("Внутренний порт (обычно 80)", cur["lan_port"])
    lan_port = int(lp) if lp.isdigit() else cur["lan_port"]
    rem_host = ask("Внешний IP/хост (для удалёнки)", cur["rem_host"])
    rem_host = rem_host if rem_host else cur["rem_host"]
    rp = ask("Внешний порт", cur["rem_port"])
    rem_port = int(rp) if rp.isdigit() else cur["rem_port"]
    dis = ask("Камеры, отключённые намеренно (IP через запятую, «-» — очистить)",
              ", ".join(cur.get("disabled") or []))
    if dis == "-":
        disabled = []
    elif dis:
        disabled = [x.strip() for x in dis.replace(" ", ",").split(",") if x.strip()]
    else:
        disabled = list(cur.get("disabled") or [])

    if not lan_ip and not rem_host:
        print(Fore.RED + "  Нужен хотя бы внутренний IP или внешний хост."
              + Style.RESET_ALL)
        press_any()
        return None
    if not name:
        name = lan_ip or rem_host or "регистратор"
    return {"name": name, "lan_ip": lan_ip, "lan_port": lan_port,
            "rem_host": rem_host, "rem_port": rem_port, "disabled": disabled}


def settings_menu(recs):
    """Меню настроек: изменить / добавить / удалить регистраторы, сохранить в exe."""
    while True:
        cls()
        banner("Настройки регистраторов")
        if not recs:
            print(Fore.YELLOW + "  Список пуст — добавьте регистратор (A)."
                  + Style.RESET_ALL)
        for i, r in enumerate(recs, 1):
            rem = (f"{r['rem_host']}:{r['rem_port']}"
                   if r["rem_host"] and r["rem_port"] else "—")
            print(f"  {Style.BRIGHT}{i}{Style.NORMAL}. {r['name']}")
            print(f"       внутренний: {r['lan_ip']}:{r['lan_port']}")
            print(f"       внешний:    {rem}")
        print()
        print("  [цифра] изменить   [A] добавить   [D] удалить")
        print("  [S] сохранить и назад   [Q] назад без сохранения")
        ans = read_line_visible(Fore.YELLOW + "  Выбор: "
                                + Style.RESET_ALL).strip().lower()

        if ans in ("q", ""):
            return
        if ans == "s":
            ok, err = save_recorders(recs)
            if ok:
                CON.add_badge("✓  SUCCESS — настройки сохранены в exe", ok=True)
            else:
                CON.add_badge(f"✗  ERROR — не сохранено: {err}", ok=False)
            press_any()
            return
        if ans == "a":
            r = _edit_recorder_dialog(None)
            if r:
                recs.append(r)
            continue
        if ans == "d":
            num = read_line_visible("  Номер для удаления: ").strip()
            if num.isdigit() and 1 <= int(num) <= len(recs):
                removed = recs.pop(int(num) - 1)
                print(Fore.GREEN + f"  Удалён: {removed['name']}"
                      + Style.RESET_ALL)
                press_any()
            continue
        if ans.isdigit() and 1 <= int(ans) <= len(recs):
            r = _edit_recorder_dialog(recs[int(ans) - 1])
            if r:
                recs[int(ans) - 1] = r
            continue


# ------------------------------------------------------------------ #
#  Excel
# ------------------------------------------------------------------ #
_thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)
HEAD_FILL = PatternFill("solid", fgColor="1F4E78")
HEAD_FONT = Font(bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(bold=True, size=14, color="1F4E78")
LABEL_FONT = Font(bold=True, color="404040")
FILL_OK = PatternFill("solid", fgColor="E2EFDA")
FILL_WARN = PatternFill("solid", fgColor="FFF2CC")
FILL_ERR = PatternFill("solid", fgColor="F8CBAD")
FILL_MAIN = PatternFill("solid", fgColor="DCE6F1")
CENTER = Alignment(horizontal="center", vertical="center")
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)


def _xl_head(ws, title, info, ncols):
    """Заголовок + блок «ключ: значение». Возвращает номер строки шапки таблицы."""
    last = get_column_letter(ncols)
    ws.merge_cells(f"A1:{last}1")
    ws["A1"] = title
    ws["A1"].font = TITLE_FONT
    row = 3
    for k, v in info:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
        c = ws.cell(row=row, column=1, value=k)
        c.font = LABEL_FONT
        ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=ncols)
        ws.cell(row=row, column=3, value=v).alignment = LEFT
        row += 1
    return row + 1


def _xl_table_head(ws, row, headers):
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=row, column=c, value=h)
        cell.fill, cell.font, cell.alignment, cell.border = \
            HEAD_FILL, HEAD_FONT, CENTER, BORDER


def _xl_row(ws, row, vals, fill, bold=False, left_cols=()):
    for c, v in enumerate(vals, 1):
        cell = ws.cell(row=row, column=c, value=v)
        cell.fill, cell.border = fill, BORDER
        cell.alignment = LEFT if c in left_cols else CENTER
        if bold:
            cell.font = Font(bold=True, color="1F3864")


def build_move_excel(path, job):
    """Карта переезда: где теперь каждая камера и где искать её старый архив."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Карта переезда"
    snap = job["snap"]["chans"]
    ver = job["results"].get("verify", {})
    sett = job["results"].get("settings", {})

    def stat(key):
        okn, fails = sett.get(key, [0, []])
        return f"{okn} — ок" + (f", ошибок: {len(fails)}" if fails else "")

    online = sum(1 for r in ver.values() if r["online"])
    is_rb = job["kind"] == "rollback"
    info = [
        ("Регистратор:", job["recorder"]),
        ("Операция:", job["title"]),
        ("Начало:", job["created"]),
        ("Окончание:", job["finished"] or "—"),
        ("Камер затронуто:", str(len(job["pairs"]))),
        ("Расписание записи:", stat("tracks")),
        ("Связки событий:", stat("triggers")),
        ("Графики охраны:", stat("scheds")),
        ("На связи после:", f"{online} из {len(ver)}"),
    ]
    hr = _xl_head(ws, "Откат перестановки камер" if is_rb else
                  "Карта переезда камер", info, 6)
    ws.merge_cells(start_row=hr - 1, start_column=1, end_row=hr - 1, end_column=6)
    c = ws.cell(row=hr - 1, column=1, value=(
        f"Архив, записанный до {job['created']}, остался под СТАРЫМИ номерами "
        "слотов — ищите его по колонке «Был слот»."))
    c.font = Font(bold=True, italic=True, color="C00000")
    hr += 1
    _xl_table_head(ws, hr, ["Слот", "Камера", "IP-адрес", "Был слот",
                            "Связь", "Примечание"])
    r = hr + 1
    first_src = job["pairs"][0][0] if job["pairs"] else None
    for o, n in sorted(job["pairs"], key=lambda p: p[1]):
        cam = snap[str(o)]
        v = ver.get(str(n), {})
        notes = []
        if v and not v.get("placed"):
            notes.append("камера в слоте не совпадает с планом")
        if v and v.get("placed") and not v.get("name_ok"):
            notes.append("имя не совпало" + (f": на регистраторе «{v['live_name']}»"
                                            if v.get("live_name") else ""))
        if v.get("pw_fixed"):
            notes.append("пароль введён отдельно")
        if v and not v.get("online") and v.get("was_online") is False:
            notes.append("была офлайн и до переноса — пароль на регистраторе "
                         "записан общий, проверить при включении")
        if int(n) in job["results"].get("skipped", []):
            notes.append("связки событий не перенесены — камера офлайн "
                         "(регистратор их не принимает)")
        if not v:
            fill, link = FILL_WARN, "не проверено"
        elif v.get("online") and v.get("placed"):
            fill, link = FILL_OK, "на связи"
        elif v.get("was_online") is False:
            fill, link = FILL_WARN, DETECT_RU.get(v.get("detect"), v.get("detect") or "офлайн")
        else:
            fill, link = FILL_ERR, DETECT_RU.get(v.get("detect"), v.get("detect") or "офлайн")
        main = (o == first_src and not is_rb)
        _xl_row(ws, r, [f"D{n}", cam["name"] or "—", cam["ip"],
                        f"D{o}" if o != n else "—", link, "; ".join(notes)],
                FILL_MAIN if main and fill is FILL_OK else fill,
                bold=main, left_cols=(2, 6))
        r += 1
    freed = sorted(set(job["touched"]) - {n for _, n in job["pairs"]})
    for s in freed:
        _xl_row(ws, r, [f"D{s}", "— слот освобождён —", "", "", "", ""],
                PatternFill("solid", fgColor="F2F2F2"), left_cols=(2,))
        r += 1
    rep = job.get("replaced")
    if rep and not is_rb:
        _xl_row(ws, r, [f"D{rep['slot']}", f"{rep['name']} — удалена с регистратора",
                        rep["ip"], "", "", f"её архив до {job['created']} остался в "
                        f"D{rep['slot']}"], FILL_WARN, left_cols=(2, 6))
        r += 1
    fails = [x for k in ("tracks", "triggers", "scheds")
             for x in sett.get(k, [0, []])[1]]
    if fails:
        r += 1
        ws.cell(row=r, column=1, value="Не перенеслось (проверить вручную):").font = \
            Font(bold=True, color="C00000")
        for f in fails:
            r += 1
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
            ws.cell(row=r, column=1, value=f).alignment = LEFT
    for i, w in enumerate([8, 30, 18, 10, 30, 42], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = ws.cell(row=hr + 1, column=1)
    wb.save(path)


def hdd_text(h):
    """Полный диск у регистратора — норма: идёт кольцевая перезапись."""
    st = h["status"]
    if st.lower() in ("ok", "normal") and h["cap_gb"] and h["free_gb"] == 0:
        return f"HDD{h['id']}: {st}, {h['cap_gb']} ГБ — заполнен, идёт перезапись (норма)"
    return f"HDD{h['id']}: {st} ({h['free_gb']} из {h['cap_gb']} ГБ свободно)"


def archive_text(x):
    """Постоянная запись — % покрытия; по событиям — число записей
    (для записи по движению % ничего не говорит)."""
    if x.get("continuous"):
        return f"{x['cov'] * 100:.0f}%" if x.get("cov") is not None else "—"
    ev = x.get("events")
    return "—" if ev is None else f"{ev} зап."


def build_check_excel(path, rec_name, rows, meta):
    wb = Workbook()
    ws = wb.active
    ws.title = "Проверка камер"
    n_ok = sum(1 for r in rows if r["level"] == 0)
    n_w = sum(1 for r in rows if r["level"] == 1)
    n_e = sum(1 for r in rows if r["level"] == 2)
    hdd = "; ".join(hdd_text(h) for h in meta["hdds"]) or "нет данных"
    info = [
        ("Регистратор:", rec_name),
        ("Время регистратора:", meta["now"].strftime("%d.%m.%Y %H:%M:%S")),
        ("Период архива:", f"последние {meta['hours']} ч"),
        ("Камер:", f"{len(rows)}   (ок {n_ok} / внимание {n_w} / ошибки {n_e})"),
        ("Диски:", hdd),
        ("Свободные слоты:", free_ranges(meta["chans"], meta["max_slot"]) or "нет"),
    ]
    hr = _xl_head(ws, "Проверка камер: связь, запись, звук", info, 10)
    _xl_table_head(ws, hr, ["Слот", "Камера", "IP-адрес", "Связь", "Запись",
                            "Последняя запись", f"Архив {meta['hours']} ч",
                            "Звук", "Видео", "Замечания"])
    r = hr + 1
    for x in rows:
        s = x["stream"] or {}
        audio = ("вкл" + (f" ({s['acodec']})" if s.get("acodec") else "")
                 if s.get("audio") else "выкл" if s.get("audio") is False
                 else "нет" if s else "—")
        video = " ".join(v for v in (s.get("codec", ""), s.get("res", ""),
                                     f"{s['fps']} к/с" if s.get("fps") else "",
                                     f"{s['kbps']} кбит/с" if s.get("kbps") else "") if v) or "—"
        rec = ("выкл" if x["rec_enabled"] is False else x["modes"])
        _xl_row(ws, r, [
            f"D{x['slot']}", x["name"] or "—", x["ip"], x["detect"], rec,
            x["last"].strftime("%d.%m.%Y %H:%M") if x["last"] else "—",
            archive_text(x),
            audio, video, "; ".join(x["notes"]) or "ок",
        ], (FILL_OK, FILL_WARN, FILL_ERR)[x["level"]], left_cols=(2, 10))
        r += 1
    for i, w in enumerate([7, 26, 16, 22, 18, 17, 10, 14, 30, 48], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = ws.cell(row=hr + 1, column=1)
    ws.auto_filter.ref = f"A{hr}:J{r - 1}"
    wb.save(path)


def open_file(path):
    try:
        os.startfile(path)   # откроет Excel автоматически (Windows)
    except Exception:
        pass


# ------------------------------------------------------------------ #
#  Экраны
# ------------------------------------------------------------------ #
def _cut(s, n):
    s = s or ""
    return s if len(s) <= n else s[:n - 1] + "…"


def load_layout(nvr):
    """Каналы + статусы + число слотов (с прогресс-баром)."""
    bar = ProgressBar("Чтение каналов").start()
    ok = False
    try:
        chans = read_channels(nvr)
        bar.set(0.5, "Статусы камер")
        status = read_status(nvr)
        max_slot = read_max_slot(nvr, chans)
        ok = True
    finally:
        bar.finish(ok=ok, prefix="Чтение каналов")
    return chans, status, max_slot


def print_channels(chans, status, max_slot, disabled=()):
    print()
    banner(f"Каналы: {len(chans)} камер из {max_slot} слотов")
    for s in sorted(chans):
        c = chans[s]
        st = status.get(s)
        if st and st["online"]:
            tag = Fore.GREEN + "● на связи" + Style.RESET_ALL
        else:
            tag = Fore.RED + "● " + detect_ru(st) + Style.RESET_ALL
        lock = Fore.MAGENTA + "  [PoE]" + Style.RESET_ALL if c["locked"] else ""
        if is_placeholder(c["ip"]):
            tag = Fore.MAGENTA + "● временная заглушка HIK-Slots" + Style.RESET_ALL
        elif c["ip"] in disabled and not (st and st["online"]):
            tag = Style.DIM + "● отключена намеренно" + Style.RESET_ALL
        print(f"  {Style.BRIGHT}D{s:<4}{Style.NORMAL} {_cut(c['name'], 28):<28} "
              f"{c['ip']:<16} {tag}{lock}")
    fr = free_ranges(chans, max_slot)
    print(Fore.CYAN + f"\n  Свободные слоты: {fr or 'нет'}" + Style.RESET_ALL)


def ask_slot(prompt, chans, max_slot, need_cam):
    """Номер слота ('115' / 'D115'). Пусто — отмена (None)."""
    while True:
        raw = read_line_visible(Fore.YELLOW + prompt + Style.RESET_ALL).strip()
        if not raw:
            return None
        s = _int(raw.lstrip("DdДд"))
        if not s or s < 1 or (max_slot and s > max_slot):
            print(Fore.RED + f"  Нет такого слота (1…{max_slot})." + Style.RESET_ALL)
            continue
        if need_cam and s not in chans:
            print(Fore.RED + f"  В слоте D{s} нет камеры." + Style.RESET_ALL)
            continue
        return s


def confirm_yes(prompt="  Запустить? [Y/n]: "):
    """Одна клавиша: Y или Enter — да, N или Esc — нет.
    Русская раскладка тоже понимается (Н = Y, Т = N)."""
    sys.stdout.write(Fore.YELLOW + prompt + Style.RESET_ALL)
    sys.stdout.flush()
    try:
        import msvcrt
    except ImportError:
        return input().strip().lower() in ("", "y", "yes", "д", "да", "н")
    while True:
        c = msvcrt.getwch()
        if c in ("\x00", "\xe0"):
            msvcrt.getwch()
            continue
        if c == "\x03":
            raise KeyboardInterrupt
        k = c.lower()
        if k in ("\r", "\n", "y", "н", "д", "\x1b", "n", "т"):
            yes = k in ("\r", "\n", "y", "н", "д")
            print("да" if yes else "нет")
            # «да» целиком + Enter: лишние буквы не должны попасть в пароль
            time.sleep(0.3)
            while msvcrt.kbhit():
                msvcrt.getwch()
            return yes


def print_plan(plan, chans, status=None):
    print()
    banner(f"План: {plan['title']}")
    first = plan["pairs"][0]
    for o, n in sorted(plan["pairs"], key=lambda p: p[1]):
        c = chans[o]
        line = (f"  D{n:<4} ← {_cut(c['name'], 26):<26} {c['ip']:<16} "
                f"(было D{o})")
        if (o, n) == tuple(first):
            print(Fore.YELLOW + Style.BRIGHT + line + Style.RESET_ALL)
        else:
            print(line)
    if plan["freed"]:
        print(Fore.CYAN + f"  D{plan['freed']:<4} — освободится" + Style.RESET_ALL)
    if plan.get("replaced"):
        r = chans[plan["replaced"]]
        print(Fore.RED + Style.BRIGHT + f"  D{plan['replaced']:<4} ✗ {_cut(r['name'], 26):<26} "
              f"{r['ip']:<16} будет удалена с регистратора" + Style.RESET_ALL)
    status = status or {}
    off_cams = [(o, chans[o]) for o, _ in plan["pairs"]
                if status.get(o) and not status[o].get("online")]
    if off_cams:
        print()
        print(Fore.YELLOW + "  Камеры офлайн (выключены / недоступны):" + Style.RESET_ALL)
        for o, c in off_cams:
            print(Fore.YELLOW + f"    D{o} {c['name']} ({c['ip']})" + Style.RESET_ALL)
        print(Fore.YELLOW + "  Их пароль не проверить — будет записан общий. "
              "Связки событий для них не переносятся." + Style.RESET_ALL)
    # по замерам на реальном регистраторе: ~8 с на шаг + ~6 с на камеру + ~40 с связь
    secs = len(plan["steps"]) * 8 + len(plan["pairs"]) * 6 + 40
    print()
    print(f"  Камер: {Style.BRIGHT}{len(plan['pairs'])}{Style.NORMAL}, займёт около "
          f"{max(1, round(secs / 60))} мин. Каждая камера будет офлайн 10–30 сек.")
    print("  Имя, расписание записи и события переедут вместе с камерой.")
    print(Fore.RED + "  Архив остаётся под старыми номерами слотов — он будет указан "
          "в карте переезда." + Style.RESET_ALL)


def check_cam_password(cam, pw):
    """Пробуем войти в саму камеру: True — пароль подходит, False — нет,
    None — камера с этого ПК недоступна (проверить нельзя). Одна попытка,
    чтобы не заблокировать учётку камеры."""
    r = cam_request(cam, pw, "GET", "/ISAPI/System/deviceInfo", timeout=3)
    if r is None:
        return None
    if r.status_code == 200:
        return True
    if r.status_code == 401:
        return False
    return None


CAM_URL = "http://{ip}{path}"       # прямой доступ к камере, минуя регистратор


def cam_request(cam, pw, method, path, body=None, timeout=4):
    """Запрос прямо к камере (её собственные настройки). None — недоступна."""
    ip = cam.get("ip") or ""
    if not ip or is_placeholder(ip):
        return None
    if body is not None and not isinstance(body, (bytes, str)):
        body = xml_bytes(body)
    try:
        return requests.request(method, CAM_URL.format(ip=ip, path=path), data=body,
                                auth=HTTPDigestAuth(cam.get("user") or "admin", pw),
                                headers={"Content-Type": "application/xml"} if body else None,
                                timeout=timeout, verify=False)
    except requests.RequestException:
        return None


def cam_vmd_trigger(ctx, cam):
    """Связка движения в самой камере: (Element | None, состояние).
    Состояние: ok / auth (пароль камеры неизвестен) / unreachable / unsupported."""
    key = cam_key(cam)
    bad = ctx.__dict__.setdefault("cam_auth_bad", set())
    if key in bad and key not in ctx.pw_by_cam:
        return None, "auth"            # не долбим камеру неверным паролем (блокировка)
    pw = ctx.pw_by_cam.get(key) or ctx.pw_common or ctx.nvr_pw
    r = cam_request(cam, pw, "GET", f"/ISAPI/Event/triggers/VMD-{cam.get('src_ch') or 1}")
    if r is None:
        return None, "unreachable"
    if r.status_code == 401:
        bad.add(key)
        return None, "auth"
    if r.status_code != 200:
        return None, "unsupported"
    try:
        return ET.fromstring(r.content), "ok"
    except ET.ParseError:
        return None, "unsupported"


def precheck_passwords(ctx, cams):
    """До перестановки: подходит ли пароль к каждой камере. Неподходящим —
    спрашиваем свой пароль. False — пользователь отменил операцию."""
    cams = [c for c in cams if c.get("ip")]
    if not cams:
        return True
    res = {}
    bar = ProgressBar("Проверка паролей").start()
    try:
        with ThreadPoolExecutor(max_workers=8) as ex:
            futs = {ex.submit(check_cam_password, c, ctx.password_for(c)): c for c in cams}
            for i, f in enumerate(as_completed(futs), 1):
                res[cam_key(futs[f])] = f.result()
                bar.set(i / len(cams), f"Проверено {i} из {len(cams)}")
    finally:
        bar.finish(ok=not any(v is False for v in res.values()), prefix="Проверка паролей")
    bad = [c for c in cams if res[cam_key(c)] is False]
    for c in bad:
        print(Fore.RED + f"  ✗ {c['name']} ({c['ip']}) — пароль не подходит" + Style.RESET_ALL)
        for _ in range(3):
            pw = read_password("    Пароль этой камеры: ")
            if not pw:
                print(Fore.GREEN + "  Отменено. На регистраторе ничего не менялось."
                      + Style.RESET_ALL)
                return False
            r = check_cam_password(c, pw)
            if r is not False:
                ctx.pw_by_cam[cam_key(c)] = pw
                print(Fore.GREEN + "    ✓ подходит" + Style.RESET_ALL)
                break
            print(Fore.RED + "    ✗ не подходит" + Style.RESET_ALL)
        else:
            print(Fore.RED + "  Пароль так и не подошёл — операция отменена, "
                  "на регистраторе ничего не менялось." + Style.RESET_ALL)
            return False
    return True


def ask_cam_password(ctx):
    print()
    pw = read_password("Пароль камер: ")
    ctx.pw_common = pw if pw else ctx.nvr_pw


def fix_passwords(ctx, job):
    """Камеры с «неверным паролем»: спрашиваем пароль именно для них."""
    res = job["results"].get("verify", {})
    snap = job["snap"]["chans"]
    bad = sorted((int(n) for n, r in res.items()
                  if is_pw_error(r)))
    if not bad:
        return
    print()
    banner(f"Другой пароль у {len(bad)} камер")
    print(Fore.YELLOW + "  Введите пароль для каждой."
          + Style.RESET_ALL)
    for n in bad:
        r = res[str(n)]
        cam = snap[str(r["from"])]
        for _ in range(3):
            pw = read_password(f"  D{n} {cam['name']}: ")
            if not pw:
                break
            if check_cam_password(cam, pw) is False:   # проверяем на самой камере
                print(Fore.RED + "    ✗ не подходит (проверено на камере)" + Style.RESET_ALL)
                continue
            ctx.pw_by_cam[cam_key(cam)] = pw
            ok, msg = do_place(ctx.nvr, n, cam, pw)
            if not ok:
                print(Fore.RED + f"    не записалось: {msg}" + Style.RESET_ALL)
                continue
            print(Fore.YELLOW + "    проверяю связь…" + Style.RESET_ALL)
            st = wait_online(ctx.nvr, [n], 30).get(n, {})
            r["online"], r["detect"] = bool(st.get("online")), st.get("detect", "")
            if r["online"]:
                r["pw_fixed"] = True
                print(Fore.GREEN + "    ✓ на связи" + Style.RESET_ALL)
                break
            print(Fore.RED + f"    {detect_ru(st)}" + Style.RESET_ALL)
        save_job(job)
    fix_names(ctx, job)          # пока камера была без пароля, имя не ставилось
    save_job(job)


def fail_menu(job):
    print()
    print(Fore.RED + Style.BRIGHT + "  Ошибка: " + job["error"] + Style.RESET_ALL)
    print(f"  Выполнено шагов: {job['done']} из {len(job['steps'])}.")
    print(Fore.YELLOW + "  1 — повторить шаг   2 — откатить сделанное   "
          "ESC — остановиться (продолжить можно позже)" + Style.RESET_ALL)
    return {"1": "retry", "2": "rollback", "\x1b": "stop"}[
        getch_choice({"1", "2", "\x1b"})]


def finish_report(job):
    fname = job["file"].replace(".json", ".xlsx")
    path = os.path.join(output_dir(), fname)
    try:
        build_move_excel(path, job)
        print(Fore.GREEN + Style.BRIGHT + f"\n  Карта переезда: {path}" + Style.RESET_ALL)
        open_file(path)
    except Exception as e:
        print(Fore.RED + f"  Excel не сохранён: {e}" + Style.RESET_ALL)


def execute_job(ctx, job):
    """Полный цикл операции по журналу: цепочка → проверка связи → настройки.
    Настройки пишем, когда камеры уже на связи: пока они переподключаются,
    регистратор отвечает «503 Device Busy».
    Можно вызывать для прерванного журнала — продолжит с нужного места."""
    if job["state"] == "chain":
        while not run_chain(ctx, job):
            act = fail_menu(job)
            if act == "retry":
                continue
            if act == "rollback":
                return rollback_job(ctx, job, ask_pw=False)
            CON.add_badge("✗  Операция остановлена — продолжить можно при следующем входе",
                          ok=False)
            return False
        job["state"] = "verify"
        save_job(job)
    verified_now = False
    if job["state"] == "verify":
        run_verify(ctx, job)
        fix_passwords(ctx, job)
        verified_now = True
        job["state"] = "settings"
        save_job(job)
    if job["state"] == "settings":
        if not verified_now:                     # продолжение после выхода
            if "verify" not in job["results"]:   # журнал старой версии
                run_verify(ctx, job)
            else:
                refresh_verify(ctx, job)
            fix_passwords(ctx, job)
        run_settings(ctx, job)
        repair_settings(ctx, job)
        job["state"] = "done"
        job["finished"] = dt.datetime.now().strftime("%d.%m.%Y %H:%M:%S")
        save_job(job)
    if job["kind"] == "rollback" and job.get("rollback_of"):
        for j in load_jobs(job["recorder"]):
            if j["file"] == job["rollback_of"]:
                j["state"] = "rolled_back"
                save_job(j)
    finish_report(job)
    ver = job["results"].get("verify", {})
    bad = [n for n, r in ver.items() if not (r["online"] and r["placed"])
           and r.get("was_online") is not False]
    sett = job["results"].get("settings", {})
    sfail = sum(len(v[1]) for v in sett.values())
    if not bad and not sfail:
        CON.add_badge("✓  SUCCESS — все камеры на своих местах и на связи", ok=True)
        return True
    msg = []
    if bad:
        msg.append(f"не на связи: {', '.join('D' + n for n in sorted(bad, key=int))}")
    if sfail:
        msg.append(f"настроек не перенеслось: {sfail}")
    CON.add_badge("✗  Внимание — " + "; ".join(msg) + " (см. Excel)", ok=False)
    return False


def cams_to_place(steps, snap):
    """Камеры, которые эти шаги будут записывать в слоты (без повторов)."""
    seen, out = set(), []
    for st in steps:
        if st[0] == "place":
            cam = snap["chans"].get(str(st[2]))
            if cam and cam_key(cam) not in seen:
                seen.add(cam_key(cam))
                out.append(cam)
    return out


def rollback_job(ctx, orig, ask_pw=True):
    """Откат: обратная цепочка сделанных шагов + исходные настройки слотов."""
    done = orig["steps"][:orig["done"]]
    if not done and orig["state"] == "chain":
        print(Fore.YELLOW + "  Изменений не было — откатывать нечего." + Style.RESET_ALL)
        orig["state"] = "rolled_back"
        save_job(orig)
        return True
    snap = orig["snap"]
    pairs = [(s, s) for s in orig["touched"] if str(s) in snap["chans"]]
    occupied = [x for x in orig["touched"] if str(x) in snap["chans"]]
    steps = rollback_steps(done, occupied, orig["touched"])
    if ask_pw:
        ask_cam_password(ctx)
        if not precheck_passwords(ctx, cams_to_place(steps, snap)):
            return False
    rb = new_job(orig["recorder"], "rollback", f"Откат: {orig['title']}",
                 steps, snap, orig["off"], orig["touched"], pairs)
    rb["rollback_of"] = orig["file"]
    save_job(rb)
    orig["state"] = "rolling_back"
    save_job(orig)
    return execute_job(ctx, rb)


# ------------------------------------------------------------------ #
#  Пункты меню
# ------------------------------------------------------------------ #
def ui_list(ctx):
    chans, status, max_slot = load_layout(ctx.nvr)
    print_channels(chans, status, max_slot, disabled_ips(ctx))


def ui_single(ctx):
    """Одна камера без сдвига: в свободный слот — просто переезд; в занятый —
    выбор: заменить (старая камера удаляется) или поменять местами."""
    chans, status, max_slot = load_layout(ctx.nvr)
    print_channels(chans, status, max_slot, disabled_ips(ctx))
    print()
    banner("Перенести камеру в другой слот")
    src = ask_slot("  Какую камеру перенести: ",
                   chans, max_slot, True)
    if src is None:
        return
    dst = ask_slot("  В какой слот (свободный или занятый): ", chans, max_slot, False)
    if dst is None:
        return
    kind = "move"
    if dst in chans and dst != src:
        c = chans[dst]
        print()
        print(Fore.YELLOW + f"  D{dst} занят: {c['name']} ({c['ip']}, "
              f"{detect_ru(status.get(dst))})." + Style.RESET_ALL)
        print(f"  {Style.BRIGHT}1{Style.NORMAL} — заменить: {c['name']} будет "
              f"{Fore.RED}удалена с регистратора{Style.RESET_ALL}, D{src} освободится")
        print(f"  {Style.BRIGHT}2{Style.NORMAL} — поменять местами")
        print(Fore.YELLOW + "  ESC — отмена" + Style.RESET_ALL)
        ch = getch_choice({"1", "2", "\x1b"})
        if ch == "\x1b":
            return
        if ch == "1":
            plan, err = plan_replace(chans, src, dst)
            kind = "replace"
        else:
            plan, err = plan_swap(chans, src, dst)
            kind = "swap"
    else:
        plan, err = plan_move(chans, src, dst)
    run_plan_ui(ctx, plan, err, chans, status, kind)


def ui_move(ctx, swap=False):
    chans, status, max_slot = load_layout(ctx.nvr)
    print_channels(chans, status, max_slot, disabled_ips(ctx))
    print()
    if swap:
        banner("Поменять две камеры местами")
        a = ask_slot("  Первый слот: ", chans, max_slot, False)
        if a is None:
            return
        b = ask_slot("  Второй слот: ", chans, max_slot, False)
        if b is None:
            return
        plan, err = plan_swap(chans, a, b)
    else:
        banner("Перенести камеру со сдвигом остальных")
        src = ask_slot("  Какую камеру перенести: ",
                       chans, max_slot, True)
        if src is None:
            return
        dst = ask_slot("  В какой слот поставить: ", chans, max_slot, False)
        if dst is None:
            return
        plan, err = plan_move(chans, src, dst)
    run_plan_ui(ctx, plan, err, chans, status, "swap" if swap else "move")


def run_plan_ui(ctx, plan, err, chans, status, kind):
    """План на экран → подтверждение → снимок → выполнение."""
    if err:
        print(Fore.RED + "  " + err + Style.RESET_ALL)
        CON.add_badge("✗  ERROR — операция невозможна", ok=False)
        return

    print_plan(plan, chans, status)
    if not confirm_yes():
        print(Fore.GREEN + "  Отменено. На регистраторе ничего не менялось."
              + Style.RESET_ALL)
        return
    ask_cam_password(ctx)
    if not precheck_passwords(ctx, [chans[o] for o, _ in plan["pairs"]]):
        return

    touched = {s for p in plan["pairs"] for s in p}
    snap, off = take_snapshot(ctx.nvr, chans, status, touched)
    job = new_job(ctx.rec["name"], kind, plan["title"],
                  plan["steps"], snap, off, touched, plan["pairs"])
    if plan.get("replaced"):
        r = chans[plan["replaced"]]
        job["replaced"] = {"slot": plan["replaced"], "name": r["name"], "ip": r["ip"]}
    save_job(job)
    execute_job(ctx, job)


def disabled_ips(ctx):
    return set(ctx.rec.get("disabled") or [])


def save_disabled(ctx, ips):
    """Список «отключена намеренно» — в настройках регистратора (внутри exe)."""
    ctx.rec["disabled"] = sorted(ips)
    recs = load_recorders()
    for r in recs:
        if r["name"] == ctx.rec["name"]:
            r["disabled"] = ctx.rec["disabled"]
    return save_recorders(recs)


def apply_disabled(ctx, rows):
    """Намеренно выключенная камера — не ошибка: серым, без тревоги."""
    dis = disabled_ips(ctx)
    for x in rows:
        if x["ip"] in dis and not x["online"]:
            x["detect"] = "отключена намеренно"
            x["notes"] = ["отключена намеренно (не ошибка)"]
            x["level"] = 0


def ui_check(ctx):
    print()
    banner("Проверка всех камер")
    rows, meta = run_check(ctx.nvr)
    apply_disabled(ctx, rows)
    ctx.last_check = (time.time(), rows, meta)      # пункт 3 возьмёт готовый список
    print()
    for h in meta["hdds"]:
        okh = h["status"].lower() in ("ok", "normal")
        print((Fore.GREEN if okh else Fore.RED)
              + "  " + hdd_text(h)
              + Style.RESET_ALL)
    if not meta["hdds"]:
        print(Fore.YELLOW + "  Диски: нет данных" + Style.RESET_ALL)
    print()
    sym = {0: Fore.GREEN + "✓", 1: Fore.YELLOW + "!", 2: Fore.RED + "✗"}
    for x in rows:
        notes = "; ".join(x["notes"]) if x["level"] else "ок"
        col = (Fore.GREEN, Fore.YELLOW, Fore.RED)[x["level"]]
        print(f"  {sym[x['level']]}{Style.RESET_ALL} {Style.BRIGHT}D{x['slot']:<4}"
              f"{Style.NORMAL} {_cut(x['name'], 24):<24} {x['ip']:<16} "
              f"{col}{notes}{Style.RESET_ALL}")
    n_ok = sum(1 for r in rows if r["level"] == 0)
    n_w = sum(1 for r in rows if r["level"] == 1)
    n_e = sum(1 for r in rows if r["level"] == 2)
    print(f"\n  Итого: {Fore.GREEN}ок {n_ok}{Style.RESET_ALL} / "
          f"{Fore.YELLOW}внимание {n_w}{Style.RESET_ALL} / "
          f"{Fore.RED}ошибки {n_e}{Style.RESET_ALL}")

    # выключенные камеры можно один раз отметить «отключена намеренно»
    off = [x for x in rows if not x["online"] and x["ip"] not in disabled_ips(ctx)
           and not is_placeholder(x["ip"])]
    if off:
        print(Fore.CYAN + "\n  Если какие-то камеры из «нет связи» выключены специально, их можно\n"
              "  отметить — проверка перестанет считать их ошибкой." + Style.RESET_ALL)
        raw = read_line_visible(Fore.YELLOW + "  Слоты через запятую: " + Style.RESET_ALL).strip()
        if raw:
            by_slot = {x["slot"]: x for x in off}
            picked = {_int(p.strip().lstrip("DdДд")) for p in raw.replace(" ", ",").split(",")
                      if p.strip()}
            ips = {by_slot[s]["ip"] for s in picked if s in by_slot}
            if ips:
                ok, err = save_disabled(ctx, disabled_ips(ctx) | ips)
                print((Fore.GREEN + "  ✓ Отмечено: " + ", ".join(sorted(ips))) if ok
                      else (Fore.RED + f"  ✗ Не сохранено: {err}"), Style.RESET_ALL)
                apply_disabled(ctx, rows)
                n_e = sum(1 for r in rows if r["level"] == 2)
    fname = (f"proverka_{_safe(ctx.rec['name'])}_"
             f"{dt.datetime.now().strftime('%d%m%Y_%H%M')}.xlsx")
    path = os.path.join(output_dir(), fname)
    try:
        build_check_excel(path, ctx.rec["name"], rows, meta)
        print(Fore.GREEN + Style.BRIGHT + f"\n  Отчёт: {path}" + Style.RESET_ALL)
        open_file(path)
    except Exception as e:
        print(Fore.RED + f"  Excel не сохранён: {e}" + Style.RESET_ALL)
    CON.add_badge("✓  SUCCESS — проверка завершена" if n_e == 0 else
                  f"✗  Ошибки у {n_e} камер — см. отчёт", ok=(n_e == 0))


# ------------------------------------------------------------------ #
#  Исправление одной камеры: подробный аудит → что починить
# ------------------------------------------------------------------ #
MD_PATH = "/ISAPI/System/Video/inputs/channels/{}/motionDetection"

# порядок исправлений: сначала связь (без неё камера не примет настройки)
FIX_ORDER = ("pw", "campw", "reconnect", "name", "track", "md_on", "md_zone", "md_sens",
             "vmd_link", "arming", "cam_center", "audio")


def full_grid(md):
    """gridMap «весь кадр» под сетку камеры: строки × столбцы, побитно.
    22 столбца → 3 байта на строку: ff ff fc (как у D1 на реальном регистраторе)."""
    rows = _int(first_text(md, "rowGranularity"), 18)
    cols = _int(first_text(md, "columnGranularity"), 22)
    nbytes = (cols + 7) // 8
    bits = "1" * cols + "0" * (nbytes * 8 - cols)
    row_hex = "".join(f"{int(bits[i:i + 8], 2):02x}" for i in range(0, len(bits), 8))
    return row_hex * rows


DAY_IDX = {"Monday": 0, "Tuesday": 1, "Wednesday": 2, "Thursday": 3,
           "Friday": 4, "Saturday": 5, "Sunday": 6}


def _hours(t):
    try:
        h, m, sec = (int(x) for x in (t or "0:0:0").split(":"))
    except ValueError:
        return 0.0
    return h + m / 60 + sec / 3600


def track_hours(t):
    """Сколько часов в неделю расписание записи вообще пишет."""
    total = 0.0
    if t is None:
        return 0.0
    for sa in findall_local(t, "ScheduleAction"):
        st = child(sa, "ScheduleActionStartTime")
        en = child(sa, "ScheduleActionEndTime")
        if st is None or en is None or first_text(sa, "Record").lower() == "false":
            continue
        a = DAY_IDX.get(ctext(st, "DayOfWeek"), 0) * 24 + _hours(ctext(st, "TimeOfDay"))
        b = DAY_IDX.get(ctext(en, "DayOfWeek"), 0) * 24 + _hours(ctext(en, "TimeOfDay"))
        if b <= a:                 # «воскресенье 00:00 → понедельник 00:00» — через конец недели
            b += 7 * 24
        total += b - a
    return min(total, 7 * 24)


def last_record_before(nvr, tid, t_end, days):
    """Последняя запись раньше t_end (ищем по дням назад, до days дней)."""
    for d in range(days):
        b = t_end - dt.timedelta(days=d)
        segs = search_records(nvr, tid, b - dt.timedelta(days=1), b, max_pages=3)
        if segs:
            return max(e for _, e in segs)
        if segs is None:
            return None
    return None


def audit_camera(ctx, s):
    """Подробный аудит одной камеры. Возвращает словарь: факты для экрана,
    проблемы (с ключом исправления или без) и всё прочитанное для исправлений."""
    nvr = ctx.nvr
    bar = ProgressBar("Аудит камеры").start()
    ok = False
    try:
        chans = read_channels(nvr)
        status = read_status(nvr)
        cam = chans[s]
        st = status.get(s) or {}
        bar.set(0.15, "Запись")
        tracks = read_tracks(nvr)
        off = detect_offset(chans, tracks)
        v = s + off
        tid = track_id(s, off)
        track = nvr.get_xml(f"/ISAPI/ContentMgmt/record/tracks/{tid}")
        ti = track_info(track)
        bar.set(0.3, "Архив за сутки")
        now, tz = nvr_now(nvr)
        t0 = now - dt.timedelta(hours=24)
        segs = search_records(nvr, tid, t0, now)
        if segs and tz:
            lag = now - max(b for _, b in segs)
            if lag > LAG_OK and abs(lag - dt.timedelta(minutes=tz)) < LAG_OK:
                sh = dt.timedelta(minutes=tz)
                segs = [(a + sh, b + sh) for a, b in segs]
        online = bool(st.get("online"))
        bar.set(0.5, "Детекция движения")
        md = nvr.get_xml(MD_PATH.format(v)) if online else None
        bar.set(0.6, "Связки событий")
        vmd = None
        for tx in read_triggers(nvr):
            e = parse_xml(tx)
            if own_slot(e, off) == s and ctext(e, "eventType").upper() == "VMD":
                vmd = e
                break
        bar.set(0.75, "График охраны")
        sched = read_schedule(nvr, "motionDetections", f"VMD_video{v}", v)
        bar.set(0.85, "Поток и звук")
        stream_root, stream_path = None, None
        if online:
            for path in (f"/ISAPI/ContentMgmt/StreamingProxy/channels/{s}01",
                         f"/ISAPI/Streaming/channels/{v}01"):
                stream_root = nvr.get_xml(path)
                if stream_root is not None:
                    stream_path = path
                    break
        ok = True
    finally:
        bar.finish(ok=ok, prefix="Аудит камеры")

    facts, issues = [], []

    def fact(level, text):
        facts.append((level, text))

    def issue(key, level, text, fix_text=None):
        issues.append({"key": key, "level": level, "text": text, "fix": fix_text})
        fact(level, text)

    # связь и имя
    if online:
        fact(0, "Связь: на связи")
    elif is_pw_error(st):
        issue("pw", 2, "Связь: неверный пароль камеры — камера не пишет",
              "ввести пароль камеры (проверяется на самой камере)")
    elif cam["ip"] in disabled_ips(ctx):
        fact(0, "Связь: отключена намеренно")
    else:
        issue("offline", 2, f"Связь: {detect_ru(st)} — проверьте питание и сеть камеры "
              f"({cam['ip']}); если выключена специально — отметьте в проверке (пункт 5)")
    if not cam["name"] or DEFAULT_NAME.match(cam["name"]):
        issue("name", 1, f"Имя: «{cam['name']}» — заводское", "задать имя")
    else:
        fact(0, f"Имя: {cam['name']}")
    fact(0, f"Адрес: {cam['ip']}:{cam['port']}  ({cam['proto'] or '—'}, пользователь "
            f"{cam['user'] or '—'})")

    # расписание записи и архив
    continuous = "CMR" in ti["modes"]
    motion = "MOTION" in ti["modes"] and not continuous
    if track is None:
        fact(1, f"Расписание записи: дорожка {tid} не читается")
    elif ti["enabled"] is False or ti["blocks"] == 0:
        issue("track", 2, "Расписание записи: " + ("выключено" if ti["enabled"] is False
                                                    else "пустое"),
              "скопировать расписание записи с другой камеры")
    else:
        rh = track_hours(track)
        if 0 < rh < 7 * 24:
            issue("track", 1, f"Расписание записи: {modes_ru(ti['modes'])}, "
                  f"только {rh:.0f} ч в неделю", "скопировать расписание записи с другой камеры")
        else:
            fact(0, f"Расписание записи: {modes_ru(ti['modes'])}, круглосуточно"
                 if rh >= 7 * 24 else f"Расписание записи: {modes_ru(ti['modes'])}")
    last = max((b for _, b in segs), default=None) if segs else None
    if segs is None:
        fact(1, "Архив: поиск не работает")
    elif continuous:
        cov = coverage(segs, t0, now)
        if last is None or now - last > LAG_OK:
            issue("reconnect" if online else "noarch", 2,
                  "Архив: постоянная запись, но архив не пополняется",
                  "переподключить камеру к регистратору" if online else None)
        else:
            fact(0 if cov >= 0.95 else 1, f"Архив: покрыто {cov * 100:.0f}% за 24 ч, "
                 f"последняя запись {last:%d.%m %H:%M}")
    else:
        if segs:
            fact(0, f"Архив: {len(segs)} записей за 24 ч, последняя {last:%d.%m %H:%M}")
        else:
            prev = last_record_before(nvr, tid, t0, 30)
            fact(1, "Архив: за 24 ч записей нет; " +
                 (f"последняя запись {prev:%d.%m.%Y %H:%M} "
                  f"({(now - prev).days} дн. назад)" if prev else "за 30 дней записей нет"))

    # движение: детекция в камере, связка «движение → запись», график охраны
    if motion or md is not None:
        if md is None and not online:
            fact(1, "Детекция движения: не проверить, пока камера не на связи")
        elif md is None:
            fact(1 if motion else 0, "Детекция движения: не читается")
        else:
            en = first_text(md, "enabled").lower() == "true"
            grid = first_text(md, "gridMap")
            sens = first_text(md, "sensitivityLevel")
            if not en:
                issue("md_on", 2 if motion else 1, "Детекция движения: выключена",
                      "включить детекцию движения")
            if grid and not grid.strip("0"):
                issue("md_zone", 2 if motion else 1, "Детекция движения: зона не нарисована",
                      "зона — весь кадр")
            if sens == "0":
                issue("md_sens", 2 if motion else 1, "Детекция движения: чувствительность 0",
                      "чувствительность 50")
            if en and grid.strip("0") and sens != "0":
                fact(0, f"Детекция движения: включена, чувствительность {sens}")
    if motion:
        if vmd is None:
            fact(1, "Связка движения: не найдена")
        else:
            has_rec = any(first_text(n, "notificationMethod").lower() == "record"
                          for n in findall_local(vmd, "EventTriggerNotification"))
            if has_rec:
                fact(0, "Связка движения: включает запись")
            else:
                issue("vmd_link", 2, "Связка движения: нет записи канала — "
                      "движение не включит запись", "добавить запись канала в связку")
        if sched is None:
            fact(1, "График охраны движения: не читается")
        else:
            h = schedule_hours(sched)
            if h == 0:
                issue("arming", 2, "График охраны движения: пуст", "сделать круглосуточным")
            elif h < 7 * 24:
                issue("arming", 0, f"График охраны движения: {h:.0f} ч в неделю",
                      "сделать круглосуточным")
            else:
                fact(0, "График охраны движения: круглосуточно")
        # Сама камера должна сообщать регистратору о движении
        # («Уведомить центр наблюдения»). Через регистратор этого не видно.
        if online:
            ctrig, cstate = cam_vmd_trigger(ctx, cam)
            if cstate == "ok":
                methods = {first_text(n, "notificationMethod").lower()
                           for n in findall_local(ctrig, "EventTriggerNotification")}
                if "center" in methods:
                    fact(0, "Камера: о движении сообщает регистратору")
                else:
                    issue("cam_center", 2, "Камера: о движении НЕ сообщает регистратору "
                          "(«Уведомить центр наблюдения» выключено) — запись по движению "
                          "не сработает", "включить в камере уведомление регистратора")
            elif cstate == "auth":
                issue("campw", 1, "Камера: её собственные настройки не проверить — "
                      "нужен пароль камеры", "ввести пароль камеры и проверить")
            elif cstate == "unreachable":
                fact(0, "Камера: с этого ПК недоступна — её собственные настройки не проверить")
            else:
                fact(0, "Камера: связку движения в самой камере прочитать нельзя")
            data_cam_trigger = ctrig
        else:
            data_cam_trigger = None
    else:
        data_cam_trigger = None

    # звук и поток
    if stream_root is not None:
        vid = child(stream_root, "Video")
        a = child(stream_root, "Audio")
        if vid is not None:
            fr = _int(ctext(vid, "maxFrameRate"))
            fps = fr // 100 if fr and fr >= 100 else fr
            fact(0, f"Видео: {ctext(vid, 'videoCodecType')} "
                    f"{ctext(vid, 'videoResolutionWidth')}×{ctext(vid, 'videoResolutionHeight')}"
                    f", {fps or '—'} к/с")
        if a is None:
            fact(0, "Звук: камера без звука")
        elif ctext(a, "enabled").lower() != "true":
            issue("audio", 1, "Звук: выключен", "включить звук в основном потоке")
        else:
            fact(0, f"Звук: включён ({ctext(a, 'audioCompressionType')})")
    elif online:
        fact(1, "Поток: параметры не читаются")

    dup = [x for x, c in chans.items() if x != s and cam_key(c) == cam_key(cam)]
    if dup:
        issue("dup", 1, "Та же камера стоит ещё в " + ", ".join(f"D{x}" for x in dup)
              + " — лишний слот можно убрать пунктом 2 (замена) или вручную")
    return {"slot": s, "cam": cam, "off": off, "v": v, "tid": tid, "track": track,
            "md": md, "vmd": vmd, "sched": sched, "stream": stream_root,
            "stream_path": stream_path, "cam_trigger": data_cam_trigger,
            "facts": facts, "issues": issues,
            # действия, которые не «исправляют найденное», а доступны всегда
            "actions": [] if is_placeholder(cam["ip"]) else
                       [{"key": "reboot", "fix": "Перезагрузить камеру"}]}


def print_audit(a):
    print()
    banner(f"Аудит D{a['slot']} — {a['cam']['name']}")
    sym = {0: Fore.GREEN + "✓", 1: Fore.YELLOW + "!", 2: Fore.RED + "✗"}
    for level, text in a["facts"]:
        print(f"  {sym[level]}{Style.RESET_ALL} {text}")


def fix_password(ctx, cam):
    """Пароль камеры для записи канала: известный за сеанс или спросить
    (с проверкой на самой камере). None — пользователь отказался."""
    key = cam_key(cam)
    if key in ctx.pw_by_cam:
        return ctx.pw_by_cam[key]
    for _ in range(3):
        pw = read_password(f"  Пароль камеры {cam['name']}: ")
        pw = pw or ctx.pw_common or ctx.nvr_pw
        r = check_cam_password(cam, pw)
        if r is False:
            print(Fore.RED + "    ✗ не подходит (проверено на камере)" + Style.RESET_ALL)
            continue
        if r is None:
            print(Fore.YELLOW + "    камера с этого ПК недоступна — пароль не проверить"
                  + Style.RESET_ALL)
        ctx.pw_by_cam[key] = pw
        return pw
    return None


def apply_fix(ctx, a, key, backup):
    """Одно исправление. backup — список исходных настроек (для ручного возврата)."""
    nvr, s, v, cam = ctx.nvr, a["slot"], a["v"], a["cam"]

    def put(path, elem, before):
        backup.append({"what": key, "path": path,
                       "before": xml_str(before) if before is not None else None})
        return send_channel(nvr, "PUT", path, elem)

    if key in ("pw", "reconnect", "name"):
        pw = None
        if key == "pw":
            for _ in range(3):
                p = read_password(f"  Пароль камеры {cam['name']}: ")
                if not p:
                    return False, "пароль не введён"
                if check_cam_password(cam, p) is False:
                    print(Fore.RED + "    ✗ не подходит (проверено на камере)" + Style.RESET_ALL)
                    continue
                pw = p
                ctx.pw_by_cam[cam_key(cam)] = p
                break
        else:
            pw = fix_password(ctx, cam)
        if pw is None:
            return False, "нет подходящего пароля камеры"
        body = channel_body(cam, s, pw)
        if key == "name":
            new = read_line_visible(f"  Новое имя для D{s}: ").strip()
            if not new:
                return False, "имя не введено"
            set_text(body, "name", new)
        ok, msg = put(f"{CH_PATH}/{s}", body, parse_xml(cam["xml"]))
        if not ok:
            return False, msg
        if key in ("pw", "reconnect"):
            print(Fore.YELLOW + "    жду, пока камера выйдет на связь…" + Style.RESET_ALL)
            st = wait_online(nvr, [s], 40).get(s, {})
            if not st.get("online"):
                return False, f"записано, но камера {detect_ru(st)}"
        return True, "готово"

    if key == "track":
        ref = ask_slot("  С какой камеры скопировать расписание записи: ",
                       read_channels(nvr), 0, True)
        if ref is None:
            return False, "отменено"
        src = nvr.get_xml(f"/ISAPI/ContentMgmt/record/tracks/{track_id(ref, a['off'])}")
        if src is None or a["track"] is None:
            return False, "дорожка записи не читается"
        want = copy.deepcopy(a["track"])
        transplant(want, src, TRACK_KEYS)
        return put(f"/ISAPI/ContentMgmt/record/tracks/{a['tid']}", want, a["track"])

    if key in ("md_on", "md_zone", "md_sens"):
        md = a.setdefault("md_new", copy.deepcopy(a["md"]))
        if key == "md_on":
            set_text(md, "enabled", "true")
        for e in md.iter():
            if key == "md_zone" and _local(e.tag) == "gridMap":
                e.text = full_grid(md)
            if key == "md_sens" and _local(e.tag) == "sensitivityLevel":
                e.text = "50"
        return put(MD_PATH.format(v), md, a["md"])

    if key == "vmd_link":
        e = copy.deepcopy(a["vmd"])
        ns = _ns(e)
        lst = child(e, "EventTriggerNotificationList")
        if lst is None:
            lst = ET.SubElement(e, _q(ns, "EventTriggerNotificationList"))
        dyn = child(e, "dynVideoInputChannelID") is not None
        n = ET.SubElement(lst, _q(ns, "EventTriggerNotification"))
        ET.SubElement(n, _q(ns, "id")).text = f"record-{s if dyn else v}"
        ET.SubElement(n, _q(ns, "notificationMethod")).text = "record"
        ET.SubElement(n, _q(ns, "dynVideoInputID" if dyn else "videoInputID")).text = \
            str(s if dyn else v)
        return put(f"{TRIG_PATH}/{ctext(e, 'id')}", e, a["vmd"])

    if key == "arming":
        sc = copy.deepcopy(a["sched"])
        ns = _ns(sc)
        tbl = child(sc, "TimeBlockList")
        if tbl is None:
            tbl = ET.SubElement(sc, _q(ns, "TimeBlockList"))
        for c in list(tbl):
            tbl.remove(c)
        for d in range(1, 8):
            tb = ET.SubElement(tbl, _q(ns, "TimeBlock"))
            ET.SubElement(tb, _q(ns, "dayOfWeek")).text = str(d)
            tr = ET.SubElement(tb, _q(ns, "TimeRange"))
            ET.SubElement(tr, _q(ns, "beginTime")).text = "00:00:00"
            ET.SubElement(tr, _q(ns, "endTime")).text = "24:00:00"
        tbl.set("size", "7")
        return put(f"/ISAPI/Event/schedules/motionDetections/{ctext(sc, 'id')}", sc, a["sched"])

    if key == "audio":
        sr = copy.deepcopy(a["stream"])
        set_text(child(sr, "Audio"), "enabled", "true")
        return put(a["stream_path"], sr, a["stream"])

    if key == "reboot":
        pw = fix_password(ctx, cam)
        if pw is None:
            return False, "нет подходящего пароля камеры"
        r = cam_request(cam, pw, "PUT", "/ISAPI/System/reboot")
        if r is None or r.status_code != 200:
            # камера с ПК недоступна — просим регистратор передать команду
            ok, msg, _ = nvr.send("PUT", f"{CH_PATH}/{s}/reboot")
            if not ok:
                return False, ("камера недоступна с этого ПК, а регистратор "
                               f"не передал команду ({msg})")
        print(Fore.YELLOW + "    камера перезагружается, жду выхода на связь "
              "(обычно 1–2 мин)…" + Style.RESET_ALL)
        time.sleep(20)                       # пока камера реально уйдёт в перезагрузку
        st = wait_online(nvr, [s], 180).get(s, {})
        if not st.get("online"):
            return False, f"после перезагрузки камера {detect_ru(st)}"
        return True, "готово"

    if key == "campw":
        pw = fix_password(ctx, cam)
        if pw is None:
            return False, "пароль не подошёл"
        ctx.__dict__.get("cam_auth_bad", set()).discard(cam_key(cam))
        return True, "пароль принят — проверяю камеру заново"

    if key == "cam_center":
        e = copy.deepcopy(a["cam_trigger"])
        ns = _ns(e)
        lst = child(e, "EventTriggerNotificationList")
        if lst is None:
            lst = ET.SubElement(e, _q(ns, "EventTriggerNotificationList"))
        n = ET.SubElement(lst, _q(ns, "EventTriggerNotification"))
        ET.SubElement(n, _q(ns, "id")).text = "center"
        ET.SubElement(n, _q(ns, "notificationMethod")).text = "center"
        ET.SubElement(n, _q(ns, "notificationRecurrence")).text = "beginning"
        path = f"/ISAPI/Event/triggers/{ctext(e, 'id')}"
        backup.append({"what": key, "path": f"камера {cam['ip']}{path}",
                       "before": xml_str(a["cam_trigger"])})
        pw = ctx.pw_by_cam.get(cam_key(cam)) or ctx.pw_common or ctx.nvr_pw
        r = cam_request(cam, pw, "PUT", path, e)
        if r is None:
            return False, "камера недоступна"
        if r.status_code != 200:
            return False, f"камера ответила HTTP {r.status_code}"
        return True, "готово"

    return False, "исправление не поддерживается"


CHECK_FRESH = 180          # результат общей проверки годен 3 минуты


def problem_cameras(ctx):
    """Проверка всех камер (или свежий результат пункта 2) — только проблемные."""
    cached = ctx.__dict__.get("last_check")
    if cached and time.time() - cached[0] < CHECK_FRESH:
        rows, meta = cached[1], cached[2]
    else:
        rows, meta = run_check(ctx.nvr)
        apply_disabled(ctx, rows)
        ctx.last_check = (time.time(), rows, meta)
    return rows, meta


def ui_fix(ctx):
    print()
    banner("Проверить и исправить камеру")
    rows, meta = problem_cameras(ctx)
    errs = [x for x in rows if x["level"] == 2]
    warns = [x for x in rows if x["level"] == 1]
    for title, group, col in (("Ошибки", errs, Fore.RED), ("Внимание", warns, Fore.YELLOW)):
        if not group:
            continue
        print(col + f"\n  {title}:" + Style.RESET_ALL)
        for x in group:
            print(f"  {Style.BRIGHT}D{x['slot']:<4}{Style.NORMAL} {_cut(x['name'], 24):<24} "
                  f"{x['ip']:<16} {col}{_cut('; '.join(x['notes']), 70)}{Style.RESET_ALL}")
    if not errs and not warns:
        print(Fore.GREEN + "\n  Все камеры в порядке." + Style.RESET_ALL)
    print()
    s = ask_slot("  Какую камеру проверить: ",
                 meta["chans"], meta["max_slot"], True)
    if s is None:
        return
    ctx.last_check = None            # после исправлений общий список устареет
    backup = []
    a = audit_camera(ctx, s)
    # По кругу: пока камера была без пароля, её детекция и звук не читались —
    # после исправления аудит повторяется и предлагает то, что нашлось теперь.
    for _ in range(5):
        print_audit(a)
        fixable = sorted((i for i in a["issues"] if i["fix"]),
                         key=lambda i: FIX_ORDER.index(i["key"]))
        if not fixable:
            msg = ("  Исправлять нечего — всё в порядке." if not a["issues"] else
                   "  Остальное автоматически не исправить — см. подсказки выше."
                   if [i for i in a["issues"] if i["level"]] else
                   "  Исправлять нечего.")
            print(Fore.GREEN + "\n" + msg + Style.RESET_ALL)
        actions = fixable + a["actions"]
        if not actions:
            break
        print()
        print(Fore.CYAN + ("  Можно исправить:" if fixable else "  Действия:") + Style.RESET_ALL)
        for i, it in enumerate(actions, 1):
            hint = f"   {Style.DIM}({it['text']}){Style.RESET_ALL}" if it.get("text") else ""
            print(f"  {Style.BRIGHT}{i}{Style.NORMAL}. {it['fix']}{hint}")
        raw = read_line_visible(Fore.YELLOW + "  Номера через запятую"
                                + (", A — исправить всё" if fixable else "")
                                + ": " + Style.RESET_ALL).strip().lower()
        if not raw:
            break
        if raw in ("a", "ф", "all", "все", "всё"):
            chosen = fixable                     # перезагрузка — только отдельно
        else:
            nums = {_int(x) for x in raw.replace(" ", ",").split(",") if x.strip()}
            chosen = [it for i, it in enumerate(actions, 1) if i in nums]
        if not chosen:
            break
        print()
        for it in chosen:
            print(f"  → {it['fix']}")
            try:
                ok, msg = apply_fix(ctx, a, it["key"], backup)
            except Exception as e:
                ok, msg = False, str(e)
            print((Fore.GREEN + "    ✓ готово" if ok else Fore.RED + "    ✗ " + msg)
                  + Style.RESET_ALL)
        a = audit_camera(ctx, s)
    if backup:
        p = os.path.join(backup_dir(), f"{_safe(ctx.rec['name'])}_"
                         f"{dt.datetime.now():%Y%m%d_%H%M%S}_fix_D{s}.txt")
        with open(p, "w", encoding="utf-8") as f:
            json.dump({"slot": s, "camera": a["cam"]["name"], "changes": backup},
                      f, ensure_ascii=False, indent=1)
    left = [i for i in a["issues"] if i["level"] == 2]
    if not left:
        CON.add_badge(f"✓  SUCCESS — D{s}: ошибок нет", ok=True)
    else:
        CON.add_badge(f"✗  D{s}: осталось ошибок {len(left)}", ok=False)


def ui_rollback(ctx):
    """Откат: выбор из последних операций. Более старую можно откатить, только
    если более новые не трогали те же слоты (иначе — сначала их)."""
    jobs = [j for j in load_jobs(ctx.rec["name"])
            if j["kind"] != "rollback" and j["state"] != "rolled_back"][:9]
    if not jobs:
        print(Fore.YELLOW + "\n  Отменять нечего — перестановок не было." + Style.RESET_ALL)
        return
    if jobs[0]["state"] == "rolling_back":
        rb = next((x for x in load_jobs(ctx.rec["name"])
                   if x.get("rollback_of") == jobs[0]["file"] and x["state"] != "done"), None)
        if rb:
            print(Fore.YELLOW + f"\n  Откат «{jobs[0]['title']}» был прерван — доделываю."
                  + Style.RESET_ALL)
            ask_cam_password(ctx)
            if precheck_passwords(ctx, cams_to_place(rb["steps"][rb["done"]:], rb["snap"])):
                execute_job(ctx, rb)
            return
    print()
    banner("Отменить перестановку")
    for i, j in enumerate(jobs, 1):
        st = "" if j["state"] == "done" else "  (не завершена)"
        print(f"  {Style.BRIGHT}{i}{Style.NORMAL}. {j['created']}  {j['title']}"
              f"  ({len(j['pairs'])} камер){st}")
    print(Fore.YELLOW + "\n  Цифра — какую перестановку отменить   ESC — назад" + Style.RESET_ALL)
    c = getch_choice({str(i) for i in range(1, len(jobs) + 1)} | {"\x1b"})
    if c == "\x1b":
        return
    idx = int(c) - 1
    j = jobs[idx]
    blockers = [x for x in jobs[:idx] if set(x["touched"]) & set(j["touched"])]
    if blockers:
        print(Fore.RED + "  Сначала отмените более новые перестановки с теми же слотами: "
              + "; ".join(f"{x['created']} {x['title']}" for x in blockers) + Style.RESET_ALL)
        return
    print()
    banner(f"Отменить: {j['title']}")
    print(f"  Операция от {j['created']}"
          + ("." if j["state"] == "done" else " — не была завершена."))
    # всё ли сейчас так, как оставила операция? Иначе — предупреждаем.
    diff = []
    try:
        live = read_channels(ctx.nvr)
        occupied = {x: x for x in j["touched"] if str(x) in j["snap"]["chans"]}
        exp = simulate_steps(j["steps"][:j["done"]], occupied)
        for s in j["touched"]:
            want = exp.get(s)
            have = live.get(s)
            if want is None:
                ok = have is None
            elif want == "P":
                ok = have is not None and is_placeholder(have["ip"])
            else:
                ok = have is not None and                     cam_key(have) == cam_key(j["snap"]["chans"][str(want)])
            if not ok:
                diff.append(s)
    except Exception:
        pass
    if diff:
        print(Fore.RED + "  После операции эти слоты меняли: "
              + ", ".join(f"D{s}" for s in sorted(diff))
              + ".\n  Откат вернёт их к состоянию ДО операции." + Style.RESET_ALL)
    print(Fore.YELLOW + "  Затронутые камеры будут кратко офлайн."
          + Style.RESET_ALL)
    if not confirm_yes("  Отменить? [Y/n]: "):
        print(Fore.GREEN + "  Отменено." + Style.RESET_ALL)
        return
    rollback_job(ctx, j)


def ui_repair_last(ctx):
    """У последней операции не перенеслись настройки или не совпало имя —
    предлагаем доделать (без повторной перестановки камер)."""
    jobs = load_jobs(ctx.rec["name"])
    if not jobs or jobs[0]["state"] != "done":
        return False
    j = jobs[0]
    refresh_verify(ctx, j)
    ver = j["results"].get("verify", {})
    pw_bad = [n for n, r in ver.items() if is_pw_error(r)]
    bad_names = [n for n, r in ver.items() if r.get("placed") and not r.get("name_ok")]
    if not settings_fails(j) and not bad_names and not pw_bad:
        return False
    print()
    banner("Последняя операция завершена не полностью")
    print(f"  {Style.BRIGHT}{j['title']}{Style.NORMAL} от {j['created']}:")
    if pw_bad:
        print(Fore.RED + "  • неверный пароль камеры (камера не пишет): "
              + ", ".join(f"D{n}" for n in sorted(pw_bad, key=int)) + Style.RESET_ALL)
    if settings_fails(j):
        print(f"  • не перенеслось настроек: {settings_fails(j)}")
    if bad_names:
        print("  • имя не совпало: " + ", ".join(f"D{n}" for n in sorted(bad_names, key=int)))
    print(Fore.YELLOW + "  1 — доделать сейчас   ESC — позже" + Style.RESET_ALL)
    if getch_choice({"1", "\x1b"}) == "\x1b":
        return True
    fix_passwords(ctx, j)             # спросит пароль именно этих камер
    need = [n for n, r in ver.items() if r.get("placed") and not r.get("name_ok")
            and cam_key(j["snap"]["chans"][str(r["from"])]) not in ctx.pw_by_cam]
    if need:
        ask_cam_password(ctx)
    fix_names(ctx, j)
    save_job(j)
    run_settings(ctx, j)
    repair_settings(ctx, j)
    save_job(j)
    finish_report(j)
    still = [n for n in bad_names if not j["results"]["verify"][n]["name_ok"]]
    if not settings_fails(j) and not still:
        CON.add_badge("✓  SUCCESS — операция доделана", ok=True)
    else:
        CON.add_badge("✗  Осталось: настроек " + str(settings_fails(j))
                      + (", имя " + ", ".join(f"D{n}" for n in still) if still else "")
                      + " (см. Excel)", ok=False)
    press_any()
    return True


def ui_resume(ctx):
    """Незавершённая операция после обрыва — продолжить или откатить."""
    if ui_repair_last(ctx):
        return
    for j in load_jobs(ctx.rec["name"]):
        # «rolling_back» доделывает свой журнал отката (он новее и найдётся раньше)
        if j["state"] not in ("chain", "settings", "verify"):
            continue
        is_rb = j["kind"] == "rollback"
        print()
        banner("Найден незавершённый откат" if is_rb
               else "Найдена незавершённая операция")
        print(f"  {Style.BRIGHT}{j['title']}{Style.NORMAL} от {j['created']} — "
              "прервана, не завершена.")
        if j["error"]:
            print(Fore.RED + "  Последняя ошибка: " + j["error"] + Style.RESET_ALL)
        if is_rb:
            print(Fore.YELLOW + "  1 — доделать откат   ESC — позже" + Style.RESET_ALL)
            c = getch_choice({"1", "\x1b"})
        else:
            print(Fore.YELLOW + "  1 — продолжить   2 — откатить   ESC — позже"
                  + Style.RESET_ALL)
            c = getch_choice({"1", "2", "\x1b"})
        if c == "1":
            ask_cam_password(ctx)
            if precheck_passwords(ctx, cams_to_place(j["steps"][j["done"]:], j["snap"])):
                execute_job(ctx, j)
        elif c == "2":
            rollback_job(ctx, j)
        press_any()
        return
    cleanup_placeholders(ctx)


def cleanup_placeholders(ctx):
    """Заглушки HIK-Slots, оставшиеся после обрыва (незавершённых операций нет —
    значит, они никому не нужны): свободные слоты-заполнители убираем сразу,
    заглушки в занятых слотах — после подтверждения."""
    try:
        chans = read_channels(ctx.nvr)
    except Exception:
        return
    fill = [s for s, c in chans.items() if (c["ip"] or "").startswith(FILL_NET)]
    park = [s for s, c in chans.items() if (c["ip"] or "").startswith(PARK_NET)]
    if not fill and not park:
        return
    print()
    banner("Остались временные каналы прерванной операции")
    for s in fill + park:
        print(f"  D{s:<4} {chans[s]['name']}  ({chans[s]['ip']})")
    print(Fore.YELLOW + "  Это не камеры, а пустые каналы — их можно удалить."
          + Style.RESET_ALL)
    if not confirm_yes("  Удалить их? [Y/n]: "):
        return
    for s in fill + park:
        ok, msg = do_delete(ctx.nvr, s)
        print((Fore.GREEN + f"  ✓ D{s} удалён") if ok else
              (Fore.RED + f"  ✗ D{s}: {msg}"), Style.RESET_ALL)
    press_any()


# ------------------------------------------------------------------ #
#  Главный сценарий
# ------------------------------------------------------------------ #
def choose_recorder():
    """Выбор регистратора (0 — настройки). Возвращает (rec, base_url, via)."""
    recs = load_recorders()
    while True:
        avail = check_all(recs) if recs else []
        print()
        banner("Выберите регистратор")
        if not recs:
            print(Fore.YELLOW + "  Список пуст. Нажмите 0, чтобы добавить "
                  "регистратор." + Style.RESET_ALL)
        for i, rec in enumerate(recs, 1):
            ep = avail[i - 1]
            if ep:
                tag = Fore.GREEN + f"[доступен — {ep[1]}]" + Style.RESET_ALL
            else:
                tag = Fore.RED + Style.BRIGHT + "[НЕДОСТУПЕН]" + Style.RESET_ALL
            print(f"  {Style.BRIGHT}{i}{Style.NORMAL}. {rec['name']}   {tag}")
        print(f"  {Style.BRIGHT}0{Style.NORMAL}. Настройки "
              "(изменить / добавить регистраторы)")
        print(Fore.YELLOW + "\n  Выберите цифру регистратора или 0 — настройки"
              + Style.RESET_ALL)

        valid = {"0"} | {str(i) for i in range(1, min(len(recs), 9) + 1)}
        choice = getch_choice(valid)
        if choice == "0":
            settings_menu(recs)
            recs = load_recorders()
            continue
        idx = int(choice) - 1
        rec = recs[idx]
        ep = avail[idx]
        if not ep:
            print(Fore.RED + f"\n  Регистратор {rec['name']} недоступен ни "
                  f"локально, ни по удалённому порту." + Style.RESET_ALL)
            CON.add_badge("✗  ERROR — регистратор недоступен, выберите другой",
                          ok=False)
            press_any()
            continue
        return rec, ep[0], ep[1]


def login(base_url):
    """Логин/пароль регистратора (3 попытки). Возвращает (session, pw) или (None, None)."""
    for _ in range(3):
        print()
        user = read_line_visible("Логин: ").strip()
        pw = read_password("Пароль: ")
        print(Fore.YELLOW + "Проверка учётных данных…" + Style.RESET_ALL)
        session, err = make_session(base_url, user, pw)
        if session:
            print(Fore.GREEN + "Авторизация успешна." + Style.RESET_ALL)
            return session, pw
        print(Fore.RED + f"Ошибка входа: {err}" + Style.RESET_ALL)
    return None, None


# (заголовок группы, [(клавиша, пункт, функция)])
MENU = [
    ("Камеры", [
        ("1", "Список камер", ui_list),
        ("2", "Проверить все камеры", ui_check),
        ("3", "Проверить и исправить одну камеру", ui_fix),
    ]),
    ("Перестановка", [
        ("4", "Перенести камеру в другой слот", ui_single),
        ("5", "Перенести камеру со сдвигом остальных", lambda c: ui_move(c)),
        ("6", "Поменять две камеры местами", lambda c: ui_move(c, swap=True)),
        ("7", "Отменить перестановку", ui_rollback),
    ]),
]
MENU_ITEMS = [it for _, items in MENU for it in items]


def main():
    cls()
    try:
        import ctypes
        ctypes.windll.kernel32.SetConsoleTitleW(f"HIK-Slots {VERSION}")
    except Exception:
        pass
    banner("HIK-Slots — камеры по слотам (Hikvision NVR)")

    rec, base_url, via = choose_recorder()

    print()
    steps = ["Определение адреса", "TCP-соединение",
             "Проверка ISAPI", "Канал готов"]
    bar = ProgressBar("Подключение").start()
    try:
        for i, st in enumerate(steps):
            bar.set(i / len(steps), st)
            time.sleep(0.45)
    finally:
        bar.finish(ok=True, prefix="Подключение")
    print(Fore.GREEN + f"Подключено к {rec['name']} ({via}): {base_url}"
          + Style.RESET_ALL)

    session, nvr_pw = login(base_url)
    if not session:
        print(Fore.RED + "Не удалось авторизоваться. Выход." + Style.RESET_ALL)
        pause_exit()
        return
    ctx = Ctx(Nvr(session, base_url), rec, via, nvr_pw)

    ui_resume(ctx)

    while True:
        cls()
        banner(f"HIK-Slots — {rec['name']} ({via})")
        for group, items in MENU:
            print(Fore.CYAN + f"\n  {group}" + Style.RESET_ALL)
            for key, title, _ in items:
                print(f"  {Style.BRIGHT}{key}{Style.NORMAL}. {title}")
        print(Fore.YELLOW + "\n  Выберите цифру   |   ESC — выход" + Style.RESET_ALL)
        c = getch_choice({k for k, _, _ in MENU_ITEMS} | {"\x1b"})
        if c == "\x1b":
            break
        fn = next(f for k, _, f in MENU_ITEMS if k == c)
        try:
            fn(ctx)
        except KeyboardInterrupt:
            raise
        except Exception as e:
            print(Fore.RED + f"\n  Ошибка: {e}" + Style.RESET_ALL)
            CON.add_badge("✗  ERROR", ok=False)
        press_any()


# ------------------------------------------------------------------ #
#  Обновление с GitHub (при запуске)
# ------------------------------------------------------------------ #
UPDATE_TIMEOUT = 4


def _update_urls():
    """Адреса GitHub (HIKSLOTS_UPDATE_SRC подменяет их — для проверки обновления)."""
    src = os.environ.get("HIKSLOTS_UPDATE_SRC", "").rstrip("/")
    if src:
        return f"{src}/repos/{REPO}/releases/latest", f"{src}/{REPO}/main/"
    return (f"https://api.github.com/repos/{REPO}/releases/latest",
            f"https://raw.githubusercontent.com/{REPO}/main/")


def _ver(s):
    return tuple(int(x) for x in re.findall(r"\d+", s or "")[:4])


def find_update():
    """(версия, url) самой новой версии на GitHub или None.
    exe: последний Release с приложенным .exe или exe из ветки main;
    запуск скриптом: nvr_slots.py из main."""
    api, raw = _update_urls()
    frozen = getattr(sys, "frozen", False)
    found = []
    if frozen:
        try:
            r = requests.get(api, timeout=UPDATE_TIMEOUT,
                             headers={"Accept": "application/vnd.github+json"})
            if r.status_code == 200:
                j = r.json()
                for a in j.get("assets") or []:
                    if str(a.get("name", "")).lower().endswith(".exe"):
                        found.append((j.get("tag_name", ""), a["browser_download_url"]))
                        break
        except (requests.RequestException, ValueError, KeyError):
            pass
    try:
        r = requests.get(raw + "nvr_slots.py", timeout=UPDATE_TIMEOUT,
                         headers={"Range": "bytes=0-16383"})
        mt = re.search(r'^VERSION = "([\d.]+)"', r.text, re.M) if r.status_code in (200, 206) else None
        if mt:
            found.append((mt.group(1), raw + ("HIK-Slots.exe" if frozen else "nvr_slots.py")))
    except requests.RequestException:
        pass
    newer = [f for f in found if _ver(f[0]) > _ver(VERSION)]
    return max(newer, key=lambda f: _ver(f[0])) if newer else None


def _download(url, dest, bar):
    r = requests.get(url, stream=True, timeout=30)
    r.raise_for_status()
    total = int(r.headers.get("Content-Length") or 0)
    got = 0
    with open(dest, "wb") as f:
        for chunk in r.iter_content(256 * 1024):
            f.write(chunk)
            got += len(chunk)
            if total:
                bar.set(got / total, f"Скачано {got // 1048576} из {total // 1048576} МБ")
    if total and got != total:
        raise RuntimeError("файл скачался не полностью")


def _restart(path):
    """Новая версия в том же окне; эта копия завершается."""
    args = [a for a in sys.argv[1:] if a != "--no-update"]
    cmd = [path] + args if getattr(sys, "frozen", False) else [sys.executable, path] + args
    try:
        CON.stop()
    except Exception:
        pass
    sys.stdout.flush()
    subprocess.Popen(cmd)
    os._exit(0)


def apply_exe_update(exe, url, version, bar):
    """Скачать новый exe, проверить его самопроверкой, перенести настройки
    регистраторов (они внутри exe, в NTFS-потоке) и поставить на место старого.
    Работающий exe Windows переименовать даёт — так и меняем."""
    base = exe[:-4] if exe.lower().endswith(".exe") else exe
    new, old = base + ".new.exe", base + ".old.exe"
    _download(url, new, bar)
    with open(new, "rb") as f:
        if f.read(2) != b"MZ":
            raise RuntimeError("скачан не exe-файл")
    bar.set(1.0, "Проверка новой версии")
    chk = subprocess.run([new, "--selftest"], capture_output=True, timeout=180)
    if b"SELFTEST OK" not in chk.stdout:
        raise RuntimeError("новая версия не прошла самопроверку")
    try:
        with open(exe + ":hikslots", "rb") as f:
            cfg = f.read()
        with open(new + ":hikslots", "wb") as f:
            f.write(cfg)
    except OSError:
        pass                          # настроек ещё не было
    if os.path.exists(old):
        os.remove(old)
    os.rename(exe, old)
    try:
        os.rename(new, exe)
    except OSError:
        os.rename(old, exe)           # вернуть как было
        raise


def apply_script_update(script, url, version, bar):
    """Запуск скриптом: новый nvr_slots.py пишется поверх старого
    (в тот же файл — настройки в его NTFS-потоке сохраняются)."""
    bar.set(0.3, "Скачивание")
    r = requests.get(url, timeout=30)
    r.raise_for_status()
    text = r.content.decode("utf-8")
    if f'VERSION = "{version}"' not in text or "def main(" not in text:
        raise RuntimeError("скачан повреждённый файл")
    compile(text, script, "exec")
    with open(script, "rb") as f:
        prev = f.read()
    with open(script + ".old", "wb") as f:
        f.write(prev)
    with open(script, "w", encoding="utf-8", newline="") as f:
        f.write(text)


def check_update():
    """Первым делом при запуске: есть новая версия — скачать и перезапуститься."""
    bar = ProgressBar("Проверка обновлений").start()
    upd = None
    try:
        upd = find_update()
    finally:
        bar.finish(ok=True, prefix="Проверка обновлений")
    if not upd:
        return
    version, url = upd
    version = version.lstrip("vV")
    print(Fore.CYAN + f"  Доступна версия {version} (у вас {VERSION}) — обновляю."
          + Style.RESET_ALL)
    frozen = getattr(sys, "frozen", False)
    target = sys.executable if frozen else os.path.abspath(__file__)
    bar = ProgressBar("Обновление").start()
    ok = False
    try:
        (apply_exe_update if frozen else apply_script_update)(target, url, version, bar)
        ok = True
    except Exception as e:
        err = str(e)
    finally:
        bar.finish(ok=ok, prefix="Обновление")
    if not ok:
        print(Fore.YELLOW + f"  Обновить не удалось ({err}) — работаю на версии {VERSION}."
              + Style.RESET_ALL)
        for p in (target[:-4] + ".new.exe",) if frozen else ():
            try:
                os.remove(p)
            except OSError:
                pass
        return
    CON.add_badge(f"✓  Обновлено до {version} — перезапуск", ok=True)
    time.sleep(1.2)
    _restart(target)


def cleanup_after_update():
    """Старый exe после обновления удаляется при следующем запуске
    (пока старая копия работала, Windows его не отдавала)."""
    if not getattr(sys, "frozen", False):
        return
    base = sys.executable[:-4]

    def clean():
        for _ in range(20):
            left = 0
            for p in (base + ".old.exe", base + ".new.exe"):
                if os.path.exists(p):
                    try:
                        os.remove(p)
                    except OSError:
                        left += 1
            if not left:
                return
            time.sleep(1)
    threading.Thread(target=clean, daemon=True).start()


def selftest():
    """Проверка без регистратора: зависимости, планировщик, Excel (для сборки .exe)."""
    import requests as _rq
    import openpyxl as _ox
    import colorama as _cl
    fake = {s: {"slot": s, "name": f"Cam {s}", "ip": f"10.0.0.{s}", "src_ch": "1",
                "locked": False} for s in list(range(1, 11)) + [13, 14]}
    # D14 → D3: сдвиг D3…D10 до дырки D11, D14 освобождается
    plan, err = plan_move(fake, 14, 3)
    assert not err and plan["freed"] == 14, err
    assert dict(plan["pairs"]) == {14: 3, **{s: s + 1 for s in range(3, 11)}}
    # симуляция цепочки: ни одна камера не стоит в двух слотах сразу,
    # а камера ставится в пустой слот только там, где он был пуст изначально
    def apply(lay, steps):
        for st in steps:
            if st[0] == "delete":
                lay.pop(st[1], None)
            elif st[0] == "park":
                lay[st[1]] = "P"
            else:
                assert st[2] not in lay.values()
                lay[st[1]] = st[2]
        return lay
    lay = apply({s: s for s in fake}, plan["steps"])
    assert all(lay[n] == o for o, n in plan["pairs"]) and 14 not in lay
    swap, _ = plan_swap(fake, 2, 5)
    lay = apply({s: s for s in fake}, swap["steps"])
    assert lay[2] == 5 and lay[5] == 2
    old_swap = [["delete", 2], ["delete", 5], ["place", 5, 2], ["place", 2, 5]]
    rep, _ = plan_replace(fake, 2, 5)
    lay = apply({s: s for s in fake}, rep["steps"])
    assert lay[5] == 2 and 2 not in lay and 5 not in lay.values()
    # откат с любого шага (обрыв посередине) возвращает всё как было
    for steps, touched in ((plan["steps"], set(range(3, 12)) | {14}),
                           (swap["steps"], {2, 5}), (old_swap, {2, 5}),
                           (rep["steps"], {2, 5})):
        occupied = [x for x in touched if x in fake]
        for k in range(len(steps) + 1):
            lay = apply({s: s for s in fake}, steps[:k])
            lay = apply(lay, rollback_steps(steps[:k], occupied, touched))
            assert lay == {s: s for s in fake}, (k, lay)
    plan, err = plan_move(fake, 2, 9)        # вниз: D3…D9 → D2…D8
    assert not err and dict(plan["pairs"]) == {2: 9, **{s: s - 1 for s in range(3, 10)}}
    tmp = os.path.join(output_dir(), "_selftest_tmp.xlsx")
    job = {"recorder": "test", "kind": "move", "title": "D14 → D3", "created": "-",
           "finished": "-", "pairs": [[14, 3]], "touched": [3, 14],
           "snap": {"chans": {"14": fake[14]}}, "results": {}}
    build_move_excel(tmp, job)
    build_check_excel(tmp, "test", [], {"now": dt.datetime.now(), "hours": 24,
                                        "hdds": [], "chans": fake, "max_slot": 16})
    try:
        os.remove(tmp)
    except Exception:
        pass
    frozen = "да" if getattr(sys, "frozen", False) else "нет"
    print(f"SELFTEST OK | requests {_rq.__version__} | openpyxl {_ox.__version__} "
          f"| colorama {_cl.__version__} | frozen={frozen}")


if __name__ == "__main__":
    if "--selftest" in sys.argv:
        selftest()
        sys.exit(0)
    try:
        cleanup_after_update()
        if "--no-update" not in sys.argv:
            check_update()
        if "--version" in sys.argv:
            print(VERSION)
            sys.exit(0)
        main()
    except KeyboardInterrupt:
        print(Fore.RED + "\nПрервано. Незавершённую операцию можно продолжить "
              "при следующем запуске." + Style.RESET_ALL)
    except Exception as e:
        print(Fore.RED + f"\nНеожиданная ошибка: {e}" + Style.RESET_ALL)
        pause_exit()
    finally:
        try:
            CON.stop()
        except Exception:
            pass
