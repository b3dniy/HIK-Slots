# -*- coding: utf-8 -*-
"""
Эмулятор регистратора Hikvision (ISAPI) для разработки HIK-Slots без железа.

    python dev/mock_nvr.py [порт]      # по умолчанию 18080, логин любой

Что эмулируется (упрощённо, но «строго», как настоящий регистратор):
  • 128 IP-слотов, в них ~120 камер; пустые слоты D7, D80, D116–D128;
  • одна камера (ip+канал) не может стоять в двух слотах — ошибка;
  • PUT в пустой слот запрещён, POST ставит камеру в первый свободный слот
    (так ведёт себя реальная прошивка на 150 каналов);
  • XML с префиксами ns0: отклоняется (badXmlContent);
  • пароль камеры не отдаётся при чтении; камера «на связи», только если
    при записи канала передан её настоящий пароль (у D33, D60, D97 он другой);
  • удаление канала сбрасывает его расписание, триггеры и графики;
  • дорожки записи, триггеры событий, графики охраны, поиск архива,
    параметры потоков, диски, время.
Сбой для теста продолжения: POST /mock/fail?put=N — N-й PUT канала оборвётся.
"""

import sys
import re
import time
import threading
import datetime as dt
import xml.etree.ElementTree as ET
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs

NS = "http://www.hikvision.com/ver20/XMLSchema"
MAX = 128
EMPTY = {7, 80} | set(range(116, MAX + 1))
OTHER_PW = {33, 60, 97}        # у камер из этих слотов «другой» пароль
NO_AUDIO = {12, 40}
OFFLINE = {66}                 # камера выключена ещё до всех операций

LOCK = threading.RLock()
STATE = {}


def _local(tag):
    return tag.split("}")[-1]


def x(tag, *kids, text=None, **attrs):
    e = ET.Element(tag, attrs)
    if text is not None:
        e.text = str(text)
    for k in kids:
        e.append(k)
    return e


def ser(e):
    e.set("version", "2.0")
    e.set("xmlns", NS)
    return ET.tostring(e, encoding="utf-8")


def strip_ns(e):
    for el in e.iter():
        el.tag = _local(el.tag)
    return e


def ft(e, name):
    for el in e.iter():
        if _local(el.tag) == name:
            return (el.text or "").strip()
    return ""


def default_track(s):
    return (f"<Track><id>{s}01</id><Channel>{s}01</Channel><Enable>true</Enable>"
            f"<Description>std</Description><TrackGUID>guid-{s}</TrackGUID>"
            "<DefaultRecordingMode>CMR</DefaultRecordingMode><LoopEnable>true</LoopEnable>"
            f"<SrcDescriptor><SrcGUID>src-{s}</SrcGUID></SrcDescriptor>"
            "<TrackSchedule><ScheduleBlockList><ScheduleBlock><ScheduleBlockGUID>{0}</ScheduleBlockGUID>"
            "<ScheduleBlockType>www.std-cgi.com/racm/schedule/ver10</ScheduleBlockType>"
            + "".join(
                f"<ScheduleAction><id>{i + 1}</id><ScheduleActionStartTime><DayOfWeek>{d}</DayOfWeek>"
                "<TimeOfDay>00:00:00</TimeOfDay></ScheduleActionStartTime><ScheduleActionEndTime>"
                f"<DayOfWeek>{d}</DayOfWeek><TimeOfDay>24:00:00</TimeOfDay></ScheduleActionEndTime>"
                "<Actions><Record>true</Record><ActionRecordingMode>CMR</ActionRecordingMode>"
                "</Actions></ScheduleAction>"
                for i, d in enumerate(("Monday", "Tuesday", "Wednesday", "Thursday",
                                       "Friday", "Saturday", "Sunday")))
            + "</ScheduleBlock></ScheduleBlockList></TrackSchedule>"
            "<CustomExtensionList><CustomExtension><CustomExtensionName>www.hikvision.com</CustomExtensionName>"
            "<enableSchedule>true</enableSchedule><SaveAudio>true</SaveAudio>"
            "<PreRecordTimeSeconds>5</PreRecordTimeSeconds><PostRecordTimeSeconds>5</PostRecordTimeSeconds>"
            "</CustomExtension></CustomExtensionList></Track>")


def default_trigger(s):
    return (f"<EventTrigger><id>VMD-{s}</id><eventType>VMD</eventType>"
            f"<eventDescription>VMD Event trigger Information</eventDescription>"
            f"<videoInputChannelID>{s}</videoInputChannelID>"
            f"<dynVideoInputChannelID>{s}</dynVideoInputChannelID>"
            "<EventTriggerNotificationList><EventTriggerNotification><id>center</id>"
            "<notificationMethod>center</notificationMethod>"
            "<notificationRecurrence>beginning</notificationRecurrence></EventTriggerNotification>"
            f"<EventTriggerNotification><id>record-{s}</id><notificationMethod>record</notificationMethod>"
            f"<notificationRecurrence>beginning</notificationRecurrence><videoInputID>{s}</videoInputID>"
            "</EventTriggerNotification></EventTriggerNotificationList></EventTrigger>")


def default_sched(s):
    return (f"<Schedule><id>VMD_video{s}</id><eventType>VMD</eventType>"
            f"<videoInputChannelID>{s}</videoInputChannelID><TimeBlockList>"
            + "".join(f"<TimeBlock><dayOfWeek>{d}</dayOfWeek><TimeRange><beginTime>00:00:00"
                      "</beginTime><endTime>24:00:00</endTime></TimeRange></TimeBlock>"
                      for d in range(1, 8))
            + "</TimeBlockList></Schedule>")


def reset():
    with LOCK:
        cams, real_pw = {}, {}
        for s in range(1, MAX + 1):
            if s in EMPTY:
                continue
            ip = f"192.168.10.{s}"
            real_pw[ip] = "Other999" if s in OTHER_PW else "Cam12345"
            cams[s] = {"name": f"Камера {s:03d}", "ip": ip, "port": "8000",
                       "src": "1", "user": "admin",
                       "pw": real_pw[ip]}
        cams[115]["name"] = "Склад-3"
        cams[50]["name"] = "Проходная"
        cams[3]["name"] = "IPCamera 03"                  # стандартное имя
        tracks = {s: default_track(s) for s in range(1, MAX + 1)}
        # у Склада-3 особое расписание: по движению
        tracks[115] = tracks[115].replace("<ActionRecordingMode>CMR", "<ActionRecordingMode>MOTION")
        trig = {s: default_trigger(s) for s in range(1, MAX + 1)}
        # у Склада-3 особая связка: ещё и звуковой сигнал
        trig[115] = trig[115].replace(
            "<EventTriggerNotificationList>",
            "<EventTriggerNotificationList><EventTriggerNotification><id>beep</id>"
            "<notificationMethod>beep</notificationMethod>"
            "<notificationRecurrence>beginning</notificationRecurrence></EventTriggerNotification>")
        # D10 при движении пишет ещё и D60 — перекрёстная связка
        trig[10] = trig[10].replace(
            "</EventTriggerNotificationList>",
            "<EventTriggerNotification><id>record-60</id><notificationMethod>record</notificationMethod>"
            "<notificationRecurrence>beginning</notificationRecurrence><videoInputID>60</videoInputID>"
            "</EventTriggerNotification></EventTriggerNotificationList>")
        sched = {s: default_sched(s) for s in range(1, MAX + 1)}
        sched[115] = sched[115].replace("<beginTime>00:00:00", "<beginTime>20:00:00")
        STATE.clear()
        STATE.update(cams=cams, real_pw=real_pw, tracks=tracks, trig=trig,
                     sched=sched, fail_put=0, puts=0, log=[])


def chan_xml(s, c):
    return x("InputProxyChannel",
             x("id", text=s), x("name", text=c["name"]),
             x("sourceInputPortDescriptor",
               x("adminProtocol", text="HIKVISION"),
               x("addressingFormatType", text="ipaddress"),
               x("ipAddress", text=c["ip"]), x("managePortNo", text=c["port"]),
               x("srcInputPort", text=c["src"]), x("userName", text=c["user"]),
               x("streamType", text="auto")),
             x("enableAnr", text="false"), x("devIndex", text=f"DEV-{s}"))


def status(code, sub="ok", http=200):
    st = {1: "OK", 4: "Invalid Operation", 5: "Device Busy"}.get(code, "Invalid XML Content")
    return http, ser(x("ResponseStatus", x("requestURL", text=""), x("statusCode", text=code),
                       x("statusString", text=st), x("subStatusCode", text=sub)))


def offline_cam(c):
    return int(c["ip"].rsplit(".", 1)[1]) in OFFLINE


def online(s):
    c = STATE["cams"].get(s)
    return bool(c) and not offline_cam(c) and c["pw"] == STATE["real_pw"].get(c["ip"])


class H(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def _send(self, code, body):
        self.send_response(code)
        self.send_header("Content-Type", "application/xml; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _body(self):
        n = int(self.headers.get("Content-Length") or 0)
        return self.rfile.read(n) if n else b""

    def do_GET(self):
        with LOCK:
            self._send(*self.route("GET", urlparse(self.path).path, b""))

    def do_PUT(self):
        b = self._body()
        with LOCK:
            r = self.route("PUT", urlparse(self.path).path, b)
        if r is None:                          # имитация обрыва связи
            self.close_connection = True
            self.connection.close()
            return
        self._send(*r)

    def do_POST(self):
        u = urlparse(self.path)
        b = self._body()
        with LOCK:
            if u.path == "/mock/fail":
                STATE["fail_put"] = int(parse_qs(u.query).get("put", ["0"])[0])
                STATE["puts"] = 0
                return self._send(*status(1))
            if u.path == "/mock/busy":       # после записи канала N сек отвечать 503
                STATE["busy"] = float(parse_qs(u.query).get("secs", ["0"])[0])
                return self._send(*status(1))
            if u.path == "/mock/strict":
                STATE["strict"] = parse_qs(u.query).get("on", ["0"])[0] == "1"
                return self._send(*status(1))
            if u.path == "/mock/reset":
                reset()
                return self._send(*status(1))
            self._send(*self.route("POST", u.path, b))

    def do_DELETE(self):
        with LOCK:
            self._send(*self.route("DELETE", urlparse(self.path).path, b""))

    # ------------------------------------------------------------------
    def route(self, m, p, body):
        S = STATE
        cams = S["cams"]
        if body:
            if b"ns0:" in body:
                return status(6, "badXmlContent", 400)
            try:
                req = strip_ns(ET.fromstring(body))
            except ET.ParseError:
                return status(6, "badXmlContent", 400)
        # прямой доступ к камере: /cam/<ip>/ISAPI/...  (в тестах CAM_URL ведёт сюда)
        mt = re.fullmatch(r"/cam/([\d.]+)(/ISAPI/.*)", p)
        if mt:
            ip, cp = mt.group(1), mt.group(2)
            no_center = S.setdefault("cam_no_center", set())
            if cp == "/ISAPI/System/deviceInfo":
                return 200, ser(x("DeviceInfo", x("deviceName", text=f"CAM {ip}")))
            if cp == "/ISAPI/System/reboot" and m == "PUT":
                S.setdefault("rebooted", []).append(ip)
                return status(1)
            if cp == "/ISAPI/Event/triggers/VMD-1":
                if m == "PUT":
                    if any(ft(n, "notificationMethod") == "center"
                           for n in req.iter() if _local(n.tag) == "EventTriggerNotification"):
                        no_center.discard(ip)
                    return status(1)
                notes = [] if ip in no_center else [x("EventTriggerNotification",
                         x("id", text="center"), x("notificationMethod", text="center"),
                         x("notificationRecurrence", text="beginning"))]
                return 200, ser(x("EventTrigger", x("id", text="VMD-1"),
                                  x("eventType", text="VMD"),
                                  x("videoInputChannelID", text=1),
                                  x("EventTriggerNotificationList", *notes)))
            return status(4, "notSupport", 404)
        if m == "GET" and p == "/ISAPI/System/deviceInfo":
            return 200, ser(x("DeviceInfo", x("deviceName", text="MOCK-NVR"),
                              x("model", text="DS-96128NI-MOCK")))
        if m == "GET" and p == "/ISAPI/System/time":
            return 200, ser(x("Time", x("timeMode", text="NTP"), x(
                "localTime", text=dt.datetime.now().strftime("%Y-%m-%dT%H:%M:%S+03:00"))))
        if m == "GET" and p == "/ISAPI/ContentMgmt/Storage":
            return 200, ser(x("storage", x("hddList", x(
                "hdd", x("id", text=1), x("capacity", text=8000000),
                x("freeSpace", text=0), x("status", text="ok")))))
        base = "/ISAPI/ContentMgmt/InputProxy/channels"
        if p == base and m == "GET":
            return 200, ser(x("InputProxyChannelList",
                              *[chan_xml(s, c) for s, c in sorted(cams.items())]))
        if p == base + "/status":
            lst = []
            for s, c in sorted(cams.items()):
                lst.append(x("InputProxyChannelStatus", x("id", text=s),
                             x("online", text=str(online(s)).lower()),
                             x("chanDetectResult", text="connect" if online(s) else (
                                 "netUnreachable" if offline_cam(c) else "errorUserNameOrPasswd"))))
            return 200, ser(x("InputProxyChannelStatusList", *lst))
        if p == base + "/capabilities":
            return 200, ser(x("InputProxyChannel", x("id", min="1", max=str(MAX))))
        if p == base and m == "POST":            # «как прошивка хочет» — первый свободный
            free = next(s for s in range(1, MAX + 1) if s not in cams)
            r = self._put_chan(free, req)
            if free in cams:                 # прошивка ставит своё имя
                cams[free]["name"] = f"IPCamera {free}"
            return r
        mt = re.fullmatch(base + r"/(\d+)", p)
        if mt:
            s = int(mt.group(1))
            if m == "GET":
                return (200, ser(chan_xml(s, cams[s]))) if s in cams else status(4, "invalidID", 404)
            if m == "DELETE":
                if s not in cams:
                    return status(4, "invalidID", 404)
                del cams[s]
                S["tracks"][s] = default_track(s)
                S["trig"][s] = default_trigger(s)
                S["sched"][s] = default_sched(s)
                S["log"].append(("delete", s))
                return status(1)
            if m == "PUT":
                S["puts"] += 1
                if S["fail_put"] and S["puts"] == S["fail_put"]:
                    S["fail_put"] = 0
                    return None
                if s not in cams:            # как настоящая прошивка: PUT только меняет
                    return status(4, "invalidOperation", 403)
                return self._put_chan(s, req)
        if p == "/ISAPI/ContentMgmt/record/tracks" and m == "GET":
            return 200, ser(x("TrackList", *[ET.fromstring(t) for _, t in sorted(S["tracks"].items())]))
        mt = re.fullmatch(r"/ISAPI/ContentMgmt/record/tracks/(\d+)", p)
        if mt:
            s = int(mt.group(1)) // 100
            if m == "GET":
                return 200, ser(ET.fromstring(S["tracks"][s]))
            if ft(req, "id") != mt.group(1):
                return status(4, "invalidID", 400)
            S["tracks"][s] = ET.tostring(req, encoding="unicode")
            return status(1)
        if p == "/ISAPI/Event/triggers" and m == "GET":
            return 200, ser(x("EventTriggerList", *[ET.fromstring(t) for _, t in sorted(S["trig"].items())]))
        mt = re.fullmatch(r"/ISAPI/Event/triggers/VMD-(\d+)", p)
        if mt:
            s = int(mt.group(1))
            if m == "GET":
                return 200, ser(ET.fromstring(S["trig"][s]))
            c = cams.get(s)
            bad_pw = c and not offline_cam(c) and not online(s)
            if bad_pw or time.time() < S.get("busy_until", 0):   # занят / не войти в камеру
                return status(5, "serviceUnavailable", 503)
            if ft(req, "id") != f"VMD-{s}" or ft(req, "videoInputChannelID") != str(s):
                return status(4, "invalidID", 400)
            S["trig"][s] = ET.tostring(req, encoding="unicode")
            return status(1)
        # как реальная прошивка: id «VMD_video1»; на «VMD-1» — 403, а на
        # странные id с цифрой в конце («regionEntrance-1») отдаёт VMD_video1
        mt = re.fullmatch(r"/ISAPI/Event/schedules/motionDetections/(.*?)(\d+)", p)
        if mt:
            s = int(mt.group(2))
            if mt.group(1) in ("VMD-", "VMD_"):
                return status(4, "invalidOperation", 403)
            if m == "GET":
                return 200, ser(ET.fromstring(S["sched"][s]))
            if mt.group(1) != "VMD_video" or ft(req, "id") != f"VMD_video{s}":
                return status(4, "invalidID", 400)
            S["sched"][s] = ET.tostring(req, encoding="unicode")
            return status(1)
        mt = re.fullmatch(r"/ISAPI/System/Video/inputs/channels/(\d+)/motionDetection", p)
        if mt:
            s = int(mt.group(1))
            mdc = S.setdefault("md", {}).setdefault(s, {
                "enabled": "true", "sens": "80",
                "grid": S.get("grid", {}).get(s, "fffffc" * 15)})
            if m == "PUT":
                mdc.update(enabled=ft(req, "enabled"), sens=ft(req, "sensitivityLevel"),
                           grid=ft(req, "gridMap"))
                S.setdefault("grid", {})[s] = mdc["grid"]
                return status(1)
            return 200, ser(x("MotionDetection", x("enabled", text=mdc["enabled"]),
                              x("Grid", x("rowGranularity", text=15),
                                x("columnGranularity", text=22)),
                              x("MotionDetectionLayout", x("sensitivityLevel", text=mdc["sens"]),
                                x("layout", x("gridMap", text=mdc["grid"])))))
        mt = re.fullmatch(r"/ISAPI/ContentMgmt/StreamingProxy/channels/(\d+)01", p)
        if mt:
            s = int(mt.group(1))
            if s not in cams or not online(s):
                return status(4, "deviceOffline", 403)
            no_audio = S.setdefault("no_audio", set(NO_AUDIO))
            if m == "PUT":
                if ft(req, "enabled") == "true":
                    no_audio.discard(s)
                return status(1)
            kids = [x("id", text=f"{s}01"), x("Video",
                      x("videoCodecType", text="H.265"), x("videoResolutionWidth", text=2560),
                      x("videoResolutionHeight", text=1440), x("videoQualityControlType", text="VBR"),
                      x("vbrUpperCap", text=4096), x("maxFrameRate", text=2000))]
            if s != 40:
                kids.append(x("Audio", x("enabled", text=str(s not in no_audio).lower()),
                              x("audioCompressionType", text="G.711ulaw")))
            return 200, ser(x("StreamingChannel", *kids))
        if p == "/ISAPI/ContentMgmt/search" and m == "POST":
            s = int(ft(req, "trackID")) // 100
            now = dt.datetime.now()
            items = []
            if s in cams and online(s):
                mode = "MOTION" if "MOTION" in S["tracks"][s] else "CMR"
                if mode == "CMR":
                    a = now - dt.timedelta(hours=30)
                    gap = s == 21                 # у D21 «дыра» в архиве
                    spans = [(a, now - dt.timedelta(hours=6)), (now - dt.timedelta(hours=4), now)] \
                        if gap else [(a, now)]
                else:
                    spans = [(now - dt.timedelta(hours=3), now - dt.timedelta(hours=2, minutes=50))]
                for a, b in spans:
                    items.append(x("searchMatchItem", x("trackID", text=f"{s}01"), x(
                        "timeSpan", x("startTime", text=a.strftime("%Y-%m-%dT%H:%M:%SZ")),
                        x("endTime", text=b.strftime("%Y-%m-%dT%H:%M:%SZ")))))
            return 200, ser(x("CMSearchResult", x("responseStatus", text="true"),
                              x("responseStatusStrg", text="OK" if items else "NO MATCHES"),
                              x("numOfMatches", text=len(items)), x("matchList", *items)))
        return status(4, "notSupport", 404)

    def _put_chan(self, s, req):
        S = STATE
        cams = S["cams"]
        if s < 1 or s > MAX:
            return status(4, "invalidID", 400)
        if S.get("strict") and any(_local(e.tag) == "enableAnr" for e in req.iter()):
            return status(6, "badXmlContent", 400)    # «капризная» прошивка
        ip, src = ft(req, "ipAddress"), ft(req, "srcInputPort") or "1"
        for o, c in cams.items():
            if o != s and c["ip"] == ip and c["src"] == src:
                return status(4, "IPCameraExists", 403)
        pw = ft(req, "password")
        old = cams.get(s)
        if not pw and old and old["ip"] == ip:
            pw = old["pw"]
        cams[s] = {"name": ft(req, "name"), "ip": ip, "port": ft(req, "managePortNo"),
                   "src": src, "user": ft(req, "userName"), "pw": pw}
        S["log"].append(("place", s, ip))
        S["busy_until"] = time.time() + S.get("busy", 0)
        return status(1)


def serve(port=18080):
    reset()
    srv = ThreadingHTTPServer(("127.0.0.1", port), H)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    return srv


if __name__ == "__main__":
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 18080
    reset()
    print(f"MOCK NVR: http://127.0.0.1:{port}  (Ctrl+C — стоп)")
    ThreadingHTTPServer(("127.0.0.1", port), H).serve_forever()
