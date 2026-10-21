# -*- coding: utf-8 -*-
"""
Сквозной тест HIK-Slots на эмуляторе (dev/mock_nvr.py) — без настоящего регистратора.

    python dev/test_against_mock.py

Проверяет: перенос со сдвигом до дырки, перенос имени/расписания/связок/графиков,
перекрёстные связки, пароли «не как у всех», откат до исходного состояния,
обмен местами, обрыв связи посередине и продолжение по журналу, проверку камер.
"""

import os
import re
import sys
import copy
import shutil
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.dirname(HERE))

import mock_nvr                       # noqa: E402
import nvr_slots as ns                # noqa: E402

PORT = 18081
TMP = tempfile.mkdtemp(prefix="hikslots_test_")
ns.output_dir = lambda: TMP           # журналы и Excel — во временную папку
ns.VERIFY_MIN = 1
ns.VERIFY_WAIT = 5
ns.open_file = lambda p: None
ns.press_any = lambda: None
ns.BUSY_PAUSE = 0.5                   # быстрые повторы на «503» в тестах
ns.CAM_URL = f"http://127.0.0.1:{PORT}/cam/{{ip}}{{path}}"   # «камеры» — в эмуляторе
ns.getch_choice = lambda valid: "\x1b"      # на любые вопросы — «позже»

# пароль «по запросу» для конкретной камеры — её настоящий пароль из эмулятора
def _pw(prompt=""):
    """Вопрос «Пароль камеры <имя>:» → настоящий пароль этой камеры в эмуляторе."""
    cams = [c for c in mock_nvr.STATE["cams"].values() if c["name"] and c["name"] in prompt]
    if not cams:
        return ""
    cam = max(cams, key=lambda c: len(c["name"]))
    return mock_nvr.STATE["real_pw"].get(cam["ip"], "")


ns.read_password = _pw

FAILS = []


def check(cond, msg):
    print(("  OK   " if cond else "  FAIL ") + msg)
    if not cond:
        FAILS.append(msg)


def state():
    with mock_nvr.LOCK:
        return copy.deepcopy({k: mock_nvr.STATE[k] for k in ("cams", "tracks", "trig", "sched")})


def connect():
    s, err = ns.make_session(f"http://127.0.0.1:{PORT}", "admin", "Cam12345")
    assert s, err
    nvr = ns.Nvr(s, f"http://127.0.0.1:{PORT}")
    return ns.Ctx(nvr, {"name": "MOCK"}, "локальная сеть", "Cam12345")


def make_job(ctx, plan, kind="move"):
    chans = ns.read_channels(ctx.nvr)
    status = ns.read_status(ctx.nvr)
    touched = {s for p in plan["pairs"] for s in p}
    snap, off = ns.take_snapshot(ctx.nvr, chans, status, touched)
    job = ns.new_job("MOCK", kind, plan["title"], plan["steps"], snap, off,
                     touched, plan["pairs"])
    ns.save_job(job)
    return job


def main():
    mock_nvr.serve(PORT)
    ctx = connect()
    ctx.pw_common = "Cam12345"
    before = state()

    print("\n=== 1. Перенос D115 → D50 (дырка в D80) ===")
    chans = ns.read_channels(ctx.nvr)
    plan, err = ns.plan_move(chans, 115, 50)
    check(err is None, "план построен")
    check(plan["freed"] == 115, "D115 освобождается (сдвиг до дырки D80)")
    check(len(plan["pairs"]) == 31, f"переезжает 31 камера (D50…D79 + Склад-3): {len(plan['pairs'])}")
    job = make_job(ctx, plan)
    ok = ns.execute_job(ctx, job)
    after = state()
    cams = after["cams"]
    check(cams[50]["ip"] == before["cams"][115]["ip"] and cams[50]["name"] == "Склад-3",
          "в D50 стоит Склад-3 с именем")
    check(all(cams[s + 1]["ip"] == before["cams"][s]["ip"] and
              cams[s + 1]["name"] == before["cams"][s]["name"] for s in range(50, 80)),
          "D50…D79 сдвинулись в D51…D80 с именами")
    check(115 not in cams, "D115 пуст")
    check(all(cams[s] == before["cams"][s] for s in range(81, 115) if s in before["cams"]),
          "D81…D114 не тронуты")
    check("MOTION" in after["tracks"][50], "расписание Склада-3 (по движению) переехало в D50")
    check("MOTION" not in after["tracks"][51], "в D51 расписание Проходной (постоянное)")
    check("beep" in after["trig"][50] and "<id>VMD-50</id>" in after["trig"][50],
          "связка Склада-3 (beep) переехала в VMD-50")
    check("<videoInputID>61</videoInputID>" in after["trig"][10]
          and "record-61" in after["trig"][10],
          "перекрёстная связка D10 → «запись D60» стала D61")
    check("20:00:00" in after["sched"][50], "график охраны Склада-3 переехал в D50")
    check(cams[61]["pw"] == "Other999" and mock_nvr.online(61),
          "камера с другим паролем (была D60) на связи после ввода её пароля")
    check(all(mock_nvr.online(s) for s in range(50, 81) if s != 67),
          "все перенесённые камеры на связи (кроме D67 — она была офлайн и до)")
    check(job["state"] == "done", f"журнал: done (итог {ok})")
    check(67 in job["results"].get("skipped", []) and ns.settings_fails(job) == 0,
          "офлайн-камера (была D66) — связки честно пропущены, без ошибок и ожидания")
    check(os.path.exists(os.path.join(TMP, job["file"].replace(".json", ".xlsx"))),
          "карта переезда в Excel сохранена")

    print("\n=== 2. Откат последней операции ===")
    ok = ns.rollback_job(ctx, job, ask_pw=False)
    after = state()
    for k in ("cams", "tracks", "trig", "sched"):
        diff = [s for s in set(before[k]) | set(after[k])
                if _norm(before[k].get(s)) != _norm(after[k].get(s))]
        check(not diff, f"{k}: всё как до операции" + (f" — отличия: {sorted(diff)[:8]}" if diff else ""))

    print("\n=== 3. Поменять местами D10 ⇄ D20 ===")
    chans = ns.read_channels(ctx.nvr)
    plan, err = ns.plan_swap(chans, 10, 20)
    job = make_job(ctx, plan, "swap")
    ns.execute_job(ctx, job)
    after = state()
    check(after["cams"][10]["ip"] == before["cams"][20]["ip"]
          and after["cams"][20]["ip"] == before["cams"][10]["ip"], "камеры поменялись")
    check("<videoInputID>60</videoInputID>" in after["trig"][20],
          "связка D10 (запись D60) уехала вместе с камерой в D20")
    ns.rollback_job(ctx, job, ask_pw=False)
    check(state()["cams"] == before["cams"], "откат обмена — всё на месте")

    print("\n=== 4. Обрыв связи посередине и продолжение по журналу ===")
    import requests
    requests.post(f"http://127.0.0.1:{PORT}/mock/fail?put=3")
    chans = ns.read_channels(ctx.nvr)
    plan, err = ns.plan_move(chans, 30, 25)
    job = make_job(ctx, plan)
    ns.fail_menu = lambda j: "stop"
    ok = ns.execute_job(ctx, job)
    check(not ok and job["state"] == "chain" and job["done"] < len(job["steps"]),
          f"операция остановилась на шаге {job['done'] + 1} из {len(job['steps'])}, журнал сохранён")
    saved = [j for j in ns.load_jobs("MOCK") if j["state"] == "chain"]
    check(len(saved) == 1, "незавершённая операция находится при следующем запуске")
    ok = ns.execute_job(ctx, saved[0])
    after = state()
    check(after["cams"][25]["ip"] == before["cams"][30]["ip"]
          and all(after["cams"][s + 1]["ip"] == before["cams"][s]["ip"] for s in range(25, 30)),
          "после продолжения D30 стоит в D25, D25…D29 → D26…D30")
    ns.rollback_job(ctx, ns.load_jobs("MOCK")[0], ask_pw=False)
    check(state()["cams"] == before["cams"], "откат — всё на месте")

    print("\n=== 4б. Реальный случай: старый журнал, обе камеры удалены, PUT в пустой слот запрещён ===")
    a, b = 109, 110
    chans = ns.read_channels(ctx.nvr)
    snap, off = ns.take_snapshot(ctx.nvr, chans, ns.read_status(ctx.nvr), {a, b})
    job = ns.new_job("MOCK", "swap", f"D{a} ⇄ D{b}",
                     [["delete", a], ["delete", b], ["place", b, a], ["place", a, b]],
                     snap, off, {a, b}, [(a, b), (b, a)])
    ns.do_delete(ctx.nvr, a)
    ns.do_delete(ctx.nvr, b)
    job["done"] = 2
    ns.save_job(job)
    check(a not in state()["cams"] and b not in state()["cams"],
          f"D{a}, D{b} пусты (как у пользователя)")
    # и прошивка капризничает: badXmlContent на необязательные поля
    requests.post(f"http://127.0.0.1:{PORT}/mock/strict?on=1")
    ns.rollback_job(ctx, job, ask_pw=False)
    requests.post(f"http://127.0.0.1:{PORT}/mock/strict?on=0")
    after = state()
    check(after["cams"].get(a, {}).get("ip") == before["cams"][a]["ip"]
          and after["cams"].get(b, {}).get("ip") == before["cams"][b]["ip"],
          "откат вернул обе камеры в свои пустые слоты (через заглушки)")
    check(not any(c["ip"].startswith(("198.51.100.", "192.0.2.")) for c in after["cams"].values()),
          "временных заглушек не осталось")
    check(after["cams"] == before["cams"], "все каналы как до операции")
    check(after["tracks"][a] == before["tracks"][a] and after["trig"][a] == before["trig"][a],
          "расписание и события D109 восстановлены")

    print("\n=== 4г. Камеры удалены, потом добавлены вручную (встали в D7 и D80) → доделать откат ===")
    a, b = 109, 110
    chans = ns.read_channels(ctx.nvr)
    snap, off = ns.take_snapshot(ctx.nvr, chans, ns.read_status(ctx.nvr), {a, b})
    job = ns.new_job("MOCK", "swap", f"D{a} ⇄ D{b}",
                     [["delete", a], ["delete", b], ["place", b, a], ["place", a, b]],
                     snap, off, {a, b}, [(a, b), (b, a)])
    ns.do_delete(ctx.nvr, a)
    ns.do_delete(ctx.nvr, b)
    job["done"] = 2
    ns.save_job(job)
    for s in (a, b):                      # «вручную через веб» — в первый свободный
        ctx.nvr.post(ns.CH_PATH, ns.channel_body(chans[s], s, "Cam12345"))
    st = state()["cams"]
    check(st.get(7, {}).get("ip") == before["cams"][a]["ip"]
          and st.get(80, {}).get("ip") == before["cams"][b]["ip"],
          "вручную добавленные камеры встали в D7 и D80")
    ns.rollback_job(ctx, job, ask_pw=False)
    after = state()
    check(after["cams"] == before["cams"],
          "доделанный откат вернул камеры в D109/D110, D7 и D80 снова свободны, дублей нет")
    check(_norm(after["tracks"][a]) == _norm(before["tracks"][a]),
          "расписание D109 восстановлено из снимка")

    print("\n=== 4д. Регистратор «занят» (503) несколько секунд после смены каналов ===")
    requests.post(f"http://127.0.0.1:{PORT}/mock/busy?secs=3")
    chans = ns.read_channels(ctx.nvr)
    plan, err = ns.plan_move(chans, 115, 116)
    job = make_job(ctx, plan)
    ns.execute_job(ctx, job)
    after = state()
    check(ns.settings_fails(job) == 0 and "beep" in after["trig"][116],
          "связки событий перенеслись, несмотря на 503 (повтор с паузой)")
    ns.rollback_job(ctx, job, ask_pw=False)
    requests.post(f"http://127.0.0.1:{PORT}/mock/busy?secs=0")
    check(state()["cams"] == before["cams"], "откат — всё на месте")

    print("\n=== 4е. Как у пользователя: общий пароль с ошибкой, перенос в пустой слот ===")
    ctx2 = connect()
    ctx2.pw_common = "WRONG-pass"
    chans = ns.read_channels(ctx2.nvr)
    plan, err = ns.plan_move(chans, 115, 116)
    job = make_job(ctx2, plan)
    ns.execute_job(ctx2, job)
    after = state()
    check(mock_nvr.online(116), "пароль камеры спрошен отдельно и исправлен — на связи")
    check(after["cams"][116]["name"] == "Склад-3",
          f"имя выставлено, а не «IPCamera 116»: {after['cams'][116]['name']}")
    check(ns.settings_fails(job) == 0 and "beep" in after["trig"][116],
          "связки перенеслись после исправления пароля")
    ns.rollback_job(ctx2, job, ask_pw=False)
    check(state()["cams"] == before["cams"], "откат — всё на месте")

    print("\n=== 4в. Одна камера: на свободный слот и с заменой ===")
    chans = ns.read_channels(ctx.nvr)
    plan, err = ns.plan_move(chans, 115, 116)
    check(err is None and len(plan["pairs"]) == 1 and plan["freed"] == 115,
          "D115 → свободный D116: без сдвига, D115 освобождается")
    job = make_job(ctx, plan)
    ns.execute_job(ctx, job)
    after = state()
    check(after["cams"].get(116, {}).get("name") == "Склад-3" and 115 not in after["cams"]
          and "MOTION" in after["tracks"][116],
          "Склад-3 в D116 с именем и расписанием, D115 пуст")
    ns.rollback_job(ctx, job, ask_pw=False)
    check(state()["cams"] == before["cams"], "откат — всё на месте")

    chans = ns.read_channels(ctx.nvr)
    plan, err = ns.plan_replace(chans, 115, 48)
    check(err is None and plan["replaced"] == 48, "D115 → занятый D48 с заменой")
    job = make_job(ctx, plan, "replace")
    job["replaced"] = {"slot": 48, "name": chans[48]["name"], "ip": chans[48]["ip"]}
    ns.execute_job(ctx, job)
    after = state()
    check(after["cams"][48]["ip"] == before["cams"][115]["ip"]
          and after["cams"][48]["name"] == "Склад-3" and 115 not in after["cams"]
          and not any(c["ip"] == before["cams"][48]["ip"] for c in after["cams"].values()),
          "Склад-3 в D48, старая камера D48 удалена, D115 пуст")
    check("beep" in after["trig"][48] and "MOTION" in after["tracks"][48],
          "связки и расписание Склада-3 переехали в D48")
    ns.rollback_job(ctx, job, ask_pw=False)
    after = state()
    check(after["cams"] == before["cams"], "откат замены — обе камеры на своих местах")
    check(_norm(after["tracks"][48]) == _norm(before["tracks"][48]),
          "откат замены — расписание D48 восстановлено")

    print("\n=== 4ж. Остались заглушки после обрыва — уборка при входе ===")
    chans = ns.read_channels(ctx.nvr)
    ctx.nvr.post(ns.CH_PATH, ns.placeholder_body(chans[1]["xml"], 7, "198.51.100.9",
                                                 "HIK-Slots fill 7", "x"))
    ns.do_park(ctx.nvr, 5, 3, "x")
    ns.confirm_yes = lambda prompt="": True
    ns.cleanup_placeholders(ctx)
    cams = state()["cams"]
    check(7 not in cams and 5 not in cams, "заглушки D7 (fill) и D5 (park) удалены")
    with mock_nvr.LOCK:                       # вернуть D5 как было
        mock_nvr.STATE["cams"][5] = copy.deepcopy(before["cams"][5])

    print("\n=== 5. Ошибки планировщика ===")
    chans = ns.read_channels(ctx.nvr)
    check(ns.plan_move(chans, 116, 5)[1] is not None, "пустой источник — отказ")
    check(ns.plan_move(chans, 5, 5)[1] is not None, "тот же слот — отказ")
    p, _ = ns.plan_move(chans, 79, 81)
    check(p and len(p["pairs"]) == 2 and p["freed"] == 79,
          "D79 → D81 вниз: сдвиг D81 → D80 (дырка), D79 освобождается")
    p, _ = ns.plan_move(chans, 5, 7)
    check(p and len(p["pairs"]) == 1, "перенос в пустой слот без сдвига")

    print("\n=== 6. Проверка камер ===")
    # D30: запись «по движению», но в связке движения нет записи канала
    with mock_nvr.LOCK:
        S = mock_nvr.STATE
        S["tracks"][30] = S["tracks"][30].replace("<ActionRecordingMode>CMR",
                                                  "<ActionRecordingMode>MOTION")
        S["trig"][30] = re.sub(r"<EventTriggerNotification><id>record-30</id>.*?"
                               r"</EventTriggerNotification>", "", S["trig"][30])
        # D31: по движению, но зона детекции пустая; D32: по движению, график охраны пуст
        for s in (31, 32):
            S["tracks"][s] = S["tracks"][s].replace("<ActionRecordingMode>CMR",
                                                    "<ActionRecordingMode>MOTION")
        S.setdefault("grid", {})[31] = "0" * 90
        S["sched"][32] = re.sub(r"<TimeBlockList>.*</TimeBlockList>",
                                "<TimeBlockList />", S["sched"][32])
    rows, meta = ns.run_check(ctx.nvr)
    by = {r["slot"]: r for r in rows}
    check(any("связке движения нет записи" in n for n in by[30]["notes"]),
          f"D30 — движение не включает запись: {by[30]['notes']}")
    check(any("зона не нарисована" in n for n in by[31]["notes"]) and by[31]["level"] == 2,
          f"D31 — зона детекции пустая: {by[31]['notes']}")
    check(any("график охраны движения пуст" in n for n in by[32]["notes"]),
          f"D32 — пустой график охраны: {by[32]['notes']}")
    check(by[115]["level"] == 0, f"Склад-3 (по движению, всё настроено) — ок: {by[115]['notes']}")
    check(ns.archive_text(by[115]).endswith("зап.") and ns.archive_text(by[1]).endswith("%"),
          "архив: по движению — число записей, постоянная — %")
    check(len(rows) == len(before["cams"]), f"проверено камер: {len(rows)}")
    check(by[1]["level"] == 0, f"обычная камера — ок {by[1]['notes']}")
    check(by[66]["level"] == 2 and not by[66]["online"], "D66 — нет связи (ошибка)")
    check(any("звук выключен" in n for n in by[12]["notes"]), "D12 — звук выключен")
    check(any("пропуски" in n for n in by[21]["notes"]), f"D21 — пропуски в архиве: {by[21]['notes']}")
    check(any("стандартное" in n for n in by[3]["notes"]), "D3 — стандартное имя")
    check(by[115]["modes"] == "по движению" and by[115]["level"] == 0,
          "Склад-3 — запись по движению, без ложной тревоги")
    path = os.path.join(TMP, "check.xlsx")
    ns.build_check_excel(path, "MOCK", rows, meta)
    check(os.path.exists(path), "Excel проверки сохранён")

    print("\n=== 7. Исправить камеру: D45 сломана во всём ===")
    with mock_nvr.LOCK:
        S = mock_nvr.STATE
        c45 = S["cams"][45]
        c45["name"], c45["pw"] = "IPCamera 45", "bad"          # заводское имя, неверный пароль
        S["tracks"][45] = S["tracks"][45].replace("<ActionRecordingMode>CMR",
                                                  "<ActionRecordingMode>MOTION")
        S["trig"][45] = re.sub(r"<EventTriggerNotification><id>record-45</id>.*?"
                               r"</EventTriggerNotification>", "", S["trig"][45])
        S["sched"][45] = re.sub(r"<TimeBlockList>.*</TimeBlockList>",
                                "<TimeBlockList />", S["sched"][45])
        S.setdefault("md", {})[45] = {"enabled": "false", "sens": "0", "grid": "0" * 90}
        S.setdefault("no_audio", set(mock_nvr.NO_AUDIO)).add(45)
    answers = iter(["D45", "A", "Склад-45", "A", ""])
    orig_rl = ns.read_line_visible
    ns.read_line_visible = lambda p: next(answers, "")
    ctx3 = connect()
    ns.ui_fix(ctx3)
    ns.read_line_visible = orig_rl
    st = state()
    a = ns.audit_camera(ctx3, 45)
    check(mock_nvr.online(45), "пароль исправлен — камера на связи")
    check(st["cams"][45]["name"] == "Склад-45", "имя задано")
    check("record-45" in st["trig"][45], "в связку движения добавлена запись канала")
    check(st["sched"][45].count("<TimeBlock>") == 7, "график охраны движения — круглосуточно")
    md = mock_nvr.STATE["md"][45]
    check(md["enabled"] == "true" and md["sens"] == "50" and md["grid"] == "fffffc" * 15,
          f"детекция: включена, чувствительность 50, зона — весь кадр ({md})")
    check(45 not in mock_nvr.STATE["no_audio"], "звук включён")
    check(not [i for i in a["issues"] if i["level"] == 2], f"ошибок не осталось: {a['issues']}")
    fixfiles = [f for f in os.listdir(os.path.join(TMP, "HIK-Slots_backup")) if "_fix_D45" in f]
    check(bool(fixfiles), "исходные настройки сохранены в резервную копию")

    print("\n=== 8. Как D92: всё настроено, но камера не сообщает регистратору о движении ===")
    with mock_nvr.LOCK:
        S = mock_nvr.STATE
        S["tracks"][46] = S["tracks"][46].replace("<ActionRecordingMode>CMR",
                                                  "<ActionRecordingMode>MOTION")
        S.setdefault("cam_no_center", set()).add("192.168.10.46")
    ctx4 = connect()
    ctx4.pw_common = "Cam12345"
    a = ns.audit_camera(ctx4, 46)
    keys = [i["key"] for i in a["issues"]]
    check("cam_center" in keys, f"аудит нашёл: камера не уведомляет регистратор ({keys})")
    answers = iter(["D46", "A", ""])
    ns.read_line_visible = lambda p: next(answers, "")
    ns.ui_fix(ctx4)
    ns.read_line_visible = orig_rl
    check("192.168.10.46" not in mock_nvr.STATE["cam_no_center"],
          "в камере включено «Уведомить центр наблюдения»")
    a = ns.audit_camera(ctx4, 46)
    check("cam_center" not in [i["key"] for i in a["issues"]], "повторный аудит — проблемы нет")

    print("\n=== 9. Список проблемных камер и перезагрузка ===")
    rows, meta = ns.problem_cameras(ctx4)
    bad = {x["slot"] for x in rows if x["level"]}
    check(66 in bad and 1 not in bad, f"в списке проблемные (D66), без исправных (D1): {sorted(bad)}")
    answers = iter(["D46", "1", ""])           # у D46 исправлять нечего — 1 = перезагрузить
    ns.read_line_visible = lambda p: next(answers, "")
    ns.ui_fix(ctx4)
    ns.read_line_visible = orig_rl
    check("192.168.10.46" in mock_nvr.STATE.get("rebooted", []), "камера перезагружена")

    print("\nФайлы теста:", TMP)
    print("\nИТОГ:", "ВСЕ ПРОВЕРКИ ПРОЙДЕНЫ" if not FAILS else f"ПРОВАЛЕНО: {len(FAILS)}")
    return 0 if not FAILS else 1


def _norm(xml):
    if xml is None:
        return None
    if isinstance(xml, dict):
        return xml
    import xml.etree.ElementTree as ET
    return ns.canon(mock_nvr.strip_ns(ET.fromstring(xml)))


if __name__ == "__main__":
    code = 1
    try:
        code = main()
    finally:
        try:
            ns.CON.stop()
        except Exception:
            pass
        if "--keep" not in sys.argv:
            shutil.rmtree(TMP, ignore_errors=True)
    sys.exit(code)
