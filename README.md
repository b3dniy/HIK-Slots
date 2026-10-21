<a id="top"></a>

<p align="center">
  <img src="hik.png" width="300" alt="HIK-Slots">
</p>

<h1 align="center">HIK-Slots</h1>

<p align="center">
  <b>Move cameras between Hikvision NVR slots, audit every camera and fix what is wrong — from the terminal.</b><br>
  <sub>Перестановка камер по слотам регистратора Hikvision, проверка и исправление камер — прямо из терминала.</sub>
</p>

<p align="center">
  <img src="https://img.shields.io/badge/platform-Windows%2010%20%2F%2011-0078D6?style=flat-square&logo=windows&logoColor=white">
  <img src="https://img.shields.io/badge/python-3.x-3776AB?style=flat-square&logo=python&logoColor=white">
  <img src="https://img.shields.io/badge/single-.exe%20(no%20install)-2ea44f?style=flat-square">
  <img src="https://img.shields.io/badge/API-Hikvision%20ISAPI-FFC20E?style=flat-square">
  <img src="https://img.shields.io/badge/license-MIT-green?style=flat-square">
</p>

<p align="center">
  <a href="#english"><img src="https://img.shields.io/badge/🇬🇧%20Read%20in-English-2ea44f?style=for-the-badge"></a>
  &nbsp;
  <a href="#russian"><img src="https://img.shields.io/badge/🇷🇺%20Читать%20на-Русском-1f6feb?style=for-the-badge"></a>
</p>

---

<a id="english"></a>

## 🇬🇧 English

When a camera is added to a Hikvision NVR it lands in the **first free slot**, not where you want it.
Moving it by hand means shifting every camera in between, re-typing names and re-checking recording
on each one. **HIK-Slots** does it in one step over **ISAPI**: the camera moves to the slot you pick,
and its **name, recording schedule, event linkage and arming schedules move with it**.

It also **audits every camera** (online, recording, motion detection, audio) and can **fix a single
camera** — wrong password, motion recording that never fires, empty schedules, a camera that does not
notify the NVR about motion — or reboot it.

Same style as [HIK-Report](https://github.com/b3dniy/HIK-Report): a single `.exe`, rainbow progress
bar, blinking SUCCESS / ERROR badges, recorders stored inside the `.exe`, reports in Excel.

### 🧭 Menu

| | |
|---|---|
| **Cameras** | |
| 1. Camera list | all slots, online status, free slots |
| 2. Check all cameras | online, recording, archive, motion detection, audio, disks → Excel |
| 3. Check & fix one camera | only problem cameras are listed → detailed audit → pick what to fix, or reboot |
| **Moving** | |
| 4. Move a camera to another slot | to a free slot, or onto an occupied one — **replace** or **swap** |
| 5. Move with shift | `D115 → D50`: cameras in between shift by one, up to the first free slot |
| 6. Swap two cameras | only two channels are touched |
| 7. Undo a move | pick one of the recent moves and restore everything as it was |

### ✨ Moving cameras

- 🏷️ **Everything follows the camera** — name, recording schedule, event linkage (including
  cross-links like “motion on D10 records D60”), arming schedules.
- 🔗 **No channel is ever deleted mid-way.** The NVR refuses the same camera in two slots, and many
  firmwares can only add a camera to the first free slot. HIK-Slots parks a placeholder (non-existent
  IP) in the slot instead, writes cameras by *editing* channels, and fills an empty slot by briefly
  occupying the free slots below it. Each camera is offline only for 10–30 s.
- 🔑 **Camera passwords are checked on the camera itself before anything changes** (one attempt, so
  the account is not locked). Cameras with a different password are asked for individually.
- 👀 **The plan is shown first** — nothing changes until you confirm with `Y` / `Enter`.
- 💾 **Snapshot and step journal** before every change: an interrupted move can be resumed or undone
  on the next start; leftover placeholders are cleaned up.
- 🗺️ **Relocation map** in Excel — where each camera is now and **in which old slot its earlier
  archive is** (archive stays with the slot, that is how NVRs work).

### 🩺 Checking & fixing cameras

**Check all cameras** — online status, recording schedule, whether the archive actually grows
(coverage for continuous recording, number of recordings for motion), motion detection on / zone /
sensitivity, “motion recording but motion linkage has no record”, arming schedule, audio, codec,
default names, duplicate cameras, disks. Cameras switched off on purpose can be marked once and are no
longer counted as errors.

**Check & fix one camera** — detailed audit, including what can only be seen **on the camera itself**
(does it notify the NVR about motion) and when the camera last recorded (up to 30 days back).
Fixes on offer:

| Problem | Fix |
|---|---|
| Wrong camera password | enter the password (checked on the camera) |
| Factory name (`IPCamera 22`) | set a name |
| Recording schedule off / empty / partial | copy from another camera |
| Motion linkage without recording | add recording of the channel |
| Motion detection off · no zone · sensitivity 0 | enable · whole frame · 50 |
| Empty arming schedule | 24/7 |
| Camera does not notify the NVR about motion | enable “Notify Surveillance Center” on the camera |
| Audio off | enable in the main stream |
| Continuous recording but archive is not growing | reconnect the camera |
| — | **reboot the camera** (always available, never part of “fix all”) |

The audit repeats after the fixes; original settings are saved to `HIK-Slots_backup`.

### 🚀 Quick start

1. Run **`HIK-Slots.exe`** (SmartScreen → *More info → Run anyway* — the file is unsigned).
2. Press **`0`** → add a recorder (internal IP + port, optional external host + port).
3. Pick the recorder, enter the NVR login and password, choose a menu item.

Reports (`.xlsx`) and the `HIK-Slots_backup` folder (journals, original settings) are created
**next to the program**. Keep the folder: it is what *Undo a move* uses.

### 🔄 Updates

On every start the program first checks GitHub for a newer version — the latest **Release** with
`HIK-Slots.exe` attached, and the `VERSION` line in **`main`** (taking whichever is newer). If there is
one, it downloads it, runs the new file's self-test, moves your recorder settings into it, replaces
the `.exe` and restarts. Without internet access it simply starts the current version.

**Publishing an update:** raise `VERSION` in `nvr_slots.py`, rebuild `HIK-Slots.exe`, commit and
push to `main` — or create a Release tagged `v<version>` with `HIK-Slots.exe` attached.

### 🛠️ Build from source

```bash
pip install requests openpyxl colorama pyinstaller
pyinstaller --onefile --console --name "HIK-Slots" --icon hik.ico nvr_slots.py
```

The script also runs directly: `python nvr_slots.py` (or `Запуск.bat`).
`dev/` contains an NVR emulator (`python dev/mock_nvr.py`) and an end-to-end test suite
(`python dev/test_against_mock.py`) — everything can be developed without real hardware.

### 🔒 Security

- No addresses in the code — recorders are added at runtime.
- NVR and camera passwords are kept **in memory only** — never written to disk, journals or reports.
- Use only on systems you own or are authorised to administer.

### 📜 License

MIT — see [`LICENSE`](LICENSE).

<p align="right"><a href="#top">↑ back to top</a></p>

---

<a id="russian"></a>

## 🇷🇺 Русский

При добавлении на регистратор Hikvision камера встаёт в **первый свободный слот**, а не туда, куда
нужно. Переставить её вручную — значит сдвинуть все камеры между слотами, заново вписать имена и
проверить запись у каждой. **HIK-Slots** делает это одним действием по **ISAPI**: камера переезжает
в выбранный слот, а её **имя, расписание записи, связки событий и графики охраны — вместе с ней**.

Ещё программа **проверяет все камеры** (связь, запись, детекция движения, звук) и умеет **исправить
одну камеру** — неверный пароль, запись по движению, которая не срабатывает, пустые графики, камера,
которая не сообщает регистратору о движении — или перезагрузить её.

Стиль — как у [HIK-Report](https://github.com/b3dniy/HIK-Report): один `.exe`, радужный
прогресс-бар, мигающие бейджи SUCCESS / ERROR, регистраторы внутри `.exe`, отчёты в Excel.

### 🧭 Меню

| | |
|---|---|
| **Камеры** | |
| 1. Список камер | все слоты, связь, свободные слоты |
| 2. Проверить все камеры | связь, запись, архив, детекция движения, звук, диски → Excel |
| 3. Проверить и исправить одну камеру | в списке только проблемные → подробный аудит → что исправить или перезагрузка |
| **Перестановка** | |
| 4. Перенести камеру в другой слот | в свободный — или на занятый: **замена** или **обмен** |
| 5. Перенести камеру со сдвигом остальных | `D115 → D50`: камеры между ними сдвигаются на один, до первого свободного слота |
| 6. Поменять две камеры местами | меняются только два канала |
| 7. Отменить перестановку | выбор из последних перестановок, всё возвращается как было |

### ✨ Перестановка

- 🏷️ **Всё едет за камерой** — имя, расписание записи, связки событий (в том числе перекрёстные:
  «движение на D10 включает запись D60»), графики охраны.
- 🔗 **Каналы по ходу не удаляются.** Регистратор не пускает одну камеру в два слота, а многие
  прошивки добавляют камеру только в первый свободный слот. Поэтому HIK-Slots ставит в слот временную
  заглушку (несуществующий IP), записывает камеры *изменением* канала, а пустой слот занимает,
  на секунды заполнив свободные слоты ниже него. Каждая камера офлайн всего 10–30 сек.
- 🔑 **Пароль камер проверяется на самой камере до любых изменений** (одна попытка, чтобы не
  заблокировать учётку). У камер с другим паролем он спрашивается отдельно.
- 👀 **Сначала план** — ничего не меняется, пока не подтвердите `Y` / `Enter`.
- 💾 **Снимок и журнал шагов** перед каждым изменением: прерванную перестановку при следующем запуске
  можно продолжить или отменить; оставшиеся заглушки убираются.
- 🗺️ **Карта переезда** в Excel — где теперь каждая камера и **в каком старом слоте её прежний архив**
  (архив привязан к слоту — так устроены регистраторы).

### 🩺 Проверка и исправление камер

**Проверить все камеры** — связь, расписание записи, пополняется ли архив на самом деле (покрытие для
постоянной записи, число записей — для записи по движению), детекция движения (включена / зона /
чувствительность), «пишет по движению, но в связке движения нет записи», график охраны, звук, кодек,
заводские имена, дубли камер, диски. Камеры, выключенные специально, можно один раз отметить —
ошибкой они больше не считаются.

**Проверить и исправить одну камеру** — подробный аудит, в том числе того, что видно **только в самой
камере** (сообщает ли она регистратору о движении), и когда камера писала в последний раз (до 30 дней
назад). Что можно исправить:

| Проблема | Исправление |
|---|---|
| Неверный пароль камеры | ввести пароль (проверяется на камере) |
| Заводское имя (`IPCamera 22`) | задать имя |
| Расписание записи выключено / пустое / не круглосуточное | скопировать с другой камеры |
| В связке движения нет записи канала | добавить запись канала |
| Детекция выключена · зона не нарисована · чувствительность 0 | включить · весь кадр · 50 |
| Пустой график охраны | круглосуточно |
| Камера не сообщает регистратору о движении | включить в камере «Уведомить центр наблюдения» |
| Звук выключен | включить в основном потоке |
| Постоянная запись, но архив не пополняется | переподключить камеру |
| — | **перезагрузить камеру** (доступно всегда, в «исправить всё» не входит) |

После исправлений аудит повторяется; исходные настройки сохраняются в `HIK-Slots_backup`.

### 🚀 Быстрый старт

1. Запустите **`HIK-Slots.exe`** (SmartScreen → *Подробнее → Выполнить в любом случае* — файл без подписи).
2. Нажмите **`0`** → добавьте регистратор (внутренний IP + порт, при необходимости внешний).
3. Выберите регистратор, введите логин и пароль, выберите пункт меню.

Отчёты (`.xlsx`) и папка `HIK-Slots_backup` (журналы, исходные настройки) создаются **рядом с
программой**. Папку не удаляйте — по ней работает «Отменить перестановку».

### 🔄 Обновления

При каждом запуске программа первым делом проверяет GitHub: последний **Release** с приложенным
`HIK-Slots.exe` и строку `VERSION` в ветке **`main`** (берётся более новая). Если есть новая версия —
скачивает её, запускает самопроверку нового файла, переносит в него ваши регистраторы, подменяет
`.exe` и перезапускается. Без интернета просто запускается текущая версия.

**Как выпустить обновление:** увеличить `VERSION` в `nvr_slots.py`, пересобрать `HIK-Slots.exe`,
закоммитить и отправить в `main` — или создать Release с тегом `v<версия>` и приложить `HIK-Slots.exe`.

### 🛠️ Сборка из исходников

```bash
pip install requests openpyxl colorama pyinstaller
pyinstaller --onefile --console --name "HIK-Slots" --icon hik.ico nvr_slots.py
```

Скрипт можно запускать и напрямую: `python nvr_slots.py` (или `Запуск.bat`).
В `dev/` — эмулятор регистратора (`python dev/mock_nvr.py`) и сквозные тесты
(`python dev/test_against_mock.py`): всё можно разрабатывать без настоящего железа.

### 🔒 Безопасность

- Никаких адресов в коде — регистраторы добавляются при запуске.
- Пароли регистратора и камер живут **только в памяти** — не пишутся ни на диск, ни в журналы, ни в отчёты.
- Используйте только на системах, которыми владеете или имеете право администрировать.

### 📜 Лицензия

MIT — см. [`LICENSE`](LICENSE).

<p align="right"><a href="#top">↑ наверх</a></p>
