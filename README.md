<a id="top"></a>

<p align="center">
  <img src="hik.png" width="300" alt="HIK-Slots">
</p>

<h1 align="center">HIK-Slots</h1>

<p align="center">
  Camera slot management and camera checks for Hikvision NVRs.<br>
  <sub>Перестановка камер по слотам и проверка камер на регистраторах Hikvision.</sub>
</p>

<p align="center">
  <img src="https://img.shields.io/badge/platform-Windows%2010%20%2F%2011-0078D6?style=flat-square&logo=windows&logoColor=white">
  <img src="https://img.shields.io/badge/python-3.x-3776AB?style=flat-square&logo=python&logoColor=white">
  <img src="https://img.shields.io/badge/API-Hikvision%20ISAPI-FFC20E?style=flat-square">
  <img src="https://img.shields.io/badge/license-MIT-green?style=flat-square">
</p>

<p align="center">
  <a href="#english">English</a> · <a href="#russian">Русский</a>
</p>

---

<a id="english"></a>

## English

A Hikvision NVR always puts a newly added camera into the first free slot. To get it into a specific
slot by hand you have to shift every camera in between and then restore names and recording settings.
HIK-Slots does this over ISAPI in one operation. It also checks all cameras and can fix common problems
on a single camera.

### Menu

| Item | What it does |
|---|---|
| 1. Camera list | All slots with online status and free slots |
| 2. Check all cameras | Online status, recording, archive, motion detection, audio, disks. Saves an Excel report |
| 3. Check and fix one camera | Lists only cameras with problems, runs a detailed check of the selected one and offers fixes or a reboot |
| 4. Move a camera to another slot | To a free slot, or to an occupied one with replace or swap |
| 5. Move with shift | For example `D115 → D50`: the cameras in between move up by one slot, up to the first free slot |
| 6. Swap two cameras | Only the two channels are changed |
| 7. Undo a move | Restores one of the recent moves |

### Moving cameras

The camera keeps its name, recording schedule, event linkage and arming schedules in the new slot.
Cross-links are updated too, so if motion on D10 started recording on D60, it will point to the new slot.

Channels are not deleted during a move. The NVR does not allow the same camera in two slots, and some
firmware adds a camera only to the first free slot. The program writes a temporary channel with a
non-existent IP into the slot instead, and fills empty slots below the target for a few seconds when it
needs to add a camera to an empty slot. Each camera is offline for about 10–30 seconds.

Before anything is changed:

- the plan is shown and has to be confirmed;
- the camera password is checked by logging in to the camera once;
- a snapshot of the settings and a step journal are saved, so an interrupted move can be resumed or undone.

The recording archive stays in the old slot. The Excel move report lists where each camera was before,
so older footage can be found.

### Checking and fixing a camera

Item 3 checks the camera through the NVR and directly on the camera, and shows when it last recorded
(up to 30 days back). It can fix:

| Problem | Fix |
|---|---|
| Wrong camera password | Enter the password, it is checked on the camera |
| Default name (`IPCamera 22`) | Set a name |
| Recording schedule off, empty or not 24/7 | Copy it from another camera |
| Motion event does not start recording | Add recording to the motion linkage |
| Motion detection off, no detection area, sensitivity 0 | Enable, full frame, sensitivity 50 |
| Empty arming schedule | Set to 24/7 |
| Camera does not notify the NVR about motion | Enable "Notify Surveillance Center" on the camera |
| Audio off | Enable audio in the main stream |
| Continuous recording but the archive is not growing | Reconnect the camera |

A reboot is always available. After the fixes the check runs again. The original settings are saved
to `HIK-Slots_backup`.

Cameras that are switched off on purpose can be marked after the check in item 2 and are no longer
reported as errors.

### Getting started

1. Run `HIK-Slots.exe`. Windows SmartScreen may warn about an unknown publisher: More info → Run anyway.
2. Press `0` and add a recorder: internal IP and port, and optionally an external address and port.
3. Select the recorder and log in with the NVR account.

Reports and the `HIK-Slots_backup` folder are created next to the program. Do not delete the folder:
it is needed to undo moves.

### Updates

On start the program checks GitHub for a newer version: the latest release with `HIK-Slots.exe`
attached and the `VERSION` line in `main`. A newer version is downloaded, checked with its self-test,
given the saved recorder list, and started in place of the old one. Without internet access the
current version starts as usual.

To publish an update, increase `VERSION` in `nvr_slots.py`, rebuild `HIK-Slots.exe` and push it to
`main`, or create a release `v<version>` with the exe attached.

### Building

```bash
pip install requests openpyxl colorama pyinstaller
pyinstaller --onefile --console --name "HIK-Slots" --icon hik.ico nvr_slots.py
```

The script can also be run with `python nvr_slots.py` or `Запуск.bat`.
`dev/` contains an NVR emulator (`python dev/mock_nvr.py`) and tests (`python dev/test_against_mock.py`).

### Security

- The code contains no addresses. Recorders are added in the settings.
- NVR and camera passwords are kept in memory only and are not written to disk, journals or reports.
- Use the program only on systems you are allowed to administer.

### License

MIT, see [LICENSE](LICENSE).

<p align="right"><a href="#top">↑ top</a></p>

---

<a id="russian"></a>

## Русский

Регистратор Hikvision ставит новую камеру в первый свободный слот. Чтобы перенести её в нужный слот
вручную, приходится сдвигать все камеры между ними, а потом заново настраивать имена и запись.
HIK-Slots делает это одной операцией через ISAPI. Кроме того, программа проверяет все камеры и умеет
исправлять типичные проблемы на отдельной камере.

### Меню

| Пункт | Что делает |
|---|---|
| 1. Список камер | Все слоты, связь, свободные слоты |
| 2. Проверить все камеры | Связь, запись, архив, детекция движения, звук, диски. Сохраняет отчёт в Excel |
| 3. Проверить и исправить одну камеру | Показывает только камеры с проблемами, подробно проверяет выбранную и предлагает исправления или перезагрузку |
| 4. Перенести камеру в другой слот | В свободный слот или в занятый с заменой либо обменом |
| 5. Перенести камеру со сдвигом остальных | Например `D115 → D50`: камеры между ними сдвигаются на один слот, до первого свободного |
| 6. Поменять две камеры местами | Меняются только два канала |
| 7. Отменить перестановку | Возвращает одну из последних перестановок |

### Перестановка

В новом слоте у камеры сохраняются имя, расписание записи, связки событий и графики охраны.
Перекрёстные связки тоже обновляются: если движение на D10 включало запись D60, связка будет указывать
на новый слот.

Каналы во время перестановки не удаляются. Регистратор не даёт поставить одну камеру в два слота, а
некоторые прошивки добавляют камеру только в первый свободный слот. Поэтому программа временно
записывает в слот канал с несуществующим IP, а если камеру нужно поставить в пустой слот, на несколько
секунд занимает свободные слоты ниже него. Каждая камера бывает офлайн примерно 10–30 секунд.

Перед изменениями:

- показывается план, его нужно подтвердить;
- пароль камер проверяется одним входом на камеру;
- сохраняются снимок настроек и журнал шагов, поэтому прерванную перестановку можно продолжить или отменить.

Архив записей остаётся в старом слоте. В отчёте о перестановке (Excel) указано, где камера стояла
раньше, по нему можно найти старые записи.

### Проверка и исправление камеры

Пункт 3 проверяет камеру через регистратор и напрямую на камере и показывает, когда она писала в
последний раз (до 30 дней назад). Что можно исправить:

| Проблема | Исправление |
|---|---|
| Неверный пароль камеры | Ввести пароль, он проверяется на камере |
| Заводское имя (`IPCamera 22`) | Задать имя |
| Расписание записи выключено, пустое или не круглосуточное | Скопировать с другой камеры |
| Движение не включает запись | Добавить запись в связку движения |
| Детекция выключена, нет зоны, чувствительность 0 | Включить, весь кадр, чувствительность 50 |
| Пустой график охраны | Сделать круглосуточным |
| Камера не сообщает регистратору о движении | Включить в камере «Уведомить центр наблюдения» |
| Звук выключен | Включить звук в основном потоке |
| Постоянная запись, но архив не пополняется | Переподключить камеру |

Перезагрузка камеры доступна всегда. После исправлений проверка повторяется. Исходные настройки
сохраняются в `HIK-Slots_backup`.

Камеры, выключенные специально, можно отметить после проверки в пункте 2, и они больше не считаются
ошибкой.

### Начало работы

1. Запустите `HIK-Slots.exe`. Windows SmartScreen может предупредить о неизвестном издателе:
   «Подробнее» → «Выполнить в любом случае».
2. Нажмите `0` и добавьте регистратор: внутренний IP и порт, при необходимости внешний адрес и порт.
3. Выберите регистратор и войдите под учётной записью регистратора.

Отчёты и папка `HIK-Slots_backup` создаются рядом с программой. Папку не удаляйте, она нужна для
отмены перестановок.

### Обновления

При запуске программа проверяет на GitHub последний релиз с приложенным `HIK-Slots.exe` и строку
`VERSION` в ветке `main`. Если есть версия новее, она скачивается, проходит самопроверку, получает
сохранённый список регистраторов и запускается вместо старой. Без интернета запускается текущая версия.

Чтобы выпустить обновление, увеличьте `VERSION` в `nvr_slots.py`, пересоберите `HIK-Slots.exe` и
отправьте в `main` или создайте релиз `v<версия>` с приложенным exe.

### Сборка

```bash
pip install requests openpyxl colorama pyinstaller
pyinstaller --onefile --console --name "HIK-Slots" --icon hik.ico nvr_slots.py
```

Скрипт можно запускать и напрямую: `python nvr_slots.py` или `Запуск.bat`.
В `dev/` лежат эмулятор регистратора (`python dev/mock_nvr.py`) и тесты (`python dev/test_against_mock.py`).

### Безопасность

- В коде нет адресов, регистраторы добавляются в настройках.
- Пароли регистратора и камер хранятся только в памяти и не пишутся на диск, в журналы и отчёты.
- Используйте программу только на системах, которые вам разрешено администрировать.

### Лицензия

MIT, см. [LICENSE](LICENSE).

<p align="right"><a href="#top">↑ наверх</a></p>
