<a id="top"></a>

<p align="center">
  <img src="hik.png" width="300" alt="HIK-Slots">
</p>

<h1 align="center">HIK-Slots</h1>

<p align="center">
  <b>Move cameras between NVR slots, check and fix cameras — in one terminal app.</b><br>
  <sub>Перестановка камер по слотам, проверка и исправление камер — в одной программе.</sub>
</p>

<p align="center">
  <img src="https://img.shields.io/badge/Windows-10%20%2F%2011-0078D6?style=flat-square&logo=windows&logoColor=white">
  <img src="https://img.shields.io/badge/Hikvision-ISAPI-FFC20E?style=flat-square">
  <img src="https://img.shields.io/badge/single-.exe-2ea44f?style=flat-square">
  <img src="https://img.shields.io/badge/auto-update-8957e5?style=flat-square">
  <img src="https://img.shields.io/badge/license-MIT-green?style=flat-square">
</p>

<p align="center">
  <a href="#english"><img src="https://img.shields.io/badge/🇬🇧%20English-2ea44f?style=for-the-badge"></a>
  &nbsp;
  <a href="#russian"><img src="https://img.shields.io/badge/🇷🇺%20Русский-1f6feb?style=for-the-badge"></a>
</p>

---

<a id="english"></a>

## 🇬🇧 English

A new camera always lands in the first free NVR slot. HIK-Slots moves it where you need it —
together with its name, recording and event settings.

### 🧭 Menu

| | |
|---|---|
| 📋 **Camera list** | all slots and free slots |
| 🩺 **Check all cameras** | online, recording, motion, audio, disks → Excel |
| 🔧 **Check & fix a camera** | shows only problem cameras, fixes them, reboots |
| ➡️ **Move to another slot** | free slot, or replace / swap with an occupied one |
| ⏩ **Move with shift** | `D115 → D50`, the cameras in between shift by one |
| 🔁 **Swap two cameras** | only these two slots change |
| ↩️ **Undo a move** | any of the recent moves |

### 🔀 Moving cameras

- 🏷️ Name, recording schedule and events move with the camera
- 👀 You see the plan first — nothing changes until you confirm
- ⚡ Each camera is offline for just 10–30 seconds
- 🗺️ Excel map shows where each camera was — the old footage stays in the old slot

### 🔧 What can be fixed

| Problem | Fix |
|---|---|
| 🔑 Wrong camera password | enter the right one |
| 🏷️ Default name | rename |
| 🎬 Recording off or not 24/7 | copy the schedule from another camera |
| 🏃 Motion does not trigger recording | detection, zone, sensitivity, linkage, arming schedule |
| 📡 Camera does not notify the NVR | enable it on the camera |
| 🔇 Audio off | turn on |
| 🔄 Anything else | reboot the camera |

### 🚀 Quick start

1. Run **`HIK-Slots.exe`**
2. Press **`0`** and add your recorder
3. Pick the recorder, log in, choose a menu item

> Keep the **`HIK-Slots_backup`** folder next to the program — it is used to undo moves.

### 🔄 Updates

The program updates itself on start when a new version is published here.

### 🛠️ Build

```bash
pip install requests openpyxl colorama pyinstaller
pyinstaller --onefile --console --name "HIK-Slots" --icon hik.ico nvr_slots.py
```

### 📜 License

MIT — see [LICENSE](LICENSE).

<p align="right"><a href="#top">↑ top</a></p>

---

<a id="russian"></a>

## 🇷🇺 Русский

Новая камера всегда встаёт в первый свободный слот регистратора. HIK-Slots переносит её куда нужно —
вместе с именем, записью и настройками событий.

### 🧭 Меню

| | |
|---|---|
| 📋 **Список камер** | все слоты и свободные слоты |
| 🩺 **Проверить все камеры** | связь, запись, движение, звук, диски → Excel |
| 🔧 **Проверить и исправить камеру** | только проблемные камеры, исправление, перезагрузка |
| ➡️ **Перенести в другой слот** | в свободный или на занятый с заменой / обменом |
| ⏩ **Перенести со сдвигом** | `D115 → D50`, камеры между ними сдвигаются на один |
| 🔁 **Поменять две камеры местами** | меняются только эти два слота |
| ↩️ **Отменить перестановку** | любую из последних |

### 🔀 Перестановка

- 🏷️ Имя, расписание записи и события переезжают вместе с камерой
- 👀 Сначала план — ничего не меняется без подтверждения
- ⚡ Каждая камера офлайн всего 10–30 секунд
- 🗺️ Карта в Excel покажет, где камера стояла раньше — старый архив остаётся в старом слоте

### 🔧 Что исправляется

| Проблема | Исправление |
|---|---|
| 🔑 Неверный пароль камеры | ввести правильный |
| 🏷️ Заводское имя | переименовать |
| 🎬 Запись выключена или не круглосуточно | скопировать расписание с другой камеры |
| 🏃 Движение не включает запись | детекция, зона, чувствительность, связка, график охраны |
| 📡 Камера не сообщает регистратору | включить в камере |
| 🔇 Звук выключен | включить |
| 🔄 Всё остальное | перезагрузить камеру |

### 🚀 Быстрый старт

1. Запустите **`HIK-Slots.exe`**
2. Нажмите **`0`** и добавьте регистратор
3. Выберите регистратор, войдите, выберите пункт меню

> Папку **`HIK-Slots_backup`** рядом с программой не удаляйте — она нужна для отмены перестановок.

### 🔄 Обновления

Программа сама обновляется при запуске, когда здесь выходит новая версия.

### 🛠️ Сборка

```bash
pip install requests openpyxl colorama pyinstaller
pyinstaller --onefile --console --name "HIK-Slots" --icon hik.ico nvr_slots.py
```

### 📜 Лицензия

MIT — см. [LICENSE](LICENSE).

<p align="right"><a href="#top">↑ наверх</a></p>
