# GNSS Navigator Trimble Style

GNSS Navigator Trimble Style — це Python GUI-застосунок для Raspberry Pi з інтерфейсом у стилі GNSS receiver / Trimble-style для моніторингу даних GNSS-модуля в реальному часі.

Програма орієнтована на невеликі сенсорні дисплеї Raspberry Pi і в цій версії оптимізована під **720×480**.  
Підтримується запуск як через **USB GNSS**, так і через **UART GNSS** через окремі стартові shell-скрипти.

---

## Можливості

- відображення навігаційних параметрів у реальному часі:
  - FIX
  - Latitude / Longitude
  - Altitude
  - Speed
  - Course
  - UTC
  - USED satellites
- вкладки:
  - **NAV**
  - **SKY**
  - **SNR**
  - **DIAG**
  - **SATS**
  - **Сузір'я**
  - **MAP**
- SkyPlot супутників
- SNR bars
- таблиця супутників SATS
- online/offline карта
- трек руху на карті
- запис таблиці SATS у **Excel (.xlsx)**
- керування режимами сузір’їв через **PCAS-команди**
- підтримка запуску:
  - через **USB** (`ttyUSB0`, `ttyUSB1`, `ttyACM0`)
  - через **UART** (`serial0`, `ttyAMA0`, `ttyS0`)

---

## Для чого призначена програма

Це не просто NMEA viewer, а повноцінний компактний GNSS-інтерфейс для Raspberry Pi, який може використовуватись для:

- польових тестів GNSS-модуля
- відображення стану приймача на сенсорному дисплеї
- контролю супутників і якості прийому
- перевірки режимів сузір’їв
- локального логування SATS в Excel
- візуалізації треку

---

## Основні функціональні блоки

### 1. Прийом NMEA
Окремий потік читає NMEA з serial-порту.

### 2. Парсер GNSS
Підтримуються речення:
- `GGA`
- `RMC`
- `GSA`
- `GSV`

### 3. Відображення в GUI
Tkinter GUI оновлює вкладки в реальному часі.

### 4. Локальний map server
Програма піднімає локальний HTTP server для:
- **online map**
- **offline map**

### 5. Логування SATS у Excel
При натисканні **«Запис»** створюється `.xlsx`, при **«Стоп»** файл зберігається.

### 6. Перемикання сузір’їв
Додана вкладка для керування режимами GNSS через PCAS-команди.

---

## Підтримуване обладнання

### Raspberry Pi
- Raspberry Pi 3 / 4 / Zero 2 W
- дисплей 720×480

### GNSS
- L76X GPS HAT
- інші модулі, що віддають стандартний NMEA через serial

### Підключення GNSS
- **USB**
- **UART**

---

## Залежності

Встановіть потрібні пакети:

```bash
sudo apt update
sudo apt install python3-serial python3-openpyxl
```

Якщо потрібно:

```bash
sudo apt install python3-tk
```

---

## Структура файлів

```text
gnss_navigator_trimble_style.py
start_gnss_usb.sh
start_gnss_uart.sh
```

### Призначення
- `gnss_navigator_trimble_style.py` — основна програма
- `start_gnss_usb.sh` — запуск через USB
- `start_gnss_uart.sh` — запуск через UART

---

## Логіка вибору порту

У цій версії вибір типу підключення винесено в shell-скрипти.

### USB запуск
Скрипт шукає:
- `/dev/ttyUSB0`
- `/dev/ttyUSB1`
- `/dev/ttyACM0`

і передає знайдений порт у Python через змінну:
```bash
GNSS_SERIAL_PORT
```

### UART запуск
Скрипт шукає:
- `/dev/serial0`
- `/dev/ttyAMA0`
- `/dev/ttyS0`

і також передає порт через:
```bash
GNSS_SERIAL_PORT
```

### Всередині Python
Python не вибирає інтерфейс самостійно, а лише читає:
```python
SERIAL_PORT = os.environ.get("GNSS_SERIAL_PORT", "/dev/serial0")
```

---

## Запуск

### 1. Запуск через USB
```bash
chmod +x start_gnss_usb.sh
./start_gnss_usb.sh
```

### 2. Запуск через UART
```bash
chmod +x start_gnss_uart.sh
./start_gnss_uart.sh
```

---

## Вкладки інтерфейсу

### NAV
Показує:
- FIX
- LAT
- LON
- ALT
- SPEED
- COURSE
- USED
- UTC

### SKY
SkyPlot видимих супутників.

### SNR
Гістограма рівнів SNR.

### DIAG
Діагностична інформація:
- serial port
- останній NMEA рядок
- fix
- DOP
- кількість супутників
- статус UART

### SATS
Таблиця супутників:
- SYS
- PRN
- USED
- EL
- AZ
- SNR
- STATUS

Кнопки:
- **Запис**
- **Стоп**

### Сузір'я
Керування режимами GNSS-модуля через PCAS-команди.

### MAP
Інформація про карту і відкриття:
- online map
- offline map

---

## Трек

У програмі реалізовано накопичення треку.

Алгоритм:
- перша точка записується одразу
- нова точка додається, якщо координати змінились
- точки обмежуються за максимальною кількістю
- трек відображається на карті як polyline

---

## Online / Offline карта

### Online map
Використовується OpenStreetMap через Leaflet.

### Offline map
Програма читає локальні тайли з каталогу:

```text
~/gnss_tiles/z/x/y.png
```

---

## Excel лог SATS

При записі створюється файл у каталозі:

```text
~/gnss_logs
```

Формат імені:
```text
sats_YYYYMMDD_HHMMSS.xlsx
```

Лог містить:
- local timestamp
- UTC
- filter
- fix
- lat
- lon
- altitude
- total satellites used
- system
- PRN
- used
- elevation
- azimuth
- snr
- status

---

## Приклад сценарію використання

### USB GNSS
1. Підключити GNSS через USB
2. Запустити:
```bash
./start_gnss_usb.sh
```

### UART GNSS
1. Підключити GNSS до UART Raspberry Pi
2. Запустити:
```bash
./start_gnss_uart.sh
```

---

## Калібрування тача

Для сенсорного дисплея рекомендується:
- X11 режим
- калібрування через `xinput_calibrator`

Приклад:
```bash
xinput list
xinput_calibrator --device 10
```

---

## Типові проблеми

### Не знаходиться USB GNSS
Перевірити:
```bash
ls /dev/ttyUSB* /dev/ttyACM*
```

### Не знаходиться UART GNSS
Перевірити:
```bash
ls -l /dev/serial0 /dev/ttyAMA0 /dev/ttyS0
```

### Конфлікт із gpsd
Якщо `cgps` працює, а програма не може відкрити порт:
```bash
sudo systemctl stop gpsd.socket gpsd.service
```

### Немає треку
Перевірити:
- чи реально змінюються координати
- чи є валідний `GGA/RMC`
- чи відкрито карту
- чи оновлюється `track` у `/api/position`

---

## Плани розвитку

- автозапуск через `systemd`
- ще компактніший інтерфейс для маленьких дисплеїв
- розширення режимів сузір’їв
- покращений offline map режим
- журнал подій/команд у GUI

---

## Ліцензія

За потреби додайте вашу ліцензію, наприклад:

```text
MIT License
```

---

## Автор / примітка

Проєкт адаптований під Raspberry Pi та компактний сенсорний GNSS-інтерфейс у стилі Trimble / receiver-style GUI.
