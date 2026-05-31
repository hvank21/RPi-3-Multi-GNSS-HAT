# GNSS Navigator Trimble Style

[![Python](https://img.shields.io/badge/Python-3.9%2B-blue.svg)](#requirements)
[![Platform](https://img.shields.io/badge/Platform-Raspberry%20Pi-C51A4A.svg)](#hardware)
[![Display](https://img.shields.io/badge/Display-720x480-2E8B57.svg)](#ui-layout)
[![GNSS](https://img.shields.io/badge/GNSS-NMEA%20GGA%2FRMC%2FGSA%2FGSV-orange.svg)](#features)
[![License](https://img.shields.io/badge/License-MIT-lightgrey.svg)](#license)

Compact Raspberry Pi GNSS monitor with a **receiver-style / Trimble-style UI** for small touch displays.

This project provides a full-screen Python/Tkinter application for real-time GNSS monitoring, satellite visualization, track display, constellation control, and SATS logging to Excel. It is optimized for **720×480** displays and supports GNSS input over **USB** or **UART**, with interface selection handled by shell launchers.

---

## Screenshots

> Add your screenshots to the repository, for example in `docs/screenshots/`, then update the paths below.

### Main NAV screen
![NAV screen](docs/screenshots/nav.png)

### SATS table
![SATS screen](docs/screenshots/sats.png)

### SKY / SNR
![SKY screen](docs/screenshots/sky.png)

### MAP view
![MAP screen](docs/screenshots/map.png)

---

## Features

- Real-time GNSS navigation data display:
  - FIX / NO FIX / 2D / 3D
  - Latitude / Longitude
  - Altitude
  - Speed
  - Course
  - UTC time and date
  - USED satellites
- Compact touchscreen-oriented multi-tab interface:
  - **NAV**
  - **SKY**
  - **SNR**
  - **DIAG**
  - **SATS**
  - **Constellations**
  - **MAP**
- Satellite **SkyPlot**
- **SNR bar graph**
- **SATS table** with visibility and USED logic
- **Track** accumulation and polyline display on map
- **Online** and **offline** map modes
- **Excel logging** of SATS table snapshots
- **Constellation switching** via queued **PCAS commands**
- Support for both:
  - **USB GNSS**
  - **UART GNSS**

---

## Project structure

```text
.
├── gnss_navigator_trimble_style.py
├── start_gnss_usb.sh
├── start_gnss_uart.sh
├── docs/
│   └── screenshots/
└── README.md
```

### Files

- `gnss_navigator_trimble_style.py` — main application
- `start_gnss_usb.sh` — launcher for USB-connected GNSS
- `start_gnss_uart.sh` — launcher for UART-connected GNSS

---

## How it works

The application has three main execution paths running in parallel:

1. **GNSSReader thread**
   - reads NMEA from serial
   - sends queued PCAS commands through the same serial connection

2. **NMEA parser**
   - parses `GGA`, `RMC`, `GSA`, `GSV`
   - updates shared GNSS state
   - updates track
   - recomputes USED satellites from `GSA`

3. **GUI + local map servers**
   - updates tabs in real time
   - serves online/offline map pages
   - optionally logs SATS data to `.xlsx`

---

## Hardware

### Raspberry Pi

Tested / intended for:
- Raspberry Pi 3
- Raspberry Pi 4
- Raspberry Pi Zero 2 W

### GNSS modules

Works with GNSS modules that output standard **NMEA** over serial.  
Typical target hardware for this project:

- **L76X GPS HAT**
- similar USB/UART GNSS modules

### Display

Optimized for:
- **720×480** touchscreen displays

---

## UI layout

This repository version is tailored for **720×480** resolution.

The layout is intentionally compact:
- reduced fonts
- reduced paddings
- compact SATS table
- compact diagnostics and map panels

For larger displays, a separate layout profile may be preferable.

---

## Requirements

### System packages

```bash
sudo apt update
sudo apt install python3-serial python3-openpyxl python3-tk
```

### Optional tools

For touchscreen calibration on X11:

```bash
sudo apt install xinput-calibrator
```

For VNC / remote administration, use your preferred Raspberry Pi OS setup.

---

## Installation

### 1. Clone the repository

```bash
git clone https://github.com/yourname/gnss-navigator-trimble-style.git
cd gnss-navigator-trimble-style
```

### 2. Copy files to the target directory

Example target on Raspberry Pi:

```bash
mkdir -p ~/gnss
cp gnss_navigator_trimble_style.py ~/gnss/
cp start_gnss_usb.sh ~/gnss/
cp start_gnss_uart.sh ~/gnss/
chmod +x ~/gnss/start_gnss_usb.sh ~/gnss/start_gnss_uart.sh
```

### 3. Verify Python dependencies

```bash
python3 -c "import serial, openpyxl, tkinter; print('OK')"
```

---

## Serial interface selection

This project intentionally keeps **interface selection outside Python**.

### Why

The Python app should not decide on its own whether GNSS is connected through:
- USB
- UART

That decision is handled in shell launchers.

### USB launcher

`start_gnss_usb.sh` typically probes:
- `/dev/ttyUSB0`
- `/dev/ttyUSB1`
- `/dev/ttyACM0`

and exports:

```bash
GNSS_SERIAL_PORT=/dev/ttyUSB0
```

### UART launcher

`start_gnss_uart.sh` typically probes:
- `/dev/serial0`
- `/dev/ttyAMA0`
- `/dev/ttyS0`

and exports:

```bash
GNSS_SERIAL_PORT=/dev/serial0
```

### Python side

The application only reads:

```python
SERIAL_PORT = os.environ.get("GNSS_SERIAL_PORT", "/dev/serial0")
```

---

## Running the application

### Run over USB

```bash
cd ~/gnss
./start_gnss_usb.sh
```

### Run over UART

```bash
cd ~/gnss
./start_gnss_uart.sh
```

---

## Tabs overview

### NAV
Displays:
- FIX
- LAT
- LON
- ALT
- SPEED
- COURSE
- USED
- UTC

### SKY
SkyPlot of visible satellites.

### SNR
Bar graph of signal strength.

### DIAG
Diagnostic information:
- selected serial port
- fix status
- DOP values
- visible satellites
- last NMEA sentence
- UART error state
- command status

### SATS
Satellite table with fields such as:
- system
- PRN
- USED
- elevation
- azimuth
- SNR
- status

Includes:
- **Record**
- **Stop**

### Constellations
Used to switch GNSS constellation modes via PCAS commands.

### MAP
Provides:
- online map access
- offline map access
- track visualization
- map status/info

---

## Track logic

Track is updated from valid `GGA` / `RMC` position data.

At a high level:
- the first valid point is stored immediately
- a new point is appended when coordinates change
- point storage is time-limited / size-limited
- track is exposed through `/api/position`
- Leaflet renders the track as a polyline

If the map shows no track:
- verify that coordinates actually change
- verify that valid `GGA` / `RMC` messages are present
- verify that the map is fetching `/api/position`

---

## SATS Excel logging

When **Record** is pressed:
- an `.xlsx` workbook is created

When **Stop** is pressed:
- the workbook is saved to:

```text
~/gnss_logs
```

Typical filename:
```text
sats_YYYYMMDD_HHMMSS.xlsx
```

The workbook contains rows with:
- local timestamp
- UTC
- selected filter
- fix state
- lat / lon / alt
- total used satellites
- system
- PRN
- USED
- elevation
- azimuth
- SNR
- status

---

## Online and offline map

### Online map
Uses **Leaflet + OpenStreetMap**.

### Offline map
Uses local tiles stored as:

```text
~/gnss_tiles/z/x/y.png
```

Example:
```text
~/gnss_tiles/12/2203/1345.png
```

---

## systemd autostart

### USB GNSS service

Create:

```bash
sudo nano /etc/systemd/system/gnss-navigator-usb.service
```

Paste:

```ini
[Unit]
Description=GNSS Navigator (USB)
After=graphical.target

[Service]
Type=simple
User=val
WorkingDirectory=/home/val/gnss
ExecStart=/home/val/gnss/start_gnss_usb.sh
Restart=always
RestartSec=3
Environment=DISPLAY=:0
Environment=XAUTHORITY=/home/val/.Xauthority

[Install]
WantedBy=graphical.target
```

Enable it:

```bash
sudo systemctl daemon-reload
sudo systemctl enable gnss-navigator-usb.service
sudo systemctl start gnss-navigator-usb.service
```

### UART GNSS service

Create:

```bash
sudo nano /etc/systemd/system/gnss-navigator-uart.service
```

Paste:

```ini
[Unit]
Description=GNSS Navigator (UART)
After=graphical.target

[Service]
Type=simple
User=val
WorkingDirectory=/home/val/gnss
ExecStart=/home/val/gnss/start_gnss_uart.sh
Restart=always
RestartSec=3
Environment=DISPLAY=:0
Environment=XAUTHORITY=/home/val/.Xauthority

[Install]
WantedBy=graphical.target
```

Enable it:

```bash
sudo systemctl daemon-reload
sudo systemctl enable gnss-navigator-uart.service
sudo systemctl start gnss-navigator-uart.service
```

### Check service status

```bash
sudo systemctl status gnss-navigator-usb.service
sudo systemctl status gnss-navigator-uart.service
```

---

## Touchscreen calibration

Recommended workflow:
- use **X11**
- calibrate with `xinput_calibrator`

Example:

```bash
xinput list
xinput_calibrator --device 10
```

Save calibration to:

```text
/etc/X11/xorg.conf.d/99-calibration.conf
```

---

## Troubleshooting

### USB GNSS not found

```bash
ls /dev/ttyUSB* /dev/ttyACM*
```

### UART GNSS not found

```bash
ls -l /dev/serial0 /dev/ttyAMA0 /dev/ttyS0
```

### gpsd conflict

If another application such as `cgps` is already using the port:

```bash
sudo systemctl stop gpsd.socket gpsd.service
```

### Low voltage on Raspberry Pi

Use a stable power supply and check:

```bash
vcgencmd get_throttled
```

### No track visible

Check:
- valid `GGA` / `RMC`
- changing coordinates
- map page open
- `/api/position` updates

### X11 / VNC issues after switching from Wayland

Verify:
- X11 is enabled
- VNC is enabled
- desktop session is started
- correct VNC resolution is configured

---

## Customization

Common things you may want to adjust:
- screen resolution constants
- fonts and paddings
- SATS table size
- map ports
- track limits
- GNSS command presets

---

## Roadmap

- cleaner profile switching for multiple display resolutions
- more polished map embedding
- richer constellation management
- improved event / command log in GUI
- optional data export formats beyond Excel
- better screenshot/documentation set

---

## License

Choose the license you want for the repository, for example:

```text
MIT License
```

---

## Notes

This README is written in English for GitHub repository use.  
If needed, you can keep a separate `README_uk.md` for the Ukrainian version.
