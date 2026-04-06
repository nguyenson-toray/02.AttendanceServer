# Attendance Server V2

Server chạy nền trên Windows, đồng bộ dữ liệu chấm công từ máy ZK → MongoDB và xử lý đăng ký OT qua QR code PDF.

---

## Kiến trúc tổng quan

```
functions.py  ← toàn bộ business logic (shared)
│
├── Config Layer
│   └── read_config()              ← đọc 02.Config/config.json
│
├── Employee Layer
│   ├── get_list_emp()             ← load dict từ MongoDB vào RAM
│   └── find_*()                   ← O(1) lookup qua dict
│
├── Excel Sync Layer               ← chạy theo schedule
│   ├── excel_aio_to_db()
│   ├── excel_maternity_to_db()
│   └── excel_resign_to_db()
│
├── Attendance Layer               ← 1 thread / máy chấm công
│   ├── get_att_log_one_time()     ← pull định kỳ
│   ├── live_capture_attendance()  ← realtime + auto-reconnect
│   └── get_att_log()              ← thread entry point (one-time → live/poll)
│
├── OT Register Layer              ← chạy theo schedule
│   ├── ot_register_detect_qr_and_save()
│   ├── qr_code_ot_register_to_db()
│   └── ot_register_append_excel()
│
├── QR Decode Layer
│   ├── decode_qr_from_image()     ← multi-pipeline robust decode
│   ├── _preprocess()              ← 5 pipelines: gray/otsu/adaptive/sharpen/upscale2x
│   └── _decode_image()            ← pyzbar + cv2.QRCodeDetector
│
└── Scheduler
    └── setup_schedule()           ← đăng ký tất cả jobs từ schedule_config


attServerV2.py    ← CLI entry point  (import functions, chỉ có __main__)
attServerV2UI.py  ← GUI entry point  (import functions as server, tkinter)
```

---

## Cấu trúc thư mục

```
02.AttendanceServer/
├── 01.AppRunning/             # PyInstaller output (exe + _internal/)
├── 02.Config/
│   ├── config.json            # cấu hình chính (IP, paths, bypass names)
│   └── config.xlsx            # giữ lại nếu team khác còn dùng
├── 03.Logs/
│   ├── info_YYYYMMDD.txt
│   ├── error_YYYYMMDD.txt
│   └── debug_YYYYMMDD.txt
└── src/
    ├── functions.py           # business logic (shared)
    ├── attServerV2.py         # CLI wrapper (~50 dòng)
    ├── attServerV2UI.py       # GUI wrapper (tkinter)
    ├── attServer.py           # v1 (gốc, giữ lại tham khảo)
    └── readme.md
```

---

## Config (02.Config/config.json)

```json
{
  "att_machines": ["10.0.1.41", "10.0.1.42"],
  "poppler_path": "C:\\Program Files\\poppler-24.02.0\\Library\\bin",
  "bypass_names": ["Shoji Izumi", "Amagata Osamu"],
  "paths": {
    "aio":       "\\\\10.0.1.5\\tiqn\\...\\Toray's employees information All in one.xlsx",
    "resign":    "\\\\10.0.1.5\\tiqn\\...\\Resigned report.xlsx",
    "maternity": "\\\\10.0.1.5\\tiqn\\...\\Danh sách nhân viên nữ mang thai 1.xlsx",
    "ot_folder": "\\\\10.0.1.5\\tiqn\\...\\7.OT request"
  },
  "sheets": {
    "aio":                "Sum",
    "resign":             "Resigned",
    "maternity_leave":    "Thai sản",
    "maternity_pregnant": "Mang thai",
    "maternity_child":    "Con nhỏ dưới 12 tháng"
  },
  "schedule": {
    "sync_time_day":            "sunday",
    "sync_time_at":             "06:00",
    "excel_sync_times":         ["07:00", "09:00", "11:50", "15:00", "18:00", "22:00"],
    "ot_scan_interval_minutes":  10,
    "att_log_interval_minutes":  6
  }
}
```

| Field | Mô tả |
|---|---|
| `att_machines` | Danh sách IP máy chấm công ZK, thứ tự → machineNo 1, 2, ... |
| `poppler_path` | Đường dẫn poppler để convert PDF → ảnh |
| `bypass_names` | Nhân viên bỏ qua khi sync Employee |
| `paths.*` | Đường dẫn các file Excel trên file server |
| `sheets.*` | Tên sheet tương ứng trong từng file Excel |
| `schedule.*` | Cấu hình lịch chạy định kỳ |

---

## MongoDB Collections

| Collection | Mô tả |
|---|---|
| `Employee` | Thông tin nhân viên (empId, attFingerId, name, workStatus, ...) |
| `AttLog` | Bản ghi chấm công (machineNo, attFingerId, empId, name, timestamp) |
| `HistoryGetAttLogs` | Lần cuối lấy log mỗi máy (`{machine, lastTimeGetAttLogs, lastCount}`) |
| `OtRegister` | Đăng ký OT từ QR code (requestNo, empId, otDate, otTimeBegin, otTimeEnd) |
| `MaternityTracking` | Theo dõi thai sản |

---

## Logic & Thuật toán chính

### 1. Employee lookup — O(1) dict

Sau mỗi lần sync Excel, build hai dict trong RAM:

```
emp_by_finger_id[attFingerId] = {empId, name}
emp_by_emp_id[empId]          = {attFingerId, name}
```

Tra cứu: `emp_by_finger_id.get(finger_id, {}).get('name', 'Not found')` — O(1).

### 2. Excel → MongoDB — Bulk Write

Các hàm `excel_*_to_db()` gom tất cả `UpdateOne` operations rồi gọi `bulk_write()` một lần duy nhất, giảm round-trips từ N xuống còn 1.

### 3. Live Capture — Auto-Reconnect với Exponential Backoff

```
delay = min(delay × 2, 300s)
```

Bắt đầu từ 10s, tăng gấp đôi mỗi lần thất bại, tối đa 5 phút. Reset về 10s nếu connection sống được > 60s.

**Zombie detection:** `live_capture()` yield `None` khi socket timeout (10s). Nếu nhận `None` liên tiếp ≥ 6 lần (= 60s không có data), force reconnect dù không có exception.

```
start thread
    │
    ▼
[while True] ←──────────────────────────────────┐
    │                                            │
    ├─► connect() ──fail──► log + sleep(delay×2) ┤
    │                                            │
    ├─► live_capture() loop                      │
    │       ├── None × 6 → force break ──────────┤
    │       ├── attendance → insert DB            │
    │       └── KeyboardInterrupt → EXIT thread  │
    │                                            │
    └─► Exception → log + sleep(delay×2) ────────┘
```

### 4. QR Decode — Multi-Pipeline

`decode_qr_from_image()` thử 5 preprocessing pipeline theo thứ tự, dừng ngay khi decode được:

| Thứ tự | Pipeline | Mục đích |
|---|---|---|
| 1 | `gray` | Baseline, nhanh nhất |
| 2 | `otsu` | Ảnh nhiễu / tương phản cao |
| 3 | `adaptive` | Ánh sáng không đều |
| 4 | `sharpen` | Ảnh mờ / nét kém |
| 5 | `upscale2x` | QR nhỏ / độ phân giải thấp |
| fallback | `zxingcpp` | QR hỏng / cắt xén (nếu cài) |

Mỗi pipeline thử 2 decoder: **pyzbar** + **cv2.QRCodeDetector**. DPI convert PDF = 400 (cao hơn 300 để cải thiện chất lượng scan mờ).

### 5. History lookup — Direct MongoDB Query

`find_one({"machine": machineNo})` thay vì scan toàn bộ collection.

### 6. OT Request dedup — Set

`ot_request_no_on_db` là `set` → kiểm tra trùng `in` là O(1).

---

## Schedule

| Thời gian | Hành động |
|---|---|
| Sunday 06:00 (configurable) | `sync_time_devices()` — đồng bộ giờ máy chấm công |
| Hàng ngày theo `excel_sync_times` | `update_excel_to_mongoDb()` |
| Mỗi `ot_scan_interval_minutes` phút | `ot_register_detect_qr_and_save()` |

---

## Hướng dẫn cài đặt

### Yêu cầu

- Python 3.10+
- MongoDB 6+ chạy local (`localhost:27017`)
- [Poppler for Windows](https://github.com/oschwartz10612/poppler-windows/releases)

### Cài thư viện

```bash
pip install pymongo pandas openpyxl pyzk opencv-python pyzbar pdf2image schedule numpy
# tuỳ chọn — decode QR hỏng tốt hơn:
pip install zxingcpp
```

### Tạo config.json

1. Copy mẫu ở mục **Config** phía trên vào `02.Config/config.json`
2. Điền IP máy chấm công vào `att_machines`
3. Điền đường dẫn các file Excel vào `paths`
4. Điền tên sheet vào `sheets`
5. Điền đường dẫn poppler vào `poppler_path`

### Khởi động (CLI)

```bash
cd src
python attServerV2.py
```

### Khởi động (GUI)

```bash
cd src
python attServerV2UI.py
```

### Khởi động không có console (chạy nền)

```bash
pythonw attServerV2.py
# hoặc
pythonw attServerV2UI.py
```

Hoặc dùng file `.exe` trong `01.AppRunning/`.

### Khởi tạo HistoryGetAttLogs

Trước lần chạy đầu tiên, tạo document cho mỗi máy:

```javascript
// MongoDB shell
db.HistoryGetAttLogs.insertMany([
  { machine: 1, lastTimeGetAttLogs: new Date("2020-01-01"), lastCount: 0 },
  { machine: 2, lastTimeGetAttLogs: new Date("2020-01-01"), lastCount: 0 }
  // ... thêm theo số máy trong att_machines
])
```

---

## Bugs đã fix

| Bug | Mô tả | Fix |
|---|---|---|
| BUG-01 | `time = datetime.now()` shadow module `time` → `NameError` ở `time.sleep()` | Đổi thành `now_time` |
| BUG-02 | `sync_time_devices` không `disconnect()` khi exception → resource leak | Thêm `finally: conn.disconnect()` |
| BUG-03 | `enable_device()` không gọi khi exception → máy chấm công bị disable mãi | Chuyển vào `finally` block |
| BUG-04 | `enable_print` chỉ khai báo trong `__main__` → `NameError` khi import | Khai báo `enable_print = False` ở top-level |

## Cải tiến hiệu năng

| ID | Vấn đề | Cải tiến |
|---|---|---|
| PERF-01 | Linear search O(n) mỗi lần tra cứu nhân viên | Dict lookup O(1) |
| PERF-02 | N round-trips MongoDB cho mỗi row Excel | `bulk_write()` 1 lần |
| PERF-03 | `list.__contains__()` O(n) cho OT dedup | `set` O(1) |
| PERF-04 | Scan toàn bộ `HistoryGetAttLogs` | `find_one({"machine": N})` |
| PERF-05 | pyzbar decode đơn lẻ thất bại với QR mờ / nhỏ | Multi-pipeline + zxingcpp fallback |
