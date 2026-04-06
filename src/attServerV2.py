# -*- coding: utf-8 -*-
"""
attServerV2.py — CLI entry point for Attendance Server V2.
All business logic lives in functions.py (shared with attServerV2UI.py).
"""
import threading
import time as time_module
import schedule
from zk import ZK
from datetime import datetime

from functions import (
    read_config,
    update_excel_to_mongoDb,
    get_att_log,
    ot_register_detect_qr_and_save,
    setup_schedule,
    write_log_info,
    ip_att_machines,
    schedule_config,
    REAL_TIME,
)

if __name__ == "__main__":
    import functions
    functions.enable_print = True

    read_config()
    sc = schedule_config
    startup_msg = (
        '=' * 80 + '\n'
        '*** ATTENDANCE SERVER V2 STARTUP ***\n'
        f'START AT : {datetime.now():%d-%m-%Y %H:%M:%S}\n'
        'Schedules:\n'
        f'  Sync time  : {sc.get("sync_time_day")} {sc.get("sync_time_at")} → sync_time_devices\n'
        f'  Excel sync : {sc.get("excel_sync_times")} → update_excel_to_mongoDb\n'
        f'  OT scan    : every {sc.get("ot_scan_interval_minutes")} min → ot_register_detect_qr_and_save\n'
        f'  Att log    : every {sc.get("att_log_interval_minutes")} min per machine\n'
        + '=' * 80
    )
    print(startup_msg)
    write_log_info(startup_msg, 'MAIN')

    update_excel_to_mongoDb()

    machine_no = 1
    for ip in ip_att_machines:
        machine = ZK(ip, port=4370, timeout=5, password=0, force_udp=False, ommit_ping=False)
        threading.Thread(
            target=get_att_log,
            args=(machine, machine_no, REAL_TIME),
            name=f'att-machine-{ip}',
            daemon=True
        ).start()
        machine_no += 1

    ot_register_detect_qr_and_save()
    setup_schedule()

    while True:
        schedule.run_pending()
        time_module.sleep(1)
