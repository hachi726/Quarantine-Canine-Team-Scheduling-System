from calendar import monthrange
from collections import defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path

from openpyxl import load_workbook
from ortools.sat.python import cp_model

TIME_LIMIT_SECONDS = 60
NEXT_DAY_MIN_AVAILABLE = 5
PASS, FAIL, PENDING = "通過", "違規", "待確認"

def as_date(value):
    if isinstance(value, datetime): return value.date()
    if isinstance(value, date): return value
    if isinstance(value, str):
        try: return date.fromisoformat(value.strip())
        except ValueError: return None
    return None

def text_value(value):
    return "" if value is None else str(value).strip()

def is_night_text(s):
    return "大夜" in s or "夜班" in s or "20:30" in s

def is_late_text(s):
    return "晚班" in s or "15:00" in s or s == "晚"

def shift_workload(shifts):
    result = 0
    for shift in shifts:
        s = text_value(shift)
        if is_night_text(s): result = max(result, 2)
        elif "早班" in s or is_late_text(s) or s == "早" or "09:00" in s: result = max(result, 1)
    return result

def run_solver(current_file_path: Path, previous_file_path: Path = None):
    wb = load_workbook(current_file_path)
    ws_settings = wb["設定"]
    settings = {text_value(ws_settings.cell(r, 1).value): ws_settings.cell(r, 2).value for r in range(2, ws_settings.max_row + 1) if ws_settings.cell(r, 1).value is not None}
    
    year, month = int(settings["排班年份"]), int(settings["排班月份"])
    default_basic = int(settings.get("每人每月基本休假天數", 9))
    default_holiday = int(settings.get("每人每月國定假日休假天數", 0))
    cycle_min = int(settings.get("每14天基本休假下限", 4))
    cycle_anchor = as_date(settings.get("週期起始日"))
    terminal_nights = int(settings.get("每月航廈大夜班數", 10))
    cargo_nights = int(settings.get("每月貨運大夜班數", 1))

    first_day = date(year, month, 1)
    last_day = date(year, month, monthrange(year, month)[1])
    days = [first_day + timedelta(days=i) for i in range((last_day - first_day).days + 1)]
    day_count = len(days)

    ws_people = wb["人員"]
    people = []
    group_counters = {"B": 0, "C": 0}
    for row in range(2, ws_people.max_row + 1):
        name = text_value(ws_people.cell(row, 1).value)
        group = text_value(ws_people.cell(row, 2).value).upper()
        if not name and not group: continue
        basic_days = int(ws_people.cell(row, 3).value) if ws_people.cell(row, 3).value is not None else default_basic
        holiday_days = int(ws_people.cell(row, 4).value) if ws_people.cell(row, 4).value is not None else default_holiday
        
        start_raw = ws_people.cell(row, 5).value
        start_idx = 0
        if start_raw not in (None, ""):
            sd = as_date(start_raw)
            start_idx = (sd - first_day).days if sd > first_day else 0 if sd is not None else max(0, int(start_raw) - 1)
        
        group_counters[group] += 1
        current_id = f"{group}{group_counters[group]}"
        
        # 【除錯日誌】印出每個人的起始上班日讀取狀況
        print(f"讀取人員: {name} (ID: {current_id}) | 第5欄原始值: {repr(start_raw)} | 計算出的 start_idx: {start_idx}")
        
        people.append({"name": name, "group": group, "id": current_id, "basic_days": basic_days, "holiday_days": holiday_days, "start_idx": start_idx})

    name_set = {p["name"] for p in people}
    requested_off = defaultdict(set)
    if "休假需求" in wb.sheetnames:
        ws_requests = wb["休假需求"]
        for row in range(2, ws_requests.max_row + 1):
            req_name = text_value(ws_requests.cell(row, 1).value)
            req_value = ws_requests.cell(row, 2).value
            if not req_name or req_name not in name_set or req_value is None: continue
            values = req_value if isinstance(req_value, (list, tuple)) else [req_value]
            for value in values:
                if isinstance(value, (datetime, date)):
                    if value.year == year and value.month == month: requested_off[req_name].add(value.day)
                else:
                    for part in str(value).replace("、", ",").split(","):
                        if part.strip():
                            try: requested_off[req_name].add(int(part.strip()))
                            except ValueError: pass

    n = len(people)
    
    # 自動讀取上個月排班結果
    known = defaultdict(set)
    if previous_file_path and Path(previous_file_path).exists():
        try:
            prev_wb = load_workbook(previous_file_path, data_only=True)
            if "排班結果" in prev_wb.sheetnames:
                prev_ws = prev_wb["排班結果"]
                header_row = [cell.value for cell in prev_ws[1]]
                date_col_indices = []
                for idx, h in enumerate(header_row):
                    if h and "/" in str(h) and "\n" in str(h):
                        date_col_indices.append(idx)
                
                target_col_indices = date_col_indices[-5:] if len(date_col_indices) >= 5 else date_col_indices
                
                for row in range(2, prev_ws.max_row + 1):
                    name = text_value(prev_ws.cell(row, 1).value)
                    if not name: continue
                    
                    for col_idx in target_col_indices:
                        header_str = str(header_row[col_idx])
                        date_part = header_str.split("\n")[0].strip()
                        m_str, d_str = date_part.split("/")
                        
                        prev_month = month - 1 if month > 1 else 12
                        prev_year = year if month > 1 else year - 1
                        
                        try:
                            target_date = date(prev_year, int(m_str), int(d_str))
                            shift_val = text_value(prev_ws.cell(row, col_idx + 1).value)
                            if shift_val:
                                known[(target_date, name)].add(shift_val)
                        except ValueError:
                            pass
                print("✅ 已成功自動載入上個月的跨月排班結果！")
        except Exception as e:
            print(f"⚠️ 自動讀取上個月檔案發生錯誤：{e}")

    if "月外已知班表" in wb.sheetnames:
        ws_outside = wb["月外已知班表"]
        for row in range(2, ws_outside.max_row + 1):
            d, name, shift = as_date(ws_outside.cell(row, 1).value), text_value(ws_outside.cell(row, 2).value), text_value(ws_outside.cell(row, 3).value)
            if d and name and shift: known[(d, name)].add(shift)

    def known_flag(name, offset, test):
        return any(test(text_value(s)) for s in known.get((first_day - timedelta(days=offset), name), set()))

    model = cp_model.CpModel()
    states = ["basic", "holiday", "early", "late", "night", "postnight", "away"]
    x = {(p, d, s): model.NewBoolVar(f"x_{p}_{d}_{s}") for p in range(n) for d in range(day_count) for s in states}
    night_event = {d: model.NewBoolVar(f"night_event_{d}") for d in range(day_count)}
    def rest(p, d): return x[p, d, "basic"] + x[p, d, "holiday"]

    penalty_vars, wk_viol_vars = [], []
    for p in range(n):
        name, group = people[p]["name"], people[p]["group"]
        for d in range(day_count):
            model.Add(sum(x[p, d, s] for s in states) == 1)
            first_half = days[d].day <= 15
            if not ((group == "B" and first_half) or (group == "C" and not first_half)): model.Add(x[p, d, "early"] == 0)
            if not ((group == "C" and first_half) or (group == "B" and not first_half)): model.Add(x[p, d, "late"] == 0)
            if d < people[p]["start_idx"]: model.Add(x[p, d, "basic" if days[d].weekday() >= 5 else "away"] == 1)
            else: model.Add(x[p, d, "away"] == 0)

            if d > 0: model.Add(x[p, d, "postnight"] == x[p, d - 1, "night"])
            else: model.Add(x[p, 0, "postnight"] == (1 if known_flag(name, 1, is_night_text) else 0))
            if d >= 2: model.Add(x[p, d - 2, "night"] <= rest(p, d))
            if d > 0: model.Add(x[p, d - 1, "late"] + x[p, d, "early"] <= 1)

        if known_flag(name, 1, is_night_text) and day_count > 1: model.Add(rest(p, 1) == 1)
        if known_flag(name, 2, is_night_text): model.Add(rest(p, 0) == 1)
        if known_flag(name, 1, is_late_text):
            model.Add(x[p, 0, "early"] == 0)
            if group == "B": model.Add(rest(p, 0) == 1)
        if group == "C" and day_count >= 16: model.Add(x[p, 14, "late"] <= rest(p, 15))

        for rd in requested_off.get(name, set()):
            if 1 <= rd <= day_count and rd - 1 >= people[p]["start_idx"]: model.Add(rest(p, rd - 1) == 1)

        model.Add(sum(x[p, d, "basic"] for d in range(day_count)) == people[p]["basic_days"])
        model.Add(sum(x[p, d, "holiday"] for d in range(day_count)) == people[p]["holiday_days"])
        nights = sum(x[p, d, "night"] for d in range(day_count))
        model.Add(nights >= (0 if people[p]["start_idx"] > 0 else 1))
        model.Add(nights <= 2)

        weekend_idx = [d for d in range(day_count) if days[d].weekday() >= 5 and d >= people[p]["start_idx"]]
        if weekend_idx:
            wk_viol = model.NewIntVar(0, day_count, f"wk_viol_{p}")
            model.Add(sum(rest(p, d) for d in weekend_idx) - wk_viol <= len(weekend_idx) // 2)
            wk_viol_vars.append((p, wk_viol))
            penalty_vars.append(wk_viol)
        else:
            wk_viol_vars.append((p, None))

    for d in range(day_count):
        model.Add(sum(x[p, d, "early"] for p in range(n)) >= 4)
        model.Add(sum(x[p, d, "early"] for p in range(n)) <= 7)
        model.Add(sum(x[p, d, "late"] for p in range(n)) >= 4)
        model.Add(sum(x[p, d, "late"] for p in range(n)) <= 7)
        model.Add(sum(x[p, d, "night"] for p in range(n)) == 2 * night_event[d])

    model.Add(night_event[day_count - 1] == 0)
    model.Add(night_event[day_count - 2] == 0)
    model.Add(sum(night_event[d] for d in range(day_count)) == terminal_nights + cargo_nights)

    for p in range(n):
        for start_offset in range(-5, day_count - 5):
            window_dates = [first_day + timedelta(days=start_offset + k) for k in range(6)]
            outside_load = sum(shift_workload(known.get((od, people[p]["name"]), set())) for od in window_dates if od < first_day or od > last_day)
            current_load = [x[p, (dt - first_day).days, "early"] + x[p, (dt - first_day).days, "late"] + 2 * x[p, (dt - first_day).days, "night"] for dt in window_dates if first_day <= dt <= last_day]
            wl_viol = model.NewIntVar(0, 10, f"wl_viol_{p}_{start_offset}")
            model.Add(sum(current_load) + outside_load - wl_viol <= 5)
            penalty_vars.append(wl_viol)

    cycle_start = cycle_anchor
    while cycle_start + timedelta(days=14) <= first_day: cycle_start += timedelta(days=14)
    while cycle_start <= last_day:
        if cycle_start >= first_day and cycle_start + timedelta(days=13) <= last_day:
            indices = [(cycle_start - first_day).days + i for i in range(14)]
            for p in range(n):
                cy_viol = model.NewIntVar(0, 14, f"cy_viol_{p}_{cycle_start}")
                model.Add(sum(x[p, d, "basic"] for d in indices) + cy_viol >= cycle_min)
                penalty_vars.append(cy_viol)
        cycle_start += timedelta(days=14)

    model.Maximize(sum(x[p, d, s] for p in range(n) for d in range(day_count) for s in ("early", "late")) - 100 * sum(penalty_vars))
    
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = TIME_LIMIT_SECONDS
    solver.parameters.num_search_workers = 8

    print(f"開始排班：{year} 年 {month} 月，共 {n} 人。")
    status = solver.Solve(model)
    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        print("無法產生符合條件的班表。")
    else:
        if penalty_vars:
            print("\n===== 衝突診斷報告 =====")
            for p, wk_viol in wk_viol_vars:
                if wk_viol is not None and solver.Value(wk_viol) > 0:
                    print(f"[週末違規] {people[p]['name']} ({people[p]['id']}) 多休了 {solver.Value(wk_viol)} 天")
            print("======================\n")

        assignment = {}
        night_dates = []
        for d in range(day_count):
            for p in range(n):
                for s in states:
                    if solver.Value(x[p, d, s]): assignment[p, d] = s; break
            if solver.Value(night_event[d]): night_dates.append(days[d])

        night_types = {d: "航廈" for d in night_dates}
        free_dates = list(night_dates)
        for k in range(cargo_nights):
            target = day_count * (k + 1) / (cargo_nights + 1)
            pick = min(free_dates, key=lambda dt: abs((dt - first_day).days + 1 - target))
            night_types[pick] = "貨運"
            free_dates.remove(pick)

        role_model = cp_model.CpModel()
        role_vars = {}
        for d in range(day_count):
            for shift in ("early", "late"):
                assigned = [p for p in range(n) if assignment[p, d] == shift]
                count = len(assigned)
                for p in assigned:
                    for r in ("T1", "T2", "統派"): role_vars[p, d, r] = role_model.NewBoolVar(f"role_{p}_{d}_{r}")
                    role_model.Add(sum(role_vars[p, d, r] for r in ("T1", "T2", "統派")) == 1)
                role_model.Add(sum(role_vars[p, d, "T1"] for p in assigned) == 2)
                role_model.Add(sum(role_vars[p, d, "T2"] for p in assigned) == (2 if count == 4 else 3))
                role_model.Add(sum(role_vars[p, d, "統派"] for p in assigned) == max(0, count - 5))

        role_solver = cp_model.CpSolver()
        role_solver.Solve(role_model)

        ws_result = wb["排班結果"]
        ws_result.delete_rows(1, ws_result.max_row)
        weekday_names = ["一", "二", "三", "四", "五", "六", "日"]
        ws_result.append(["姓名", "組別", "編號"] + [f"{d.month}/{d.day}\n{weekday_names[d.weekday()]}" for d in days])
        
        for p in range(n):
            row_vals = [people[p]["name"], people[p]["group"], people[p]["id"]]
            for d in range(day_count):
                state = assignment[p, d]
                if state in ("early", "late"):
                    role = next((r for r in ("T1", "T2", "統派") if role_solver.Value(role_vars.get((p, d, r), 0))), "")
                    row_vals.append(f"{'早班' if state=='early' else '晚班'}\n{role}")
                elif state == "night": row_vals.append("大夜" + night_types[days[d]])
                elif state == "postnight": row_vals.append("大夜後休")
                elif state == "away": row_vals.append("轉訓")
                else: row_vals.append("休假")
            ws_result.append(row_vals)

        wb.save(current_file_path)
        print("🎉 排班完成並已成功寫入 Excel！")