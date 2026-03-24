# -*- coding: utf-8 -*-
"""
Haftalık Makine Dağılımı (GÜNLÜK KAPASİTE, 1+2. VARDİYA) - STREAMLIT REVİZE v4

✅ Ekler (senin isteklerin):
1) Streamlit: Öncelikli SKU seçimi (seçilen SKU'lar plan sırasının en başına alınır)
2) 2. vardiya için ayrı Gün / Makine / Kapasite seçimi
3) 1. Vardiya ve 2. Vardiya ekranları default açık (expanded=True)

✅ Patch (son istek):
- Öncelikli SKU'lar takvimde de öne gelsin (Pazartesi’ye yasla)
- Makinede boş gün bırakma (SIVI DOLUM 3 Salı boş kalmasın)

Temel Kurallar:
- Aynı ürün hafta boyunca tek makinede çalışır (1+2 vardiya)
- 1. vardiya: ardışık günler
- 2. vardiya:
    * sadece SHIFT2 günlerinde
    * sadece SHIFT2 makinelerinde
    * 1. vardiyada o makine+gün slotu TAM doluysa açılır (used1 >= cap1)
    * 2. vardiya kapasitesi > 0 olmalı
- 2. vardiyada forbidden makineler ve pair kuralları geçerli
"""

import os
from collections import Counter
import pandas as pd
from datetime import datetime
import streamlit as st

today = datetime.today().strftime("%Y-%m-%d")

# ==== Dosya adları ====
PLAN_XLSX = "Üretim Planı.xlsx"
MAP_XLSX = "Üretim Plan Ayrıştırma Data.xlsx"
OUTPUT_XLSX = f"haftalik_makine_dagilimi_{today}.xlsx"

# ==== MASTER Günler (sabit sıra) ====
MASTER_DAYS = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "Cumartesi", "Pazar"]
MASTER_DAY_INDEX = {d: i for i, d in enumerate(MASTER_DAYS)}

# ==== Varsayılan kapasiteler (vardiya başına) ====
DEFAULT_CAPS_SHIFT1 = {
    "KOLONYA": 2000,
    "Makine 1": 4000,
    "Makine 2": 3000,
    "Makine 3": 5000,
    "Makine 4": 5500,
    "Makine 5": 2500,
    "Makine 6": 2500,
    "Makine 7": 2500,
    "Makine 8": 2500,
    "Makine 1-1": 12000,
    "Makine 2-1": 12000,
    "Makine 3-1": 12000,
    "Makine 4-1": 10000,
    "SIVI DOLUM 1": 5000,
    "SIVI DOLUM 2": 2000,
    "SIVI DOLUM 3": 5000,
    "ASO - 1": 3000,
    "ASO - 2": 3000,
}
DEFAULT_CAPS_SHIFT2 = dict(DEFAULT_CAPS_SHIFT1)

# Öncelikli makineler (makine seçimi sırasında öncelik) - 1. vardiya için
DEFAULT_PRIORITY_MACHINES = ["Makine 1-1", "Makine 2-1", "Makine 3-1", "Makine 4-1"]

# 2. vardiyada KULLANILMAYACAK makineler
DEFAULT_SECOND_SHIFT_FORBIDDEN = {"Makine 1", "Makine 3", "Makine 4"}

# 2. vardiyada özel ilişki grupları
DEFAULT_PAIR_12 = {"Makine 1-1", "Makine 2-1"}
DEFAULT_PAIR_34 = {"Makine 3-1", "Makine 4-1"}

# ===================== GLOBAL (UI ile set edilecek) =====================
SHIFT1_DAYS = MASTER_DAYS[:]     # UI ile değişecek
SHIFT2_DAYS = MASTER_DAYS[:]     # UI ile değişecek
SHIFT1_MACHINES = sorted(DEFAULT_CAPS_SHIFT1.keys())
SHIFT2_MACHINES = sorted(DEFAULT_CAPS_SHIFT2.keys())

CAPACITIES_SHIFT1 = dict(DEFAULT_CAPS_SHIFT1)
CAPACITIES_SHIFT2 = dict(DEFAULT_CAPS_SHIFT2)

PRIORITY_MACHINES = list(DEFAULT_PRIORITY_MACHINES)
SECOND_SHIFT_FORBIDDEN = set(DEFAULT_SECOND_SHIFT_FORBIDDEN)
PAIR_12 = set(DEFAULT_PAIR_12)
PAIR_34 = set(DEFAULT_PAIR_34)

PRIORITY_SKU_ORDER = []       # UI ile seçilecek (seçim sırası korunur)
PRIORITY_SKUS = set()            # UI ile seçilecek


# ===================== Yardımcılar =====================

def _union_machines():
    return sorted(set(SHIFT1_MACHINES) | set(SHIFT2_MACHINES))


def _union_days_in_master_order():
    return [d for d in MASTER_DAYS if (d in SHIFT1_DAYS) or (d in SHIFT2_DAYS)]


def read_plan(path=PLAN_XLSX):
    if not os.path.exists(path):
        raise FileNotFoundError(f"Plan dosyası bulunamadı: {path}")
    plan = pd.read_excel(path, sheet_name="Sheet1")
    plan.columns = [str(c).strip() for c in plan.columns]

    req_plan = ["Ürün Kodu", "Ürün Adı", "Tür", "Üretim Planı", "Ml", "Hammadde Kodu"]
    miss = [c for c in req_plan if c not in plan.columns]
    if miss:
        raise ValueError(f"Üretim Planı sheet eksik kolonlar: {miss}")

    plan["Ürün Kodu"] = plan["Ürün Kodu"].astype(str).str.strip()
    plan["Ürün Adı"] = plan["Ürün Adı"].astype(str).str.strip()
    plan["Tür"] = plan["Tür"].astype(str).str.strip()
    plan["Hammadde Kodu"] = plan["Hammadde Kodu"].astype(str).str.strip()
    plan["Üretim Planı"] = pd.to_numeric(plan["Üretim Planı"], errors="coerce").fillna(0).astype(int)

    plan = plan[(plan["Ürün Kodu"] != "") & (plan["Üretim Planı"] > 0)].copy()
    return plan


def read_mapping(path=MAP_XLSX):
    if not os.path.exists(path):
        raise FileNotFoundError(f"Mapping dosyası bulunamadı: {path}")
    mapping = pd.read_excel(path, sheet_name="DATA")
    mapping.columns = [str(c).strip() for c in mapping.columns]

    req_map = ["Makine Adı", "Ürün Kodu"]
    miss = [c for c in req_map if c not in mapping.columns]
    if miss:
        raise ValueError(f"Mapping sheet eksik kolonlar: {miss}")

    mapping["Makine Adı"] = mapping["Makine Adı"].astype(str).str.strip()
    mapping["Ürün Kodu"] = mapping["Ürün Kodu"].astype(str).str.strip()

    allowed_machines = set(_union_machines())
    mapping = mapping[mapping["Makine Adı"].isin(allowed_machines)].copy()

    prod_to_machines = (
        mapping.groupby("Ürün Kodu")["Makine Adı"]
        .apply(lambda s: sorted(set(s)))
        .to_dict()
    )
    return prod_to_machines


def sort_plan_with_priority(plan: pd.DataFrame) -> pd.DataFrame:
    """Öncelikli SKU'lar en üste; öncelikli SKU'ların kendi içinde seçim sırası korunur.
    Sonrasında diğerleri Üretim Planı büyükten küçüğe.
    """
    prio_order = [str(x).strip() for x in (PRIORITY_SKU_ORDER or []) if str(x).strip()]
    tmp = plan.copy()

    # 1) Öncelikli olanlar (seçim sırası)
    by_code = {str(r["Ürün Kodu"]).strip(): i for i, r in tmp.iterrows()}
    prio_rows = []
    used_codes = set()
    for c in prio_order:
        if c in by_code:
            prio_rows.append(tmp.loc[by_code[c]])
            used_codes.add(c)

    # 2) Kalanlar (Üretim Planı desc)
    rest = tmp[~tmp["Ürün Kodu"].astype(str).str.strip().isin(used_codes)].copy()
    rest = rest.sort_values(["Üretim Planı"], ascending=[False])

    out = pd.concat([pd.DataFrame(prio_rows), rest], ignore_index=True)
    out = out.reset_index(drop=True)
    return out


def choose_product_machine(pcode, machines, remaining_cap_shift1, is_priority=False):
    """
    1. vardiyada ürün için sabit makine seçimi.

    Normal SKU:
    - Sadece SHIFT1_MACHINES içindeki makinelerden seçer
    - SHIFT1 günlerinde haftalık toplam kalan kapasitesi > 0 olanlar viable
    - PRIORITY_MACHINES içindekilere öncelik
    - En yüksek haftalık kapasiteyi seçer

    Öncelikli SKU (is_priority=True):
    - Önce en erken başlayabildiği makineyi seç (MASTER gün sırasına göre; mümkünse Pazartesi)
    - Eşitlikte PRIORITY_MACHINES öncelikli
    - Sonra o en erken gündeki kalan kapasite (yüksek olan) ve haftalık kapasite ile tie-break
    """
    machines = [m for m in machines if m in SHIFT1_MACHINES]
    viable = []
    for m in machines:
        # Bu makinede ilk kez kapasite olan gün (MASTER sırası ile)
        first_day_idx = None
        first_day_cap = 0
        weekly_cap = 0
        for d in SHIFT1_DAYS:
            c = int(remaining_cap_shift1.get(m, {}).get(d, 0))
            weekly_cap += c
            if first_day_idx is None and c > 0:
                first_day_idx = MASTER_DAY_INDEX.get(d, 10**9)
                first_day_cap = c
        if weekly_cap > 0 and first_day_idx is not None:
            viable.append((m, first_day_idx, first_day_cap, weekly_cap))

    if not viable:
        return None

    if is_priority:
        # sort: earliest day, prefer priority machines, higher first_day_cap, higher weekly_cap
        viable.sort(key=lambda x: (x[1], 0 if x[0] in PRIORITY_MACHINES else 1, -x[2], -x[3]))
        return viable[0][0]

    # Normal davranış
    prio = [x for x in viable if x[0] in PRIORITY_MACHINES]
    non_prio = [x for x in viable if x[0] not in PRIORITY_MACHINES]
    if prio:
        return max(prio, key=lambda x: x[3])[0]
    return max(non_prio, key=lambda x: x[3])[0]


# ===================== PATCH'li 1. Vardiya Dağıtımı =====================

def distribute_product_shift1(
    product_row,
    prod_to_machines,
    remaining_cap_shift1,
    remaining_cap_shift2,
    day_second_shift_machines,
    day_pair_primary_12,
    day_pair_primary_34,
    day_totals,
    schedule_rows,
    product_machine,
    product_first_day_master_idx,
    product_last_day_master_idx,
):
    """
    ✅ Revize kural:
    1) Ürün bir makineye alındı mı -> ürün bitene kadar araya SKU sokma.
    2) Sıralama: 1. vardiya (aynı gün) -> 2. vardiya (aynı gün, varsa) -> ertesi gün 1. vardiya.
    3) 2. vardiya kısıtları:
       - gün SHIFT2_DAYS içinde olmalı
       - makine SHIFT2_MACHINES içinde olmalı
       - forbidden makineler kullanılmaz
       - gün başına max 3 makine
       - pair kuralları (PAIR_12, PAIR_34)
    """
    pcode = product_row["Ürün Kodu"]
    pname = product_row["Ürün Adı"]
    ptype = product_row["Tür"]
    hm = product_row["Hammadde Kodu"]
    qty_left = int(product_row["Üretim Planı"])

    machines = prod_to_machines.get(pcode, [])
    if not machines or qty_left <= 0:
        return qty_left

    # Ürün için sabit makine seç (hafta boyu)
    if pcode in product_machine:
        chosen_machine = product_machine[pcode]
    else:
        chosen_machine = choose_product_machine(pcode, machines, remaining_cap_shift1, is_priority=(pcode in PRIORITY_SKUS))
        if chosen_machine is None:
            return qty_left
        product_machine[pcode] = chosen_machine

    # Bu makinede 1. vardiyada kapasitesi olan günler
    days_with_cap = [
        d for d in SHIFT1_DAYS
        if remaining_cap_shift1.get(chosen_machine, {}).get(d, 0) > 0
    ]
    if not days_with_cap:
        return qty_left

    # ✅ START DAY: en erken gün (priority/normal farkı burada yok; hepsi öne yaslı)
    start_day = min(days_with_cap, key=lambda d: MASTER_DAY_INDEX[d])
    start_idx = SHIFT1_DAYS.index(start_day)

    # ✅ Gün gün ilerle: SHIFT1 -> aynı gün SHIFT2 -> sonraki gün SHIFT1
    for i in range(start_idx, len(SHIFT1_DAYS)):
        if qty_left <= 0:
            break

        day = SHIFT1_DAYS[i]

        # --- 1. VARDIYA ---
        cap1_rem = remaining_cap_shift1.get(chosen_machine, {}).get(day, 0)
        if cap1_rem > 0 and qty_left > 0:
            assign1 = min(qty_left, cap1_rem)
            schedule_rows.append(
                {
                    "Gün": day,
                    "Makine": chosen_machine,
                    "Ürün Kodu": pcode,
                    "Ürün Adı": pname,
                    "Tür": ptype,
                    "Hammadde Kodu": hm,
                    "Adet": assign1,
                    "Not": "",
                }
            )
            master_idx = MASTER_DAY_INDEX[day]
            if pcode not in product_first_day_master_idx:
                product_first_day_master_idx[pcode] = master_idx
            product_last_day_master_idx[pcode] = master_idx

            remaining_cap_shift1[chosen_machine][day] -= assign1
            day_totals[day] = day_totals.get(day, 0) + assign1
            qty_left -= assign1

        # --- 2. VARDIYA (aynı gün devam) ---
        if qty_left > 0 and (day in SHIFT2_DAYS) and (chosen_machine in SHIFT2_MACHINES):
            if chosen_machine not in SECOND_SHIFT_FORBIDDEN:
                # Gün başına max 3 makine
                used_machines_today = day_second_shift_machines[day]
                if chosen_machine in used_machines_today or len(used_machines_today) < 3:
                    # Pair kuralı
                    pair_ok = True
                    if chosen_machine in PAIR_12:
                        primary = day_pair_primary_12[day]
                        if primary is None:
                            day_pair_primary_12[day] = chosen_machine
                        elif primary != chosen_machine and remaining_cap_shift2.get(primary, {}).get(day, 0) > 0:
                            pair_ok = False
                    elif chosen_machine in PAIR_34:
                        primary = day_pair_primary_34[day]
                        if primary is None:
                            day_pair_primary_34[day] = chosen_machine
                        elif primary != chosen_machine and remaining_cap_shift2.get(primary, {}).get(day, 0) > 0:
                            pair_ok = False

                    if pair_ok:
                        cap2_rem = remaining_cap_shift2.get(chosen_machine, {}).get(day, 0)
                        if cap2_rem and cap2_rem > 0:
                            assign2 = min(qty_left, cap2_rem)
                            schedule_rows.append(
                                {
                                    "Gün": day,
                                    "Makine": chosen_machine,
                                    "Ürün Kodu": pcode,
                                    "Ürün Adı": pname,
                                    "Tür": ptype,
                                    "Hammadde Kodu": hm,
                                    "Adet": assign2,
                                    "Not": "2.VARDIYA",
                                }
                            )
                            remaining_cap_shift2[chosen_machine][day] -= assign2
                            qty_left -= assign2
                            used_machines_today.add(chosen_machine)

    return qty_left

def simulate_machine_slot_schedule(plan_sorted: pd.DataFrame, prod_to_machines: dict):
    """
    Makine bazlı kuyruk + slot doldurma (non-preemptive) motoru.

    Kurallar:
    1) Makineye giren ürün bitmeden çıkmaz (queue head).
    2) Vardiya slotu kapasite dolana kadar çalışır; ürün biterse aynı slot içinde sıradaki ürüne geçilir.
       (1. vardiya ve 2. vardiya için aynı)

    2. vardiya kısıtları (korunur):
    - gün SHIFT2_DAYS içinde olmalı
    - makine SHIFT2_MACHINES içinde olmalı
    - SECOND_SHIFT_FORBIDDEN kullanılmaz
    - gün başına max 3 makine
    - pair kuralları (PAIR_12 / PAIR_34)
    """

    machines_union = _union_machines()
    days_union = _union_days_in_master_order()

    # ---------- 0) Ürün -> sabit makine atama (shadow capacity ile) ----------
    remaining_cap_shift1_shadow = {
        m: {d: int(CAPACITIES_SHIFT1.get(m, 0)) for d in SHIFT1_DAYS}
        for m in SHIFT1_MACHINES
    }

    product_machine = {}
    for _, r in plan_sorted.iterrows():
        pcode = str(r["Ürün Kodu"]).strip()
        qty = int(r.get("Üretim Planı", 0) or 0)
        machines = prod_to_machines.get(pcode, [])
        if not machines or qty <= 0:
            continue

        m = choose_product_machine(pcode, machines, remaining_cap_shift1_shadow, is_priority=(pcode in PRIORITY_SKUS))
        if m is None:
            candidates = [x for x in machines if x in SHIFT1_MACHINES]
            m = candidates[0] if candidates else None
        if m is None:
            continue

        product_machine[pcode] = m

        # ✅ Shadow capacity tüket: bu SKU'nun talebini seçilen makinenin 1. vardiya haftalık kapasitesinden düş
        shadow_qty = qty
        if shadow_qty > 0 and m in remaining_cap_shift1_shadow and SHIFT1_DAYS:
            for d in SHIFT1_DAYS:
                avail = int(remaining_cap_shift1_shadow[m].get(d, 0))
                if avail <= 0:
                    continue
                take = shadow_qty if shadow_qty < avail else avail
                remaining_cap_shift1_shadow[m][d] = avail - take
                shadow_qty -= take
                if shadow_qty <= 0:
                    break

    # ---------- 1) Makine kuyruklarını oluştur ----------
    queue = {m: [] for m in machines_union}          # makine -> [pcode, pcode, ...]
    remaining_qty = {}                               # pcode -> qty_left
    row_by_code = {str(r["Ürün Kodu"]).strip(): r for _, r in plan_sorted.iterrows()}
    order_idx = {c: i for i, c in enumerate(plan_sorted["Ürün Kodu"].astype(str).tolist())}

    for _, r in plan_sorted.iterrows():
        pcode = str(r["Ürün Kodu"]).strip()
        remaining_qty[pcode] = int(r.get("Üretim Planı", 0) or 0)
        m = product_machine.get(pcode)
        if m is None:
            continue
        queue[m].append(pcode)

    # ---------- 2) REBALANCE (boş kapasite varken overload'ı azalt) ----------
    def machine_week_capacity(m: str) -> int:
        cap = 0
        if m in SHIFT1_MACHINES:
            cap += int(CAPACITIES_SHIFT1.get(m, 0)) * len(SHIFT1_DAYS)
        if m in SHIFT2_MACHINES:
            cap += int(CAPACITIES_SHIFT2.get(m, 0)) * len(SHIFT2_DAYS)
        return int(cap)

    cap_total = {m: machine_week_capacity(m) for m in machines_union}
    assigned_total = {m: 0 for m in machines_union}
    for m in machines_union:
        assigned_total[m] = sum(int(remaining_qty.get(pc, 0)) for pc in queue.get(m, []))

    def slack(m: str) -> int:
        return int(cap_total.get(m, 0)) - int(assigned_total.get(m, 0))

    moved_any = True
    safety = 0
    while moved_any and safety < 5000:
        safety += 1
        moved_any = False

        overloads = [m for m in machines_union if slack(m) < 0 and queue.get(m)]
        if not overloads:
            break

        overloads.sort(key=lambda m: slack(m))  # en negatif slack önce

        for src_m in overloads:
            if slack(src_m) >= 0 or not queue.get(src_m):
                continue

            # kuyruk sonundan dene (düşük öncelik)
            pc = queue[src_m][-1]
            qty = int(remaining_qty.get(pc, 0))
            if qty <= 0:
                queue[src_m].pop()
                continue

            # alternatif makineler: aynı SKU'nun mapping'inde olan ve SHIFT1_MACHINES içinde olan makineler
            candidates = [
                m for m in prod_to_machines.get(pc, [])
                if m in machines_union and m != src_m and m in SHIFT1_MACHINES
            ]
            if not candidates:
                continue

            # en boş olana taşı (slack en yüksek)
            candidates.sort(key=lambda m: slack(m), reverse=True)
            dst_m = None
            for m in candidates:
                if slack(m) > 0:
                    dst_m = m
                    break
            if dst_m is None:
                continue

            # taşı
            queue[src_m].pop()
            queue[dst_m].append(pc)
            queue[dst_m].sort(key=lambda c: order_idx.get(c, 10**9))

            product_machine[pc] = dst_m
            assigned_total[src_m] -= qty
            assigned_total[dst_m] += qty
            moved_any = True

    # ---------- 3) Slot simülasyonu ----------
    day_second_shift_machines = {d: set() for d in SHIFT2_DAYS}
    day_pair_primary_12 = {d: None for d in SHIFT2_DAYS}
    day_pair_primary_34 = {d: None for d in SHIFT2_DAYS}

    def _machine_has_remaining_work(m: str) -> bool:
        for pc in queue.get(m, []):
            if int(remaining_qty.get(pc, 0)) > 0:
                return True
        return False

    def allow_shift2(day: str, machine: str) -> bool:
        if day not in SHIFT2_DAYS:
            return False
        if machine not in SHIFT2_MACHINES:
            return False
        if int(CAPACITIES_SHIFT2.get(machine, 0)) <= 0:
            return False
        if machine in SECOND_SHIFT_FORBIDDEN:
            return False

        used = day_second_shift_machines[day]
        if machine not in used and len(used) >= 3:
            return False

        # Pair kuralı
        if machine in PAIR_12:
            primary = day_pair_primary_12[day]
            if primary is None:
                day_pair_primary_12[day] = machine
            elif primary != machine and _machine_has_remaining_work(primary):
                return False

        if machine in PAIR_34:
            primary = day_pair_primary_34[day]
            if primary is None:
                day_pair_primary_34[day] = machine
            elif primary != machine and _machine_has_remaining_work(primary):
                return False

        return True

    schedule_rows = []
    product_first_day_master_idx = {}
    product_last_day_master_idx = {}
    day_hammadde_owner = {d: {} for d in days_union}  # gün -> hammadde -> makine
    raw_material_block_warnings = set()

    def _clean_hammadde(value) -> str:
        if pd.isna(value):
            return ""
        return str(value).strip()

    def _raw_material_conflicts(day: str, machine: str, pcode: str) -> bool:
        prow = row_by_code.get(pcode)
        if prow is None:
            return False
        hm = _clean_hammadde(prow.get("Hammadde Kodu", ""))
        if not hm:
            return False
        owner = day_hammadde_owner.setdefault(day, {}).get(hm)
        return owner is not None and owner != machine

    def _reserve_raw_material(day: str, machine: str, pcode: str):
        prow = row_by_code.get(pcode)
        if prow is None:
            return
        hm = _clean_hammadde(prow.get("Hammadde Kodu", ""))
        if hm:
            day_hammadde_owner.setdefault(day, {}).setdefault(hm, machine)

    def _remember_raw_material_warning(day: str, machine: str, pcode: str):
        prow = row_by_code.get(pcode)
        if prow is None:
            return
        hm = _clean_hammadde(prow.get("Hammadde Kodu", ""))
        owner = day_hammadde_owner.setdefault(day, {}).get(hm)
        if hm and owner and owner != machine:
            raw_material_block_warnings.add(
                f"{day}: {pcode} ({machine}) planlanamadı; {hm} hammaddesi aynı gün {owner} üzerinde zaten kullanılıyor."
            )

    for day in days_union:
        for machine in machines_union:
            if not queue.get(machine):
                continue

            # ---- 1. VARDIYA ----
            cap1 = int(CAPACITIES_SHIFT1.get(machine, 0)) if (day in SHIFT1_DAYS and machine in SHIFT1_MACHINES) else 0
            cap_rem = cap1
            while cap_rem > 0 and queue[machine]:
                pcode = queue[machine][0]
                left = int(remaining_qty.get(pcode, 0))
                if left <= 0:
                    queue[machine].pop(0)
                    continue

                if _raw_material_conflicts(day, machine, pcode):
                    _remember_raw_material_warning(day, machine, pcode)
                    break

                assign = left if left < cap_rem else cap_rem
                prow = row_by_code.get(pcode)
                if prow is None:
                    # safety
                    queue[machine].pop(0)
                    continue

                _reserve_raw_material(day, machine, pcode)
                schedule_rows.append({
                    "Gün": day,
                    "Makine": machine,
                    "Ürün Kodu": pcode,
                    "Ürün Adı": prow["Ürün Adı"],
                    "Tür": prow["Tür"],
                    "Hammadde Kodu": prow["Hammadde Kodu"],
                    "Adet": int(assign),
                    "Not": "",
                })

                mi = MASTER_DAY_INDEX.get(day, 0)
                if pcode not in product_first_day_master_idx:
                    product_first_day_master_idx[pcode] = mi
                product_last_day_master_idx[pcode] = mi

                remaining_qty[pcode] = left - assign
                cap_rem -= assign
                if remaining_qty[pcode] <= 0:
                    queue[machine].pop(0)

            # ---- 2. VARDIYA ----
            if not allow_shift2(day, machine):
                continue

            cap2 = int(CAPACITIES_SHIFT2.get(machine, 0))
            cap_rem = cap2
            produced_any = False
            while cap_rem > 0 and queue[machine]:
                pcode = queue[machine][0]
                left = int(remaining_qty.get(pcode, 0))
                if left <= 0:
                    queue[machine].pop(0)
                    continue

                if _raw_material_conflicts(day, machine, pcode):
                    _remember_raw_material_warning(day, machine, pcode)
                    break

                assign = left if left < cap_rem else cap_rem
                prow = row_by_code.get(pcode)
                if prow is None:
                    queue[machine].pop(0)
                    continue

                _reserve_raw_material(day, machine, pcode)
                schedule_rows.append({
                    "Gün": day,
                    "Makine": machine,
                    "Ürün Kodu": pcode,
                    "Ürün Adı": prow["Ürün Adı"],
                    "Tür": prow["Tür"],
                    "Hammadde Kodu": prow["Hammadde Kodu"],
                    "Adet": int(assign),
                    "Not": "2.VARDIYA",
                })

                produced_any = True
                remaining_qty[pcode] = left - assign
                cap_rem -= assign
                if remaining_qty[pcode] <= 0:
                    queue[machine].pop(0)

            if produced_any:
                day_second_shift_machines[day].add(machine)

    # ---------- 4) Unassigned & warnings ----------
    unassigned = []
    warnings = []
    for _, r in plan_sorted.iterrows():
        pcode = str(r["Ürün Kodu"]).strip()
        if pcode not in prod_to_machines:
            unassigned.append({
                "Ürün Kodu": pcode,
                "Ürün Adı": r["Ürün Adı"],
                "Tür": r["Tür"],
                "Hammadde Kodu": r["Hammadde Kodu"],
                "Adet": int(r["Üretim Planı"]),
                "Sebep": "Uygun makine bulunamadı (eşleşme yok)",
            })
            continue

        leftover = int(remaining_qty.get(pcode, int(r["Üretim Planı"])))
        if leftover > 0:
            unassigned.append({
                "Ürün Kodu": pcode,
                "Ürün Adı": r["Ürün Adı"],
                "Tür": r["Tür"],
                "Hammadde Kodu": r["Hammadde Kodu"],
                "Adet": leftover,
                "Sebep": "Vardiya/gün/makine kapasitesi yetersiz (ürün bitirilemedi)",
            })
            warnings.append(f"{pcode} için {leftover} adet sığmadı.")

    warnings.extend(sorted(raw_material_block_warnings))

    return schedule_rows, unassigned, warnings, product_first_day_master_idx, product_last_day_master_idx


# ===================== Ana Akış =====================

def main():
    plan = read_plan()
    prod_to_machines = read_mapping()

    # Öncelikli SKU'lar en üste + sonra büyükten küçüğe
    plan_sorted = sort_plan_with_priority(plan)


    plan_row_by_code = {r["Ürün Kodu"]: r for _, r in plan_sorted.iterrows()}
    machines_union = _union_machines()
    days_union = _union_days_in_master_order()

    # 1. vardiya kalanları (sadece SHIFT1 gün/makine)
    remaining_cap_shift1 = {m: {d: int(CAPACITIES_SHIFT1.get(m, 0)) for d in SHIFT1_DAYS} for m in SHIFT1_MACHINES}

    # 2. vardiya kalanları
    remaining_cap_shift2 = {
        m: {d: int(CAPACITIES_SHIFT2.get(m, 0)) for d in SHIFT2_DAYS}
        for m in SHIFT2_MACHINES
    }

    # 2. vardiya gün bazlı makine takibi (max 3 makine kuralı için)
    day_second_shift_machines = {d: set() for d in SHIFT2_DAYS}

    # Pair kuralları için gün bazlı ana makine takibi
    day_pair_primary_12 = {d: None for d in SHIFT2_DAYS}
    day_pair_primary_34 = {d: None for d in SHIFT2_DAYS}

    # tüketim takip (union gün+makine)
    consumed_shift1 = {m: {d: 0 for d in days_union} for m in machines_union}
    consumed_shift2 = {m: {d: 0 for d in days_union} for m in machines_union}

    # (artık dengeleme ana hedef değil ama metrik olarak tutuyoruz)
    day_totals = {d: 0 for d in SHIFT1_DAYS}

    # ✅ Makine bazlı kuyruk + slot simülasyonu (non-preemptive + vardiya slot doldurma)
    schedule_rows, unassigned, warnings, product_first_day_master_idx, product_last_day_master_idx = simulate_machine_slot_schedule(plan_sorted, prod_to_machines)

    # === schedule_df + Açıklama ===
    schedule_df = pd.DataFrame(
        schedule_rows,
        columns=["Gün", "Makine", "Ürün Kodu", "Ürün Adı", "Tür", "Hammadde Kodu", "Adet", "Not"],
    )
    schedule_df["Açıklama"] = ""

    # ✅ FIX: 1./2. vardiya tüketimlerini schedule_df'den yeniden hesapla
    machines_union = _union_machines()
    days_union = _union_days_in_master_order()
    consumed_shift1 = {m: {d: 0 for d in days_union} for m in machines_union}
    consumed_shift2 = {m: {d: 0 for d in days_union} for m in machines_union}

    _tmp = schedule_df.copy()
    _tmp["Not"] = _tmp["Not"].fillna("")
    _tmp["Adet"] = pd.to_numeric(_tmp["Adet"], errors="coerce").fillna(0).astype(int)
    _g = _tmp.groupby(["Gün", "Makine", "Not"], as_index=False)["Adet"].sum()
    for _, _r in _g.iterrows():
        _d = _r["Gün"]
        _m = _r["Makine"]
        _q = int(_r["Adet"])
        if _r["Not"] == "2.VARDIYA":
            consumed_shift2[_m][_d] += _q
        else:
            consumed_shift1[_m][_d] += _q

    # açıklama sadece 1. vardiya satırları için ardışık gün (MASTER düzeninde)
    first_shift_df = schedule_df[schedule_df["Not"] != "2.VARDIYA"].copy()
    if not first_shift_df.empty:
        for (pcode, machine), grp in first_shift_df.groupby(["Ürün Kodu", "Makine"]):
            grp = grp.copy()
            grp["master_idx"] = grp["Gün"].map(MASTER_DAY_INDEX)
            grp = grp.sort_values("master_idx")

            idx_by_master = {int(row["master_idx"]): row.name for _, row in grp.iterrows()}
            for mi, row_idx in idx_by_master.items():
                next_mi = mi + 1
                if next_mi not in idx_by_master:
                    continue
                next_row_idx = idx_by_master[next_mi]
                next_day_name = MASTER_DAYS[next_mi]
                next_qty = schedule_df.at[next_row_idx, "Adet"]
                schedule_df.at[row_idx, "Açıklama"] = f"{next_day_name} günü {machine} {next_qty} adet üretim devam edecektir."

    # Haftalık Çizelge grid (union gün + union makine)
    days_union = _union_days_in_master_order()
    machines_union = _union_machines()

    weekly_rows = []
    for day in days_union:
        day_df = schedule_df[schedule_df["Gün"] == day]
        for m in machines_union:
            dfm = day_df[day_df["Makine"] == m]
            if dfm.empty:
                weekly_rows.append(
                    {"Gün": day, "Makine": m, "Ürün Kodu": "", "Ürün Adı": "", "Tür": "", "Hammadde Kodu": "",
                     "Adet": "", "Not": "", "Açıklama": ""}
                )
            else:
                weekly_rows.extend(dfm.to_dict("records"))

    weekly_grid = pd.DataFrame(
        weekly_rows,
        columns=["Gün", "Makine", "Ürün Kodu", "Ürün Adı", "Tür", "Hammadde Kodu", "Adet", "Not", "Açıklama"],
    )

    # Günlük kullanım & özetler
    usage_rows = []

    weekly_cap1 = {m: 0 for m in machines_union}
    weekly_cap2 = {m: 0 for m in machines_union}
    weekly_used1 = {m: 0 for m in machines_union}
    weekly_used2 = {m: 0 for m in machines_union}
    weekly_total_cap = {m: 0 for m in machines_union}
    weekly_total_used = {m: 0 for m in machines_union}

    for day in days_union:
        for m in machines_union:
            cap1 = int(CAPACITIES_SHIFT1.get(m, 0)) if (day in SHIFT1_DAYS and m in SHIFT1_MACHINES) else 0
            used1 = int(consumed_shift1.get(m, {}).get(day, 0))

            cap2_base = int(CAPACITIES_SHIFT2.get(m, 0)) if (day in SHIFT2_DAYS and m in SHIFT2_MACHINES) else 0
            used2 = int(consumed_shift2.get(m, {}).get(day, 0))
            cap2 = cap2_base

            cap_total = cap1 + cap2
            used_total = used1 + used2

            weekly_cap1[m] += cap1
            weekly_cap2[m] += cap2
            weekly_used1[m] += used1
            weekly_used2[m] += used2
            weekly_total_cap[m] += cap_total
            weekly_total_used[m] += used_total

            usage_rows.append(
                {
                    "Gün": day,
                    "Makine": m,
                    "Kapasite 1. Vardiya": cap1,
                    "Tüketilen 1. Vardiya": used1,
                    "Kullanım % 1. Vardiya": (used1 / cap1) if cap1 > 0 else 0.0,
                    "Kapasite 2. Vardiya": cap2,
                    "Tüketilen 2. Vardiya": used2,
                    "Kullanım % 2. Vardiya": (used2 / cap2) if cap2 > 0 else 0.0,
                    "Toplam Kapasite": cap_total,
                    "Toplam Tüketim": used_total,
                    "Toplam Kullanım %": (used_total / cap_total) if cap_total > 0 else 0.0,
                }
            )

    usage_df = pd.DataFrame(usage_rows)

    weekly_df = pd.DataFrame(
        [
            {
                "Makine": m,
                "1. Vardiya Kapasite": int(weekly_cap1[m]),
                "1. Vardiya Üretim Planı": int(weekly_used1[m]),
                "2. Vardiya Kapasite": int(weekly_cap2[m]),
                "2. Vardiya Üretim Planı": int(weekly_used2[m]),
                "Toplam Kapasite": int(weekly_total_cap[m]),
                "Toplam Üretim Planı": int(weekly_total_used[m]),
                "Toplam Doluluk %": (weekly_total_used[m] / weekly_total_cap[m]) if weekly_total_cap[m] > 0 else 0.0,
                "2. Vardiya İhtiyaç mı": "Evet" if weekly_used2[m] > 0 else "",
            }
            for m in machines_union
        ]
    ).sort_values("Makine")

    daily_summary = (
        usage_df.groupby("Gün", as_index=False)
        .agg({
            "Kapasite 1. Vardiya": "sum",
            "Tüketilen 1. Vardiya": "sum",
            "Kapasite 2. Vardiya": "sum",
            "Tüketilen 2. Vardiya": "sum",
            "Toplam Kapasite": "sum",
            "Toplam Tüketim": "sum",
        })
    )
    daily_summary["2. Vardiya İhtiyaç mı"] = daily_summary["Tüketilen 2. Vardiya"].apply(lambda x: "Evet" if x > 0 else "")
    daily_summary["Toplam Doluluk %"] = daily_summary.apply(
        lambda r: (r["Toplam Tüketim"] / r["Toplam Kapasite"]) if r["Toplam Kapasite"] > 0 else 0.0, axis=1
    )

    daily_df = pd.DataFrame({
        "Gün": daily_summary["Gün"],
        "1. Vardiya Kapasite": daily_summary["Kapasite 1. Vardiya"].astype(int),
        "1. Vardiya Üretim Planı": daily_summary["Tüketilen 1. Vardiya"].astype(int),
        "2. Vardiya Kapasite": daily_summary["Kapasite 2. Vardiya"].astype(int),
        "2. Vardiya Üretim Planı": daily_summary["Tüketilen 2. Vardiya"].astype(int),
        "Toplam Kapasite": daily_summary["Toplam Kapasite"].astype(int),
        "Toplam Üretim Planı": daily_summary["Toplam Tüketim"].astype(int),
        "2. Vardiya İhtiyaç mı": daily_summary["2. Vardiya İhtiyaç mı"],
        "Toplam Doluluk %": daily_summary["Toplam Doluluk %"],
    })

    # Dağıtılamayanlar & Uyarılar
    un_df = pd.DataFrame(unassigned, columns=["Ürün Kodu", "Ürün Adı", "Tür", "Hammadde Kodu", "Adet", "Sebep"])
    if un_df.empty:
        un_df = pd.DataFrame([{"Ürün Kodu": "", "Ürün Adı": "", "Tür": "", "Hammadde Kodu": "", "Adet": "", "Sebep": "(Boş)"}])

    warn_df = pd.DataFrame(warnings, columns=["Uyarı"]) if warnings else pd.DataFrame([{"Uyarı": "(Uyarı yok)"}])

    # Excel'e yaz
    with pd.ExcelWriter(OUTPUT_XLSX, engine="xlsxwriter") as writer:
        weekly_grid.to_excel(writer, sheet_name="Haftalık Çizelge", index=False)
        usage_df.to_excel(writer, sheet_name="Günlük Kullanım", index=False)
        weekly_df.to_excel(writer, sheet_name="Haftalık Özet", index=False)
        daily_df.to_excel(writer, sheet_name="Günlük Özet", index=False)
        un_df.to_excel(writer, sheet_name="Dağıtılamayanlar", index=False)
        warn_df.to_excel(writer, sheet_name="Uyarılar", index=False)

        wb = writer.book
        pct_fmt = wb.add_format({"num_format": "0.00%"})
        fmt_red_bold = wb.add_format({"font_color": "red", "bold": True})
        fmt_red = wb.add_format({"font_color": "red"})

        # Haftalık Çizelge format
        ws_grid = writer.sheets["Haftalık Çizelge"]
        ws_grid.set_column(0, 0, 12)
        ws_grid.set_column(1, 1, 16)
        ws_grid.set_column(2, 2, 16)
        ws_grid.set_column(3, 3, 32)
        ws_grid.set_column(4, 4, 14)
        ws_grid.set_column(5, 5, 16)
        ws_grid.set_column(6, 6, 10)
        ws_grid.set_column(7, 7, 12)
        ws_grid.set_column(8, 8, 46)

        # HM tekrarlarını kırmızı yap (aynı gün + aynı HM)
        non_empty_hm = weekly_grid[weekly_grid["Hammadde Kodu"].astype(str).str.strip() != ""][["Gün", "Hammadde Kodu"]]
        counts = Counter([tuple(x) for x in non_empty_hm.values])
        red_keys = {k for k, v in counts.items() if v >= 2}
        for i in range(len(weekly_grid)):
            day = weekly_grid.iloc[i]["Gün"]
            hm = str(weekly_grid.iloc[i]["Hammadde Kodu"]).strip()
            if hm and (day, hm) in red_keys:
                ws_grid.set_row(i + 1, None, fmt_red)

        # Günlük kullanım format
        ws_usage = writer.sheets["Günlük Kullanım"]
        ws_usage.set_column(0, 1, 14)
        ws_usage.set_column(2, 3, 18)
        ws_usage.set_column(5, 6, 18)
        ws_usage.set_column(8, 9, 18)
        ws_usage.set_column(4, 4, 14, pct_fmt)
        ws_usage.set_column(7, 7, 14, pct_fmt)
        ws_usage.set_column(10, 10, 14, pct_fmt)

        # Özetler + conditional
        ws_week = writer.sheets["Haftalık Özet"]
        ws_week.set_column(0, 0, 16)
        ws_week.set_column(1, 6, 22)
        ws_week.set_column(7, 7, 18)
        ws_week.set_column(8, 8, 18, pct_fmt)
        ws_week.conditional_format(1, 7, len(weekly_df), 7, {"type": "text", "criteria": "containing", "value": "Evet", "format": fmt_red_bold})

        ws_day = writer.sheets["Günlük Özet"]
        ws_day.set_column(0, 0, 14)
        ws_day.set_column(1, 6, 22)
        ws_day.set_column(7, 7, 18)
        ws_day.set_column(8, 8, 18, pct_fmt)
        ws_day.conditional_format(1, 7, len(daily_df), 7, {"type": "text", "criteria": "containing", "value": "Evet", "format": fmt_red_bold})

    print(f"Tamamlandı. Çıktı: {os.path.abspath(OUTPUT_XLSX)}")


# ===================== STREAMLIT ARAYÜZÜ =====================
try:
    import streamlit as st
    import tempfile
    from io import BytesIO
except ImportError:
    st = None

if st is not None:
    st.set_page_config(page_title="Haftalık Makine Dağılımı", layout="wide")
    st.title("📦 Haftalık Makine Dağılımı Planlayıcı")
    st.write("created by **Cavit Karakuzu**")

    st.header("1️⃣ Çalışma Dosyaları")
    col1, col2 = st.columns(2)
    with col1:
        plan_file = st.file_uploader("Üretim Planı", type=["xlsx"], key="plan")
    with col2:
        map_file = st.file_uploader("Ürün-Makine Eşleştirme", type=["xlsx"], key="map")

    st.markdown("---")

    # SKU listesi için plan preview
    sku_options = []
    if plan_file is not None:
        try:
            plan_preview = pd.read_excel(BytesIO(plan_file.getvalue()), sheet_name="Sheet1")
            plan_preview.columns = [str(c).strip() for c in plan_preview.columns]
            if "Ürün Kodu" in plan_preview.columns:
                sku_options = (
                    plan_preview["Ürün Kodu"].astype(str).str.strip()
                    .replace("nan", "")
                )
                sku_options = sku_options[sku_options != ""].dropna().unique().tolist()
                sku_options = sorted(set(sku_options))
        except Exception:
            sku_options = []

    st.header("2️⃣ Öncelikli SKU Seçimi")
    st.caption("Seçtiğin SKU'lar plan sırasının en başına alınır (sonra Üretim Planı büyükten küçüğe).")
    priority_skus_ui = st.multiselect(
        "Öncelikli SKU'lar",
        options=sku_options,
        default=[],
        help="Üretimde önce görmek istediğin SKU'ları seç."
    )

    # ✅ Seçim sırasını koru (1-seçilen, 2-seçilen, ...)
    if "priority_sku_order" not in st.session_state:
        st.session_state["priority_sku_order"] = []

    current_sel = [str(x).strip() for x in (priority_skus_ui or []) if str(x).strip()]
    prev_order = [str(x).strip() for x in (st.session_state["priority_sku_order"] or []) if str(x).strip()]

    # listeden çıkarılanları sil
    prev_order = [x for x in prev_order if x in current_sel]
    # yeni eklenenleri, UI'dan gelen sıraya göre ekle
    newly_added = [x for x in current_sel if x not in prev_order]
    st.session_state["priority_sku_order"] = prev_order + newly_added
    st.markdown("---")

    # ===== 1. Vardiya Seçimleri =====
    st.header("3️⃣ 1. Vardiya Ayarları")
    all_machine_options = sorted(DEFAULT_CAPS_SHIFT1.keys())

    with st.expander("🕐 1. Vardiya: Gün / Makine / Kapasite Seçimi", expanded=True):
        c1, c2 = st.columns(2)
        with c1:
            shift1_days_ui = st.multiselect(
                "1. Vardiya Günleri",
                options=MASTER_DAYS,
                default=MASTER_DAYS,
                help="1. vardiyada çalışılacak günleri seç."
            )
        with c2:
            shift1_machines_ui = st.multiselect(
                "1. Vardiya Makineleri",
                options=all_machine_options,
                default=all_machine_options,
                help="1. vardiyada aktif makineleri seç."
            )

        st.subheader("1. Vardiya Kapasiteleri")
        caps1_ui = {}
        for m in all_machine_options:
            if m not in shift1_machines_ui:
                continue
            caps1_ui[m] = st.number_input(
                f"{m} - 1. Vardiya Kapasitesi",
                min_value=0,
                value=int(DEFAULT_CAPS_SHIFT1.get(m, 0)),
                step=500,
                key=f"cap1_{m}"
            )

    st.markdown("---")

    # ===== 2. Vardiya Seçimleri =====
    st.header("4️⃣ 2. Vardiya Ayarları")

    with st.expander("🌙 2. Vardiya: Gün / Makine / Kapasite Seçimi", expanded=True):
        c3, c4 = st.columns(2)
        with c3:
            shift2_days_ui = st.multiselect(
                "2. Vardiya Günleri",
                options=MASTER_DAYS,
                default=MASTER_DAYS,
                help="2. vardiyada çalışılacak günleri seç."
            )
        with c4:
            shift2_machines_ui = st.multiselect(
                "2. Vardiya Makineleri",
                options=all_machine_options,
                default=all_machine_options,
                help="2. vardiyada aktif makineleri seç (çalışmayanı çıkar ya da kapasiteyi 0 yap)."
            )

        st.subheader("2. Vardiya Kapasiteleri")
        st.caption("2. vardiyada çalışmayan makineye 0 girebilirsin.")
        caps2_ui = {}
        for m in all_machine_options:
            if m not in shift2_machines_ui:
                continue
            caps2_ui[m] = st.number_input(
                f"{m} - 2. Vardiya Kapasitesi",
                min_value=0,
                value=int(DEFAULT_CAPS_SHIFT2.get(m, 0)),
                step=500,
                key=f"cap2_{m}"
            )

    st.markdown("---")

    st.header("5️⃣ Planı Çalıştır")
    run_button = st.button("▶️ Planı Makinelere Dağıt")

    if run_button:
        if not plan_file or not map_file:
            st.error("Önce her iki Excel dosyasını da yüklemelisin.")
            st.stop()

        if not shift1_days_ui:
            st.error("1. vardiya için en az bir gün seçmelisin.")
            st.stop()

        if not shift1_machines_ui:
            st.error("1. vardiya için en az bir makine seçmelisin.")
            st.stop()

        if not shift2_days_ui:
            st.error("2. vardiya için en az bir gün seçmelisin.")
            st.stop()

        if not shift2_machines_ui:
            st.error("2. vardiya için en az bir makine seçmelisin.")
            st.stop()

        with tempfile.TemporaryDirectory() as tmpdir:
            # dosyaları temp’e yaz
            plan_path = os.path.join(tmpdir, PLAN_XLSX)
            map_path = os.path.join(tmpdir, MAP_XLSX)

            with open(plan_path, "wb") as f:
                f.write(plan_file.getvalue())
            with open(map_path, "wb") as f:
                f.write(map_file.getvalue())

            # UI -> global parametreler

            PRIORITY_SKU_ORDER = list(st.session_state.get('priority_sku_order', []))
            PRIORITY_SKUS = set(PRIORITY_SKU_ORDER)

            SHIFT1_DAYS = [d for d in MASTER_DAYS if d in shift1_days_ui]
            SHIFT2_DAYS = [d for d in MASTER_DAYS if d in shift2_days_ui]

            SHIFT1_MACHINES = list(shift1_machines_ui)
            SHIFT2_MACHINES = list(shift2_machines_ui)

            CAPACITIES_SHIFT1 = {m: int(caps1_ui.get(m, 0)) for m in SHIFT1_MACHINES}
            CAPACITIES_SHIFT2 = {m: int(caps2_ui.get(m, 0)) for m in SHIFT2_MACHINES}

            PRIORITY_MACHINES = [m for m in DEFAULT_PRIORITY_MACHINES if m in SHIFT1_MACHINES]

            # forbidden/pair setleri 2. vardiya makinelerine göre daralt
            SECOND_SHIFT_FORBIDDEN = {m for m in DEFAULT_SECOND_SHIFT_FORBIDDEN if m in SHIFT2_MACHINES}
            PAIR_12 = {m for m in DEFAULT_PAIR_12 if m in SHIFT2_MACHINES}
            PAIR_34 = {m for m in DEFAULT_PAIR_34 if m in SHIFT2_MACHINES}

            old_cwd = os.getcwd()
            os.chdir(tmpdir)

            st.info("Plan makinelere ve günlere bölünüyor... ⏳")
            try:
                main()
            except Exception as e:
                os.chdir(old_cwd)
                st.error("Kod çalışırken hata oluştu.")
                st.exception(e)
                st.stop()
            finally:
                os.chdir(old_cwd)

            out_path = os.path.join(tmpdir, OUTPUT_XLSX)
            if not os.path.exists(out_path):
                st.error("Çıktı dosyası bulunamadı.")
                st.stop()

            st.success("Plan hazır! 🎉")
            with open(out_path, "rb") as f:
                data = f.read()

            st.download_button(
                label="📥 Haftalık Makine Dağılımı Excel'ini İndir",
                data=data,
                file_name=OUTPUT_XLSX,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            # Hızlı bakış
            try:
                st.markdown("---")
                st.subheader("📊 Haftalık Özet (Hızlı Bakış)")
                df_week = pd.read_excel(out_path, sheet_name="Haftalık Özet")

                def style_second_shift(col):
                    return ["font-weight: bold; color: red" if val == "Evet" else "" for val in col]

                st.dataframe(df_week.style.apply(style_second_shift, subset=["2. Vardiya İhtiyaç mı"]), use_container_width=True)

                st.subheader("📅 Günlük Özet")
                df_day_summary = pd.read_excel(out_path, sheet_name="Günlük Özet")
                st.dataframe(df_day_summary.style.apply(style_second_shift, subset=["2. Vardiya İhtiyaç mı"]), use_container_width=True)

                st.subheader("📅 Günlük Kullanım (Detay)")
                df_day = pd.read_excel(out_path, sheet_name="Günlük Kullanım")
                st.dataframe(df_day, use_container_width=True)
            except Exception:
                pass
