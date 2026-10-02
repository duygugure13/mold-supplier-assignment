import math
from pathlib import Path
import pandas as pd
import gurobipy as gp
from gurobipy import GRB

# 1. DOSYA AYARLARI
BASE_DIR = Path(__file__).resolve().parent

# Proje klasörü:
# Her yeni projede sadece bu klasörün içindeki teklif/rule dosyalarını değiştir.
PROJECT_DIR = BASE_DIR / "Proje"

SUPPLIER_FILE = PROJECT_DIR / "Supplier_Master.xlsx"
QUOTE_FOLDER = PROJECT_DIR / "Teklifler"
HARMONY_FILE = PROJECT_DIR / "Harmony_Groups.xlsx"
RULES_FILE = PROJECT_DIR / "Project_Rules.xlsx"
OUTPUT_FILE = PROJECT_DIR / "optimizasyon_sonuclari.xlsx"

# Geliştirme aşamasında eski dosya adını da kabul et.
if not SUPPLIER_FILE.exists():
    old_supplier_file = BASE_DIR / "master_supplier_dataset_30_suppliers_controlled_tradeoff.xlsx"
    if old_supplier_file.exists():
        SUPPLIER_FILE = old_supplier_file

# 2. PROJE AYARLARI
BASE_CAPACITY = 6

# Dengeli senaryo için başlangıç ağırlıkları.Daha sonra kullanıcı tarafından değiştirilebilir.
BALANCED_WEIGHTS = {
    "cost": 1 / 3,
    "delivery": 1 / 3,
    "supplier": 1 / 3,
}

# Domestic / Abroad senaryolarında sadece ilgili tedarikçiler kullanılabilir.
SCENARIOS = [
    "Minimum Cost",
    "Minimum Delivery",
    "Balanced",
    "100% Domestic",
    "100% Abroad",
]

# 3. VERİ OKUMA
def load_quotations_from_folder():
    """
    Teklifler/F001.xlsx -> Supplier_ID = F001
    Kalıp sayısı sabit değildir. Excel'de kaç farklı Mold_ID varsa
    sistem o kadar kalıbı otomatik algılar.
    """
    if not QUOTE_FOLDER.exists():
        raise FileNotFoundError(
            f"'{QUOTE_FOLDER}' klasörü bulunamadı. "
            f"Proje klasörü ve Teklifler klasörünü oluşturun."
        )

    files = sorted(
        p for p in QUOTE_FOLDER.iterdir()
        if p.is_file() and p.suffix.lower() in {".xlsx", ".xls"}
        and not p.name.startswith("~$")
    )

    if not files:
        raise FileNotFoundError(
            f"'{QUOTE_FOLDER}' klasöründe hiç teklif Excel dosyası bulunamadı."
        )

    frames = []
    for file in files:
        supplier_id = file.stem.strip()

        try:
            df = pd.read_excel(file)
        except Exception as e:
            raise ValueError(f"Teklif dosyası okunamadı: {file.name} -> {e}")

        # Sütun isimlerindeki boşlukları temizle.
        df.columns = [str(c).strip() for c in df.columns]

        required = {
            "Mold_ID",
            "Price_Excl_VAT_TL",
            "Delivery_Time_Days",
        }
        missing = required - set(df.columns)
        if missing:
            raise ValueError(
                f"{file.name} eksik sütun içeriyor: {sorted(missing)}. "
                f"Gerekli sütunlar: Mold_ID, Price_Excl_VAT_TL, Delivery_Time_Days"
            )

        # Supplier_ID dosya adından belirlenir; böylece dosyanın içindeki
        # Supplier_ID yanlış olsa bile proje içinde tutarlı kalır.
        df["Supplier_ID"] = supplier_id

        frames.append(df)

    quotations = pd.concat(frames, ignore_index=True)

    # Dinamik kalıp listesi.
    quotations["Mold_ID"] = quotations["Mold_ID"].astype(str).str.strip()

    if quotations["Mold_ID"].eq("").any():
        raise ValueError("Teklif dosyalarında boş Mold_ID bulundu.")

    # Sayısal alanları güvenli biçimde sayıya çevir.
    for col in ["Price_Excl_VAT_TL", "Delivery_Time_Days"]:
        quotations[col] = pd.to_numeric(quotations[col], errors="coerce")
        if quotations[col].isna().any():
            bad_rows = quotations.index[quotations[col].isna()].tolist()[:10]
            raise ValueError(
                f"{col} alanında sayısal olmayan değer bulundu. "
                f"Örnek satırlar: {bad_rows}"
            )

    # Aynı tedarikçiden aynı kalıp için iki teklif varsa model belirsizleşir.
    duplicates = quotations[
        quotations.duplicated(["Supplier_ID", "Mold_ID"], keep=False)
    ]
    if not duplicates.empty:
        pairs = (
            duplicates[["Supplier_ID", "Mold_ID"]]
            .drop_duplicates()
            .astype(str)
            .agg(" + ".join, axis=1)
            .tolist()
        )
        raise ValueError(
            "Aynı Supplier_ID + Mold_ID için birden fazla teklif bulundu: "
            + ", ".join(pairs[:10])
        )

    return quotations


def load_harmony():
    """
    Harmony bilgisi proje klasöründeki Harmony_Groups.xlsx dosyasından okunur.
    Dosya yoksa tüm kalıplar No_Harmony kabul edilir.
    """
    if HARMONY_FILE.exists():
        harmony = pd.read_excel(HARMONY_FILE)
        harmony.columns = [str(c).strip() for c in harmony.columns]

        required = {"Mold_ID", "Harmony_Group"}
        missing = required - set(harmony.columns)
        if missing:
            raise ValueError(
                f"Harmony_Groups.xlsx eksik sütunlar: {sorted(missing)}"
            )

        harmony = harmony[["Mold_ID", "Harmony_Group"]].copy()
        harmony["Mold_ID"] = harmony["Mold_ID"].astype(str).str.strip()
        harmony["Harmony_Group"] = (
            harmony["Harmony_Group"].fillna("No_Harmony").astype(str).str.strip()
        )
    else:
        # Harmony dosyası verilmezse hiçbir kalıp gruplu değildir.
        harmony = None

    return harmony


def load_project_rules():
    """
    Project_Rules.xlsx opsiyoneldir.
    Şimdilik proje ayarlarının gelecekte genişletilebilmesi için okunur.
    """
    if not RULES_FILE.exists():
        return None
    return pd.read_excel(RULES_FILE)


def load_data():
    # İlk sayfayı oku; sayfa adının "Supplier_Master" olması zorunlu değildir.
    supplier_master = pd.read_excel(SUPPLIER_FILE)

    quotations = load_quotations_from_folder()

    harmony = load_harmony()

    if harmony is None:
        harmony = pd.DataFrame({
            "Mold_ID": sorted(quotations["Mold_ID"].unique()),
            "Harmony_Group": "No_Harmony"
        })

    return supplier_master, quotations, harmony


# 4. TEDARİKÇİ SKORLARINI HESAPLAMA
def calculate_supplier_scores(df):
    df = df.copy()

    # 0-10 olan teknik kriterler
    df["Technical_Credibility"] = (
        0.30 * df["Technical_Team"] +
        0.25 * df["Manufacturing_Capability"] +
        0.25 * df["Design_Engineering_Capability"] +
        0.20 * df["Environmental_Certification"]
    ) * 10

    # Genel geçmiş performans
    df["Past_Performance"] = (
        0.30 * df["Quality_Performance"] +
        0.25 * df["Problem_Solving_Score"] +
        0.25 * df["Project_Management_Score"] +
        0.20 * df["Customer_Satisfaction"]
    )

    # Esneklik
    df["Flexibility_Score"] = (
        0.25 * df["Design_Change_Flexibility"] +
        0.25 * df["Delivery_Change_Flexibility"] +
        0.25 * df["Production_Flexibility"] +
        0.25 * df["Rush_Order_Capability"]
    ) * 10

    # Sürdürülebilirlik
    df["Sustainability_Score"] = (
        0.30 * df["Environmental_Management"] +
        0.25 * df["Waste_Management"] +
        0.25 * df["Energy_Efficiency"] +
        0.20 * df["Environmental_Certification"]
    ) * 10

    # Benzer proje deneyimi 0-10'dan 0-100'e çevrilir.
    df["Similar_Experience_Score"] = (
        df["Similar_Project_Experience"] * 10
    )

    # Tek bir tedarikçi performans skoru.Yukarıda hesaplanan tüm kriterler birleştirilir.
    df["Supplier_Performance_Score"] = (
        0.30 * df["Technical_Credibility"] +
        0.25 * df["Past_Performance"] +
        0.20 * df["Similar_Experience_Score"] +
        0.15 * df["Flexibility_Score"] +
        0.10 * df["Sustainability_Score"]
    )

    # Kapasite hard constraint'i
    df["Max_Capacity"] = (
        BASE_CAPACITY * df["Capacity_Coefficient"]
    ).apply(math.floor).astype(int)

    return df

# 5. TEKLİF SKORLARINI HESAPLAMA
def calculate_quote_scores(quotes):
    quotes = quotes.copy()

    # Daha düşük maliyet = daha yüksek skor
    min_cost = quotes["Price_Excl_VAT_TL"].min()
    quotes["Cost_Score"] = (
        min_cost / quotes["Price_Excl_VAT_TL"] * 100
    )
    # Daha kısa teslim süresi = daha yüksek skor
    min_delivery = quotes["Delivery_Time_Days"].min()
    quotes["Delivery_Score"] = (
        min_delivery / quotes["Delivery_Time_Days"] * 100
    )

    return quotes

# 6. VERİ KONTROLÜ
def validate_data(suppliers, quotes, harmony):
    required_supplier_cols = {
        "Supplier_ID",
        "On_Time_Delivery_Rate",
        "Capacity_Coefficient",
        "Technical_Team",
        "Manufacturing_Capability",
        "Design_Engineering_Capability",
        "Similar_Project_Experience",
        "Quality_Performance",
        "Problem_Solving_Score",
        "Project_Management_Score",
        "Customer_Satisfaction",
        "Distance_km",
        "Design_Change_Flexibility",
        "Delivery_Change_Flexibility",
        "Production_Flexibility",
        "Rush_Order_Capability",
        "Domestic_Abroad",
        "Environmental_Management",
        "Waste_Management",
        "Energy_Efficiency",
        "Environmental_Certification",
    }

    required_quote_cols = {
        "Supplier_ID",
        "Mold_ID",
        "Price_Excl_VAT_TL",
        "Delivery_Time_Days",
    }

    missing_supplier = required_supplier_cols - set(suppliers.columns)
    missing_quote = required_quote_cols - set(quotes.columns)

    if missing_supplier:
        raise ValueError(
            f"Supplier Master eksik sütunlar: {sorted(missing_supplier)}"
        )

    if missing_quote:
        raise ValueError(
            f"Teklif dosyası eksik sütunlar: {sorted(missing_quote)}"
        )

    if suppliers["Supplier_ID"].duplicated().any():
        raise ValueError("Supplier_ID değerleri benzersiz olmalıdır.")

    if quotes[["Supplier_ID", "Mold_ID"]].duplicated().any():
        raise ValueError(
            "Aynı Supplier_ID + Mold_ID için birden fazla teklif var."
        )

    supplier_set = set(suppliers["Supplier_ID"])
    quote_supplier_set = set(quotes["Supplier_ID"])

    unknown_suppliers = quote_supplier_set - supplier_set
    if unknown_suppliers:
        raise ValueError(
            f"Supplier Master'da bulunmayan tedarikçiler: {unknown_suppliers}"
        )

    mold_set = set(quotes["Mold_ID"])
    harmony_mold_set = set(harmony["Mold_ID"])

    unknown_harmony_molds = harmony_mold_set - mold_set
    if unknown_harmony_molds:
        raise ValueError(
            f"Tekliflerde bulunmayan harmony kalıpları: {unknown_harmony_molds}"
        )

# 7. HARMONY UYGUNLUK KONTROLÜ VE AÇIKLAMASI

def audit_harmony_eligibility(suppliers, quotes, harmony, allowed_suppliers, scenario):
    """
    Her harmony grubunda hangi tedarikçinin neden aday olduğunu/elenğini açıklar.
    Bu fonksiyon optimizasyon sonucunu değiştirmez; sadece denetim ve raporlama yapar.
    """
    rows = []
    harmonized = harmony[harmony["Harmony_Group"] != "No_Harmony"].copy()

    if harmonized.empty:
        return pd.DataFrame(columns=[
            "Scenario", "Harmony_Group", "Supplier_ID", "Status",
            "Group_Mold_Count", "Quoted_Mold_Count", "Missing_Molds", "Reason"
        ])

    quote_map = quotes.groupby("Supplier_ID")["Mold_ID"].apply(set).to_dict()

    for group_name, group_df in harmonized.groupby("Harmony_Group"):
        group_molds = set(group_df["Mold_ID"])

        for s in sorted(suppliers["Supplier_ID"].unique()):
            supplier_type = suppliers.loc[
                suppliers["Supplier_ID"] == s, "Domestic_Abroad"
            ].iloc[0]

            if s not in allowed_suppliers:
                rows.append({
                    "Scenario": scenario,
                    "Harmony_Group": group_name,
                    "Supplier_ID": s,
                    "Status": "ELENDİ",
                    "Group_Mold_Count": len(group_molds),
                    "Quoted_Mold_Count": len(group_molds & quote_map.get(s, set())),
                    "Missing_Molds": ", ".join(sorted(group_molds - quote_map.get(s, set()))),
                    "Reason": f"Senaryo kısıtı: {supplier_type} tedarikçi bu senaryoda kullanılamaz."
                })
                continue

            quoted = quote_map.get(s, set())
            missing = sorted(group_molds - quoted)

            if missing:
                rows.append({
                    "Scenario": scenario,
                    "Harmony_Group": group_name,
                    "Supplier_ID": s,
                    "Status": "ELENDİ",
                    "Group_Mold_Count": len(group_molds),
                    "Quoted_Mold_Count": len(group_molds & quoted),
                    "Missing_Molds": ", ".join(missing),
                    "Reason": (
                        f"Harmony şartı sağlanmıyor: grubun {len(group_molds)} kalıbından "
                        f"{len(missing)} tanesine teklif yok. "
                        f"Eksik kalıplar: {', '.join(missing)}."
                    )
                })
            else:
                rows.append({
                    "Scenario": scenario,
                    "Harmony_Group": group_name,
                    "Supplier_ID": s,
                    "Status": "ADAY",
                    "Group_Mold_Count": len(group_molds),
                    "Quoted_Mold_Count": len(group_molds & quoted),
                    "Missing_Molds": "",
                    "Reason": "Harmony şartı sağlanıyor: grubun tüm kalıplarına teklif mevcut."
                })

    return pd.DataFrame(rows)


# 8. OPTİMİZASYON MODELİ
def optimize(
    suppliers,
    quotes,
    harmony,
    scenario="Balanced",
    balanced_weights=None,
):
    suppliers = suppliers.copy()
    quotes = quotes.copy()
    harmony = harmony.copy()

    if balanced_weights is None:
        balanced_weights = BALANCED_WEIGHTS

    # Sadece teklif bulunan kalıp/tedarikçi kombinasyonları karar değişkenidir.
    pairs = list(
        quotes[["Mold_ID", "Supplier_ID"]]
        .itertuples(index=False, name=None)
    )

    molds = sorted(quotes["Mold_ID"].unique())
    supplier_ids = sorted(suppliers["Supplier_ID"].unique())

    supplier_info = suppliers.set_index("Supplier_ID").to_dict("index")
    quote_info = (
        quotes.set_index(["Mold_ID", "Supplier_ID"])
        .to_dict("index")
    )

    # Senaryoya göre izin verilen tedarikçiler

    if scenario == "100% Domestic":
        allowed_suppliers = {
            s for s in supplier_ids
            if supplier_info[s]["Domestic_Abroad"] == "Domestic"
        }
    elif scenario == "100% Abroad":
        allowed_suppliers = {
            s for s in supplier_ids
            if supplier_info[s]["Domestic_Abroad"] == "Abroad"
        }
    else:
        allowed_suppliers = set(supplier_ids)

    # Harmony uygunluk raporu: hangi tedarikçi neden elendi?
    harmony_audit = audit_harmony_eligibility(
        suppliers=suppliers,
        quotes=quotes,
        harmony=harmony,
        allowed_suppliers=allowed_suppliers,
        scenario=scenario,
    )

    # Gurobi modeli
    model = gp.Model(f"MoldSupplier_{scenario.replace(' ', '_')}")

    # x[m,s] = 1 ise mold m supplier s'ye atanır.
    x = {
        (m, s): model.addVar(
            vtype=GRB.BINARY,
            name=f"x_{m}_{s}"
        )
        for m, s in pairs
        if s in allowed_suppliers
    }

    model.update()

    # Minimum Delivery senaryosunda projenin tamamlanma süresini
    # (seçilen tekliflerin en yüksek teslimat süresini) temsil eder.
    max_delivery = None
    if scenario == "Minimum Delivery":
        max_delivery = model.addVar(
            vtype=GRB.CONTINUOUS,
            name="Max_Delivery_Time"
        )
        model.update()

    # Kısıt 1: Her kalıp tam olarak bir tedarikçiye

    for m in molds:
        vars_for_mold = [
            x[(m, s)]
            for s in supplier_ids
            if (m, s) in x
        ]

        if not vars_for_mold:
            raise ValueError(
                f"{m} kalıbı için uygun teklif/tedarikçi bulunamadı."
            )

        model.addConstr(
            gp.quicksum(vars_for_mold) == 1,
            name=f"OneSupplier_{m}"
        )

    # Kısıt 2: Harmony grupları aynı tedarikçiye

    harmonized = harmony[
        harmony["Harmony_Group"] != "No_Harmony"
    ].copy()

    for group_name, group_df in harmonized.groupby("Harmony_Group"):
        group_molds = group_df["Mold_ID"].tolist()

        eligible_suppliers = []
        for s in sorted(allowed_suppliers):
            quoted_molds = set(quotes.loc[quotes["Supplier_ID"] == s, "Mold_ID"])
            if set(group_molds).issubset(quoted_molds):
                eligible_suppliers.append(s)

        if not eligible_suppliers:
            detail = harmony_audit[
                (harmony_audit["Harmony_Group"] == group_name) &
                (harmony_audit["Status"] == "ELENDİ")
            ][["Supplier_ID", "Missing_Molds", "Reason"]]

            lines = [
                f"{group_name} ({', '.join(group_molds)}) için uygun tedarikçi kalmadı.",
                "Harmony grubundaki tüm kalıplara aynı tedarikçinin teklif vermesi gerekiyor.",
            ]
            for _, r in detail.iterrows():
                reason = r["Reason"]
                if r["Missing_Molds"]:
                    reason = f"Eksik teklifler: {r['Missing_Molds']}"
                lines.append(f"- {r['Supplier_ID']}: {reason}")
            raise ValueError("\n".join(lines))

        # Her tedarikçi için grubun bütün kalıpları aynı anda seçilir.
        for s in allowed_suppliers:
            group_vars = [
                x[(m, s)]
                for m in group_molds
                if (m, s) in x
            ]

            # Bir harmony grubunun tüm kalıpları o firmada
            # tekliflendiyse aynı firmaya birlikte atanabilir.
            # Eksik teklif varsa bu firma o harmony grubuna aday olamaz.
            if len(group_vars) == len(group_molds):
                first_var = group_vars[0]
                for v in group_vars[1:]:
                    model.addConstr(
                        v == first_var,
                        name=f"Harmony_{group_name}_{s}_{v.VarName}"
                    )
            else:
                # Eksik teklif nedeniyle firma bu harmony grubunu alamaz.
                for m in group_molds:
                    if (m, s) in x:
                        model.addConstr(
                            x[(m, s)] == 0,
                            name=f"HarmonyIncomplete_{group_name}_{s}_{m}"
                        )

    # Kısıt 3: Tedarikçi kapasitesi
    for s in allowed_suppliers:
        capacity = int(supplier_info[s]["Max_Capacity"])

        assigned = [
            x[(m, s)]
            for m in molds
            if (m, s) in x
        ]

        model.addConstr(
            gp.quicksum(assigned) <= capacity,
            name=f"Capacity_{s}"
        )

    # Amaç fonksiyonu
    # Balanced senaryoda her kriteri 0-100 ölçeğinde normalize ediyoruz.
    objective_terms = []

    # Minimum Delivery için seçilen her teklifin teslim süresi,
    # projenin tamamlanma süresini temsil eden Max_Delivery_Time'dan
    # küçük/eşit olmalıdır.
    if scenario == "Minimum Delivery":
        for m in molds:
            model.addConstr(
                max_delivery >= gp.quicksum(
                    quote_info[(m, s)]["Delivery_Time_Days"] * x[(m, s)]
                    for s in supplier_ids
                    if (m, s) in x
                ),
                name=f"ProjectDelivery_{m}"
            )

        model.setObjective(max_delivery, GRB.MINIMIZE)

    for (m, s), var in x.items():
        q = quote_info[(m, s)]
        supplier_score = supplier_info[s]["Supplier_Performance_Score"]

        if scenario == "Minimum Cost":
            coefficient = q["Price_Excl_VAT_TL"]

        elif scenario == "Minimum Delivery":
            # Objective yukarıda Max_Delivery_Time üzerinden tanımlandı.
            coefficient = q["Delivery_Time_Days"]
            continue

        elif scenario == "Balanced":
            coefficient = (
                balanced_weights["cost"] * (100 - q["Cost_Score"]) +
                balanced_weights["delivery"] * (100 - q["Delivery_Score"]) +
                balanced_weights["supplier"] * (100 - supplier_score)
            )

        elif scenario in {"100% Domestic", "100% Abroad"}:
            # Bu senaryolarda ana amaç maliyettir.
            # Domestic/Abroad uygunluğu zaten hard constraint ile sağlanır.
            coefficient = q["Price_Excl_VAT_TL"]

        else:
            raise ValueError(f"Bilinmeyen senaryo: {scenario}")

        objective_terms.append(coefficient * var)

    # Minimum Delivery dışında kalan senaryolarda yukarıdaki
    # teklif bazlı objective kullanılır.
    if scenario != "Minimum Delivery":
        model.setObjective(
            gp.quicksum(objective_terms),
            GRB.MINIMIZE
        )

    # Çöz
    model.optimize()

    if model.Status != GRB.OPTIMAL:
        raise RuntimeError(
            f"Model optimal çözüm bulamadı. Gurobi status: {model.Status}"
        )

    # Sonuçlar
    selected = []

    for (m, s), var in x.items():
        if var.X > 0.5:
            q = quote_info[(m, s)]
            sp = supplier_info[s]

            selected.append({
                "Mold_ID": m,
                "Supplier_ID": s,
                "Domestic_Abroad": sp["Domestic_Abroad"],
                "Price_Excl_VAT_TL": q["Price_Excl_VAT_TL"],
                "Delivery_Time_Days": q["Delivery_Time_Days"],
                "Cost_Score": q["Cost_Score"],
                "Delivery_Score": q["Delivery_Score"],
                "Supplier_Performance_Score": sp["Supplier_Performance_Score"],
                "Capacity_Coefficient": sp["Capacity_Coefficient"],
                "Max_Capacity": sp["Max_Capacity"],
                "Mold_Complexity": q.get("Mold_Complexity", ""),
            })

    result = pd.DataFrame(selected)

    # Harmony bilgisini tekrar ekle.
    result = result.merge(
        harmony[["Mold_ID", "Harmony_Group"]],
        on="Mold_ID",
        how="left"
    )

    # Supplier bazında özet
    supplier_summary = (
        result.groupby("Supplier_ID")
        .agg(
            Mold_Count=("Mold_ID", "count"),
            Total_Cost_TL=("Price_Excl_VAT_TL", "sum"),
            Average_Delivery_Days=("Delivery_Time_Days", "mean"),
            Max_Capacity=("Max_Capacity", "first"),
            Domestic_Abroad=("Domestic_Abroad", "first"),
            Supplier_Performance_Score=(
                "Supplier_Performance_Score", "first"
            ),
        )
        .reset_index()
    )

    return model, result, supplier_summary, harmony_audit

# 8. TÜM SENARYOLARI ÇALIŞTIR
def run_all_scenarios():
    suppliers, quotes, harmony = load_data()

    validate_data(suppliers, quotes, harmony)

    suppliers = calculate_supplier_scores(suppliers)
    quotes = calculate_quote_scores(quotes)

    scenario_results = {}
    scenario_summaries = []
    harmony_audits = []

    for scenario in SCENARIOS:
        print(f"\nÇözülüyor: {scenario}")

        try:
            model, result, summary, harmony_audit = optimize(
                suppliers=suppliers,
                quotes=quotes,
                harmony=harmony,
                scenario=scenario,
            )

            total_cost = result["Price_Excl_VAT_TL"].sum()
            avg_delivery = result["Delivery_Time_Days"].mean()
            max_delivery = result["Delivery_Time_Days"].max()
            supplier_count = result["Supplier_ID"].nunique()

            scenario_results[scenario] = result
            harmony_audits.append(harmony_audit)

            # Konsolda harmony nedeniyle elenen tedarikçileri açıkça göster.
            excluded = harmony_audit[harmony_audit["Status"] == "ELENDİ"]
            harmony_excluded = excluded[
                excluded["Reason"].str.startswith("Harmony şartı")
            ]
            if not harmony_excluded.empty:
                print("  Harmony nedeniyle elenen tedarikçiler:")
                for _, row in harmony_excluded.iterrows():
                    print(
                        f"    - {row['Supplier_ID']} / {row['Harmony_Group']}: "
                        f"eksik teklif = {row['Missing_Molds']}"
                    )

            scenario_summaries.append({
                "Scenario": scenario,
                "Total_Cost_TL": total_cost,
                "Average_Delivery_Days": avg_delivery,
                "Maximum_Delivery_Days": max_delivery,
                "Supplier_Count": supplier_count,
                "Objective_Value": model.ObjVal,
            })

            if scenario == "Minimum Delivery":
                print(
                    f"Toplam maliyet: {total_cost:,.0f} TL | "
                    f"Proje tamamlanma süresi: {max_delivery:.1f} gün | "
                    f"Ortalama teslimat: {avg_delivery:.1f} gün | "
                    f"Tedarikçi sayısı: {supplier_count}"
                )
            else:
                print(
                    f"Toplam maliyet: {total_cost:,.0f} TL | "
                    f"Ortalama teslimat: {avg_delivery:.1f} gün | "
                    f"Tedarikçi sayısı: {supplier_count}"
                )

        except RuntimeError as e:
            print(f"  Çözülemedi: {e}")

    scenario_summary_df = pd.DataFrame(scenario_summaries)
    # Excel çıktı

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:

        # Tedarikçi skorlarını da çıktı olarak ver.
        supplier_output_cols = [
            "Supplier_ID",
            "Domestic_Abroad",
            "On_Time_Delivery_Rate",
            "Capacity_Coefficient",
            "Max_Capacity",
            "Technical_Credibility",
            "Similar_Experience_Score",
            "Past_Performance",
            "Flexibility_Score",
            "Sustainability_Score",
            "Supplier_Performance_Score",
            "Distance_km",
        ]

        suppliers[supplier_output_cols].to_excel(
            writer,
            sheet_name="Supplier_Scores",
            index=False
        )

        # Python tarafından otomatik birleştirilen ham teklifler.
        quotes.to_excel(
            writer,
            sheet_name="All_Quotations",
            index=False
        )

        # Skorları hesaplanmış teklif tablosu.
        quotes.to_excel(
            writer,
            sheet_name="Quotation_Scores",
            index=False
        )

        scenario_summary_df.to_excel(
            writer,
            sheet_name="Scenario_Comparison",
            index=False
        )

        if harmony_audits:
            pd.concat(harmony_audits, ignore_index=True).to_excel(
                writer,
                sheet_name="Harmony_Eligibility",
                index=False
            )

        for scenario, result in scenario_results.items():
            sheet_name = scenario[:31].replace("%", "Pct")
            result.to_excel(
                writer,
                sheet_name=sheet_name,
                index=False
            )

    print("\nTamamlandı.")
    print(f"Çıktı dosyası: {OUTPUT_FILE}")

if __name__ == "__main__":
    run_all_scenarios()