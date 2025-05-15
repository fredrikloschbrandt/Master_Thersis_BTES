import pandas as pd

# --- Innstillinger ---
filsti = "BTES_simulering.xlsx"
arknavn = "Inputdata til Python simulering"

# Effektdata
VP_max_kW = 537.1 # INPUT - ikke spesifikk for brønnparken
sirkulasjonspumpe = 23.40 # 
tørrkjøler_kW = 10.12
ant_tørrkjølere = 4
tot_tørrkjøler_kW = tørrkjøler_kW * ant_tørrkjølere
System_max_kW = VP_max_kW + sirkulasjonspumpe + tot_tørrkjøler_kW
System_min_kW = System_max_kW * 0.15
COP = 3.7
varmekapasitet_kWh_per_K = 166176
start_temp = 6.3  # °C
temperaturkrav_tørrkjøler = 13

# Økonomi
energiledd_kjop = 0.05
energiledd_salg = -0.05
fastledd_salg = 0.0198

# Solcelledegradering
PV_ar_1 = 0.9850
PV_ar_25 = 0.8890
steg_reduksjon = (PV_ar_1 - PV_ar_25) / 24

# Årlige varmetap (kWh)
tap_per_ar_liste = [
    0, 0, 0, 1344465, 1081465, 965630, 897047, 850594, 816581, 790371, 769433, 752252, 737857, 725596, 715012, 
    705772, 697630, 690398, 683930, 678109, 672843, 668058, 663691, 659691, 656016, 652627, 649496, 646594
]
# Årlig varmebehov i bygg
varmebehov_bygg = 1_528_173  # kWh

# --- Last inn data ---
df = pd.read_excel(filsti, sheet_name=arknavn)
df["Tid"] = pd.to_datetime(df["Tid"])
df = df[(df["Tid"].dt.month >= 4) & (df["Tid"].dt.month <= 10)].copy()
opprinnelig_PV = df["GridExport [kWh]"].copy()

arsresultater = []
slutt_temp = start_temp

for ar in range(25):
    faktor = 1 - steg_reduksjon * ar
    tap_per_ar_kWh = tap_per_ar_liste[ar]

    # Beregn starttemperatur for året: trekk tap ut først
    temp_start = slutt_temp - tap_per_ar_kWh / varmekapasitet_kWh_per_K
    temp_start = max(temp_start, start_temp)

    # Sjekk om lager allerede er fullt ved start (55 °C)
    lager_full = temp_start >= 55

    # Uttak aktiveres fra år 4 (indeks 3)
    uttak_aktivert = (ar >= 3)

    # Forbered DataFrame-kolonner
    df["PV_korrigert"] = opprinnelig_PV * faktor
    for col in ["VP_drift_kW", "VP_varme_kWh", "Lager_Temp_C",
                "PV_til_bygg", "PV_til_nett", "Besparelse_i_bygg", "Salgsinntekt_nett"]:
        df[col] = 0.0

    temp = temp_start
    total_varme = 0.0
    total_driftstimer = 0
    total_el_forbruk = 0.0  # Inkl. VP drift
    total_el_forbruk_VP = 0.0

    # --- Lading og drift per time ---
    for i, row in df.iterrows():
        PV = row["PV_korrigert"]
        T = row["Temperatur"]
        bygglast = row["Bygglast"]
        strompris = row["Strømpris"]

        VP_drift = 0.0
        VP_varme = 0.0
        # Kun lad når lager ikke er fullt
        if not lager_full and T >= temperaturkrav_tørrkjøler and PV >= System_min_kW:
            VP_drift = min(PV, System_max_kW)
            VP_varme = VP_drift * (VP_max_kW / System_max_kW) * COP
            total_varme += VP_varme
            total_el_forbruk += VP_drift
            total_el_forbruk_VP += VP_varme / COP
            total_driftstimer += 1

            # Oppdater temperatur og sett flagg om fullt
            temp += VP_varme / varmekapasitet_kWh_per_K
            if temp >= 55:
                temp = 55
                lager_full = True

        # PV-distribusjon når VP er kjørt (eller lager fullt)
        PV_etter_VP = max(0, PV - VP_drift)
        PV_til_bygg = min(PV_etter_VP, bygglast)
        PV_til_nett = PV_etter_VP - PV_til_bygg
        besparelse_bygg = PV_til_bygg * (strompris + energiledd_kjop)

        df.at[i, "VP_drift_kW"] = VP_drift
        df.at[i, "VP_varme_kWh"] = VP_varme
        df.at[i, "Lager_Temp_C"] = temp
        df.at[i, "PV_til_bygg"] = PV_til_bygg
        df.at[i, "PV_til_nett"] = PV_til_nett
        df.at[i, "Besparelse_i_bygg"] = besparelse_bygg

    temp_etter_lading = temp

    # --- Uttak fra lager ---
    uttak = 0
    if uttak_aktivert and temp > 35:
        maks_mulig_uttak = varmekapasitet_kWh_per_K * (temp - 35)
        uttak = min(varmebehov_bygg, maks_mulig_uttak)
        temp -= uttak / varmekapasitet_kWh_per_K

    slutt_temp = temp

    # --- Salgsinntekt ---
    midlere_innmating = df["PV_til_nett"].mean()
    fastledd_time = midlere_innmating * fastledd_salg
    df["Salgsinntekt_nett"] = df["PV_til_nett"] * (df["Strømpris"] - energiledd_salg) - fastledd_time

    arsresultater.append({
        "År": ar,
        "PV_faktor": faktor,
        "Varme tilført": df["VP_varme_kWh"].sum(),
        "Tap i lager": tap_per_ar_kWh,
        "Starttemperatur etter tap (°C)": round(temp_start, 2),
        "Temp etter lading (°C)": round(temp_etter_lading, 2),
        "Slutt-temperatur (°C)": round(slutt_temp, 2),
        "Driftstimer VP": total_driftstimer,
        "El-forbruk VP inkl. pumpe (kWh)": total_el_forbruk,
        "El-forbruk VP (kWh)": total_el_forbruk_VP,
        "Uttak (kWh)": uttak,
        "PV til bygg (kWh)": df["PV_til_bygg"].sum(),
        "Besparelse i bygg (kr)": df["Besparelse_i_bygg"].sum(),
        "PV til nett (kWh)": df["PV_til_nett"].sum(),
        "Salgsinntekt fra nett (kr)": df["Salgsinntekt_nett"].sum()
    })

# --- Oppsummering ---
resultater_df = pd.DataFrame(arsresultater)
print(resultater_df)

# --- Eksport til Excel ---
with pd.ExcelWriter(filsti, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    resultater_df.to_excel(writer, sheet_name="SIM_BP2", index=False)
