import pandas as pd  # Importerer pandas for databehandling

# --- Innstillinger ---
filsti = "BTES_simulering.xlsx"  # Filsti til Excel-arbeidsbok
arknavn = "Inputdata til Python simulering"  # Navn på arket med inputdata

# Effektdata
VP_max_kW = 537.1  # Maksimal effekt fra varmepumpen (ikke brønnspesifikk)
sirkulasjonspumpe = 23.40  # Effektforbruk sirkulasjonspumpe (kW)
tørrkjøler_kW = 10.12  # Effektforbruk per tørrkjøler (kW)
ant_tørrkjølere = 4  # Antall tørrkjølere
tot_tørrkjøler_kW = tørrkjøler_kW * ant_tørrkjølere  # Total effekt fra tørrkjølere
System_max_kW = VP_max_kW + sirkulasjonspumpe + tot_tørrkjøler_kW  # Maksimal samlet effekt i systemet
System_min_kW = System_max_kW * 0.15  # Minimum effekt for å kunne drive systemet
COP = 3.7  # Varmefaktor (Coefficient of Performance)
varmekapasitet_kWh_per_K = 166176  # Varmelagringskapasitet i kWh per °C
start_temp = 6.3  # Starttemperatur i lageret (°C)
temperaturkrav_tørrkjøler = 13  # Minimum utetemp. for tørrkjølerdrift

# Økonomiske parametre
energiledd_kjop = 0.05  # Energiledd ved kjøp av strøm (kr/kWh)
energiledd_salg = -0.05  # Energiledd ved salg til nett (kr/kWh)
fastledd_salg = 0.0198  # Fastledd for innmating til nett (kr/kWh)

# Solcelledegradering per år (lineær)
PV_ar_1 = 0.9850  # Effektfaktor år 1
PV_ar_25 = 0.8890  # Effektfaktor år 25
steg_reduksjon = (PV_ar_1 - PV_ar_25) / 24  # Lineært reduksjon per år

# Estimerte årlige varmetap fra lager (kWh)
tap_per_ar_liste = [
    0, 0, 0, 1344465, 1081465, 965630, 897047, 850594, 816581, 790371, 769433, 752252, 737857, 725596, 715012, 
    705772, 697630, 690398, 683930, 678109, 672843, 668058, 663691, 659691, 656016, 652627, 649496, 646594
]

# Årlig varmebehov i bygget (kWh)
varmebehov_bygg = 1_528_173

# --- Last inn data ---
df = pd.read_excel(filsti, sheet_name=arknavn)  # Leser inn data fra Excel
df["Tid"] = pd.to_datetime(df["Tid"])  # Konverterer tid til datetime-format
df = df[(df["Tid"].dt.month >= 4) & (df["Tid"].dt.month <= 10)].copy()  # Filtrerer april–oktober
opprinnelig_PV = df["GridExport [kWh]"].copy()  # Lagrer original PV-produksjon for hvert år

arsresultater = []  # Liste for å lagre resultater per år
slutt_temp = start_temp  # Initialiserer lagertemperatur

# Simulerer for hvert år i 25-årsperioden
for ar in range(25):
    faktor = 1 - steg_reduksjon * ar  # Reduksjonsfaktor for solproduksjon
    tap_per_ar_kWh = tap_per_ar_liste[ar]  # Årlig varmetap

    # Beregner starttemperatur etter å ha trukket fra varmetap
    temp_start = slutt_temp - tap_per_ar_kWh / varmekapasitet_kWh_per_K
    temp_start = max(temp_start, start_temp)  # Begrens til minimum starttemp

    lager_full = temp_start >= 55  # Sjekker om lageret allerede er fullt ved start
    uttak_aktivert = (ar >= 3)  # Uttak skjer kun fra år 4 og utover

    # Initialiserer kolonner for simulering
    df["PV_korrigert"] = opprinnelig_PV * faktor
    for col in ["VP_drift_kW", "VP_varme_kWh", "Lager_Temp_C",
                "PV_til_bygg", "PV_til_nett", "Besparelse_i_bygg", "Salgsinntekt_nett"]:
        df[col] = 0.0

    temp = temp_start  # Init. temperatur
    total_varme = 0.0
    total_driftstimer = 0
    total_el_forbruk = 0.0
    total_el_forbruk_VP = 0.0

    # --- Lading og drift per time ---
    for i, row in df.iterrows():
        PV = row["PV_korrigert"]
        T = row["Temperatur"]
        bygglast = row["Bygglast"]
        strompris = row["Strømpris"]

        VP_drift = 0.0
        VP_varme = 0.0

        # Lading skjer kun hvis lager ikke er fullt, temperaturen er høy nok, og tilstrekkelig PV er tilgjengelig
        if not lager_full and T >= temperaturkrav_tørrkjøler and PV >= System_min_kW:
            VP_drift = min(PV, System_max_kW)
            VP_varme = VP_drift * (VP_max_kW / System_max_kW) * COP
            total_varme += VP_varme
            total_el_forbruk += VP_drift
            total_el_forbruk_VP += VP_varme / COP
            total_driftstimer += 1

            temp += VP_varme / varmekapasitet_kWh_per_K
            if temp >= 55:
                temp = 55
                lager_full = True

        # Fordeler resterende PV til bygg og nett
        PV_etter_VP = max(0, PV - VP_drift)
        PV_til_bygg = min(PV_etter_VP, bygglast)
        PV_til_nett = PV_etter_VP - PV_til_bygg
        besparelse_bygg = PV_til_bygg * (strompris + energiledd_kjop)

        # Registrerer timesverdier i datasettet
        df.at[i, "VP_drift_kW"] = VP_drift
        df.at[i, "VP_varme_kWh"] = VP_varme
        df.at[i, "Lager_Temp_C"] = temp
        df.at[i, "PV_til_bygg"] = PV_til_bygg
        df.at[i, "PV_til_nett"] = PV_til_nett
        df.at[i, "Besparelse_i_bygg"] = besparelse_bygg

    temp_etter_lading = temp  # Temperatur etter lading

    # --- Uttak fra lager ---
    uttak = 0
    if uttak_aktivert and temp > 35:
        maks_mulig_uttak = varmekapasitet_kWh_per_K * (temp - 35)
        uttak = min(varmebehov_bygg, maks_mulig_uttak)
        temp -= uttak / varmekapasitet_kWh_per_K

    slutt_temp = temp  # Slutt-temperatur etter drift

    # --- Salgsinntekt ---
    midlere_innmating = df["PV_til_nett"].mean()
    fastledd_time = midlere_innmating * fastledd_salg
    df["Salgsinntekt_nett"] = df["PV_til_nett"] * (df["Strømpris"] - energiledd_salg) - fastledd_time

    # Legger årsresultater i listen
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
resultater_df = pd.DataFrame(arsresultater)  # Samler alle årsresultater i én DataFrame
print(resultater_df)  # Skriver til terminal for visuell sjekk

# --- Eksport til Excel ---
# Skriver resultatene til nytt ark i eksisterende Excel-fil
with pd.ExcelWriter(filsti, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    resultater_df.to_excel(writer, sheet_name="SIM_BP2", index=False)
