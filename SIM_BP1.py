import pandas as pd  # Importerer pandas for databehandling

# --- Innstillinger ---
filsti = "BTES_simulering.xlsx"  # Filsti til Excel-arbeidsbok
arknavn = "Inputdata til Python simulering"  # Ark med timesdata for simulering

""" Effektdata """
VP_max_kW = 537.1  # Maksimal effekt fra varmepumpen (kW)
sirkulasjonspumpe = 17.18  # Effektforbruk sirkulasjonspumpe (kW)
tørrkjøler_kW = 10.12  # Effekt per tørrkjøler (kW)
ant_tørrkjølere = 4  # Antall tørrkjølere i systemet

""" Drift og systemdata """
tot_tørrkjøler_kW = tørrkjøler_kW * ant_tørrkjølere  # Total effekt fra alle tørrkjølere (kW)
System_max_kW = VP_max_kW + sirkulasjonspumpe + tot_tørrkjøler_kW  # Total makseffekt fra hele systemet (kW)
System_min_kW = System_max_kW * 0.15  # Minimum effekt for at systemet skal kunne operere (kW)
COP = 3.7  # Varmefaktor (Coefficient of Performance) for varmepumpen
varmekapasitet_kWh_per_K = 153461  # Varmelagringskapasitet per grad temperaturendring (kWh/°C)
start_temp = 7.2  # Starttemperatur i lageret ved begynnelsen av simuleringen (°C)
temperaturkrav_tørrkjøler = 13  # Minimum utetemperatur for drift av tørrkjøler (°C)

""" Nettleie og avgifter """
energiledd_kjop = 0.05  # Energiledd ved kjøp av strøm (kr/kWh)
energiledd_salg = -0.05  # Energiledd ved salg til nettet (kr/kWh)
fastledd_salg = 0.0198  # Fastledd ved innmating (kr/kWh gjennomsnittlig innmatingseffekt)

""" Kalkulering av solcelledegradering over 25 år """
PV_ar_1 = 0.9850  # Effektfaktor etter år 1
PV_ar_25 = 0.8890  # Effektfaktor etter år 25
steg_reduksjon = (PV_ar_1 - PV_ar_25) / 24  # Lineær reduksjon i produksjon per år

""" Årlige varmetap (kWh) fra termisk lager """
tap_per_ar_liste = [
    0, 0, 0, 1990622, 1572758, 1386540, 1275190, 1199096, 1142917, 1099286, 1064169, 1035140, 1010646, 989637, 
    971375, 955323, 941081, 928344, 916874, 906482, 897017, 888356, 880397, 873056, 866261, 859952, 854080, 848598
]

""" Årlig varmebehov i bygget """
varmebehov_bygg = 1_528_173  # Total varmeleveransebehov for bygget per år (kWh)

""" Innhenting av data fra Excel """
df = pd.read_excel(filsti, sheet_name=arknavn)  # Leser inn timesdata fra Excel-arket

# --- Dataforberedelse ---
df["Tid"] = pd.to_datetime(df["Tid"])  # Konverterer tid til datetime-format
df = df[(df["Tid"].dt.month >= 4) & (df["Tid"].dt.month <= 10)].copy()  # Filtrerer kun sommerhalvåret (april–oktober)
opprinnelig_PV = df["GridExport [kWh]"].copy()  # Lagrer opprinnelig solproduksjon for bruk i hver simulert årgang

# --- Simulering av varmepumpe og varmelagring ---
arsresultater = []  # Liste for å lagre resultater per år
slutt_temp = start_temp  # Initierer lagertemperatur med valgt startverdi

# --- Simulerer 25-årig drift ---
for ar in range(25):
    faktor = 1 - steg_reduksjon * ar  # Reduksjon i solproduksjon pga. degradering
    tap_per_ar_kWh = tap_per_ar_liste[ar]  # Varme som går tapt fra lageret i år n

    # Beregner starttemperatur etter at varmetapet er trukket fra
    temp_start = slutt_temp - tap_per_ar_kWh / varmekapasitet_kWh_per_K
    temp_start = max(temp_start, start_temp)  # Unngår at temperaturen faller under startverdi

    lager_full = temp_start >= 55  # Lageret regnes som fullt ved 55 °C
    uttak_aktivert = (ar >= 3)  # Uttak fra lager er aktivert fra og med år 4

    # Oppdaterer solproduksjon med degraderingsfaktor og nullstiller resultatkolonner
    df["PV_korrigert"] = opprinnelig_PV * faktor
    for col in ["VP_drift_kW", "VP_varme_kWh", "Lager_Temp_C",
                "PV_til_bygg", "PV_til_nett", "Besparelse_i_bygg", "Salgsinntekt_nett"]:
        df[col] = 0.0

    # Initierer variable for akkumulert energiflyt
    temp = temp_start
    total_varme = 0.0
    total_driftstimer = 0
    total_el_forbruk = 0.0
    total_el_forbruk_VP = 0.0

    # --- Lading og drift per time ---
    for i, row in df.iterrows():
        # Leser inn timeverdier
        PV = row["PV_korrigert"]
        T = row["Temperatur"]
        bygglast = row["Bygglast"]
        strompris = row["Strømpris"]

        VP_drift = 0.0
        VP_varme = 0.0

        # Lagring skjer kun dersom lageret ikke er fullt og alle betingelser er oppfylt
        if not lager_full and T >= temperaturkrav_tørrkjøler and PV >= System_min_kW:
            VP_drift = min(PV, System_max_kW)
            VP_varme = VP_drift * (VP_max_kW / System_max_kW) * COP
            total_varme += VP_varme
            total_el_forbruk += VP_drift
            total_el_forbruk_VP += VP_varme / COP
            total_driftstimer += 1

            # Temperaturøkning i lageret
            temp += VP_varme / varmekapasitet_kWh_per_K
            if temp >= 55:
                temp = 55
                lager_full = True

        # Resterende solenergi fordeles til bygg og eventuell overskuddsmating
        PV_etter_VP = max(0, PV - VP_drift)
        PV_til_bygg = min(PV_etter_VP, bygglast)
        PV_til_nett = PV_etter_VP - PV_til_bygg
        besparelse_bygg = PV_til_bygg * (strompris + energiledd_kjop)

        # Skriver resultatene til DataFrame
        df.at[i, "VP_drift_kW"] = VP_drift
        df.at[i, "VP_varme_kWh"] = VP_varme
        df.at[i, "Lager_Temp_C"] = temp
        df.at[i, "PV_til_bygg"] = PV_til_bygg
        df.at[i, "PV_til_nett"] = PV_til_nett
        df.at[i, "Besparelse_i_bygg"] = besparelse_bygg

    temp_etter_lading = temp  # Temperatur etter lading er gjennomført

    # --- Uttak fra lager ---
    uttak = 0
    if uttak_aktivert and temp > 35:
        maks_mulig_uttak = varmekapasitet_kWh_per_K * (temp - 35)  # Beregner hvor mye som kan tas ut
        uttak = min(varmebehov_bygg, maks_mulig_uttak)
        temp -= uttak / varmekapasitet_kWh_per_K  # Reduserer temperatur tilsvarende uttaket

    slutt_temp = temp  # Lagertemperatur ved slutten av året

    # --- Salgsinntekt fra overskudd til nett ---
    midlere_innmating = df["PV_til_nett"].mean()
    fastledd_time = midlere_innmating * fastledd_salg
    df["Salgsinntekt_nett"] = df["PV_til_nett"] * (df["Strømpris"] - energiledd_salg) - fastledd_time

    # Legger resultater for året til listen
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

# --- Oppsummering av resultater ---
resultater_df = pd.DataFrame(arsresultater)  # Samler alle resultater i én DataFrame
print(resultater_df)  # Skriver resultatene til konsoll

# --- Eksport til Excel ---
with pd.ExcelWriter(filsti, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    resultater_df.to_excel(writer, sheet_name="SIM_BP1", index=False)  # Skriver til nytt ark i regnearket
