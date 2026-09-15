import streamlit as st
import pandas as pd
from datetime import datetime
from PIL import Image
import os
import openpyxl

# Set overall page options
st.set_page_config(page_title="Magacin Obuće", layout="wide")

# Custom CSS styling
st.markdown("""
<style>
    /* Styling headers */
    .stHeader { font-family: 'Helvetica', sans-serif; color: #1E3A8A; }
    
    /* Input fields styling */
    div[data-baseweb="input"] { border-radius: 8px; }
    
    /* Table headers fixed visibility */
    .stDataFrame table thead th {
        background-color: #f0f2f6 !important;
        color: #333333 !important;
    }
</style>
""", unsafe_allow_html=True)

# Application Header
st.title("👟 Magacin Obuće - Upravljanje Zalihama")

# ==========================================
# CONSTANTS & CONFIGURATION
# ==========================================
G_SHEETS_MAGACIN = "https://docs.google.com/spreadsheets/d/1Sj0bX19p9J-uO5N6eT7SstH1Rstx2_J-B1N99_L4078/export?format=csv&gid=0"
G_SHEETS_DNEVNI_ULAZ = "https://docs.google.com/spreadsheets/d/1Sj0bX19p9J-uO5N6eT7SstH1Rstx2_J-B1N99_L4078/export?format=csv&gid=111111111"  # Zameniti pravim GID-om
G_SHEETS_DNEVNI_IZLAZ = "https://docs.google.com/spreadsheets/d/1Sj0bX19p9J-uO5N6eT7SstH1Rstx2_J-B1N99_L4078/export?format=csv&gid=222222222" # Zameniti pravim GID-om

LOKALNI_MAGACIN = "magacin_obuce.xlsx"
LOKALNI_ULAZ = "dnevni_ulaz.xlsx"
LOKALNI_IZLAZ = "dnevni_izlaz.xlsx"

SVETLE_BOJE = ["BELA", "BEŽ", "BEZ", "SREBRNA", "ŽUTA", "ZUTA", "SVETLO PLAVA", "SVETLO SIVA", "ROZE"]
SEZONE = ["PROLEĆE/LETO", "JESEN/ZIMA"]

# Initializing Session State for inputs
if 'reset_kolicina' not in st.session_state:
    st.session_state['reset_kolicina'] = 0
if 'reset_nabavna' not in st.session_state:
    st.session_state['reset_nabavna'] = 0
if 'reset_prodajna' not in st.session_state:
    st.session_state['reset_prodajna'] = 0
if 'reset_internet' not in st.session_state:
    st.session_state['reset_internet'] = 0

if 'reset_izlaz_kolicina' not in st.session_state:
    st.session_state['reset_izlaz_kolicina'] = 0
if 'reset_izlaz_nabavna' not in st.session_state:
    st.session_state['reset_izlaz_nabavna'] = 0

# ==========================================
# HELPER FUNCTIONS
# ==========================================
@st.cache_data(ttl=60)
def ucitaj_podatke():
    try:
        df_m = pd.read_csv(G_SHEETS_MAGACIN)
        df_u = pd.read_csv(G_SHEETS_DNEVNI_ULAZ)
        df_i = pd.read_csv(G_SHEETS_DNEVNI_IZLAZ)
    except Exception:
        if os.path.exists(LOKALNI_MAGACIN):
            df_m = pd.read_excel(LOKALNI_MAGACIN)
        else:
            df_m = pd.DataFrame(columns=[
                "sifra_modela", "sezona", "boja", "svetao_djon", "velicina", 
                "stanje", "nabavna_cena", "prodajna_cena", "internet_cena", 
                "grad", "pripada_radnji", "slika_path"
            ])
            
        if os.path.exists(LOKALNI_ULAZ):
            df_u = pd.read_excel(LOKALNI_ULAZ)
        else:
            df_u = pd.DataFrame(columns=[
                "datum", "sifra_modela", "sezona", "boja", "velicina", 
                "kolicina", "nabavna_cena", "prodajna_cena", "internet_cena", "grad"
            ])
            
        if os.path.exists(LOKALNI_IZLAZ):
            df_i = pd.read_excel(LOKALNI_IZLAZ)
        else:
            df_i = pd.DataFrame(columns=[
                "datum", "sifra_modela", "sezona", "boja", "velicina", 
                "kolicina", "prodajna_cena", "nabavna_cena", "grad"
            ])
            
    # Standardizacija naziva i tipova
    for df in [df_m, df_u, df_i]:
        if not df.empty:
            if "boja" in df.columns:
                df["boja"] = df["boja"].astype(str).str.upper()
            if "sifra_modela" in df.columns:
                df["sifra_modela"] = df["sifra_modela"].astype(str)
            if "sezona" in df.columns:
                df["sezona"] = df["sezona"].astype(str)
                
    return df_m, df_u, df_i

def sacuvaj_podatke(df_magacin, df_ulaz, df_izlaz):
    df_magacin.to_excel(LOKALNI_MAGACIN, index=False)
    df_ulaz.to_excel(LOKALNI_ULAZ, index=False)
    df_izlaz.to_excel(LOKALNI_IZLAZ, index=False)
    st.cache_data.clear()

# Učitavanje inicijalnih podataka
df_magacin, df_ulaz, df_izlaz = ucitaj_podatke()

# Priprema objedinjenih podataka za selekcije
if not df_magacin.empty:
    df_svi_artikli = df_magacin[["sifra_modela", "sezona", "boja", "prodajna_cena", "internet_cena"]].drop_duplicates()
    df_svi_artikli["Sifra_Ispis"] = df_svi_artikli["sifra_modela"] + " (" + df_svi_artikli["sezona"] + ")"
    
    mapa_sifra_sezona = dict(zip(df_svi_artikli["Sifra_Ispis"], df_svi_artikli["sifra_modela"]))
    opcije_sifre = sorted(df_svi_artikli["Sifra_Ispis"].unique().tolist())
else:
    df_svi_artikli = pd.DataFrame()
    mapa_sifra_sezona = {}
    opcije_sifre = []

lista_gradova = ["Internet", "Beograd", "Novi Sad", "Niš", "Kragujevac", "Mladenovac"]

# ==========================================
# CALLBACK FUNKCIJE
# ==========================================
def azuriraj_cenu_izlaza():
    """Aktivira se na promenu šifre, boje ili grada u izlazu i postavlja fabričku cenu."""
    izabrani_label = st.session_state.get("jedinstveni_izlaz_sifra")
    izabrani_g = st.session_state.get("jedinstveni_izlaz_grad", "Internet")
    izabrana_b = st.session_state.get("jedinstveni_izlaz_boja")

    if izabrani_label and not df_svi_artikli.empty:
        artikli_filtrirani = df_svi_artikli[
            (df_svi_artikli["Sifra_Ispis"] == izabrani_label) & 
            (df_svi_artikli["boja"] == izabrana_b)
        ]
        if not artikli_filtrirani.empty:
            if izabrani_g == "Internet":
                nova_c = float(artikli_filtrirani.iloc[0]["internet_cena"])
            else:
                nova_c = float(artikli_filtrirani.iloc[0]["prodajna_cena"])
            st.session_state["prodajna_cena_input"] = nova_c

# Inicijalizacija ključa cene
if "prodajna_cena_input" not in st.session_state:
    st.session_state["prodajna_cena_input"] = 0.0

# ==========================================
# TABOVI APLIKACIJE
# ==========================================
tab1, tab2, tab3 = st.tabs(["📋 Unos Novog Modela / Ulaz", "📤 Evidencija Izlaza", "📊 Stanje na Magacinu"])

# ------------------------------------------
# TAB 1: UNOS ULAZA
# ------------------------------------------
with tab1:
    st.header("Unos novog modela ili dopuna zaliha")
    
    col1, col2 = st.columns(2)
    with col1:
        sifra_modela = st.text_input("Šifra modela:").strip().upper()
        sezona = st.selectbox("Sezona:", SEZONE)
        boja = st.text_input("Boja:").strip().upper()
        
        svetao_djon = st.checkbox("Svetao đon?") if boja in SVETLE_BOJE else False
        velicina = st.number_input("Veličina:", min_value=35, max_value=48, step=1)

    with col2:
        kljuc_kolicina = f"kolicina_{st.session_state['reset_kolicina']}"
        kljuc_nabavna = f"nabavna_{st.session_state['reset_nabavna']}"
        kljuc_prodajna = f"prodajna_{st.session_state['reset_prodajna']}"
        kljuc_internet = f"internet_{st.session_state['reset_internet']}"
        
        kolicina = st.number_input("Količina:", min_value=1, step=1, value=None, key=kljuc_kolicina)
        nabavna_cena = st.number_input("Nabavna cena (RSD):", min_value=0.0, step=50.0, value=None, key=kljuc_nabavna)
        prodajna_cena = st.number_input("Prodajna cena (RSD):", min_value=0.0, step=50.0, value=None, key=kljuc_prodajna)
        internet_cena = st.number_input("Internet cena (RSD):", min_value=0.0, step=50.0, value=None, key=kljuc_internet)
        grad = st.selectbox("Grad / Lokacija:", lista_gradova)
        
    slika_file = st.file_uploader("Otpremi sliku modela (opciono):", type=["jpg", "jpeg", "png"])

    if st.button("Sačuvaj Ulaz", type="primary"):
        if not sifra_modela or not boja or kolicina is None or nabavna_cena is None or prodajna_cena is None or internet_cena is None:
            st.error("Molimo popunite sva obavezna polja!")
        else:
            slika_path = ""
            if slika_file is not None:
                os.makedirs("slike", exist_ok=True)
                slika_path = os.path.join("slike", f"{sifra_modela}_{boja}.png")
                img = Image.open(slika_file)
                img.save(slika_path)

            pripada_radnji = grad if grad != "Internet" else "Magacin"

            # 1. Provera i ažuriranje glavnog magacina
            maska = (
                (df_magacin["sifra_modela"] == sifra_modela) & 
                (df_magacin["sezona"] == sezona) & 
                (df_magacin["boja"] == boja) & 
                (df_magacin["velicina"] == velicina) & 
                (df_magacin["grad"] == grad)
            )
            
            if not df_magacin[maska].empty:
                df_magacin.loc[maska, "stanje"] += kolicina
                df_magacin.loc[maska, "nabavna_cena"] = nabavna_cena
                df_magacin.loc[maska, "prodajna_cena"] = prodajna_cena
                df_magacin.loc[maska, "internet_cena"] = internet_cena
                if slika_path:
                    df_magacin.loc[maska, "slika_path"] = slika_path
            else:
                novi_red_m = {
                    "sifra_modela": sifra_modela, "sezona": sezona, "boja": boja,
                    "svetao_djon": svetao_djon, "velicina": velicina, "stanje": kolicina,
                    "nabavna_cena": nabavna_cena, "prodajna_cena": prodajna_cena,
                    "internet_cena": internet_cena, "grad": grad,
                    "pripada_radnji": pripada_radnji, "slika_path": slika_path
                }
                df_magacin = pd.concat([df_magacin, pd.DataFrame([novi_red_m])], ignore_index=True)

            # 2. Dnevni ulaz
            novi_red_u = {
                "datum": datetime.now().strftime("%Y-%m-%d %H:%M"),
                "sifra_modela": sifra_modela, "sezona": sezona, "boja": boja,
                "velicina": velicina, "kolicina": kolicina, "nabavna_cena": nabavna_cena,
                "prodajna_cena": prodajna_cena, "internet_cena": internet_cena, "grad": grad
            }
            df_ulaz = pd.concat([df_ulaz, pd.DataFrame([novi_red_u])], ignore_index=True)

            # Resetovanje formi
            st.session_state['reset_kolicina'] += 1
            st.session_state['reset_nabavna'] += 1
            st.session_state['reset_prodajna'] += 1
            st.session_state['reset_internet'] += 1

            sacuvaj_podatke(df_magacin, df_ulaz, df_izlaz)
            st.success(f"Uspešno dodat ulaz za model {sifra_modela} ({boja}, vel. {velicina})!")
            st.rerun()

# ------------------------------------------
# TAB 2: EVIDENCIJA IZLAZA
# ------------------------------------------
with tab2:
    st.header("Evidencija izlaza robe (Prodaja / Transfer)")

    if opcije_sifre:
        col1, col2 = st.columns(2)
        with col1:
            izabrani_datum = st.date_input("Izaberi datum izlaza:", datetime.now())
            
            # Selectbox za Šifru sa ON_CHANGE
            izabrana_sifra_label = st.selectbox(
                "Izaberi šifru modela:", 
                opcije_sifre, 
                key="jedinstveni_izlaz_sifra",
                on_change=azuriraj_cenu_izlaza
            )
            prava_sifra = mapa_sifra_sezona[izabrana_sifra_label]
            
            artikli_za_sifru = df_svi_artikli[df_svi_artikli["Sifra_Ispis"] == izabrana_sifra_label]
            boje_za_sifru = sorted(artikli_za_sifru["boja"].unique().tolist())
            
            # Selectbox za Boju sa ON_CHANGE
            izabrana_boja = st.selectbox(
                "Izaberi boju modela:", 
                boje_za_sifru, 
                key="jedinstveni_izlaz_boja",
                on_change=azuriraj_cenu_izlaza
            )
            
            # Selectbox za Grad sa ON_CHANGE
            izabrani_grad = st.selectbox(
                "Izaberi grad:", 
                lista_gradova,
                key="jedinstveni_izlaz_grad",
                on_change=azuriraj_cenu_izlaza
            )
            
            maska_dostupne_vel = (
                (df_magacin["sifra_modela"] == prava_sifra) & 
                (df_magacin["boja"] == izabrana_boja) & 
                (df_magacin["grad"] == izabrani_grad) & 
                (df_magacin["stanje"] > 0)
            )
            dostupne_velicine = sorted(df_magacin[maska_dostupne_vel]["velicina"].unique().tolist())
            
            izabrana_velicina = st.selectbox("Izaberi veličinu:", dostupne_velicine if dostupne_velicine else ["Nema na stanju"])

        with col2:
            kljuc_kolicina_izlaza = f"kolicina_izlaz_{st.session_state['reset_izlaz_kolicina']}"
            kljuc_nabavna_cena = f"nabavna_cena_{st.session_state['reset_izlaz_nabavna']}"
            
            kolicina_izlaza = st.number_input("Količina za izlaz:", min_value=1, step=1, value=None, key=kljuc_kolicina_izlaza)
            
            # Prodajna cena spregnuta sa session_state
            prodajna_cena_par = st.number_input(
                "Prodajna cena (RSD):", 
                min_value=0.0, 
                step=50.0, 
                key="prodajna_cena_input"
            )
            
            nabavna_cena_par = st.number_input("Nabavna cena (Opciono):", min_value=0.0, step=50.0, value=None, key=kljuc_nabavna_cena)

        if st.button("Proknjiži Izlaz", type="primary"):
            if izabrana_velicina == "Nema na stanju":
                st.error("Izabrani artikal nema na stanju na toj lokaciji!")
            elif kolicina_izlaza is None:
                st.error("Unesite količinu za izlaz!")
            else:
                maska_zaliha = (
                    (df_magacin["sifra_modela"] == prava_sifra) & 
                    (df_magacin["boja"] == izabrana_boja) & 
                    (df_magacin["velicina"] == izabrana_velicina) & 
                    (df_magacin["grad"] == izabrani_grad)
                )
                
                trenutno_stanje = df_magacin.loc[maska_zaliha, "stanje"].values[0] if not df_magacin[maska_zaliha].empty else 0
                
                if trenutno_stanje < kolicina_izlaza:
                    st.error(f"Nedovoljno na stanju! Trenutno stanje: {trenutno_stanje}")
                else:
                    # Skini sa stanja
                    df_magacin.loc[maska_zaliha, "stanje"] -= kolicina_izlaza
                    
                    # Upisi u dnevni izlaz
                    sezona_val = df_magacin.loc[maska_zaliha, "sezona"].values[0]
                    novi_izlaz = {
                        "datum": izabrani_datum.strftime("%Y-%m-%d"),
                        "sifra_modela": prava_sifra,
                        "sezona": sezona_val,
                        "boja": izabrana_boja,
                        "velicina": izabrana_velicina,
                        "kolicina": kolicina_izlaza,
                        "prodajna_cena": prodajna_cena_par,
                        "nabavna_cena": nabavna_cena_par if nabavna_cena_par is not None else 0.0,
                        "grad": izabrani_grad
                    }
                    df_izlaz = pd.concat([df_izlaz, pd.DataFrame([novi_izlaz])], ignore_index=True)
                    
                    st.session_state['reset_izlaz_kolicina'] += 1
                    st.session_state['reset_izlaz_nabavna'] += 1
                    
                    sacuvaj_podatke(df_magacin, df_ulaz, df_izlaz)
                    st.success("Izlaz uspešno proknjižen!")
                    st.rerun()
    else:
        st.info("Magacin je trenutno prazan. Unesite prve artikle u Tabu 1.")

# ------------------------------------------
# TAB 3: STANJE NA MAGACINU
# ------------------------------------------
with tab3:
    st.header("Pregled trenutnog stanja")
    
    if not df_magacin.empty:
        col_f1, col_f2, col_f3 = st.columns(3)
        with col_f1:
            filter_grad = st.multiselect("Filtriraj po lokaciji:", df_magacin["grad"].unique())
        with col_f2:
            filter_sezona = st.multiselect("Filtriraj po sezoni:", df_magacin["sezona"].unique())
        with col_f3:
            filter_sifra = st.text_input("Pretraži po šifri:").strip().upper()
            
        prikaz_df = df_magacin.copy()
        
        if filter_grad:
            prikaz_df = prikaz_df[prikaz_df["grad"].isin(filter_grad)]
        if filter_sezona:
            prikaz_df = prikaz_df[prikaz_df["sezona"].isin(filter_sezona)]
        if filter_sifra:
            prikaz_df = prikaz_df[prikaz_df["sifra_modela"].str.contains(filter_sifra, na=False)]
            
        st.dataframe(prikaz_df, use_container_width=True)
        
        # Prikaz ukupnih suma
        st.markdown("---")
        m_col1, m_col2 = st.columns(2)
        m_col1.metric("Ukupno pari na stanju:", int(prikaz_df["stanje"].sum()))
        m_col2.metric("Ukupna nabavna vrednost:", f"{float((prikaz_df['stanje'] * prikaz_df['nabavna_cena']).sum()):,.2f} RSD")
    else:
        st.info("Nema podataka za prikaz.")