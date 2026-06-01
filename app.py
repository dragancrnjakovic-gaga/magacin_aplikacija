import streamlit as st
import psycopg2
from psycopg2.extras import RealDictCursor
import cloudinary
import cloudinary.uploader
import pandas as pd
from datetime import datetime
import io

# --- KONFIGURACIJA CLOUDINARY-JA ---
cloudinary.config(
    cloud_name = st.secrets["cloudinary"]["cloud_name"],
    api_key = st.secrets["cloudinary"]["api_key"],
    api_secret = st.secrets["cloudinary"]["api_secret"],
    secure = True
)

# --- PODEŠAVANJE NEON POSTGRES BAZE ---
def uzmi_vezu_sa_bazom():
    return psycopg2.connect(st.secrets["postgres"]["url"])

def kreiraj_tabele():
    conn = uzmi_vezu_sa_bazom()
    cursor = conn.cursor()
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS artikli (
            sifra TEXT,
            boja TEXT,
            sezona TEXT DEFAULT 'Proleće-Leto',
            broj_pari INTEGER,
            pari_u_kutiji INTEGER,
            prodajna_cena REAL,
            internet_cena REAL,
            slika_putanja TEXT,
            PRIMARY KEY (sifra, boja)
        )
    ''')
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS izlaz_robe (
            id SERIAL PRIMARY KEY,
            datum TEXT,
            sifra_artikla TEXT,
            boja_artikla TEXT,
            kolicina_izlaz INTEGER
        )
    ''')
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS sifrarnik_boja (
            boja TEXT PRIMARY KEY
        )
    ''')
    
    cursor.execute("SELECT COUNT(*) FROM sifrarnik_boja")
    if cursor.fetchone()[0] == 0:
        pocetne_boje = [("Black",), ("Blue",), ("Red",), ("Gray",), ("White",), ("Beige",)]
        cursor.executemany("INSERT INTO sifrarnik_boja (boja) VALUES (%s)", pocetne_boje)
        
    conn.commit()
    conn.close()

# Automatski kreiramo strukturu na internetu ako ne postoji
kreiraj_tabele()

# --- NAPREDNO KEŠIRANJE PODATAKA ---
@st.cache_data(ttl=300)
def ucitaj_artikle_za_sezonu(sezona):
    conn = uzmi_vezu_sa_bazom()
    df = pd.read_sql_query("SELECT * FROM artikli WHERE sezona = %s ORDER BY sifra ASC, boja ASC", conn, params=(sezona,))
    conn.close()
    return df

@st.cache_data(ttl=600)
def ucitaj_boje():
    conn = uzmi_vezu_sa_bazom()
    cursor = conn.cursor()
    cursor.execute("SELECT boja FROM sifrarnik_boja ORDER BY boja ASC")
    boje = [red[0] for red in cursor.fetchall()]
    conn.close()
    return boje

def konvertuj_u_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Magacin')
    return output.getvalue()

# --- POMOĆNA FUNKCIJA: Optimizovana pretraga slika u memoriji ---
def pronadji_sliku_u_df(df, sifra):
    if df.empty:
        return ""
    filtrirano = df[(df["sifra"] == sifra) & (df["slika_putanja"] != "") & (df["slika_putanja"].notna())]
    if not filtrirano.empty:
        return filtrirano.iloc[0]["slika_putanja"]
    return ""

# --- IZGLED I STILIZACIJA APLIKACIJE ---
st.set_page_config(page_title="Magacin", layout="wide")

st.markdown("""
    <style>
    .block-container {
        padding-top: 3.5rem !important;
        padding-bottom: 2rem !important;
    }
    h1 { font-size: 1.8rem !important; padding-bottom: 10px !important; margin: 0px !important; }
    h2 { font-size: 1.35rem !important; padding-bottom: 15px !important; margin: 0px !important; }
    h3 { font-size: 1.05rem !important; font-weight: bold !important; }
    [data-testid="stMetricValue"] { font-size: 1.05rem !important; }
    [data-testid="stMetricLabel"] { font-size: 0.75rem !important; }
    .stTextInput p, .stNumberInput p, .stSelectbox p, .stDateInput p, label p { font-size: 0.85rem !important; }
    .stAlert p { font-size: 0.85rem !important; }
    .stExpander p { font-size: 0.8rem !important; }
    div[data-testid="stHorizontalBlock"] { background: #1e2229; padding: 15px; border-radius: 6px; margin-bottom: 10px; border: 1px solid #2d3139; }
    
    /* Stilizacija dugmadi za stranice da izgledaju lepše, ujednačenije i kompaktnije */
    div.stButton > button {
        padding: 2px 10px !important;
        font-size: 0.85rem !important;
        min-width: 40px !important;
        text-align: center !important;
    }
    </style>
""", unsafe_allow_html=True)

# ⚡ ZVANIČNI SKRIPT ZA SKROL NA VRH
if "skroluj_na_vrh" in st.session_state and st.session_state["skroluj_na_vrh"]:
    st.components.v1.html(
        "<script>window.parent.document.querySelector('.stMain').scrollTo(0, 0);</script>",
        height=0,
        width=0,
    )
    st.session_state["skroluj_na_vrh"] = False

st.title("📦 Višekorisnički sistem za praćenje stanja u magacinu")

# 1. KATEGORIJA / SEZONA
izabrana_sezona = st.sidebar.radio("🌸 IZABERI KATEGORIJU / SEZONU:", ["Proleće-Leto", "Jesen-Zima", "Torbe"])
st.sidebar.markdown("---")

# 2. FUNKCIJA
meni = st.sidebar.selectbox("Izaberi opciju:", ["Trenutno stanje", "Unos nove robe", "Evidencija izlaza (Po danima)"])
st.sidebar.info(f"Trenutno radite u sekciji:\n**{izabrana_sezona}**")

# Resetujemo stranicu na 1 ako se promeni meni
if meni != "Trenutno stanje":
    st.session_state["trenutna_stranica"] = 1

if "reset_brojac" not in st.session_state:
    st.session_state["reset_brojac"] = 0

# --- PREFINJENA POMOĆNA FUNKCIJA ZA PRIKAZ PAGINACIJE ---
def prikazi_brojeve_stranica(broj_stranica, trenutna, kljuc_prefiks):
    vidljivi_brojevi = set()
    
    # Ako je ukupan broj stranica mali (npr. do 12), prikaži ih sve bez ikakvog skraćivanja
    if broj_stranica <= 12:
        for i in range(1, broj_stranica + 1):
            vidljivi_brojevi.add(i)
    else:
        # Uvek prikaži prve tri stranice stabilno (1, 2, 3)
        for i in range(1, 4):
            vidljivi_brojevi.add(i)
            
        # Ako smo blizu početka (stranice 1-6), prikaži stabilno ceo blok do 8 da nema "skakutanja" boks-a
        if trenutna <= 6:
            for i in range(1, 9):
                vidljivi_brojevi.add(i)
        # Ako smo negde u sredini ili prema kraju, dinamički širi opseg oko trenutne stranice
        else:
            for i in range(max(1, trenutna - 3), min(broj_stranica + 1, trenutna + 4)):
                vidljivi_brojevi.add(i)
                
        # Uvek prikaži poslednju stranicu na kraju niza
        vidljivi_brojevi.add(broj_stranica)
    
    sortirani_brojevi = sorted(list(vidljivi_brojevi))
    
    # Formatiranje prikaza sa tri tačke na prirodnim mestima
    ekran_lista = []
    prethodni = 0
    for br in sortirani_brojevi:
        if br - prethodni > 1:
            ekran_lista.append("...")
        ekran_lista.append(br)
        prethodni = br

    # Generisanje kolona u Streamlit-u na osnovu broja elemenata
    cols = st.columns(len(ekran_lista) + 2)
    
    # 1. Leva strelica
    with cols[0]:
        if st.button("⬅️", disabled=(trenutna == 1), key=f"{kljuc_prefiks}_prev"):
            st.session_state["trenutna_stranica"] = trenutna - 1
            st.session_state["skroluj_na_vrh"] = True
            st.rerun()
            
    # 2. Brojevi i separator (...)
    trenutna_kol_indeks = 1
    for stavka in ekran_lista:
        with cols[trenutna_kol_indeks]:
            if stavka == "...":
                st.write("<p style='margin-top:5px; text-align:center; color:gray; font-weight:bold;'>...</p>", unsafe_allow_html=True)
            else:
                tip_dugmeta = "primary" if stavka == trenutna else "secondary"
                if st.button(str(stavka), type=tip_dugmeta, key=f"{kljuc_prefiks}_str_{stavka}"):
                    st.session_state["trenutna_stranica"] = stavka
                    st.session_state["skroluj_na_vrh"] = True
                    st.rerun()
        trenutna_kol_indeks += 1
        
    # 3. Desna strelica
    with cols[trenutna_kol_indeks]:
        if st.button("➡️", disabled=(trenutna == broj_stranica), key=f"{kljuc_prefiks}_next"):
            st.session_state["trenutna_stranica"] = trenutna + 1
            st.session_state["skroluj_na_vrh"] = True
            st.rerun()


# --- OPCIJA 1: UNOS NOVE ROBE ---
if meni == "Unos nove robe":
    st.header(f"➕ Unos novog artikla ({izabrana_sezona})")
    lista_boja = ucitaj_boje()
    
    if "unos_sifra" not in st.session_state: st.session_state["unos_sifra"] = ""
    if "unos_boja" not in st.session_state: st.session_state["unos_boja"] = lista_boja[0] if lista_boja else ""
    if "unos_kolicina" not in st.session_state: st.session_state["unos_kolicina"] = None
    if "unos_kutija" not in st.session_state: st.session_state["unos_kutija"] = None
    if "unos_prodajna" not in st.session_state: st.session_state["unos_prodajna"] = None
    if "unos_internet" not in st.session_state: st.session_state["unos_internet"] = None
    
    col1, col2 = st.columns(2)
    with col1:
        sifra = st.text_input("Šifra modela:", value=st.session_state["unos_sifra"]).strip().upper()
        boja = st.selectbox("Boja modela:", lista_boja, index=lista_boja.index(st.session_state["unos_boja"]) if st.session_state["unos_boja"] in lista_boja else 0)
        labela_kol = "Količina (komada/pari):" if izabrana_sezona == "Torbe" else "Količina pari:"
        labela_kut = "Broj komada u jednoj kutiji/pakovanju:" if izabrana_sezona == "Torbe" else "Broj pari u jednoj kutiji:"
        
        broj_pari = st.number_input(labela_kol, min_value=0, step=1, value=st.session_state["unos_kolicina"])
        pari_u_kutiji = st.number_input(labela_kut, min_value=1, step=1, value=st.session_state["unos_kutija"])
    with col2:
        prodajna_cena = st.number_input("Prodajna cena (RSD):", min_value=0.0, step=50.0, value=st.session_state["unos_prodajna"])
        internet_cena = st.number_input("Internet cena (RSD):", min_value=0.0, step=50.0, value=st.session_state["unos_internet"])
        slika = st.file_uploader("Ubaci sliku modela (Ostavi prazno ako šifra već ima sliku):", type=["jpg", "jpeg", "png"], key=f"slika_unos_{st.session_state['reset_brojac']}")
        
    podaci_nedostaju = (
        sifra == "" or 
        boja is None or 
        boja == "" or 
        broj_pari is None or 
        pari_u_kutiji is None or 
        prodajna_cena is None or 
        internet_cena is None
    )
    
    st.write("")
    dugme_potvrdi = st.button("Sačuvaj artikal u bazu", type="primary", disabled=podaci_nedostaju)
    
    if podaci_nedostaju:
        st.caption("⚠️ Dugme će postati aktivno kada popunite Šifru, Boju, Količinu, Pakovanje i obe Cene.")
        
    if dugme_potvrdi:
        df_trenutni = ucitaj_artikle_za_sezonu(izabrana_sezona)
        url_slike = ""
        if slika is not None:
            with st.spinner("Slanje slike na Cloudinary..."):
                try:
                    rezultat_slike = cloudinary.uploader.upload(
                        slika, 
                        folder="magacin/",
                        public_id=f"{sifra}_{boja}",
                        transformation=[
                            {"width": 800, "crop": "limit"},
                            {"quality": "auto", "fetch_format": "auto"}
                        ]
                    )
                    url_slike = rezultat_slike["secure_url"]
                except Exception as e:
                    st.error(f"Greška pri slanju slike: {e}")
        else:
            url_slike = pronadji_sliku_u_df(df_trenutni, sifra)
            if url_slike != "":
                st.info("💡 Automatski je preuzeta postojeća slika za ovu šifru modela!")
        
        try:
            conn = uzmi_vezu_sa_bazom()
            cursor = conn.cursor()
            cursor.execute('''
                INSERT INTO artikli (sifra, boja, sezona, broj_pari, pari_u_kutiji, prodajna_cena, internet_cena, slika_putanja)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
            ''', (sifra, boja, izabrana_sezona, broj_pari, pari_u_kutiji, prodajna_cena, internet_cena, url_slike))
            conn.commit()
            conn.close()
            
            ucitaj_artikle_za_sezonu.clear()
            
            st.session_state["unos_sifra"] = ""
            st.session_state["unos_kolicina"] = None
            st.session_state["unos_kutija"] = None
            st.session_state["unos_prodajna"] = None
            st.session_state["unos_internet"] = None
            st.session_state["reset_brojac"] += 1
            
            st.success(f"Uspešno sačuvan model: Šifra '{sifra}' - Boja '{boja}'!")
            st.rerun()
        except psycopg2.IntegrityError:
            st.error(f"Greška: Model sa šifrom '{sifra}' u boji '{boja}' već postoji u ovoj sekciji!")

    st.markdown("---")
    st.subheader("🎨 Upravljanje listom boja")
    col_nova_boja, col_dugme_boja = st.columns([3, 1])
    with col_nova_boja:
        nova_boja_unos = st.text_input("Unesi naziv nove boje:", "").strip().capitalize()
    with col_dugme_boja:
        st.write(""); st.write("")
        if st.button("➕ Dodaj boju u listu"):
            if nova_boja_unos != "":
                try:
                    conn = uzmi_vezu_sa_bazom()
                    cursor = conn.cursor()
                    cursor.execute("INSERT INTO sifrarnik_boja (boja) VALUES (%s)", (nova_boja_unos,))
                    conn.commit()
                    conn.close()
                    ucitaj_boje.clear()
                    st.success(f"Boja '{nova_boja_unos}' je dodata!")
                    st.rerun()
                except psycopg2.IntegrityError:
                    st.warning("Boja već postoji u listi.")

# --- OPCIJA 2: TRENUTNO STANJE ---
elif meni == "Trenutno stanje":
    st.header(f"📋 Stanje robe - Sekcija: {izabrana_sezona}")
    lista_boja = ucitaj_boje()
    
    df = ucitaj_artikle_za_sezonu(izabrana_sezona)
    
    if df.empty:
        st.info(f"U sekciji {izabrana_sezona} trenutno nema unete robe.")
    else:
        df_excel = df.copy()
        df_excel["Broj kutija"] = df_excel["broj_pari"] // df_excel["pari_u_kutiji"]
        df_excel["Ostatak"] = df_excel["broj_pari"] % df_excel["pari_u_kutiji"]
        
        relabel_kom = "Ukupno komada" if izabrana_sezona == "Torbe" else "Ukupno pari"
        relabel_kut = "Kutija/Pakovanja" if izabrana_sezona == "Torbe" else "Pari u kutiji"
        
        df_excel = df_excel.rename(columns={
            "sifra": "Šifra modela", "boja": "Boja", "sezona": "Kategorija", 
            "broj_pari": relabel_kom, "pari_u_kutiji": relabel_kut,
            "prodajna_cena": "Prodajna cena (RSD)", "internet_cena": "Internet cena (RSD)"
        }).drop(columns=["slika_putanja"], errors="ignore")
        
        excel_podaci = konvertuj_u_excel(df_excel)
        st.download_button(
            label="🟢 Preuzmi kompletno stanje kao Excel tabelu (.xlsx)",
            data=excel_podaci,
            file_name=f"stanje_{izabrana_sezona}_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        st.markdown("---")
        
        pretraga = st.text_input("🔍 Pretraži ovu sekciju po šifri modela (Pretražuje sve stranice):", "").strip().upper()
        if pretraga:
            df_prikaz = df[df["sifra"].str.contains(pretraga, na=False)]
        else:
            df_prikaz = df.copy()
        
        if df_prikaz.empty:
            st.warning(f"Nema rezultata za šifru '{pretraga}'")
        else:
            # KONFIGURACIJA STRANIČENJA
            BROJ_ARTIKALA_PO_STRANICI = 20
            ukupno_artikala = len(df_prikaz)
            broj_stranica = (ukupno_artikala // BROJ_ARTIKALA_PO_STRANICI) + (1 if ukupno_artikala % BROJ_ARTIKALA_PO_STRANICI > 0 else 0)
            
            if "trenutna_stranica" not in st.session_state or pretraga != "":
                st.session_state["trenutna_stranica"] = 1
                
            # 1. KONTROLE STRANICA NA VRHU
            if broj_stranica > 1 and not pretraga:
                st.caption(f"Ukupno pronađeno: {ukupno_artikala} modela raspoređenih na {broj_stranica} stranica.")
                prikazi_brojeve_stranica(broj_stranica, st.session_state["trenutna_stranica"], "vrh")
                st.write("")
            
            start_indeks = (st.session_state["trenutna_stranica"] - 1) * BROJ_ARTIKALA_PO_STRANICI
            kraj_indeks = start_indeks + BROJ_ARTIKALA_PO_STRANICI
            df_za_prikaz = df_prikaz.iloc[start_indeks:kraj_indeks]
            
            # Prikaz artikala na trenutnoj stranici
            for index, row in df_za_prikaz.iterrows():
                sif = row['sifra']
                boj = row['boja']
                kljuc_id = f"{sif}_{boj}"
                trenutna_slika = row["slika_putanja"]
                
                if not trenutna_slika or trenutna_slika == "":
                    trenutna_slika = pronadji_sliku_u_df(df, sif)
                
                br_kutija = row["broj_pari"] // row["pari_u_kutiji"]
                ost_pari = row["broj_pari"] % row["pari_u_kutiji"]
                
                p_cena_int = int(row['prodajna_cena'])
                i_cena_int = int(row['internet_cena'])
                
                with st.container():
                    col_slika, col_detalji, col_akcije = st.columns([1.2, 3, 1.5])
                    with col_slika:
                        if trenutna_slika:
                            mala_slika_url = trenutna_slika.replace("/upload/", "/upload/w_150,c_limit,q_auto,f_auto/")
                            st.image(mala_slika_url, width=120)
                            with st.expander("🔍 Vidi veliku sliku"):
                                st.image(trenutna_slika, use_container_width=True)
                        else:
                            st.write("❌ Nema slike")
                            
                    with col_detalji:
                        st.subheader(f"Šifra modela: {sif} | Boja: {boj}")
                        c1, c2, c3, c4 = st.columns(4)
                        
                        m_kol_label = "Ukupno komada" if izabrana_sezona == "Torbe" else "Ukupno pari na stanju"
                        m_pak_label = "Pakovanje (kut. + kom)" if izabrana_sezona == "Torbe" else "Pakovanje"
                        m_sufiks = "kom" if izabrana_sezona == "Torbe" else "kom"
                        m_pak_sufiks = f"{br_kutija} pak. + {ost_pari} kom" if izabrana_sezona == "Torbe" else f"{br_kutija} kut. + {ost_pari} par"
                        
                        c1.metric(m_kol_label, f"{row['broj_pari']} {m_sufiks}")
                        c2.metric(m_pak_label, m_pak_sufiks)
                        c3.metric("Prodajna cena", f"{p_cena_int} din")
                        c4.metric("Internet cena", f"{i_cena_int} din")
                        
                    with col_akcije:
                        ekspander = st.expander("🛠️ Izmeni / Obriši")
                        with ekspander:
                            st.write("**Uredi podatke:**")
                            
                            indeks_trenutne_boje = lista_boja.index(boj) if boj in lista_boja else 0
                            nova_boja_izmena = st.selectbox("Izmeni boju artikla:", lista_boja, index=indeks_trenutne_boje, key=f"boja_{kljuc_id}")
                            
                            labela_izmena_kol = "Novo ukupno komada:" if izabrana_sezona == "Torbe" else "Novo ukupno pari:"
                            nova_kol = st.number_input(labela_izmena_kol, min_value=0, value=int(row['broj_pari']), step=1, key=f"kol_{kljuc_id}")
                            nova_p_cena = st.number_input("Prodajna cena (RSD):", min_value=0.0, value=float(row['prodajna_cena']), step=50.0, key=f"pc_{kljuc_id}")
                            nova_i_cena = st.number_input("Internet cena (RSD):", min_value=0.0, value=float(row['internet_cena']), step=50.0, key=f"ic_{kljuc_id}")
                            nova_slika_file = st.file_uploader("Zameni sliku artikla:", type=["jpg", "jpeg", "png"], key=f"img_{kljuc_id}")
                            
                            col_b1, col_b2 = st.columns(2)
                            with col_b1:
                                if st.button("💾 Snimi", key=f"Snimi_{kljuc_id}"):
                                    finalna_putanja_slike = row["slika_putanja"]
                                    if nova_slika_file is not None:
                                        with st.spinner("Menjanje slike..."):
                                            try:
                                                rez_nove_slike = cloudinary.uploader.upload(
                                                    nova_slika_file,
                                                    folder="magacin/",
                                                    public_id=f"{sif}_{nova_boja_izmena}",
                                                    transformation=[
                                                        {"width": 800, "crop": "limit"},
                                                        {"quality": "auto", "fetch_format": "auto"}
                                                    ]
                                                )
                                                finalna_putanja_slike = rez_nove_slike["secure_url"]
                                            except:
                                                pass
                                            
                                    try:
                                        conn = uzmi_vezu_sa_bazom()
                                        cursor = conn.cursor()
                                        cursor.execute('''
                                            UPDATE artikli 
                                            SET boja = %s, broj_pari = %s, prodajna_cena = %s, internet_cena = %s, slika_putanja = %s
                                            WHERE sifra = %s AND boja = %s AND sezona = %s
                                        ''', (nova_boja_izmena, nova_kol, nova_p_cena, nova_i_cena, finalna_putanja_slike, sif, boj, izabrana_sezona))
                                        conn.commit()
                                        conn.close()
                                        
                                        ucitaj_artikle_za_sezonu.clear()
                                        st.success("Izmenjeno!")
                                        st.rerun()
                                    except psycopg2.IntegrityError:
                                        st.error(f"Greška: Šifra '{sif}' u boji '{nova_boja_izmena}' već postoji u ovoj sekciji!")
                                    
                            with col_b2:
                                if st.button("🗑️ Obriši", key=f"Obr_{kljuc_id}"):
                                    conn = uzmi_vezu_sa_bazom()
                                    cursor = conn.cursor()
                                    cursor.execute("DELETE FROM artikli WHERE sifra = %s AND boja = %s AND sezona = %s", (sif, boj, izabrana_sezona))
                                    conn.commit()
                                    conn.close()
                                    
                                    ucitaj_artikle_za_sezonu.clear()
                                    st.warning("Obrisano!")
                                    st.rerun()
                st.markdown("---")
            
            # 2. KONTROLE STRANICA NA DNU
            if broj_stranica > 1 and not pretraga:
                prikazi_brojeve_stranica(broj_stranica, st.session_state["trenutna_stranica"], "dole")

# --- OPCIJA 3: EVIDENCIJA IZLAZA ---
elif meni == "Evidencija izlaza (Po danima)":
    st.header(f"📆 Dnevni izlaz robe - Sekcija: {izabrana_sezona}")
    
    df_artikli = ucitaj_artikle_za_sezonu(izabrana_sezona)
    sve_sifre = sorted(df_artikli["sifra"].unique().tolist()) if not df_artikli.empty else []
    
    if not sve_sifre:
        st.info(f"Nema unete robe u sekciji {izabrana_sezona} da biste zabeležili izlaz.")
    else:
        col1, col2 = st.columns(2)
        with col1:
            izabrani_datum = st.date_input("Izaberi datum izlaza:", datetime.now(), key="datum_izlaza_main")
            izabrana_sifra = st.selectbox("Izaberi šifru modela:", sve_sifre, key="izlaz_sifra_select")
            
            dostupne_boje = sorted(df_artikli[df_artikli["sifra"] == izabrana_sifra]["boja"].tolist())
            izabrana_boja = st.selectbox("Izaberi boju modela:", dostupne_boje, key="izlaz_boja_select")
        
        with col2:
            current_stanje = 0
            if izabrana_boja:
                filtriran_red = df_artikli[(df_artikli["sifra"] == izabrana_sifra) & (df_artikli["boja"] == izabrana_boja)]
                if not filtriran_red.empty:
                    current_stanje = int(filtriran_red.iloc[0]["broj_pari"])
            
            st.write("")
            sufiks_stanje = "komada" if izabrana_sezona == "Torbe" else "pari"
            st.info(f"💡 Trenutno stanje za **{izabrana_sifra}** (**{izabrana_boja}**) je: **{current_stanje} {sufiks_stanje}**")
            
            dinamicki_kljuc = f"izlaz_pari_input_{st.session_state['reset_brojac']}"
            labela_izlaz_unos = "Letimičan unos količine za izlaz (kom):" if izabrana_sezona == "Torbe" else "Letimičan unos količine za izlaz:"
            kolicina_izlaza = st.number_input(
                labela_izlaz_unos, 
                min_value=1, max_value=max(1, current_stanje), 
                step=1, value=None, key=dinamicki_kljuc
            )
        
        dugme_onemoguceno = kolicina_izlaza is None or kolicina_izlaza <= 0
        
        if st.button("Zapiši izlaz robe", type="primary", key="dugme_zapisi_izlaz", disabled=dugme_onemoguceno):
            if kolicina_izlaza is None or current_stanje < kolicina_izlaza or current_stanje == 0:
                st.error("Greška: Nemate dovoljno količine na stanju!")
            else:
                with st.spinner("Zapisivanje u toku..."):
                    try:
                        conn = uzmi_vezu_sa_bazom()
                        cursor = conn.cursor()
                        
                        cursor.execute('''
                            INSERT INTO izlaz_robe (datum, sifra_artikla, boja_artikla, kolicina_izlaz)
                            VALUES (%s, %s, %s, %s)
                        ''', (izabrani_datum.strftime("%Y-%m-%d"), izabrana_sifra, izabrana_boja, kolicina_izlaza))
                        
                        novo_stanje = current_stanje - kolicina_izlaza
                        cursor.execute('''
                            UPDATE artikli SET broj_pari = %s WHERE sifra = %s AND boja = %s AND sezona = %s
                        ''', (novo_stanje, izabrana_sifra, izabrana_boja, izabrana_sezona))
                        
                        conn.commit()
                        conn.close()
                        
                        ucitaj_artikle_za_sezonu.clear()
                        
                        st.session_state["reset_brojac"] += 1
                        st.success(f"Uspešno proknjižen izlaz! Novo stanje je {novo_stanje} {sufiks_stanje}.")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Sistemska greška pri upisu: {e}")

        st.markdown("---")
        st.subheader(f"📋 Istorija dnevnih izlaza robe za sekciju: {izabrana_sezona}")
        
        conn = uzmi_vezu_sa_bazom()
        upit_istorija = '''
            SELECT 
                ir.datum AS "Datum", 
                ir.sifra_artikla AS "Šifra modela", 
                ir.boja_artikla AS "Boja", 
                ir.kolicina_izlaz AS "Izašlo" 
            FROM izlaz_robe ir
            INNER JOIN artikli a ON ir.sifra_artikla = a.sifra AND ir.boja_artikla = a.boja
            WHERE a.sezona = %s
            ORDER BY ir.id ASC
        '''
        df_izlazi = pd.read_sql_query(upit_istorija, conn, params=(izabrana_sezona,))
        conn.close()
        
        if not df_izlazi.empty:
            st.write(f"📅 **Izaberi period za preuzimanje Excel tabele ({izabrana_sezona}):**")
            col_d1, col_d2 = st.columns(2)
            with col_d1:
                od_datuma = st.date_input("Od datuma:", datetime.strptime(df_izlazi['Datum'].min(), "%Y-%m-%d") if not df_izlazi.empty else datetime.now())
            with col_d2:
                do_datuma = st.date_input("Do datuma:", datetime.now())
            
            od_str = od_datuma.strftime("%Y-%m-%d")
            do_str = do_datuma.strftime("%Y-%m-%d")
            
            df_filtrirano = df_izlazi[(df_izlazi['Datum'] >= od_str) & (df_izlazi['Datum'] <= do_str)]
            
            excel_izlazi = konvertuj_u_excel(df_filtrirano)
            st.download_button(
                label=f"🟢 Preuzmi Excel za period ({od_datuma.strftime('%d.%m.%Y.')} - {do_datuma.strftime('%d.%m.%Y.')})",
                data=excel_izlazi,
                file_name=f"izlazi_{izabrana_sezona}_{od_str}_do_{do_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dugme_download_excel_izlazi"
            )
            st.dataframe(df_izlazi, use_container_width=True)
        else:
            st.write(f"Još uvek nema zabeleženih izlaza robe za sekciju {izabrana_sezona}.")