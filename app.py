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
@st.cache_resource
def inicijalizuj_bazu():
    return psycopg2.connect(st.secrets["postgres"]["url"])

def uzmi_vezu_sa_bazom():
    try:
        conn = inicijalizuj_bazu()
        if conn.closed != 0:
            st.cache_resource.clear()
            conn = inicijalizuj_bazu()
        return conn
    except Exception:
        st.cache_resource.clear()
        return inicijalizuj_bazu()

def kreiraj_tabele():
    try:
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
        
        try:
            cursor.execute("ALTER TABLE izlaz_robe ADD COLUMN IF NOT EXISTS grad TEXT;")
            cursor.execute("ALTER TABLE izlaz_robe ADD COLUMN IF NOT EXISTS prodajna_cena REAL;")
            cursor.execute("ALTER TABLE izlaz_robe ADD COLUMN IF NOT EXISTS nabavna_cena REAL;")
            conn.commit()
        except Exception:
            conn.rollback()

        try:
            cursor.execute("ALTER TABLE artikli ADD COLUMN IF NOT EXISTS datum_unosa TEXT;")
            conn.commit()
        except Exception:
            conn.rollback()
        
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
    except Exception:
        pass

kreiraj_tabele()

# --- NAPREDNO I UBRZANO KEŠIRANJE PODATAKA ---
@st.cache_data(ttl=60)
def ucitaj_artikle_za_sezonu(sezona):
    conn = uzmi_vezu_sa_bazom()
    df = pd.read_sql_query("SELECT * FROM artikli WHERE sezona = %s ORDER BY sifra ASC, boja ASC", conn, params=(sezona,))
    return df

@st.cache_data(ttl=300)
def ucitaj_boje():
    conn = uzmi_vezu_sa_bazom()
    cursor = conn.cursor()
    cursor.execute("SELECT boja FROM sifrarnik_boja ORDER BY boja ASC")
    boje = [red[0] for red in cursor.fetchall()]
    return boje

@st.cache_data(ttl=30)
def ucitaj_istoriju_izlaza_za_sezonu(sezona):
    conn = uzmi_vezu_sa_bazom()
    upit_istorija = '''
        SELECT ir.id AS "ID Zapisa", 
               ir.datum AS "Datum", 
               ir.sifra_artikla AS "Šifra modela", 
               ir.boja_artikla AS "Boja proizvoda", 
               ir.grad AS "Grad", 
               ir.kolicina_izlaz AS "Izašlo",
               ir.prodajna_cena AS "Prodajna cena po paru", 
               (ir.kolicina_izlaz * ir.prodajna_cena) AS "Ukupno prodajna",
               ir.nabavna_cena AS "Nabavna cena po paru", 
               (ir.kolicina_izlaz * ir.nabavna_cena) AS "Ukupno nabavna"
        FROM izlaz_robe ir INNER JOIN artikli a ON ir.sifra_artikla = a.sifra AND ir.boja_artikla = a.boja
        WHERE a.sezona = %s 
        ORDER BY ir.id DESC
    '''
    df = pd.read_sql_query(upit_istorija, conn, params=(sezona,))
    return df

def konvertuj_u_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Magacin')
    return output.getvalue()

def pronadji_sliku_u_df(df, sifra):
    if df.empty:
        return ""
    filtrirano = df[(df["sifra"] == sifra) & (df["slika_putanja"] != "") & (df["slika_putanja"].notna())]
    if not filtrirano.empty:
        return filtrirano.iloc[0]["slika_putanja"]
    return ""

# --- IZGLED I STILIZACIJA APLIKACIJE ---
st.set_page_config(page_title="Magacin", layout="wide")

# Omogućavamo glatko skrolovanje kroz CSS i eliminišemo podvlačenje teksta na linku
st.markdown("""
    <style>
    html {
        scroll-behavior: smooth !important;
    }
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
    
    div[data-testid="stHorizontalBlock"] { 
        background: var(--secondary-background-color);
        padding: 15px; 
        border-radius: 8px; 
        margin-bottom: 12px; 
        border: 1px solid var(--border-color);
    }
    
    div.stButton > button, div[data-testid="stForm"] button {
        width: 100% !important;
        padding: 2px 10px !important;
        font-size: 0.9rem !important;
        height: 38px !important;
    }
    
    .indikator-stranice {
        background-color: var(--secondary-background-color);
        padding: 6px 12px;
        border-radius: 4px;
        font-weight: bold;
        color: #26a69a;
        font-size: 0.95rem;
        border-left: 4px solid #26a69a;
        margin-bottom: 15px;
        display: inline-block;
        border-top: 1px solid var(--border-color);
        border-right: 1px solid var(--border-color);
        border-bottom: 1px solid var(--border-color);
    }
    
    .podsetnik-unosa {
        font-size: 1.02rem !important;
        color: #000000;
        font-style: italic;
        margin-top: -8px;
        margin-bottom: 12px;
    }

    /* Stilizacija dugmeta za povratak na vrh */
    .dugme-vrh-kontejner {
        display: flex;
        justify-content: center;
        margin: 30px 0 10px 0;
    }
    .dugme-vrh-link {
        text-decoration: none !important;
        display: inline-block;
    }
    .dugme-vrh {
        background-color: #26a69a !important;
        color: white !important;
        border: none !important;
        padding: 12px 30px !important;
        font-size: 0.95rem !important;
        font-weight: bold !important;
        border-radius: 4px !important;
        cursor: pointer !important;
        transition: background 0.2s ease !important;
        box-shadow: 0 2px 5px rgba(0,0,0,0.15) !important;
        text-align: center;
    }
    .dugme-vrh:hover {
        background-color: #208b80 !important;
    }
    </style>
""", unsafe_allow_html=True)

# --- 1. POSTAVLJAMO SIDRO NA SAM VRH STRANICE ---
st.markdown('<div id="vrh-stranice"></div>', unsafe_allow_html=True)

st.title("📦 Višekorisnički sistem za praćenje stanja u magacinu")

izabrana_sezona = st.sidebar.radio("🌸 IZABERI KATEGORIJU / SEZONU:", ["Proleće-Leto", "Jesen-Zima", "Torbe"])
st.sidebar.markdown("---")

meni = st.sidebar.selectbox("Izaberi opciju:", ["Trenutno stanje", "Unos nove robe", "Evidencija izlaza (Po danima)"])
st.sidebar.info(f"Trenutno radite u sekciji:\n**{izabrana_sezona}**")

# --- INICIJALIZACIJA STANJA ---
if "trenutna_stranica" not in st.session_state:
    st.session_state["trenutna_stranica"] = 1

if "prethodna_sezona" not in st.session_state:
    st.session_state["prethodna_sezona"] = izabrana_sezona

if "prethodni_meni" not in st.session_state:
    st.session_state["prethodni_meni"] = meni

if izabrana_sezona != st.session_state["prethodna_sezona"] or meni != st.session_state["prethodni_meni"]:
    st.session_state["trenutna_stranica"] = 1
    st.session_state["prethodna_sezona"] = izabrana_sezona
    st.session_state["prethodni_meni"] = meni

if "reset_brojac" not in st.session_state:
    st.session_state["reset_brojac"] = 0

if "reset_izlaz_kolicina" not in st.session_state:
    st.session_state["reset_izlaz_kolicina"] = 0


# --- OPCIJA 1: UNOS NOVE ROBE ---
if meni == "Unos nove robe":
    st.header(f"➕ Unos novog artikla ({izabrana_sezona})")
    lista_boja = ucitaj_boje()
    
    datum_unosa_odabir = st.date_input("📆 Izaberi datum unosa robe:", datetime.now())
    st.markdown("---")
    
    if "unos_sifra" not in st.session_state: st.session_state["unos_sifra"] = ""
    if "unos_boja" not in st.session_state: st.session_state["unos_boja"] = lista_boja[0] if lista_boja else ""
    
    kljuc_unos_pari = f"unos_kol_{st.session_state['reset_brojac']}"
    kljuc_unos_kutija = f"unos_kut_{st.session_state['reset_brojac']}"
    kljuc_unos_prodajna = f"unos_pc_{st.session_state['reset_brojac']}"
    kljuc_unos_internet = f"unos_ic_{st.session_state['reset_brojac']}"
    
    col1, col2 = st.columns(2)
    with col1:
        sifra = st.text_input("Šifra modela:", value=st.session_state["unos_sifra"]).strip().upper()
        boja = st.selectbox("Boja modela:", lista_boja, index=lista_boja.index(st.session_state["unos_boja"]) if st.session_state["unos_boja"] in lista_boja else 0)
        
        labela_kol = "Količina (komada/pari):" if izabrana_sezona == "Torbe" else "Količina pari:"
        labela_kut = "Broj komada u jednoj kutiji/pakovanju:" if izabrana_sezona == "Torbe" else "Broj pari u jednoj kutiji:"
        
        broj_pari = st.number_input(labela_kol, min_value=0, step=1, value=None, key=kljuc_unos_pari)
        pari_u_kutiji = st.number_input(labela_kut, min_value=1, step=1, value=None, key=kljuc_unos_kutija)
    with col2:
        prodajna_cena = st.number_input("Prodajna cena (RSD):", min_value=0.0, step=50.0, value=None, key=kljuc_unos_prodajna)
        internet_cena = st.number_input("Internet cena (RSD):", min_value=0.0, step=50.0, value=None, key=kljuc_unos_internet)
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
    
    if dugme_potvrdi:
        df_trenutni = ucitaj_artikle_za_sezonu(izabrana_sezona)
        url_slike = ""
        if slika is not None:
            with st.spinner("Slanje slike na Cloudinary..."):
                try:
                    rezultat_slike = cloudinary.uploader.upload(
                        slika, folder="magacin/", public_id=f"{sifra}_{boja}",
                        transformation=[{"width": 800, "crop": "limit"}, {"quality": "auto", "fetch_format": "auto"}]
                    )
                    url_slike = rezultat_slike["secure_url"]
                except Exception as e:
                    st.error(f"Greška pri slanju slike: {e}")
        else:
            url_slike = pronadji_sliku_u_df(df_trenutni, sifra)
        
        try:
            conn = uzmi_vezu_sa_bazom()
            cursor = conn.cursor()
            cursor.execute('''
                INSERT INTO artikli (sifra, boja, sezona, broj_pari, pari_u_kutiji, prodajna_cena, internet_cena, slika_putanja, datum_unosa)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
            ''', (sifra, boja, izabrana_sezona, broj_pari, pari_u_kutiji, prodajna_cena, internet_cena, url_slike, datum_unosa_odabir.strftime("%Y-%m-%d")))
            conn.commit()
            
            ucitaj_artikle_za_sezonu.clear()
            st.session_state["unos_sifra"] = ""
            st.session_state["reset_brojac"] += 1
            
            st.success(f"Uspešno sačuvan model: Šifra '{sifra}' - Boja '{boja}' sa datumom unosa {datum_unosa_odabir.strftime('%Y-%m-%d')}!")
            st.rerun()
        except psycopg2.IntegrityError:
            st.error(f"Greška: Model sa šifrom '{sifra}' u boji '{boja}' već postoji u ovoj sekciji!")

    st.markdown("---")
    
    st.subheader(f"📋 Pregled unete robe ({izabrana_sezona})")
    df_unos_pregled = ucitaj_artikle_za_sezonu(izabrana_sezona)
    
    if df_unos_pregled.empty:
        st.info("Još uvek nema unetih artikala u ovoj sekciji.")
    else:
        df_prikaz_unos = df_unos_pregled.copy()
        df_prikaz_unos["Ukupno kartona"] = (df_prikaz_unos["broj_pari"] / df_prikaz_unos["pari_u_kutiji"]).round(2)
        
        relabel_kom = "Ukupno komada" if izabrana_sezona == "Torbe" else "Ukupno pari"
        relabel_kut = "Kutija/Pakovanja" if izabrana_sezona == "Torbe" else "Pari u kutiji"
        relabel_karton = "Ukupno pakovanja" if izabrana_sezona == "Torbe" else "Ukupno kartona"
        
        df_prikaz_unos = df_prikaz_unos.rename(columns={
            "datum_unosa": "Datum unosa", "sifra": "Šifra modela", "boja": "Boja proizvoda", 
            "broj_pari": relabel_kom, "pari_u_kutiji": relabel_kut, "Ukupno kartona": relabel_karton,
            "prodajna_cena": "Prodajna cena (RSD)", "internet_cena": "Internet cena (RSD)"
        })
        
        kolone_redosted = ["Datum unosa", "Šifra modela", "Boja proizvoda", relabel_kom, relabel_kut, relabel_karton, "Prodajna cena (RSD)", "Internet cena (RSD)"]
        df_prikaz_unos = df_prikaz_unos[[c for c in kolone_redosted if c in df_prikaz_unos.columns]]
        df_prikaz_unos = df_prikaz_unos.sort_values(by="Datum unosa", ascending=False, na_position="last")
        
        st.write("#### 📆 Filter datuma za izvoz u Excel:")
        col_ex1, col_ex2 = st.columns(2)
        
        minimalni_datum_baza = datetime.now()
        sve_unete_vrednosti_datuma = df_prikaz_unos["Datum unosa"].dropna()
        if not sve_unete_vrednosti_datuma.empty:
            try:
                minimalni_datum_baza = datetime.strptime(sve_unete_vrednosti_datuma.min(), "%Y-%m-%d")
            except: pass
            
        with col_ex1:
            izvoz_od = st.date_input("Od datuma unosa:", minimalni_datum_baza, key="izvoz_datum_od")
        with col_ex2:
            izvoz_do = st.date_input("Do datuma unosa:", datetime.now(), key="izvoz_datum_do")
            
        od_str, do_str = izvoz_od.strftime("%Y-%m-%d"), izvoz_do.strftime("%Y-%m-%d")
        df_za_excel = df_prikaz_unos[
            (df_prikaz_unos["Datum unosa"] >= od_str) & 
            (df_prikaz_unos["Datum unosa"] <= do_str)
        ]
        
        staro_stanje_bez_datuma = df_prikaz_unos[df_prikaz_unos["Datum unosa"].isna()]
        if izvoz_od == minimalni_datum_baza and not staro_stanje_bez_datuma.empty:
            df_za_excel = pd.concat([df_za_excel, staro_stanje_bez_datuma], ignore_index=True)

        excel_unosa = konvertuj_u_excel(df_za_excel)
        
        st.download_button(
            label=f"🟢 Preuzmi profiltrirane artikle ({len(df_za_excel)} modela) kao Excel (.xlsx)",
            data=excel_unosa,
            file_name=f"uneseni_artikli_{izabrana_sezona}_od_{od_str}_do_{do_str}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        st.dataframe(df_prikaz_unos, use_container_width=True)

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
            "sifra": "Šifra modela", "boja": "Boja proizvoda", "sezona": "Kategorija", 
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
            BROJ_ARTIKALA_PO_STRANICI = 20
            ukupno_artikala = len(df_prikaz)
            broj_stranica = (ukupno_artikala // BROJ_ARTIKALA_PO_STRANICI) + (1 if ukupno_artikala % BROJ_ARTIKALA_PO_STRANICI > 0 else 0)
            
            if pretraga != "" or st.session_state["trenutna_stranica"] > broj_stranica:
                st.session_state["trenutna_stranica"] = 1
            
            if broj_stranica > 1 and not pretraga:
                st.caption(f"Ukupno pronađeno: {ukupno_artikala} modela raspoređenih na {broj_stranica} stranica.")
            
            # --- JEDINA NAVIGACIJA (SAMO GORE) ---
            if broj_stranica > 1:
                st.markdown(f'<div class="indikator-stranice">📄 Stranica: {st.session_state["trenutna_stranica"]} od {broj_stranica}</div>', unsafe_allow_html=True)
                
                col_nav1, col_nav2, col_nav3 = st.columns([1, 2, 1])
                with col_nav1:
                    prethodna_onemogucena = st.session_state["trenutna_stranica"] == 1
                    if st.button("⬅️ Prethodna", key="prev_gore", disabled=prethodna_onemogucena):
                        st.session_state["trenutna_stranica"] -= 1
                        st.rerun()
                with col_nav2:
                    sve_stranice = list(range(1, broj_stranica + 1))
                    
                    def promena_stranice_gore():
                        st.session_state["trenutna_stranica"] = st.session_state["izbor_str_gore"]

                    st.selectbox(
                        "Idi na stranicu:", 
                        sve_stranice, 
                        index=sve_stranice.index(st.session_state["trenutna_stranica"]),
                        key="izbor_str_gore",
                        on_change=promena_stranice_gore
                    )
                with col_nav3:
                    sledec_onemogucena = st.session_state["trenutna_stranica"] == broj_stranica
                    if st.button("Sledeća ➡️", key="next_gore", disabled=sledec_onemogucena):
                        st.session_state["trenutna_stranica"] += 1
                        st.rerun()
                st.markdown("---")
            
            start_indeks = (st.session_state["trenutna_stranica"] - 1) * BROJ_ARTIKALA_PO_STRANICI
            kraj_indeks = start_indeks + BROJ_ARTIKALA_PO_STRANICI
            df_za_prikaz = df_prikaz.iloc[start_indeks:kraj_indeks]
            
            try:
                conn = uzmi_vezu_sa_bazom()
                cursor = conn.cursor(cursor_factory=RealDictCursor)
                cursor.execute("SELECT sifra_artikla, boja_artikla, SUM(kolicina_izlaz) as ukupno_izaslo FROM izlaz_robe GROUP BY sifra_artikla, boja_artikla")
                izlaz_mapa = {f"{red['sifra_artikla']}_{red['boja_artikla']}": red['ukupno_izaslo'] for red in cursor.fetchall()}
            except Exception:
                izlaz_mapa = {}
            
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
                
                trenutno_na_stanju = int(row["broj_pari"])
                izaslo_iz_magacina = int(izlaz_mapa.get(kljuc_id, 0))
                pocetni_pari = trenutno_na_stanju + izaslo_iz_magacina
                pari_u_kut = int(row["pari_u_kutiji"])
                
                pocetni_broj_kutija = round(pocetni_pari / pari_u_kut, 2)
                datum_unosa_ispis = row.get("datum_unosa")
                if not datum_unosa_ispis or pd.isna(datum_unosa_ispis):
                    datum_unosa_ispis = "Nepoznat datum"
                
                sufiks_kartona = "pakovanja" if izabrana_sezona == "Torbe" else "kartona"
                sufiks_jedinica = "kom" if izabrana_sezona == "Torbe" else "pari"
                
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
                        
                        st.markdown(
                            f'<div class="podsetnik-unosa">📆 Uneto: {datum_unosa_ispis} | 📦 Početno ušlo: {pocetni_pari} {sufiks_jedinica} ({pocetni_broj_kutija} {sufiks_kartona})</div>', 
                            unsafe_allow_html=True
                        )
                        
                        c1, c2, c3, c4 = st.columns(4)
                        m_kol_label = "Ukupno komada" if izabrana_sezona == "Torbe" else "Ukupno pari na stanju"
                        m_pak_label = "Pakovanje (kut. + kom)" if izabrana_sezona == "Torbe" else "Pakovanje"
                        m_sufiks = "kom"
                        m_pak_sufiks = f"{br_kutija} pak. + {ost_pari} kom" if izabrana_sezona == "Torbe" else f"{br_kutija} kut. + {ost_pari} par"
                        
                        c1.metric(m_kol_label, f"{row['broj_pari']} {m_sufiks}")
                        c2.metric(m_pak_label, m_pak_sufiks)
                        c3.metric("Prodajna cena", f"{p_cena_int} din")
                        c4.metric("Internet cena", f"{i_cena_int} din")
                        
                    with col_akcije:
                        ekspander = st.expander("🛠️ Izmeni / Obriši")
                        with ekspander:
                            st.write("**Uredi podatke:**")
                            nova_sifra_izmena = st.text_input("Izmeni šifru modela:", value=sif, key=f"sifra_izm_{kljuc_id}").strip().upper()
                            indeks_trenutne_boje = lista_boja.index(boj) if boj in lista_boja else 0
                            nova_boja_izmena = st.selectbox("Izmeni boju artikla:", lista_boja, index=indeks_trenutne_boje, key=f"boja_{kljuc_id}")
                            
                            labela_izmena_kol = "Novo ukupno komada:" if izabrana_sezona == "Torbe" else "Novo ukupno pari:"
                            nova_kol = st.number_input(labela_izmena_kol, min_value=0, value=int(row['broj_pari']), step=1, key=f"kol_{kljuc_id}")
                            nova_p_cena = st.number_input("Prodajna cena (RSD):", min_value=0.0, value=float(row['prodajna_cena']), step=50.0, key=f"pc_{kljuc_id}")
                            nova_i_cena = st.number_input("Internet cena (RSD):", min_value=0.0, value=float(row['internet_cena']), step=50.0, key=f"ic_{kljuc_id}")
                            nova_slika_file = st.file_uploader("Zamani sliku artikla:", type=["jpg", "jpeg", "png"], key=f"img_{kljuc_id}")
                            
                            col_b1, col_b2 = st.columns(2)
                            with col_b1:
                                if st.button("💾 Snimi", key=f"Snimi_{kljuc_id}"):
                                    if nova_sifra_izmena == "":
                                        st.error("Šifra modela ne može biti prazna!")
                                    else:
                                        finalna_putanja_slike = row["slika_putanja"]
                                        if nova_slika_file is not None:
                                            with st.spinner("Menjanje slike..."):
                                                try:
                                                    rez_nove_slike = cloudinary.uploader.upload(
                                                        nova_slika_file, folder="magacin/", public_id=f"{nova_sifra_izmena}_{nova_boja_izmena}",
                                                        transformation=[{"width": 800, "crop": "limit"}, {"quality": "auto", "fetch_format": "auto"}]
                                                    )
                                                    finalna_putanja_slike = rez_nove_slike["secure_url"]
                                                except: pass
                                                
                                        try:
                                            conn = uzmi_vezu_sa_bazom()
                                            cursor = conn.cursor()
                                            
                                            if nova_sifra_izmena != sif:
                                                cursor.execute('UPDATE artikli SET sifra = %s WHERE sifra = %s AND sezona = %s', (nova_sifra_izmena, sif, izabrana_sezona))
                                                cursor.execute('UPDATE izlaz_robe SET sifra_artikla = %s WHERE sifra_artikla = %s', (nova_sifra_izmena, sif))
                                            
                                            cursor.execute('UPDATE artikli SET prodajna_cena = %s, internet_cena = %s WHERE sifra = %s AND sezona = %s', (nova_p_cena, nova_i_cena, nova_sifra_izmena, izabrana_sezona))
                                            cursor.execute('UPDATE artikli SET boja = %s, broj_pari = %s, slika_putanja = %s WHERE sifra = %s AND boja = %s AND sezona = %s', (nova_boja_izmena, nova_kol, finalna_putanja_slike, nova_sifra_izmena, boj, izabrana_sezona))
                                            
                                            conn.commit()
                                            ucitaj_artikle_za_sezonu.clear()
                                            st.success("Izmene uspešno sačuvane!")
                                            st.rerun()
                                        except psycopg2.IntegrityError:
                                            st.error(f"Greška: Šifra '{nova_sifra_izmena}' u boji '{nova_boja_izmena}' već postoji!")
                                        
                            with col_b2:
                                if st.button("🗑️ Obriši", key=f"Obr_{kljuc_id}"):
                                    conn = uzmi_vezu_sa_bazom()
                                    cursor = conn.cursor()
                                    cursor.execute("DELETE FROM artikli WHERE sifra = %s AND boja = %s AND sezona = %s", (sif, boj, izabrana_sezona))
                                    conn.commit()
                                    
                                    ucitaj_artikle_za_sezonu.clear()
                                    st.warning("Obrisano!")
                                    st.rerun()
                st.markdown("---")

            # --- 2. DODAJEMO ČISTU I BEZBEDNU HTML SIDRO VEZU (DUGME ZA VRH) ---
            st.markdown(
                """
                <div class="dugme-vrh-kontejner">
                    <a href="#vrh-stranice" target="_self" class="dugme-vrh-link">
                        <button class="dugme-vrh">⬆️ Idi na vrh stranice</button>
                    </a>
                </div>
                """,
                unsafe_allow_html=True
            )


# --- OPCIJA 3: EVIDENCIJA IZLAZA ---
elif meni == "Evidencija izlaza (Po danima)":
    st.header(f"📆 Dnevni izlaz robe - Sekcija: {izabrana_sezona}")
    df_artikli = ucitaj_artikle_za_sezonu(izabrana_sezona)
    sve_sifre = sorted(df_artikli["sifra"].unique().tolist()) if not df_artikli.empty else []
    
    if not sve_sifre:
        st.info(f"💡 Trenutno nema unete robe na stanju za kategoriju '{izabrana_sezona}'.")
    else:
        lista_gradova = ["Internet", "Mladenovac Gore", "Mladenovac Dole", 
                         "Smederevska Palanka", "Zaječar", "Subotica", "Aleksinac", "Loznica", "Sremska Mitrovica", "Pančevo", "Vršac", "Bečej", "Prokuplje"]
        
        st.write("### 📝 Popunite podatke for novi izlaz")
        col1, col2 = st.columns(2)
        
        with col1:
            izabrani_datum = st.date_input("Izaberi datum izlaza:", datetime.now())
            trenutna_sifra_key = f"izlaz_sifra_{izabrana_sezona}"
            izabrana_sifra = st.selectbox("Izaberi šifru modela:", sve_sifre, key=trenutna_sifra_key)
            
            boje_za_sifru = sorted(df_artikli[df_artikli["sifra"] == izabrana_sifra]["boja"].unique().tolist())
            izabrana_boja = st.selectbox("Izaberi boju modela:", boje_za_sifru, key=f"izlaz_boja_{izabrana_sezona}")
            izabrani_grad = st.selectbox("Izaberi grad:", lista_gradova)
            
        filtriran_artikal = df_artikli[(df_artikli["sifra"] == izabrana_sifra) & (df_artikli["boja"] == izabrana_boja)]
        
        fabricka_cena = 0.0
        zaliha_komada = 0
        
        if not filtriran_artikal.empty:
            zaliha_komada = int(filtriran_artikal.iloc[0]["broj_pari"])
            if izabrani_grad == "Internet":
                fabricka_cena = float(filtriran_artikal.iloc[0]["internet_cena"])
            else:
                fabricka_cena = float(filtriran_artikal.iloc[0]["prodajna_cena"])
            
        with col2:
            kljuc_kolicina_izlaza = f"kolicina_izlaz_{st.session_state['reset_izlaz_kolicina']}"
            kolicina_izlaza = st.number_input("Količina za izlaz:", min_value=1, step=1, value=None, key=kljuc_kolicina_izlaza)
            
            prodajna_cena_par = st.number_input("Prodajna cena (RSD):", min_value=0.0, step=50.0, value=fabricka_cena)
            nabavna_cena_par = st.number_input("Nabavna cena (Opciono):", min_value=0.0, step=50.0, value=None)
            
            tekst_zaliha = "komada" if izabrana_sezona == "Torbe" else "pari"
            st.info(f"📊 Trenutno stanje u magacinu za model **{izabrana_sifra} ({izabrana_boja})**: **{zaliha_komada} {tekst_zaliha}**")

        dugme_onemoguceno = (prodajna_cena_par is None or prodajna_cena_par <= 0.0 or kolicina_izlaza is None)
        
        st.write("")
        potvrdi_izlaz = st.button("Zapiši izlaz robe", type="primary", disabled=dugme_onemoguceno)
            
        if potvrdi_izlaz:
            if zaliha_komada < kolicina_izlaza:
                st.error(f"❌ Nema dovoljno robe! Na stanju ima samo {zaliha_komada} kom.")
            else:
                conn = uzmi_vezu_sa_bazom()
                cursor = conn.cursor()
                cursor.execute('''
                    INSERT INTO izlaz_robe (datum, sifra_artikla, boja_artikla, kolicina_izlaz, grad, prodajna_cena, nabavna_cena)
                    VALUES (%s, %s, %s, %s, %s, %s, %s)
                ''', (izabrani_datum.strftime("%Y-%m-%d"), izabrana_sifra, izabrana_boja, kolicina_izlaza, izabrani_grad, prodajna_cena_par, nabavna_cena_par))
                cursor.execute('UPDATE artikli SET broj_pari = broj_pari - %s WHERE sifra = %s AND boja = %s AND sezona = %s', (kolicina_izlaza, izabrana_sifra, izabrana_boja, izabrana_sezona))
                
                conn.commit()
                st.session_state["reset_izlaz_kolicina"] += 1
                
                ucitaj_artikle_za_sezonu.clear()
                ucitaj_istoriju_izlaza_za_sezonu.clear()
                st.success("✅ Izlaz uspešno proknjižen!")
                st.rerun()

    st.markdown("---")
    st.subheader(f"📋 Istorija dnevnih izlaza robe za sekciju: {izabrana_sezona}")
    
    try:
        df_izlazi = ucitaj_istoriju_izlaza_za_sezonu(izabrana_sezona)
        if not df_izlazi.empty:
            lista_gradova = ["Internet", "Mladenovac Gore", "Mladenovac Dole", "Smederevska Palanka", "Zaječar", "Subotica", "Aleksinac", "Loznica", "Sremska Mitrovica", "Pančevo", "Vršac", "Bečej", "Prokuplje"]
            col_filter1, col_filter2, col_filter3 = st.columns(3)
            
            stari_datum = datetime.now()
            try: stari_datum = datetime.strptime(df_izlazi['Datum'].min(), "%Y-%m-%d")
            except: pass
                
            with col_filter1: od_datuma = st.date_input("Od datuma:", stari_datum)
            with col_filter2: do_datuma = st.date_input("Do datuma:", datetime.now())
            with col_filter3: izabrani_grad_filter = st.selectbox("Izaberi grad za tabelu:", ["SVI GRADOVI"] + lista_gradova)
            
            od_str, do_str = od_datuma.strftime("%Y-%m-%d"), do_datuma.strftime("%Y-%m-%d")
            df_filtrirano = df_izlazi[(df_izlazi['Datum'] >= od_str) & (df_izlazi['Datum'] <= do_str)]
            if izabrani_grad_filter != "SVI GRADOVI":
                df_filtrirano = df_filtrirano[df_filtrirano['Grad'] == izabrani_grad_filter]
            
            excel_izlazi = konvertuj_u_excel(df_filtrirano)
            st.download_button(label="🟢 Preuzmi istoriju izlaza kao Excel", data=excel_izlazi, file_name=f"izlazi_{izabrana_sezona}_{datetime.now().strftime('%Y-%m-%d')}.xlsx")
            
            if not df_filtrirano.empty:
                st.dataframe(df_filtrirano, use_container_width=True)
            else:
                st.info("Nema zabeleženih izlaza za izabrani period i grad u tabeli.")
        else:
            st.write(f"Još uvek nema zabeleženih izlaza robe za sekciju {izabrana_sezona}.")
    except Exception as e:
        st.warning(f"Tabela sa istorijom ne može da se prikaže na standardan način, ali možete nesmetano stornirati zapise ispod. (Greška: {e})")

    st.markdown("---")
    st.write("### 🚨 Storniranje (Brisanje) zapisa")
    
    try:
        conn = uzmi_vezu_sa_bazom()
        cursor = conn.cursor(cursor_factory=RealDictCursor)
        upit_storno = '''
            SELECT ir.id, ir.datum, ir.sifra_artikla, ir.boja_artikla, ir.grad, ir.kolicina_izlaz
            FROM izlaz_robe ir 
            LEFT JOIN artikli a ON ir.sifra_artikla = a.sifra AND ir.boja_artikla = a.boja
            WHERE a.sezona = %s OR a.sezona IS NULL
            ORDER BY ir.id DESC LIMIT 50
        '''
        cursor.execute(upit_storno, (izabrana_sezona,))
        sirovi_izlazi = cursor.fetchall()
        
        if not sirovi_izlazi:
            st.info(f"Trenutno nema zapisa za storniranje u kategoriji '{izabrana_sezona}'.")
        else:
            opcije_za_storno = []
            mapa_zapisa = {}
            
            for red in sirovi_izlazi:
                sif_m = red['sifra_artikla'] if red['sifra_artikla'] else "OBRISAN MODEL"
                boj_m = red['boja_artikla'] if red['boja_artikla'] else "NEPOZNATO"
                kol_m = red['kolicina_izlaz'] if red['kolicina_izlaz'] else 0
                dat_m = red['datum'] if red['datum'] else "---"
                
                tekst_opcije = f"ID: {red['id']} | Model: {sif_m} | Boja: {boj_m} | Količina: {kol_m} kom | Datum: {dat_m}"
                opcije_za_storno.append(tekst_opcije)
                mapa_zapisa[tekst_opcije] = red
                
            izabrana_opcija_storno = st.selectbox(
                "Izaberi zapis koji želiš stornirati:", 
                ["--- Izaberi zapis iz baze ---"] + opcije_za_storno, 
                key="storno_izlaza_finalni_kljuc"
            )
            
            if izabrana_opcija_storno != "--- Izaberi zapis iz baze ---":
                zapis = mapa_zapisa[izabrana_opcija_storno]
                id_zapis = int(zapis["id"])
                sif_zapis = zapis["sifra_artikla"]
                boj_zapis = zapis["boja_artikla"]
                kol_zapis = int(zapis["kolicina_izlaz"])
                
                st.error(f"Upozorenje: Brisanjem zapisa ID {id_zapis}, vraćate {kol_zapis} kom na stanje modela {sif_zapis}.")
                
                if st.button("❌ POTVRDI BRISANJE I VRATI ROBU NA STANJE", type="primary", key="potvrda_storniranja_ok"):
                    conn = uzmi_vezu_sa_bazom()
                    cursor = conn.cursor()
                    cursor.execute("DELETE FROM izlaz_robe WHERE id = %s", (id_zapis,))
                    
                    cursor.execute("SELECT * FROM artikli WHERE sifra = %s AND boja = %s AND sezona = %s", (sif_zapis, boj_zapis, izabrana_sezona))
                    if cursor.fetchone() is not None:
                        cursor.execute('UPDATE artikli SET broj_pari = broj_pari + %s WHERE sifra = %s AND boja = %s AND sezona = %s', (kol_zapis, sif_zapis, boj_zapis, izabrana_sezona))
                    
                    conn.commit()
                    ucitaj_artikle_za_sezonu.clear()
                    ucitaj_istoriju_izlaza_za_sezonu.clear()
                    st.success("Zapis uspešno storniran i obrisan!")
                    st.rerun()
    except Exception as storno_err:
        st.error(f"Greška u storno modulu: {storno_err}")