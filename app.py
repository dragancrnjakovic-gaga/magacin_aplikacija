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
    
    try:
        cursor.execute("ALTER TABLE izlaz_robe ADD COLUMN IF NOT EXISTS grad TEXT;")
        cursor.execute("ALTER TABLE izlaz_robe ADD COLUMN IF NOT EXISTS prodajna_cena REAL;")
        cursor.execute("ALTER TABLE izlaz_robe ADD COLUMN IF NOT EXISTS nabavna_cena REAL;")
        conn.commit()
    except Exception as e:
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

@st.cache_data(ttl=300)
def ucitaj_istoriju_izlaza_za_sezonu(sezona):
    conn = uzmi_vezu_sa_bazom()
    upit_istorija = '''
        SELECT ir.datum AS "Datum", ir.sifra_artikla AS "Šifra modela", ir.boja_artikla AS "Boja", ir.grad AS "Grad", ir.kolicina_izlaz AS "Izašlo",
               ir.prodajna_cena AS "Prodajna cena po paru", (ir.kolicina_izlaz * ir.prodajna_cena) AS "Ukupno prodajna",
               ir.nabavna_cena AS "Nabavna cena po paru", (ir.kolicina_izlaz * ir.nabavna_cena) AS "Ukupno nabavna"
        FROM izlaz_robe ir INNER JOIN artikli a ON ir.sifra_artikla = a.sifra AND ir.boja_artikla = a.boja
        WHERE a.sezona = %s ORDER BY ir.id DESC
    '''
    df = pd.read_sql_query(upit_istorija, conn, params=(sezona,))
    conn.close()
    return df

def konvertuj_u_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Magacin')
    return output.getvalue()

# --- POMOĆNA FUNKCIJA ---
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
    
    div.skrivena-labela label {
        display: none !important;
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

    [data-testid="stHorizontalBlock"]:has(button[key^="vrh_"]),
    [data-testid="stHorizontalBlock"]:has(button[key^="dole_"]) {
        display: flex !important;
        flex-direction: row !important;
        justify-content: space-between !important;
        align-items: center !important;
        flex-wrap: nowrap !important;
    }

    div[data-testid="stColumn"]:has(button[key="vrh_next"]),
    div[data-testid="stColumn"]:has(button[key="dole_veliko_next"]) {
        display: flex !important;
        justify-content: flex-end !important;
        padding-right: 0px !important;
    }
    
    div[data-testid="stColumn"]:has(button[key="vrh_next"]) > div,
    div[data-testid="stColumn"]:has(button[key="dole_veliko_next"]) > div {
        width: auto !important;
    }

    button[key="vrh_next"], button[key="dole_veliko_next"] {
        width: max-content !important;
        margin-left: auto !important;
    }
    </style>
""", unsafe_allow_html=True)

if "skroluj_na_vrh" in st.session_state and st.session_state["skroluj_na_vrh"]:
    st.components.v1.html(
        "<script>window.parent.document.querySelector('.stMain').scrollTo(0, 0);</script>",
        height=0,
        width=0,
    )
    st.session_state["skroluj_na_vrh"] = False

st.title("📦 Višekorisnički sistem za praćenje stanja u magacinu")

izabrana_sezona = st.sidebar.radio("🌸 IZABERI KATEGORIJU / SEZONU:", ["Proleće-Leto", "Jesen-Zima", "Torbe"])
st.sidebar.markdown("---")

meni = st.sidebar.selectbox("Izaberi opciju:", ["Trenutno stanje", "Unos nove robe", "Evidencija izlaza (Po danima)"])
st.sidebar.info(f"Trenutno radite u sekciji:\n**{izabrana_sezona}**")

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


# --- OPCIJA 1: UNOS NOVE ROBE ---
if meni == "Unos nove robe":
    st.header(f"➕ Unos novog artikla ({izabrana_sezona})")
    lista_boja = ucitaj_boje()
    
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
            BROJ_ARTIKALA_PO_STRANICI = 20
            ukupno_artikala = len(df_prikaz)
            broj_stranica = (ukupno_artikala // BROJ_ARTIKALA_PO_STRANICI) + (1 if ukupno_artikala % BROJ_ARTIKALA_PO_STRANICI > 0 else 0)
            
            if pretraga != "" or st.session_state["trenutna_stranica"] > broj_stranica:
                st.session_state["trenutna_stranica"] = 1
                
            if broj_stranica > 1 and not pretraga:
                st.caption(f"Ukupno pronađeno: {ukupno_artikala} modela raspoređenih na {broj_stranica} stranica.")
                prikazi_gornju_paginaciju(broj_stranica, st.session_state["trenutna_stranica"])
                st.write("")
            
            if broj_stranica > 1:
                st.markdown(f'<div class="indikator-stranice">📄 Stranica: {st.session_state["trenutna_stranica"]} od {broj_stranica}</div>', unsafe_allow_html=True)
            
            start_indeks = (st.session_state["trenutna_stranica"] - 1) * BROJ_ARTIKALA_PO_STRANICI
            kraj_indeks = start_indeks + BROJ_ARTIKALA_PO_STRANICI
            df_za_prikaz = df_prikaz.iloc[start_indeks:kraj_indeks]
            
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
                            nova_sifra_izmena = st.text_input("Izmeni šifru modela:", value=sif, key=f"sifra_izm_{kljuc_id}").strip().upper()
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
                                                cursor.execute('''
                                                    UPDATE artikli 
                                                    SET sifra = %s
                                                    WHERE sifra = %s AND sezona = %s
                                                ''', (nova_sifra_izmena, sif, izabrana_sezona))
                                                
                                                cursor.execute('''
                                                    UPDATE izlaz_robe
                                                    SET sifra_artikla = %s
                                                    WHERE sifra_artikla = %s
                                                ''', (nova_sifra_izmena, sif))
                                            
                                            cursor.execute('''
                                                UPDATE artikli
                                                SET prodajna_cena = %s, internet_cena = %s
                                                WHERE sifra = %s AND sezona = %s
                                            ''', (nova_p_cena, nova_i_cena, nova_sifra_izmena, izabrana_sezona))
                                            
                                            cursor.execute('''
                                                UPDATE artikli SET boja = %s, broj_pari = %s, slika_putanja = %s
                                                WHERE sifra = %s AND boja = %s AND sezona = %s
                                            ''', (nova_boja_izmena, nova_kol, finalna_putanja_slike, nova_sifra_izmena, boj, izabrana_sezona))
                                            
                                            conn.commit()
                                            conn.close()
                                            ucitaj_artikle_za_sezonu.clear()
                                            st.success("Izmene uspešno sačuvane za sve varijacije modela!")
                                            st.rerun()
                                        except psycopg2.IntegrityError:
                                            st.error(f"Greška: Šifra '{nova_sifra_izmena}' u boji '{nova_boja_izmena}' već postoji u ovoj sekciji!")
                                        
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
            if broj_stranica > 1 and not pretraga:
                prikazi_donju_paginaciju(broj_stranica, st.session_state["trenutna_stranica"])


# --- OPCIJA 3: EVIDENCIJA IZLAZA (POTPUNA MUNJEVITA BRZINA KROZ st.form) ---
elif meni == "Evidencija izlaza (Po danima)":
    st.header(f"📆 Dnevni izlaz robe - Sekcija: {izabrana_sezona}")
    df_artikli = ucitaj_artikle_za_sezonu(izabrana_sezona)
    sve_sifre = sorted(df_artikli["sifra"].unique().tolist()) if not df_artikli.empty else []
     sve_boje = ucitaj_boje()
    
    if not sve_sifre:
        st.info(f"Nema unete robe u sekciji {izabrana_sezona} da biste zabeležili izlaz.")
    else:
        lista_gradova = ["Internet", "Mladenovac Gore", "Mladenovac Dole", "Smederevska Palanka", "Zaječar", "Subotica", "Aleksinac", "Loznica", "Sremska Mitrovica", "Pančevo", "Vršac", "Bečej", "Prokuplje"]
        
        # OTVARAMO FORMULAR: Sve unutar ovog bloka radi trenutno bez ikakvog seckanja i osvežavanja ekrana!
        with st.form("formular_za_izlaz_robe", clear_on_submit=False):
            st.write("### 📝 Popunite podatke za izlaz")
            
            col1, col2 = st.columns(2)
            with col1:
                izabrani_datum = st.date_input("Izaberi datum izlaza:", datetime.now())
                izabrana_sifra = st.selectbox("Izaberi šifru modela:", sve_sifre)
                izabrana_boja = st.selectbox("Izaberi boju modela:", sve_boje)
                izabrani_grad = st.selectbox("Izaberi grad:", lista_gradova)
            
            with col2:
                labela_izlaz_unos = "Količina za izlaz (broj komada/pari):"
                kolicina_izlaza = st.number_input(labela_izlaz_unos, min_value=1, step=1, value=None)
                prodajna_cena_par = st.number_input("Prodajna cena po paru/komadu (RSD):", min_value=0.0, step=50.0, value=None)
                nabavna_cena_par = st.number_input("Nabavna cena po paru/komadu (Opciono - RSD):", min_value=0.0, step=50.0, value=None)
            
            # Dugme unutar formulara koje pokreće slanje odjednom
            potvrdi_izlaz = st.form_submit_button("Zapiši izlaz robe", type="primary")
            
        if potvrdi_izlaz:
            if (prodajna_cena_par is None) or (kolicina_izlaza is None) or (kolicina_izlaza <= 0):
                st.error("❌ Greška: Morate uneti ispravnu količinu i prodajnu cenu!")
            else:
                # Provera stvarnog stanja u bazi tek nakon klika na dugme
                filtriran_red = df_artikli[(df_artikli["sifra"] == izabrana_sifra) & (df_artikli["boja"] == izabrana_boja)]
                
                if filtriran_red.empty:
                    st.error(f"❌ Greška: Model '{izabrana_sifra}' u boji '{izabrana_boja}' uopšte ne postoji na stanju u ovoj kategoriji!")
                else:
                    current_stanje = int(filtriran_red.iloc[0]["broj_pari"])
                    
                    if current_stanje < kolicina_izlaza:
                        st.error(f"❌ Greška: Nemate dovoljno robe! Na stanju ima samo {current_stanje} kom, a vi pokušavate da iznesete {kolicina_izlaza} kom.")
                    else:
                        with st.spinner("Zapisivanje u toku..."):
                            try:
                                conn = uzmi_vezu_sa_bazom()
                                cursor = conn.cursor()
                                cursor.execute('''
                                    INSERT INTO izlaz_robe (datum, sifra_artikla, boja_artikla, kolicina_izlaz, grad, prodajna_cena, nabavna_cena)
                                    VALUES (%s, %s, %s, %s, %s, %s, %s)
                                ''', (izabrani_datum.strftime("%Y-%m-%d"), izabrana_sifra, izabrana_boja, kolicina_izlaza, izabrani_grad, prodajna_cena_par, nabavna_cena_par))
                                
                                novo_stanje = current_stanje - kolicina_izlaza
                                cursor.execute('UPDATE artikli SET broj_pari = %s WHERE sifra = %s AND boja = %s AND sezona = %s', (novo_stanje, izabrana_sifra, izabrana_boja, izabrana_sezona))
                                conn.commit()
                                conn.close()
                                
                                # Čistimo keš da se tabele odmah osveže sa novim stanjem
                                ucitaj_artikle_za_sezonu.clear()
                                ucitaj_istoriju_izlaza_za_sezonu.clear()
                                
                                st.success(f"✅ Uspešno proknjižen izlaz za {izabrani_grad}! Skinuto {kolicina_izlaza} kom sa stanja.")
                                st.rerun()
                            except Exception as e:
                                st.error(f"Sistemska greška pri upisu: {e}")

        st.markdown("---")
        st.subheader(f"📋 Istorija dnevnih izlaza robe za sekciju: {izabrana_sezona}")
        
        df_izlazi = ucitaj_istoriju_izlaza_za_sezonu(izabrana_sezona)
        
        if not df_izlazi.empty:
            col_filter1, col_filter2, col_filter3 = st.columns(3)
            with col_filter1: od_datuma = st.date_input("Od datuma:", datetime.strptime(df_izlazi['Datum'].min(), "%Y-%m-%d") if not df_izlazi.empty else datetime.now())
            with col_filter2: do_datuma = st.date_input("Do datuma:", datetime.now())
            with col_filter3: izabrani_grad_filter = st.selectbox("Izaberi grad za tabelu i Excel:", ["SVI GRADOVI"] + lista_gradova)
            
            od_str, do_str = od_datuma.strftime("%Y-%m-%d"), do_datuma.strftime("%Y-%m-%d")
            df_filtrirano = df_izlazi[(df_izlazi['Datum'] >= od_str) & (df_izlazi['Datum'] <= do_str)]
            if izabrani_grad_filter != "SVI GRADOVI":
                df_filtrirano = df_filtrirano[df_filtrirano['Grad'] == izabrani_grad_filter]
            
            excel_izlazi = konvertuj_u_excel(df_filtrirano)
            st.download_button(
                label=f"🟢 Preuzmi Excel ({od_datuma.strftime('%d.%m.%Y.')} - {do_datuma.strftime('%d.%m.%Y.')}) - {izabrani_grad_filter}",
                data=excel_izlazi, file_name=f"izlazi_{izabrana_sezona}_{izabrani_grad_filter.replace(' ', '_')}_{od_str}_do_{do_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            if not df_filtrirano.empty:
                st.dataframe(df_filtrirano, use_container_width=True)
            else:
                st.info("Nema zabeleženih izlaza za izabrani period i grad.")
        else:
            st.write(f"Još uvek nema zabeleženih izlaza robe za sekciju {izabrana_sezona}.")