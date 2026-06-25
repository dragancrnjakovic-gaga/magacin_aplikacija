import streamlit as st
import sqlite3
import cloudinary
import cloudinary.uploader
import pandas as pd
from datetime import datetime, timedelta
import io

# --- KONFIGURACIJA CLOUDINARY-JA ---
cloudinary.config(
    cloud_name = st.secrets["cloudinary"]["cloud_name"],
    api_key = st.secrets["cloudinary"]["api_key"],
    api_secret = st.secrets["cloudinary"]["api_secret"],
    secure = True
)

# --- PODEŠAVANJE LOKALNE SQLITE BAZE I MIGRACIJA ---
def uzmi_vezu_sa_bazom():
    conn = sqlite3.connect("lokalni_magacin.db")
    conn.row_factory = sqlite3.Row
    return conn

def kreiraj_tabele_i_migracije():
    conn = uzmi_vezu_sa_bazom()
    cursor = conn.cursor()
    
    # Tabela artikala
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
    
    # Osnovna tabela izlaza robe
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS izlaz_robe (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            datum TEXT,
            sifra_artikla TEXT,
            boja_artikla TEXT,
            kolicina_izlaz INTEGER
        )
    ''')
    
    # AUTOMATSKA MIGRACIJA: Provera i dodavanje kolona koje fale u staroj bazi
    cursor.execute("PRAGMA table_info(izlaz_robe)")
    postojece_kolone = [kolona[1] for kolona in cursor.fetchall()]
    
    novi_stubovi = {
        "grad": "TEXT",
        "prodajna_cena": "REAL DEFAULT 0.0",
        "zbir_prodajna": "REAL DEFAULT 0.0",
        "nabavna_cena": "REAL DEFAULT 0.0",
        "zbir_nabavna": "REAL DEFAULT 0.0"
    }
    
    for kolona_ime, kolona_tip in novi_stubovi.items():
        if kolona_ime not in postojece_kolone:
            cursor.execute(f"ALTER TABLE izlaz_robe ADD COLUMN {kolona_ime} {kolona_tip}")
    
    # Šifrarnik boja
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS sifrarnik_boja (
            boja TEXT PRIMARY KEY
        )
    ''')
    
    cursor.execute("SELECT COUNT(*) FROM sifrarnik_boja")
    if cursor.fetchone()[0] == 0:
        pocetne_boje = [("Black",), ("Blue",), ("Red",), ("Gray",), ("White",), ("Beige",)]
        cursor.executemany("INSERT INTO sifrarnik_boja (boja) VALUES (?)", pocetne_boje)
        
    conn.commit()
    conn.close()

# Pokretanje baze i provera kolona
kreiraj_tabele_i_migracije()

# --- NAPREDNO KEŠIRANJE PODATAKA ---
@st.cache_data(ttl=10)
def ucitaj_artikle_za_sezonu(sezona):
    conn = uzmi_vezu_sa_bazom()
    df = pd.read_sql_query("SELECT * FROM artikli WHERE sezona = ? ORDER BY sifra ASC, boja ASC", conn, params=(sezona,))
    conn.close()
    return df

@st.cache_data(ttl=120)
def ucitaj_boje():
    conn = uzmi_vezu_sa_bazom()
    cursor = conn.cursor()
    cursor.execute("SELECT boja FROM sifrarnik_boja ORDER BY boja ASC")
    boje = [red[0] for red in cursor.fetchall()]
    conn.close()
    return boje

def konvertuj_u_excel(df, sheet_name='Magacin'):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()

def pronadji_sliku_u_df(df, sifra):
    if df.empty:
        return ""
    filtrirano = df[(df["sifra"] == sifra) & (df["slika_putanja"] != "") & (df["slika_putanja"].notna())]
    if not filtrirano.empty:
        return filtrirano.iloc[0]["slika_putanja"]
    return ""

# --- IZGLED I STILIZACIJA APLIKACIJE ---
st.set_page_config(page_title="Magacin LIVE", layout="wide")

st.markdown("""
    <style>
    .block-container { padding-top: 3.5rem !important; padding-bottom: 2rem !important; }
    h1 { font-size: 1.8rem !important; padding-bottom: 10px !important; margin: 0px !important; }
    h2 { font-size: 1.35rem !important; padding-bottom: 15px !important; margin: 0px !important; }
    h3 { font-size: 1.05rem !important; font-weight: bold !important; }
    [data-testid="stMetricValue"] { font-size: 1.05rem !important; }
    [data-testid="stMetricLabel"] { font-size: 0.75rem !important; }
    .stTextInput p, .stNumberInput p, .stSelectbox p, .stDateInput p, label p { font-size: 0.85rem !important; }
    .stAlert p { font-size: 0.85rem !important; }
    .stExpander p { font-size: 0.8rem !important; }
    
    div[data-testid="stHorizontalBlock"] { 
        background: #1e2229; padding: 15px; border-radius: 6px; margin-bottom: 10px; border: 1px solid #2d3139; 
    }
    div.stButton > button { width: 100% !important; padding: 2px 10px !important; font-size: 0.9rem !important; height: 38px !important; }
    div.skrivena-labela label { display: none !important; }
    
    .indikator-stranice {
        background-color: #2e3440; padding: 6px 12px; border-radius: 4px; font-weight: bold; color: #88c0d0; font-size: 0.95rem;
        border-left: 4px solid #88c0d0; margin-bottom: 15px; display: inline-block;
    }

    [data-testid="stHorizontalBlock"]:has(button[key^="vrh_"]),
    [data-testid="stHorizontalBlock"]:has(button[key^="dole_"]) {
        display: flex !important; flex-direction: row !important; justify-content: space-between !important; align-items: center !important; flex-wrap: nowrap !important;
    }
    </style>
""", unsafe_allow_html=True)

if "skroluj_na_vrh" in st.session_state and st.session_state["skroluj_na_vrh"]:
    st.components.v1.html("<script>window.parent.document.querySelector('.stMain').scrollTo(0, 0);</script>", height=0, width=0)
    st.session_state["skroluj_na_vrh"] = False

st.title("📦 LIVE VERZIJA - Glavni Magacin")

izabrana_sezona = st.sidebar.radio("🌸 IZABERI KATEGORIJU / SEZONU:", ["Proleće-Leto", "Jesen-Zima", "Torbe"])
st.sidebar.markdown("---")
meni = st.sidebar.selectbox("Izaberi opciju:", ["Trenutno stanje", "Unos nove robe", "Evidencija izlaza (Po danima)"])

if "trenutna_stranica" not in st.session_state: st.session_state["trenutna_stranica"] = 1
if "prethodna_sezona" not in st.session_state: st.session_state["prethodna_sezona"] = izabrana_sezona
if "prethodni_meni" not in st.session_state: st.session_state["prethodni_meni"] = meni

if izabrana_sezona != st.session_state["prethodna_sezona"] or meni != st.session_state["prethodni_meni"]:
    st.session_state["trenutna_stranica"] = 1
    st.session_state["prethodna_sezona"] = izabrana_sezona
    st.session_state["prethodni_meni"] = meni

# Inicijalizacija ID-jeva formi za pražnjenje polja
if "izlaz_form_id" not in st.session_state: st.session_state["izlaz_form_id"] = 0
if "unos_form_id" not in st.session_state: st.session_state["unos_form_id"] = 0

def prikazi_gornju_paginaciju(broj_stranica, trenutna):
    if broj_stranica <= 1: return
    pag_cols = st.columns([1, 2, 1])
    with pag_cols[0]:
        if st.button("⬅️ Prethodna", disabled=(trenutna == 1), key="vrh_prev"):
            st.session_state["trenutna_stranica"] = trenutna - 1
            st.session_state["skroluj_na_vrh"] = True
            st.rerun()
    with pag_cols[1]:
        opcije_stranica = [i for i in range(1, broj_stranica + 1)]
        st.markdown('<div class="skrivena-labela">', unsafe_allow_html=True)
        izbor = st.selectbox("Izaberi stranicu:", options=opcije_stranica, index=trenutna - 1, format_func=lambda x: f"Stranica {x} od {broj_stranica}")
        st.markdown('</div>', unsafe_allow_html=True)
        if izbor != trenutna:
            st.session_state["trenutna_stranica"] = izbor
            st.session_state["skroluj_na_vrh"] = True
            st.rerun()
    with pag_cols[2]:
        if st.button("Sledeća ➡️", disabled=(trenutna == broj_stranica), key="vrh_next"):
            st.session_state["trenutna_stranica"] = trenutna + 1
            st.session_state["skroluj_na_vrh"] = True
            st.rerun()

def prikazi_donju_paginaciju(broj_stranica, trenutna):
    if broj_stranica <= 1: return
    st.write("")
    dole_cols = st.columns([1.5, 4, 1.5])
    with dole_cols[0]:
        if st.button("⬅️ PRETHODNA STRANICA", disabled=(trenutna == 1), key="dole_veliko_prev"):
            st.session_state["trenutna_stranica"] = trenutna - 1
            st.session_state["skroluj_na_vrh"] = True
            st.rerun()
    with dole_cols[2]:
        if st.button("SLEDEĆA STRANICA ➡️", disabled=(trenutna == broj_stranica), key="dole_veliko_next"):
            st.session_state["trenutna_stranica"] = trenutna + 1
            st.session_state["skroluj_na_vrh"] = True
            st.rerun()

# --- OPCIJA: UNOS NOVE ROBE ---
if meni == "Unos nove robe":
    st.header(f"➕ Unos novog artikla ({izabrana_sezona})")
    lista_boja = ucitaj_boje()
    
    u_id = st.session_state["unos_form_id"]
    
    col1, col2 = st.columns(2)
    with col1:
        sifra = st.text_input("Šifra modela:", key=f"u_sif_{u_id}").strip().upper()
        boja = st.selectbox("Boja modela:", lista_boja, key=f"u_boj_{u_id}")
        
        labela_kol = "Količina (komada/pari):" if izabrana_sezona == "Torbe" else "Količina pari:"
        labela_kut = "Broj komada u jednoj kutiji/pakovanju:" if izabrana_sezona == "Torbe" else "Broj pari u jednoj kutiji:"
        
        broj_pari = st.number_input(labela_kol, min_value=0, step=1, value=None, key=f"u_kol_{u_id}")
        pari_u_kutiji = st.number_input(labela_kut, min_value=1, step=1, value=None, key=f"u_kut_{u_id}")
    with col2:
        prodajna_cena = st.number_input("Prodajna cena (RSD):", min_value=0.0, step=50.0, value=None, key=f"u_pc_{u_id}")
        internet_cena = st.number_input("Internet cena (RSD):", min_value=0.0, step=50.0, value=None, key=f"u_ic_{u_id}")
        slika = st.file_uploader("Ubaci sliku modela:", type=["jpg", "jpeg", "png"], key=f"slika_unos_{u_id}")
        
    podaci_nedostaju = (sifra == "" or boja is None or boja == "" or broj_pari is None or pari_u_kutiji is None or prodajna_cena is None or internet_cena is None)
    
    st.write("")
    dugme_potvrdi = st.button("Sačuvaj artikal u bazu", type="primary", disabled=podaci_nedostaju)
    
    if dugme_potvrdi:
        df_trenutni = ucitaj_artikle_za_sezonu(izabrana_sezona)
        url_slike = ""
        if slika is not None:
            with st.spinner("Slanje slike..."):
                try:
                    rezultat_slike = cloudinary.uploader.upload(
                        slika, folder="magacin_live/", public_id=f"live_{sifra}_{boja}",
                        transformation=[{"width": 800, "crop": "limit"}, {"quality": "auto", "fetch_format": "auto"}]
                    )
                    url_slike = apiKey = rezultat_slike["secure_url"]
                except Exception as e:
                    st.error(f"Greška pri slanju slike: {e}")
        else:
            url_slike = pronadji_sliku_u_df(df_trenutni, sifra)
        
        try:
            conn = uzmi_vezu_sa_bazom()
            cursor = conn.cursor()
            cursor.execute('''
                INSERT INTO artikli (sifra, boja, sezona, broj_pari, pari_u_kutiji, prodajna_cena, internet_cena, slika_putanja)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ''', (sifra, boja, izabrana_sezona, broj_pari, pari_u_kutiji, prodajna_cena, internet_cena, url_slike))
            conn.commit()
            conn.close()
            
            ucitaj_artikle_za_sezonu.clear()
            st.session_state["unos_form_id"] += 1
            st.success(f"Uspešno sačuvan model {sifra} - {boja}!")
            st.rerun()
        except sqlite3.IntegrityError:
            st.error(f"Greška: Model sa šifrom '{sifra}' u boji '{boja}' već postoji u bazi!")

# --- OPCIJA: TRENUTNO STANJE ---
elif meni == "Trenutno stanje":
    st.header(f"📋 Stanje robe - Sekcija: {izabrana_sezona}")
    lista_boja = ucitaj_boje()
    df = ucitaj_artikle_za_sezonu(izabrana_sezona)
    
    if df.empty:
        st.info(f"U bazi za sekciju {izabrana_sezona} trenutno nema podataka.")
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
        
        st.download_button(
            label="🟢 Preuzmi trenutno stanje kao Excel", data=konvertuj_u_excel(df_excel, 'Stanje'),
            file_name=f"stanje_{izabrana_sezona}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        st.markdown("---")
        pretraga = st.text_input("🔍 Pretraži po šifri modela:", "").strip().upper()
        df_prikaz = df[df["sifra"].str.contains(pretraga, na=False)] if pretraga else df.copy()
        
        if df_prikaz.empty:
            st.warning(f"Nema rezultata za šifru '{pretraga}'")
        else:
            BROJ_ARTIKALA_PO_STRANICI = 20
            ukupno_artikala = len(df_prikaz)
            broj_stranica = (ukupno_artikala // BROJ_ARTIKALA_PO_STRANICI) + (1 if ukupno_artikala % BROJ_ARTIKALA_PO_STRANICI > 0 else 0)
            
            if pretraga != "" or st.session_state["trenutna_stranica"] > broj_stranica:
                st.session_state["trenutna_stranica"] = 1
                
            if broj_stranica > 1 and not pretraga:
                prikazi_gornju_paginaciju(broj_stranica, st.session_state["trenutna_stranica"])
            
            if broj_stranica > 1:
                st.markdown(f'<div class="indikator-stranice">📄 Stranica: {st.session_state["trenutna_stranica"]} od {broj_stranica}</div>', unsafe_allow_html=True)
            
            start_indeks = (st.session_state["trenutna_stranica"] - 1) * BROJ_ARTIKALA_PO_STRANICI
            df_za_prikaz = df_prikaz.iloc[start_indeks:start_indeks + BROJ_ARTIKALA_PO_STRANICI]
            
            for index, row in df_za_prikaz.iterrows():
                sif, boj = row['sifra'], row['boja']
                kljuc_id = f"live_{sif}_{boj}"
                trenutna_slika = row["slika_putanja"] or pronadji_sliku_u_df(df, sif)
                
                br_kutija = row["broj_pari"] // row["pari_u_kutiji"]
                ost_pari = row["broj_pari"] % row["pari_u_kutiji"]
                
                with st.container():
                    col_slika, col_detalji, col_akcije = st.columns([1.2, 3, 1.5])
                    with col_slika:
                        if trenutna_slika:
                            st.image(trenutna_slika.replace("/upload/", "/upload/w_150,c_limit,q_auto,f_auto/"), width=120)
                        else:
                            st.write("❌ Nema slike")
                            
                    with col_detalji:
                        st.subheader(f"Šifra: {sif} | Boja: {boj}")
                        c1, c2, c3, c4 = st.columns(4)
                        m_kol_label = "Ukupno komada" if izabrana_sezona == "Torbe" else "Ukupno pari"
                        m_pak_sufiks = f"{br_kutija} pak. + {ost_pari} kom" if izabrana_sezona == "Torbe" else f"{br_kutija} kut. + {ost_pari} par"
                        
                        c1.metric(m_kol_label, f"{row['broj_pari']} kom")
                        c2.metric("Pakovanje", m_pak_sufiks)
                        c3.metric("Prodajna cena", f"{int(row['prodajna_cena'])} din")
                        c4.metric("Internet cena", f"{int(row['internet_cena'])} din")
                        
                    with col_akcije:
                        with st.expander("🛠️ Izmeni / Obriši"):
                            nova_boja_izmena = st.selectbox("Uredi boju:", lista_boja, index=lista_boja.index(boj) if boj in lista_boja else 0, key=f"b_{kljuc_id}")
                            nova_kol = st.number_input("Novo stanje:", min_value=0, value=int(row['broj_pari']), key=f"k_{kljuc_id}")
                            nova_p_cena = st.number_input("Prodajna cena:", min_value=0.0, value=float(row['prodajna_cena']), key=f"pc_{kljuc_id}")
                            nova_i_cena = st.number_input("Internet cena:", min_value=0.0, value=float(row['internet_cena']), key=f"ic_{kljuc_id}")
                            
                            col_b1, col_b2 = st.columns(2)
                            with col_b1:
                                if st.button("💾 Snimi", key=f"Save_{kljuc_id}"):
                                    conn = uzmi_vezu_sa_bazom()
                                    cursor = conn.cursor()
                                    try:
                                        cursor.execute('''
                                            UPDATE artikli SET boja = ?, broj_pari = ?, prodajna_cena = ?, internet_cena = ?
                                            WHERE sifra = ? AND boja = ? AND sezona = ?
                                        ''', (nova_boja_izmena, nova_kol, nova_p_cena, nova_i_cena, sif, boj, izabrana_sezona))
                                        conn.commit()
                                        ucitaj_artikle_za_sezonu.clear()
                                        st.success("Sačuvano!")
                                        st.rerun()
                                    except sqlite3.IntegrityError:
                                        st.error("Model već postoji!")
                                    finally:
                                        conn.close()
                            with col_b2:
                                if st.button("🗑️ Obriši", key=f"Del_{kljuc_id}"):
                                    conn = uzmi_vezu_sa_bazom()
                                    cursor = conn.cursor()
                                    cursor.execute("DELETE FROM artikli WHERE sifra = ? AND boja = ? AND sezona = ?", (sif, boj, izabrana_sezona))
                                    conn.commit()
                                    conn.close()
                                    ucitaj_artikle_za_sezonu.clear()
                                    st.rerun()
                st.markdown("---")
            
            if broj_stranica > 1 and not pretraga:
                prikazi_donju_paginaciju(broj_stranica, st.session_state["trenutna_stranica"])

# --- OPCIJA: EVIDENCIJA IZLAZA ROBE ---
elif meni == "Evidencija izlaza (Po danima)":
    st.header(f"📦 Dnevni izlaz i Prodaja - Sekcija: {izabrana_sezona}")
    
    df_artikli = ucitaj_artikle_za_sezonu(izabrana_sezona)
    sve_sifre = sorted(df_artikli["sifra"].unique().tolist()) if not df_artikli.empty else []
    
    gradovi = ["Internet", "Mladenovac Gore", "Mladenovac Dole", "Smederevska Palanka", 
               "Zaječar", "Subotica", "Aleksinac", "Loznica", "Sremska Mitrovica", 
               "Pančevo", "Vršac", "Bečej", "Prokuplje"]
               
    if not sve_sifre:
        st.info("Nema unete robe u ovoj kategoriji.")
    else:
        st.subheader("📝 Unos novog izlaza")
        
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            izabrani_datum = st.date_input("Datum izlaza:", datetime.now(), key="b_date")
            izabrana_sifra = st.selectbox("Šifra modela:", sve_sifre, key="b_sif")
            
            dostupne_boje = sorted(df_artikli[df_artikli["sifra"] == izabrana_sifra]["boja"].tolist())
            izabrana_boja = st.selectbox("Boja modela:", dostupne_boje, key="b_boj")
            izabrani_grad = st.selectbox("Izaberi grad / lokaciju:", gradovi, key="b_grad")

        # Stanje artikla iz baze podataka
        trenutni_artikal = df_artikli[(df_artikli["sifra"] == izabrana_sifra) & (df_artikli["boja"] == izabrana_boja)]
        baza_stanje = int(trenutni_artikal.iloc[0]["broj_pari"]) if not trenutni_artikal.empty else 0
        
        with col_f2:
            st.info(f"💡 Trenutno stanje na zalihama: **{baza_stanje} kom/par**")
            
            f_id = st.session_state["izlaz_form_id"]
            
            kolicina_izlaza = st.number_input("Količina za izlaz (broj pari):", min_value=1, max_value=max(1, baza_stanje), step=1, value=None, key=f"qty_input_{f_id}")
            p_cena_unos = st.number_input("Prodajna cena po paru (RSD):", min_value=0.0, step=50.0, value=None, key=f"pc_input_{f_id}")
            
            izracunata_kol = kolicina_izlaza if kolicina_izlaza is not None else 0
            izracunata_p_cena = p_cena_unos if p_cena_unos is not None else 0.0
            
            zbir_prodajna = izracunata_kol * izracunata_p_cena
            st.markdown(f"**💰 Zbir prodajna:** <span style='color:#88c0d0; font-size:1.1rem;'>{int(zbir_prodajna)} RSD</span>", unsafe_allow_html=True)
            
            st.write("")
            n_cena_unos = st.number_input("Nabavna cena po paru (Opciono - RSD):", min_value=0.0, step=50.0, value=None, key=f"nc_input_{f_id}")
            
            izracunata_n_cena = n_cena_unos if n_cena_unos is not None else 0.0
            zbir_nabavna = izracunata_kol * izracunata_n_cena
            st.markdown(f"**📉 Zbir nabavna:** {int(zbir_nabavna)} RSD")

        dugme_onemoguceno = (p_cena_unos is None or p_cena_unos <= 0.0)

        st.write("")
        dugme_upisi = st.button("Zapiši izlaz robe", type="primary", disabled=dugme_onemoguceno)
        
        if dugme_upisi:
            konacna_kolicina = kolicina_izlaza if kolicina_izlaza is not None else 1
            zbir_prodajna = konacna_kolicina * izracunata_p_cena
            zbir_nabavna = konacna_kolicina * izracunata_n_cena
            
            conn = uzmi_vezu_sa_bazom()
            cursor = conn.cursor()
            
            cursor.execute('''
                INSERT INTO izlaz_robe (datum, sifra_artikla, boja_artikla, kolicina_izlaz, grad, prodajna_cena, zbir_prodajna, nabavna_cena, zbir_nabavna) 
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (izabrani_datum.strftime("%Y-%m-%d"), izabrana_sifra, izabrana_boja, konacna_kolicina, izabrani_grad, izracunata_p_cena, zbir_prodajna, izracunata_n_cena, zbir_nabavna))
            
            cursor.execute('UPDATE artikli SET broj_pari = ? WHERE sifra = ? AND boja = ? AND sezona = ?',
                           (baza_stanje - konacna_kolicina, izabrana_sifra, izabrana_boja, izabrana_sezona))
            
            conn.commit()
            conn.close()
            
            st.session_state["izlaz_form_id"] += 1
            ucitaj_artikle_za_sezonu.clear()
            
            st.success(f"Uspešno proknjižen izlaz za {izabrani_grad}! Skinuto {konacna_kolicina} pari sa zaliha.")
            st.rerun()

        # --- DODATI FILTERI ZA PREUZIMANJE EXCEL TABELE ---
        st.markdown("---")
        st.subheader("📊 Pregled i evidencija proknjiženih izlaza")
        
        # Učitavanje svih sirovih podataka za odabranu sezonu
        conn = uzmi_vezu_sa_bazom()
        df_svi_izlazi = pd.read_sql_query('''
            SELECT 
                ir.datum AS "Datum", 
                ir.sifra_artikla AS "Šifra modela", 
                ir.boja_artikla AS "Boja", 
                ir.grad AS "Grad/Lokacija",
                ir.kolicina_izlaz AS "Broj pari", 
                ir.prodajna_cena AS "Prodajna cena (par)", 
                ir.zbir_prodajna AS "Zbir prodajna", 
                ir.nabavna_cena AS "Nabavna cena (par)", 
                ir.zbir_nabavna AS "Zbir nabavna"
            FROM izlaz_robe ir 
            INNER JOIN artikli a ON ir.sifra_artikla = a.sifra AND ir.boja_artikla = a.boja 
            WHERE a.sezona = ?
            ORDER BY ir.id DESC
        ''', conn, params=(izabrana_sezona,))
        conn.close()
        
        if df_svi_izlazi.empty:
            st.info("U ovoj kategoriji još uvek nema zapisanih izlaza robe.")
        else:
            # Sekcija sa filterima za Excel i prikaz
            st.markdown("##### 🔍 Filteri za tabelu i preuzimanje")
            col_filt1, col_filt2 = st.columns(2)
            
            with col_filt1:
                excel_gradovi_opcije = ["SVI GRADOVI"] + gradovi
                izabrani_grad_excel = st.selectbox("Izaberi grad za Excel:", excel_gradovi_opcije, key="excel_filter_grad")
            
            with col_filt2:
                # Automatski postavljamo period od početka meseca do danas kao podrazumevani
                danasnji_dan = datetime.now()
                pocetak_meseca = danasnji_dan.replace(day=1)
                
                izabrani_period = st.date_input(
                    "Izaberi period (Od - Do):", 
                    value=(pocetak_meseca, danasnji_dan),
                    key="excel_filter_period"
                )
            
            # Primena filtera na DataFrame
            df_filtriran = df_svi_izlazi.copy()
            
            # 1. Filter za grad
            if izabrani_grad_excel != "SVI GRADOVI":
                df_filtriran = df_filtriran[df_filtriran["Grad/Lokacija"] == izabrani_grad_excel]
            
            # 2. Filter za datumski period (provera da li su izabrana oba datuma)
            if isinstance(izabrani_period, tuple) and len(izabrani_period) == 2:
                datum_od = izabrani_period[0].strftime("%Y-%m-%d")
                datum_do = izabrani_period[1].strftime("%Y-%m-%d")
                df_filtriran = df_filtriran[(df_filtriran["Datum"] >= datum_od) & (df_filtriran["Datum"] <= datum_do)]
            
            st.write("")
            
            if df_filtriran.empty:
                st.warning("Nema podataka za izabrani grad i period.")
            else:
                # Generisanje naziva fajla u zavisnosti od filtera
                grad_sufiks = izabrani_grad_excel.lower().replace(" ", "_")
                
                st.download_button(
                    label=f"🟢 Preuzmi filtriranu tabelu ({len(df_filtriran)} zapisa) kao Excel", 
                    data=konvertuj_u_excel(df_filtriran, 'Izlazi_Robe'),
                    file_name=f"izlazi_{grad_sufiks}_{datetime.now().strftime('%Y-%m-%d')}.xlsx", 
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                st.write("")
                st.dataframe(df_filtriran, use_container_width=True)