import streamlit as st
import psycopg2
from psycopg2.extras import RealDictCursor
import cloudinary
import cloudinary.uploader
import pandas as pd
from datetime import datetime
import io

# --- KONFIGURACIJA CLOUDINARY-JA ---
# Podaci se povlače iz bezbednih Streamlit Secrets-a
cloudinary.config(
    cloud_name = st.secrets["cloudinary"]["cloud_name"],
    api_key = st.secrets["cloudinary"]["api_key"],
    api_secret = st.secrets["cloudinary"]["api_secret"],
    secure = True
)

# --- PODEŠAVANJE NEON POSTGRES BAZE ---
def uzmi_vezu_sa_bazom():
    # Povezivanje preko tajnog internet linka iz Neon-a
    return psycopg2.connect(st.secrets["postgres"]["url"])

def kreiraj_tabele():
    conn = uzmi_vezu_sa_bazom()
    cursor = conn.cursor()
    
    # Kreiranje tabele artikala na Neon-u
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
    
    # Kreiranje tabele izlaza na Neon-u
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS izlaz_robe (
            id SERIAL PRIMARY KEY,
            datum TEXT,
            sifra_artikla TEXT,
            boja_artikla TEXT,
            kolicina_izlaz INTEGER
        )
    ''')
    
    # Kreiranje šifrarnika boja
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
    </style>
""", unsafe_allow_html=True)

st.title("📦 Višekorisnički sistem za praćenje stanja u magacinu")

# 1. SEZONA
izabrana_sezona = st.sidebar.radio("🌸 IZABERI SEZONU:", ["Proleće-Leto", "Jesen-Zima"])
st.sidebar.markdown("---")

# 2. FUNKCIJA
meni = st.sidebar.selectbox("Izaberi opciju:", ["Trenutno stanje", "Unos nove robe", "Evidencija izlaza (Po danima)"])
st.sidebar.info(f"Trenutno radite u sekciji:\n**{izabrana_sezona}**")

if "reset_brojac" not in st.session_state:
    st.session_state["reset_brojac"] = 0

# --- OPCIJA 1: UNOS NOVE ROBE ---
if meni == "Unos nove robe":
    st.header(f"➕ Unos novog artikla ({izabrana_sezona})")
    lista_boja = ucitaj_boje()
    
    with st.form("forma_za_unos", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            sifra = st.text_input("Šifra modela:").strip().upper()
            boja = st.selectbox("Boja modela:", lista_boja)
            broj_pari = st.number_input("Količina pari:", min_value=0, step=1)
            pari_u_kutiji = st.number_input("Broj pari u jednoj kutiji:", min_value=1, step=1)
        with col2:
            prodajna_cena = st.number_input("Prodajna cena (RSD):", min_value=0.0, step=50.0)
            internet_cena = st.number_input("Internet cena (RSD):", min_value=0.0, step=50.0)
            slika = st.file_uploader("Ubaci sliku modela:", type=["jpg", "jpeg", "png"])
            
        dugme_potvrdi = st.form_submit_button("Sačuvaj artikal u bazu")
        
        if dugme_potvrdi:
            if sifra == "" or boja is None or boja == "":
                st.error("Greška: Šifra i boja ne smeju biti prazne!")
            else:
                url_slike = ""
                if slika is not None:
                    with st.spinner("Slanje slike na Cloudinary..."):
                        try:
                            # Slanje fajla direktno na internet u folder "magacin"
                            rezultat_slike = cloudinary.uploader.upload(
                                slika, 
                                folder="magacin/",
                                public_id=f"{sifra}_{boja}"
                            )
                            url_slike = resultado_slike = rezultat_slike["secure_url"]
                        except Exception as e:
                            st.error(f"Greška pri slanju slike: {e}")
                
                try:
                    conn = uzmi_vezu_sa_bazom()
                    cursor = conn.cursor()
                    cursor.execute('''
                        INSERT INTO artikli (sifra, boja, sezona, broj_pari, pari_u_kutiji, prodajna_cena, internet_cena, slika_putanja)
                        VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
                    ''', (sifra, boja, izabrana_sezona, broj_pari, pari_u_kutiji, prodajna_cena, internet_cena, url_slike))
                    conn.commit()
                    conn.close()
                    st.success(f"Uspešno sačuvan model: Šifra '{sifra}' - Boja '{boja}'!")
                    st.rerun()
                except psycopg2.IntegrityError:
                    st.error(f"Greška: Model sa šifrom '{sifra}' u boji '{boja}' već postoji u bazi!")

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
                    st.success(f"Boja '{nova_boja_unos}' je dodata!")
                    st.rerun()
                except psycopg2.IntegrityError:
                    st.warning("Boja već postoji u listi.")

# --- OPCIJA 2: TRENUTNO STANJE ---
elif meni == "Trenutno stanje":
    st.header(f"📋 Stanje robe - Sezona: {izabrana_sezona}")
    
    conn = uzmi_vezu_sa_bazom()
    df = pd.read_sql_query("SELECT * FROM artikli WHERE sezona = %s", conn, params=(izabrana_sezona,))
    conn.close()
    
    if df.empty:
        st.info(f"U sezoni {izabrana_sezona} trenutno nema unete robe.")
    else:
        df_excel = df.copy()
        df_excel["Broj kutija"] = df_excel["broj_pari"] // df_excel["pari_u_kutiji"]
        df_excel["Ostatak pari"] = df_excel["broj_pari"] % df_excel["pari_u_kutiji"]
        df_excel = df_excel.rename(columns={
            "sifra": "Šifra modela", "boja": "Boja", "sezona": "Sezona", 
            "broj_pari": "Ukupno pari", "pari_u_kutiji": "Pari u kutiji",
            "prodajna_cena": "Prodajna cena (RSD)", "internet_cena": "Internet cena (RSD)"
        }).drop(columns=["slika_putanja"], errors="ignore")
        
        excel_podaci = konvertuj_u_excel(df_excel)
        st.download_button(
            label="🟢 Preuzmi stanje kao Excel tabelu (.xlsx)",
            data=excel_podaci,
            file_name=f"stanje_magacina_{izabrana_sezona}_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        st.markdown("---")
        pretraga = st.text_input("🔍 Pretraži ovu sezonu po šifri modela:", "").strip().upper()
        if pretraga:
            df = df[df["sifra"].str.contains(pretraga, na=False)]
        
        if df.empty:
            st.warning(f"Nema rezultata za šifru '{pretraga}'")
        else:
            for index, row in df.iterrows():
                sif = row['sifra']
                boj = row['boja']
                kljuc_id = f"{sif}_{boj}"
                trenutna_slika = row["slika_putanja"]
                
                br_kutija = row["broj_pari"] // row["pari_u_kutiji"]
                ost_pari = row["broj_pari"] % row["pari_u_kutiji"]
                
                with st.container():
                    col_slika, col_detalji, col_akcije = st.columns([1.2, 3, 1.5])
                    with col_slika:
                        if trenutna_slika:
                            st.image(trenutna_slika, width=120)
                            with st.expander("🔍 Vidi veliku sliku"):
                                st.image(trenutna_slika, use_container_width=True)
                        else:
                            st.write("❌ Nema slike")
                            
                    with col_detalji:
                        st.subheader(f"Šifra modela: {sif} | Boja: {boj}")
                        c1, c2, c3, c4 = st.columns(4)
                        c1.metric("Ukupno pari", f"{row['broj_pari']} kom")
                        c2.metric("Pakovanje", f"{br_kutija} kut. + {ost_pari} par")
                        c3.metric("Prodajna", f"{row['prodajna_cena']} din")
                        c4.metric("Internet", f"{row['internet_cena']} din")
                        
                    with col_akcije:
                        ekspander = st.expander("🛠️ Izmeni / Obriši")
                        with ekspander:
                            st.write("**Uredi podatke:**")
                            nova_kol = st.number_input("Novo ukupno pari:", min_value=0, value=int(row['broj_pari']), step=1, key=f"kol_{kljuc_id}")
                            nova_p_cena = st.number_input("Prodajna cena (RSD):", min_value=0.0, value=float(row['prodajna_cena']), step=50.0, key=f"pc_{kljuc_id}")
                            nova_i_cena = st.number_input("Internet cena (RSD):", min_value=0.0, value=float(row['internet_cena']), step=50.0, key=f"ic_{kljuc_id}")
                            nova_slika_file = st.file_uploader("Zameni sliku artikla:", type=["jpg", "jpeg", "png"], key=f"img_{kljuc_id}")
                            
                            col_b1, col_b2 = st.columns(2)
                            with col_b1:
                                if st.button("💾 Snimi", key=f"Snimi_{kljuc_id}"):
                                    finalna_putanja_slike = trenutna_slika
                                    if nova_slika_file is not None:
                                        with st.spinner("Menjanje slike..."):
                                            try:
                                                rez_nove_slike = cloudinary.uploader.upload(
                                                    nova_slika_file,
                                                    folder="magacin/",
                                                    public_id=f"{sif}_{boj}"
                                                )
                                                finalna_putanja_slike = rez_nove_slike["secure_url"]
                                            except:
                                                pass
                                            
                                    conn = uzmi_vezu_sa_bazom()
                                    cursor = conn.cursor()
                                    cursor.execute('''
                                        UPDATE artikli 
                                        SET broj_pari = %s, prodajna_cena = %s, internet_cena = %s, slika_putanja = %s
                                        WHERE sifra = %s AND boja = %s AND sezona = %s
                                    ''', (nova_kol, nova_p_cena, nova_i_cena, finalna_putanja_slike, sif, boj, izabrana_sezona))
                                    conn.commit()
                                    conn.close()
                                    st.success("Izmenjeno!")
                                    st.rerun()
                                    
                            with col_b2:
                                if st.button("🗑️ Obriši", key=f"Obr_{kljuc_id}"):
                                    conn = uzmi_vezu_sa_bazom()
                                    cursor = conn.cursor()
                                    cursor.execute("DELETE FROM artikli WHERE sifra = %s AND boja = %s AND sezona = %s", (sif, boj, izabrana_sezona))
                                    conn.commit()
                                    conn.close()
                                    # Sliku ostavljamo na Cloudinary-ju radi istorije ili je brišemo ručno tamo po potrebi
                                    st.warning("Obrisano!")
                                    st.rerun()
                st.markdown("---")

# --- OPCIJA 3: EVIDENCIJA IZLAZA ---
elif meni == "Evidencija izlaza (Po danima)":
    st.header(f"📆 Dnevni izlaz robe - Sezona: {izabrana_sezona}")
    
    conn = uzmi_vezu_sa_bazom()
    cursor = conn.cursor()
    cursor.execute("SELECT DISTINCT sifra FROM artikli WHERE sezona = %s", (izabrana_sezona,))
    sve_sifre = [red[0] for red in cursor.fetchall()]
    conn.close()
    
    if not sve_sifre:
        st.info(f"Nema unete robe u sezoni {izabrana_sezona} da biste zabeležili izlaz.")
    else:
        col1, col2 = st.columns(2)
        with col1:
            izabrani_datum = st.date_input("Izaberi datum izlaza:", datetime.now(), key="datum_izlaza_main")
            izabrana_sifra = st.selectbox("Izaberi šifru modela:", sve_sifre, key="izlaz_sifra_select")
            
            conn = uzmi_vezu_sa_bazom()
            cursor = conn.cursor()
            cursor.execute("SELECT boja FROM artikli WHERE sifra = %s AND sezona = %s", (izabrana_sifra, izabrana_sezona))
            dostupne_boje = [red[0] for red in cursor.fetchall()]
            conn.close()
            izabrana_boja = st.selectbox("Izaberi boju modela:", dostupne_boje, key="izlaz_boja_select")
        
        with col2:
            trenutno_na_stanju = 0
            if izabrana_boja:
                conn = uzmi_vezu_sa_bazom()
                cursor = conn.cursor()
                cursor.execute("SELECT broj_pari FROM artikli WHERE sifra = %s AND boja = %s AND sezona = %s", (izabrana_sifra, izabrana_boja, izabrana_sezona))
                rezultat = cursor.fetchone()
                if rezultat:
                    trenutno_na_stanju = rezultat[0]
                conn.close()
            
            st.write("")
            st.info(f"💡 Trenutno stanje za **{izabrana_sifra}** (**{izabrana_boja}**) je: **{trenutno_na_stanju} pari**")
            
            dinamicki_kljuc = f"izlaz_pari_input_{st.session_state['reset_brojac']}"
            kolicina_izlaza = st.number_input(
                "Koliko pari izlazi iz magacina:", 
                min_value=1, max_value=max(1, trenutno_na_stanju), 
                step=1, value=None, key=dinamicki_kljuc
            )
        
        dugme_onemoguceno = kolicina_izlaza is None or kolicina_izlaza <= 0
        
        if st.button("Zapiši izlaz robe", type="primary", key="dugme_zapisi_izlaz", disabled=dugme_onemoguceno):
            if kolicina_izlaza is None or trenutno_na_stanju < kolicina_izlaza or trenutno_na_stanju == 0:
                st.error("Greška: Nemate dovoljno pari na stanju!")
            else:
                conn = uzmi_vezu_sa_bazom()
                cursor = conn.cursor()
                
                cursor.execute('''
                    INSERT INTO izlaz_robe (datum, sifra_artikla, boja_artikla, kolicina_izlaz)
                    VALUES (%s, %s, %s, %s)
                ''', (izabrani_datum.strftime("%Y-%m-%d"), izabrana_sifra, izabrana_boja, kolicina_izlaza))
                
                novo_stanje = trenutno_na_stanju - kolicina_izlaza
                cursor.execute('''
                    UPDATE artikli SET broj_pari = %s WHERE sifra = %s AND boja = %s AND sezona = %s
                ''', (novo_stanje, izabrana_sifra, izabrana_boja, izabrana_sezona))
                
                conn.commit()
                conn.close()
                st.session_state["reset_brojac"] += 1
                st.success(f"Uspešno proknjižen izlaz! Novo stanje je {novo_stanje} pari.")
                st.rerun()

        st.markdown("---")
        st.subheader(f"📋 Istorija dnevnih izlaza robe za sezonu: {izabrana_sezona}")
        
        conn = uzmi_vezu_sa_bazom()
        upit_istorija = '''
            SELECT 
                ir.datum AS "Datum", 
                ir.sifra_artikla AS "Šifra modela", 
                ir.boja_artikla AS "Boja", 
                ir.kolicina_izlaz AS "Izašlo (pari)" 
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
                file_name=f"izlazi_robe_{izabrana_sezona}_{od_str}_do_{do_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dugme_download_excel_izlazi"
            )
            st.dataframe(df_izlazi, use_container_width=True)
        else:
            st.write(f"Još uvek nema zabeleženih izlaza robe za sezonu {izabrana_sezona}.")