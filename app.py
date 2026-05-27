import streamlit as st
import sqlite3
import os
import pandas as pd
from datetime import datetime
import io

# --- PODEŠAVANJE BAZE PODATAKA ---
def kreiraj_bazu():
    conn = sqlite3.connect("magacin.db")
    cursor = conn.cursor()
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS artikli (
            sifra TEXT,
            boja TEXT,
            sezona TEXT,
            broj_pari INTEGER,
            pari_u_kutiji INTEGER,
            prodajna_cena REAL,
            internet_cena REAL,
            slika_putanja TEXT,
            PRIMARY KEY (sifra, boja)
        )
    ''')
    
    try:
        cursor.execute("ALTER TABLE artikli ADD COLUMN sezona TEXT DEFAULT 'Proleće-Leto'")
    except sqlite3.OperationalError:
        pass
        
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS izlaz_robe (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            datum TEXT,
            sifra_artikla TEXT,
            boja_artikla TEXT,
            kolicina_izlaz INTEGER,
            FOREIGN KEY (sifra_artikla, boja_artikla) REFERENCES artikli (sifra, boja)
        )
    ''')
    conn.commit()
    conn.close()

kreiraj_bazu()

if not os.path.exists("slike_modela"):
    os.makedirs("slike_modela")

# Pomoćna funkcija za Excel
def konvertuj_u_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Magacin')
    return output.getvalue()

# --- IZGLED I STILIZACIJA APLIKACIJE ---
st.set_page_config(page_title="Magacin", layout="wide")

# --- POPRAVLJENI CSS (Bez isecanja naslova) ---
st.markdown("""
    <style>
    /* Smanjivanje praznog prostora na vrhu stranice */
    .block-container {
        padding-top: 1.5rem !important;
        padding-bottom: 1rem !important;
    }
    
    /* Glavni naslov - normalna margina, smanjen padding */
    h1 {
        font-size: 1.8rem !important;
        padding-top: 0px !important;
        padding-bottom: 5px !important;
        margin-top: 0px !important;
    }
    
    /* Podnaslovi sekcija - normalna margina, smanjen padding */
    h2 {
        font-size: 1.35rem !important;
        padding-top: 0px !important;
        padding-bottom: 10px !important;
        margin-top: 0px !important;
    }
    
    /* Naziv artikla na stranici Stanje (Šifra modela i Boja) */
    h3 {
        font-size: 1.05rem !important;
        font-weight: bold !important;
    }
    
    /* Veličina tekstualnog prikaza i brojki unutar detalja artikla */
    [data-testid="stMetricValue"] {
        font-size: 1.05rem !important;
    }
    [data-testid="stMetricLabel"] {
        font-size: 0.75rem !important;
    }
    
    /* Smanjivanje običnog teksta i labela u formama */
    .stTextInput p, .stNumberInput p, .stSelectbox p, .stDateInput p, label p {
        font-size: 0.85rem !important;
    }
    
    /* Smanjivanje teksta unutar plavih/zelenih info polja */
    .stAlert p {
        font-size: 0.85rem !important;
    }
    </style>
""", unsafe_allow_html=True)

# Glavni naslov
st.title("📦 Sistem za praćenje stanja u magacinu")

# 1. SEZONA
izabrana_sezona = st.sidebar.radio("🌸 IZABERI SEZONU:", ["Proleće-Leto", "Jesen-Zima"])
st.sidebar.markdown("---")

# 2. FUNKCIJA
meni = st.sidebar.selectbox("Izaberi opciju:", ["Trenutno stanje", "Unos nove robe", "Evidencija izlaza (Po danima)"])
st.sidebar.info(f"Trenutno radite u sekciji:\n**{izabrana_sezona}**")

# --- OPCIJA 1: UNOS NOVE ROBE ---
if meni == "Unos nove robe":
    st.header(f"➕ Unos novog artikla ({izabrana_sezona})")
    
    with st.form("forma_za_unos", clear_on_submit=True):
        col1, col2 = st.columns(2)
        
        with col1:
            sifra = st.text_input("Šifra modela:").strip().upper()
            boja = st.text_input("Boja modela:").strip().capitalize()
            broj_pari = st.number_input("Broj pari:", min_value=0, step=1)
            pari_u_kutiji = st.number_input("Broj pari u jednoj kutiji:", min_value=1, step=1)
            
        with col2:
            prodajna_cena = st.number_input("Prodajna cena (RSD):", min_value=0.0, step=50.0)
            internet_cena = st.number_input("Internet cena (RSD):", min_value=0.0, step=50.0)
            slika = st.file_uploader("Ubaci sliku modela:", type=["jpg", "jpeg", "png"])
            
        dugme_potvrdi = st.form_submit_button("Sačuvaj artikal u bazu")
        
        if dugme_potvrdi:
            if sifra == "" or boja == "":
                st.error("Greška: Šifra i boja ne smeju biti prazne!")
            else:
                putanja_slike = ""
                if slika is not None:
                    ekstenzija = slika.name.split(".")[-1]
                    putanja_slike = f"slike_modela/{sifra}_{boja}.{ekstenzija}"
                    with open(putanja_slike, "wb") as f:
                        f.write(slika.getbuffer())
                
                try:
                    conn = sqlite3.connect("magacin.db")
                    cursor = conn.cursor()
                    cursor.execute('''
                        INSERT INTO artikli (sifra, boja, sezona, broj_pari, pari_u_kutiji, prodajna_cena, internet_cena, slika_putanja)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (sifra, boja, izabrana_sezona, broj_pari, pari_u_kutiji, prodajna_cena, internet_cena, putanja_slike))
                    conn.commit()
                    conn.close()
                    st.success(f"Uspešno sačuvan model u sezoni {izabrana_sezona}: Šifra '{sifra}' - Boja '{boja}'!")
                    st.rerun()
                except sqlite3.IntegrityError:
                    st.error(f"Greška: Model sa šifrom '{sifra}' u boji '{boja}' već postoji u bazi!")

# --- OPCIJA 2: TRENUTNO STANJE ---
elif meni == "Trenutno stanje":
    st.header(f"📋 Stanje robe - Sezona: {izabrana_sezona}")
    
    conn = sqlite3.connect("magacin.db")
    df = pd.read_sql_query("SELECT * FROM artikli WHERE sezona = ?", conn, params=(izabrana_sezona,))
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
                
                br_kutija = row["broj_pari"] // row["pari_u_kutiji"]
                ost_pari = row["broj_pari"] % row["pari_u_kutiji"]
                
                with st.container():
                    col_slika, col_detalji, col_akcije = st.columns([1, 3, 1.5])
                    
                    with col_slika:
                        putanja = row["slika_putanja"]
                        if putanja and os.path.exists(putanja):
                            st.image(putanja, width=120)
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
                            
                            col_b1, col_b2 = st.columns(2)
                            with col_b1:
                                if st.button("💾 Snimi", key=f"Snimi_{kljuc_id}"):
                                    conn = sqlite3.connect("magacin.db")
                                    cursor = conn.cursor()
                                    cursor.execute('''
                                        UPDATE artikli 
                                        SET broj_pari = ?, prodajna_cena = ?, internet_cena = ?
                                        WHERE sifra = ? AND boja = ? AND sezona = ?
                                    ''', (nova_kol, nova_p_cena, nova_i_cena, sif, boj, izabrana_sezona))
                                    conn.commit()
                                    conn.close()
                                    st.success("Izmenjeno!")
                                    st.rerun()
                                    
                            with col_b2:
                                if st.button("🗑️ Obriši", key=f"Obr_{kljuc_id}"):
                                    conn = sqlite3.connect("magacin.db")
                                    cursor = conn.cursor()
                                    cursor.execute("DELETE FROM artikli WHERE sifra = ? AND boja = ? AND sezona = ?", (sif, boj, izabrana_sezona))
                                    conn.commit()
                                    conn.close()
                                    if putanja and os.path.exists(putanja):
                                        os.remove(putanja)
                                    st.warning("Obrisano!")
                                    st.rerun()
                st.markdown("---")

# --- OPCIJA 3: EVIDENCIJA IZLAZA ---
elif meni == "Evidencija izlaza (Po danima)":
    st.header(f"📆 Dnevni izlaz robe - Sezona: {izabrana_sezona}")
    
    conn = sqlite3.connect("magacin.db")
    cursor = conn.cursor()
    cursor.execute("SELECT DISTINCT sifra FROM artikli WHERE sezona = ?", (izabrana_sezona,))
    sve_sifre = [red[0] for red in cursor.fetchall()]
    conn.close()
    
    if not sve_sifre:
        st.info(f"Nema unete robe u sezoni {izabrana_sezona} da biste zabeležili izlaz.")
    else:
        col1, col2 = st.columns(2)
        
        with col1:
            izabrani_datum = st.date_input("Izaberi datum izlaza:", datetime.now(), key="datum_izlaza_main")
            izabrana_sifra = st.selectbox("Izaberi šifru modela:", sve_sifre, key="izlaz_sifra_select")
            
            conn = sqlite3.connect("magacin.db")
            cursor = conn.cursor()
            cursor.execute("SELECT boja FROM artikli WHERE sifra = ? AND sezona = ?", (izabrana_sifra, izabrana_sezona))
            dostupne_boje = [red[0] for red in cursor.fetchall()]
            conn.close()
            
            izabrana_boja = st.selectbox("Izaberi boju modela:", dostupne_boje, key="izlaz_boja_select")
        
        with col2:
            trenutno_na_stanju = 0
            if izabrana_boja:
                conn = sqlite3.connect("magacin.db")
                cursor = conn.cursor()
                cursor.execute("SELECT broj_pari FROM artikli WHERE sifra = ? AND boja = ? AND sezona = ?", (izabrana_sifra, izabrana_boja, izabrana_sezona))
                rezultat = cursor.fetchone()
                if rezultat:
                    trenutno_na_stanju = rezultat[0]
                conn.close()
            
            st.write("")
            st.info(f"💡 Trenutno stanje za **{izabrana_sifra}** (**{izabrana_boja}**) je: **{trenutno_na_stanju} pari**")
            
            kolicina_izlaza = st.number_input("Koliko pari izlazi iz magacina:", min_value=1, max_value=max(1, trenutno_na_stanju), step=1, key="izlaz_kolicina_input")
        
        if st.button("Zapiši izlaz robe", type="primary", key="dugme_zapisi_izlaz"):
            if trenutno_na_stanju < kolicina_izlaza or trenutno_na_stanju == 0:
                st.error("Greška: Nemate dovoljno pari na stanju ili boja nije pravilno izabrana!")
            else:
                conn = sqlite3.connect("magacin.db")
                cursor = conn.cursor()
                
                cursor.execute('''
                    INSERT INTO izlaz_robe (datum, sifra_artikla, boja_artikla, kolicina_izlaz)
                    VALUES (?, ?, ?, ?)
                ''', (izabrani_datum.strftime("%Y-%m-%d"), izabrana_sifra, izabrana_boja, kolicina_izlaza))
                
                novo_stanje = trenutno_na_stanju - kolicina_izlaza
                cursor.execute('''
                    UPDATE artikli SET broj_pari = ? WHERE sifra = ? AND boja = ? AND sezona = ?
                ''', (novo_stanje, izabrana_sifra, izabrana_boja, izabrana_sezona))
                
                conn.commit()
                conn.close()
                
                st.success(f"Uspešno proknjižen izlaz! Novo stanje je {novo_stanje} pari.")
                st.rerun()

        st.markdown("---")
        
        st.subheader("📋 Istorija dnevnih izlaza robe")
        
        conn = sqlite3.connect("magacin.db")
        df_izlazi = pd.read_sql_query("SELECT datum AS 'Datum', sifra_artikla AS 'Šifra modela', boja_artikla AS 'Boja', kolicina_izlaz AS 'Izašlo (pari)' FROM izlaz_robe ORDER BY id ASC", conn)
        conn.close()
        
        if not df_izlazi.empty:
            st.write("📅 **Izaberi period za preuzimanje Excel tabele:**")
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
                file_name=f"izlazi_robe_{od_str}_do_{do_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dugme_download_excel_izlazi"
            )
            
            st.write("Prikaz svih zabeleženih izlaza:")
            st.dataframe(df_izlazi, use_container_width=True)
        else:
            st.write("Još uvek nema zabeleženih izlaza robe.")