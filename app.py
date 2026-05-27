import streamlit as st
import sqlite3
import os
import pandas as pd
from datetime import datetime

# --- PODEŠAVANJE BAZE PODATAKA ---
def kreiraj_bazu():
    conn = sqlite3.connect("magacin.db")
    cursor = conn.cursor()
    
    # Dodata kolona 'sezona' u tabelu artikli
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
    
    # Automatska migracija za postojeće baze (ako kolona sezona ne postoji, dodaj je)
    try:
        cursor.execute("ALTER TABLE artikli ADD COLUMN sezona TEXT DEFAULT 'Proleće-Leto'")
    except sqlite3.OperationalError:
        pass # Kolona već postoji, preskoči
        
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

# --- IZGLED APLIKACIJE ---
st.set_page_config(page_title="Magacin", layout="wide")
st.title("📦 Sistem za praćenje stanja u magacinu")

# 1. PRVO BIRAMO SEZONU NA SIDEBAR-U
izabrana_sezona = st.sidebar.radio("🌸 IZABERI SEZONU:", ["Proleće-Leto", "Jesen-Zima"])

st.sidebar.markdown("---")

# 2. ZATIM BIRAMO FUNKCIJU ZA TU SEZONU
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
            broj_pari = st.number_input("Trenutni broj pari na stanju:", min_value=0, step=1)
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
                except sqlite3.IntegrityError:
                    st.error(f"Greška: Model sa šifrom '{sifra}' u boji '{boja}' već postoji u bazi!")

# --- OPCIJA 2: TRENUTNO STANJE FILTRIRANO PO SEZONI ---
elif meni == "Trenutno stanje":
    st.header(f"📋 Stanje robe - Sezona: {izabrana_sezona}")
    
    conn = sqlite3.connect("magacin.db")
    # Čitamo iz baze samo artikle koji pripadaju izabranoj sezoni
    df = pd.read_sql_query("SELECT * FROM artikli WHERE sezona = ?", conn, params=(izabrana_sezona,))
    conn.close()
    
    if df.empty:
        st.info(f"U sezoni {izabrana_sezona} trenutno nema unete robe. Izaberi 'Unos nove robe' sa leve strane.")
    else:
        pretraga = st.text_input("🔍 Pretraži ovu sezonu po šifri modela:", "").strip().upper()
        
        if pretraga:
            df = df[df["sifra"].str.contains(pretraga, na=False)]
        
        if df.empty:
            st.warning(f"Nema rezultata za šifru '{pretraga}' u sezoni {izabrana_sezona}")
        else:
            df["Broj kutija"] = df["broj_pari"] // df["pari_u_kutiji"]
            df["Preostalo pari van kutije"] = df["broj_pari"] % df["pari_u_kutiji"]
            
            df_prikaz = df.rename(columns={
                "sifra": "Šifra modela",
                "boja": "Boja",
                "broj_pari": "Ukupno pari",
                "pari_u_kutiji": "Pari u kutiji",
                "prodajna_cena": "Prodajna cena (RSD)",
                "internet_cena": "Internet cena (RSD)"
            })
            
            for index, row in df_prikaz.iterrows():
                with st.container():
                    col_slika, col_detalji = st.columns([1, 4])
                    
                    with col_slika:
                        putanja = row["slika_putanja"]
                        if putanja and os.path.exists(putanja):
                            st.image(putanja, width=130)
                        else:
                            st.write("❌ Nema slike")
                            
                    with col_detalji:
                        st.subheader(f"Model: {row['Šifra modela']} | Boja: {row['Boja']}")
                        
                        c1, c2, c3, c4 = st.columns(4)
                        c1.metric("Ukupno pari na stanju", f"{row['Ukupno pari']} kom")
                        c2.metric("Pakovanje (Kutija / Ostatak)", f"{row['Broj kutija']} kut. + {row['Preostalo pari van kutije']} par")
                        c3.metric("Prodajna cena", f"{row['Prodajna cena (RSD)']} din")
                        c4.metric("Internet cena", f"{row['Internet cena (RSD)']} din")
                    st.markdown("---")

# --- OPCIJA 3: EVIDENCIJA IZLAZA FILTRIRANO PO SEZONI ---
elif meni == "Evidencija izlaza (Po danima)":
    st.header(f"📆 Dnevni izlaz robe - Sezona: {izabrana_sezona}")
    
    conn = sqlite3.connect("magacin.db")
    cursor = conn.cursor()
    # Nudimo samo šifre iz trenutno izabrane sezone
    cursor.execute("SELECT DISTINCT sifra FROM artikli WHERE sezona = ?", (izabrana_sezona,))
    sve_sifre = [red[0] for red in cursor.fetchall()]
    conn.close()
    
    if not sve_sifre:
        st.info(f"Nema unete robe u sezoni {izabrana_sezona} da biste zabeležili izlaz.")
    else:
        with st.form("forma_izlaz"):
            col1, col2 = st.columns(2)
            
            with col1:
                izabrani_datum = st.date_input("Izaberi datum izlaza:", datetime.now())
                izabrana_sifra = st.selectbox("Izaberi šifru modela:", sve_sifre)
                
                conn = sqlite3.connect("magacin.db")
                cursor = conn.cursor()
                cursor.execute("SELECT boja FROM artikli WHERE sifra = ? AND sezona = ?", (izabrana_sifra, izabrana_sezona))
                dostupne_boje = [red[0] for red in cursor.fetchall()]
                conn.close()
                
                izabrana_boja = st.selectbox("Izaberi boju modela:", dostupne_boje)
            
            with col2:
                conn = sqlite3.connect("magacin.db")
                cursor = conn.cursor()
                cursor.execute("SELECT broj_pari FROM artikli WHERE sifra = ? AND boja = ? AND sezona = ?", (izabrana_sifra, izabrana_boja, izabrana_sezona))
                trenutno_na_stanju = cursor.fetchone()[0]
                conn.close()
                
                st.write("")
                st.info(f"💡 Trenutno stanje za {izabrana_sifra} ({izabrana_boja}) je: **{trenutno_na_stanju} pari**")
                kolicina_izlaza = st.number_input("Koliko pari izlazi iz magacina:", min_value=1, max_value=max(1, trenutno_na_stanju), step=1, key="izlaz_kolicina_input")
            
            dugme_izlaz = st.form_submit_button("Zapiši izlaz robe")
            
            if dugme_izlaz:
                if trenutno_na_stanju < kolicina_izlaza:
                    st.error("Greška: Nemate dovoljno pari na stanju!")
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

        # Istorija prikazuje sve izlaze, ali radi lakšeg snalaženja
        st.subheader("📋 Istorija dnevnih izlaza robe (Sve sezone)")
        conn = sqlite3.connect("magacin.db")
        df_izlazi = pd.read_sql_query("SELECT datum AS 'Datum', sifra_artikla AS 'Šifra modela', boja_artikla AS 'Boja', kolicina_izlaz AS 'Izašlo (pari)' FROM izlaz_robe ORDER BY id DESC", conn)
        conn.close()
        
        if not df_izlazi.empty:
            st.dataframe(df_izlazi, use_container_width=True)
        else:
            st.write("Još uvek nema zabeleženih izlaza robe.")