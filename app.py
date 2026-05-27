import streamlit as st
import sqlite3
import os
import pandas as pd
from datetime import datetime

# --- PODEŠAVANJE BAZE PODATAKA ---
def kreiraj_bazu():
    conn = sqlite3.connect("magacin.db")
    cursor = conn.cursor()
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS artikli (
            sifra TEXT,
            boja TEXT,
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

meni = st.sidebar.selectbox("Izaberi opciju:", ["Unos nove robe", "Trenutno stanje", "Evidencija izlaza (Po danima)"])

# --- OPCIJA 1: UNOS NOVE ROBE ---
if meni == "Unos nove robe":
    st.header("➕ Unos novog artikla u bazu")
    
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
                        INSERT INTO artikli (sifra, boja, broj_pari, pari_u_kutiji, prodajna_cena, internet_cena, slika_putanja)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                    ''', (sifra, boja, broj_pari, pari_u_kutiji, prodajna_cena, internet_cena, putanja_slike))
                    conn.commit()
                    conn.close()
                    st.success(f"Uspešno sačuvan model: Šifra '{sifra}' - Boja '{boja}'!")
                except sqlite3.IntegrityError:
                    st.error(f"Greška: Model sa šifrom '{sifra}' u boji '{boja}' već postoji u bazi!")

# --- OPCIJA 2: TRENUTNO STANJE (SA PRETRAGOM) ---
elif meni == "Trenutno stanje":
    st.header("📋 Trenutno stanje robe u magacinu")
    
    conn = sqlite3.connect("magacin.db")
    df = pd.read_sql_query("SELECT * FROM artikli", conn)
    conn.close()
    
    if df.empty:
        st.info("Magacin je trenutno prazan. Unesi robu u meniju sa leve strane.")
    else:
        # POLJE ZA PRETRAGU ŠIFRE
        pretraga = st.text_input("🔍 Pretraži magacin po šifri modela:", "").strip().upper()
        
        # Filtriranje tabele na osnovu unosa
        if pretraga:
            df = df[df["sifra"].str.contains(pretraga, na=False)]
        
        if df.empty:
            st.warning(f"Nema rezultata za šifru '{pretraga}'")
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
                        st.subheader(f"Modela: {row['Šifra modela']} | Boja: {row['Boja']}")
                        
                        c1, c2, c3, c4 = st.columns(4)
                        c1.metric("Ukupno pari na stanju", f"{row['Ukupno pari']} kom")
                        c2.metric("Pakovanje (Kutija / Ostatak)", f"{row['Broj kutija']} kut. + {row['Preostalo pari van kutije']} par")
                        c3.metric("Prodajna cena", f"{row['Prodajna cena (RSD)']} din")
                        c4.metric("Internet cena", f"{row['Internet cena (RSD)']} din")
                    st.markdown("---")

# --- OPCIJA 3: EVIDENCIJA IZLAZA (AUTOMATSKO RAČUNANJE) ---
elif meni == "Evidencija izlaza (Po danima)":
    st.header("📆 Dnevni izlaz robe")
    
    conn = sqlite3.connect("magacin.db")
    cursor = conn.cursor()
    # Izvlačimo sve unikatne šifre iz baze da ih ponudimo korisniku
    cursor.execute("SELECT DISTINCT sifra FROM artikli")
    sve_sifre = [red[0] for red in cursor.fetchall()]
    conn.close()
    
    if not sve_sifre:
        st.info("Morate prvo uneti robu u magacin da biste evidentirali izlaz.")
    else:
        # FORMA ZA UPIS IZLAZA
        with st.form("forma_izlaz"):
            col1, col2 = st.columns(2)
            
            with col1:
                izabrani_datum = st.date_input("Izaberi datum izlaza:", datetime.now())
                # Korisnik bira šifru sa liste
                izabrana_sifra = st.selectbox("Izaberi šifru modela:", sve_sifre)
                
                # Na osnovu šifre, tražimo koje boje imamo u bazi za taj model
                conn = sqlite3.connect("magacin.db")
                cursor = conn.cursor()
                cursor.execute("SELECT boja FROM artikli WHERE sifra = ?", (izabrana_sifra,))
                dostupne_boje = [red[0] for red in cursor.fetchall()]
                conn.close()
                
                izabrana_boja = st.selectbox("Izaberi boju modela:", dostupne_boje)
            
            with col2:
                # Proveravamo koliko tog modela i boje trenutno ima na stanju
                conn = sqlite3.connect("magacin.db")
                cursor = conn.cursor()
                cursor.execute("SELECT broj_pari FROM artikli WHERE sifra = ? AND boja = ?", (izabrana_sifra, izabrana_boja))
                trenutno_na_stanju = cursor.fetchone()[0]
                conn.close()
                
                st.write("") # Malo razmaka
                st.info(f"💡 Trenutno stanje za {izabrana_sifra} ({izabrana_boja}) je: **{trenutno_na_stanju} pari**")
                kolicina_izlaza = st.number_input("Koliko pari izlazi iz magacina:", min_value=1, max_value=max(1, trenutno_na_stanju), step=1, key="izlaz_kolicina_input")
            
            dugme_izlaz = st.form_submit_button("Zapiši izlaz robe")
            
            if dugme_izlaz:
                if trenutno_na_stanju < kolicina_izlaza:
                    st.error("Greška: Nemate dovoljno pari na stanju za ovaj izlaz!")
                else:
                    # Izvršavanje transakcije: upis u istoriju i smanjivanje stanja u artiklima
                    conn = sqlite3.connect("magacin.db")
                    cursor = conn.cursor()
                    
                    # 1. Upis u tabelu izlaza
                    cursor.execute('''
                        INSERT INTO izlaz_robe (datum, sifra_artikla, boja_artikla, kolicina_izlaz)
                        VALUES (?, ?, ?, ?)
                    ''', (izabrani_datum.strftime("%Y-%m-%d"), izabrana_sifra, izabrana_boja, kolicina_izlaza))
                    
                    # 2. Smanjivanje broja pari u tabeli artikala
                    novo_stanje = trenutno_na_stanju - kolicina_izlaza
                    cursor.execute('''
                        UPDATE artikli SET broj_pari = ? WHERE sifra = ? AND boja = ?
                    ''', (novo_stanje, izabrana_sifra, izabrana_boja))
                    
                    conn.commit()
                    conn.close()
                    
                    st.success(f"Uspešno proknjižen izlaz! Novo stanje za {izabrana_sifra} ({izabrana_boja}) je {novo_stanje} pari.")
                    st.rerun() # Osvežavamo stranicu da odmah povuče nove podatke

        # --- PRIKAZ ISTORIJE IZLAZA NA DNU STRANICE ---
        st.subheader("📋 Istorija dnevnih izlaza robe")
        conn = sqlite3.connect("magacin.db")
        df_izlazi = pd.read_sql_query("SELECT datum AS 'Datum', sifra_artikla AS 'Šifra modela', boja_artikla AS 'Boja', kolicina_izlaz AS 'Izašlo (pari)' FROM izlaz_robe ORDER BY id DESC", conn)
        conn.close()
        
        if not df_izlazi.empty:
            st.dataframe(df_izlazi, use_container_width=True)
        else:
            st.write("Još uvek nema zabeleženih izlaza robe.")