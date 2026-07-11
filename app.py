import streamlit as st
import psycopg2
from psycopg2.extras import RealDictCursor
import cloudinary
import cloudinary.uploader
import pandas as pd
from datetime import datetime
import io

# --- KONFIGURACIJA ---
cloudinary.config(
    cloud_name = st.secrets["cloudinary"]["cloud_name"],
    api_key = st.secrets["cloudinary"]["api_key"],
    api_secret = st.secrets["cloudinary"]["api_secret"],
    secure = True
)

@st.cache_resource
def uzmi_vezu_sa_bazom():
    return psycopg2.connect(st.secrets["postgres"]["url"])

@st.cache_data(ttl=60)
def ucitaj_artikle_za_sezonu(sezona):
    conn = uzmi_vezu_sa_bazom()
    return pd.read_sql_query("SELECT * FROM artikli WHERE sezona = %s ORDER BY sifra ASC, boja ASC", conn, params=(sezona,))

@st.cache_data(ttl=300)
def ucitaj_boje():
    conn = uzmi_vezu_sa_bazom()
    cursor = conn.cursor()
    cursor.execute("SELECT boja FROM sifrarnik_boja ORDER BY boja ASC")
    return [red[0] for red in cursor.fetchall()]

@st.cache_data(ttl=30)
def ucitaj_istoriju_izlaza_za_sezonu(sezona):
    conn = uzmi_vezu_sa_bazom()
    upit = '''
        SELECT ir.id AS "ID Zapisa", ir.datum AS "Datum", ir.sifra_artikla AS "Šifra modela", 
               ir.boja_artikla AS "Boja proizvoda", ir.grad AS "Grad", ir.kolicina_izlaz AS "Izašlo",
               ir.prodajna_cena AS "Prodajna cena po paru", (ir.kolicina_izlaz * ir.prodajna_cena) AS "Ukupno prodajna",
               ir.nabavna_cena AS "Nabavna cena po paru", (ir.kolicina_izlaz * ir.nabavna_cena) AS "Ukupno nabavna"
        FROM izlaz_robe ir INNER JOIN artikli a ON ir.sifra_artikla = a.sifra AND ir.boja_artikla = a.boja
        WHERE a.sezona = %s ORDER BY ir.id DESC
    '''
    return pd.read_sql_query(upit, conn, params=(sezona,))

def kreiraj_tabele():
    conn = uzmi_vezu_sa_bazom()
    cursor = conn.cursor()
    cursor.execute('''CREATE TABLE IF NOT EXISTS artikli (sifra TEXT, boja TEXT, sezona TEXT DEFAULT 'Proleće-Leto', broj_pari INTEGER, pari_u_kutiji INTEGER, prodajna_cena REAL, internet_cena REAL, slika_putanja TEXT, PRIMARY KEY (sifra, boja))''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS izlaz_robe (id SERIAL PRIMARY KEY, datum TEXT, sifra_artikla TEXT, boja_artikla TEXT, kolicina_izlaz INTEGER, grad TEXT, prodajna_cena REAL, nabavna_cena REAL)''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS sifrarnik_boja (boja TEXT PRIMARY KEY)''')
    cursor.execute("SELECT COUNT(*) FROM sifrarnik_boja")
    if cursor.fetchone()[0] == 0:
        cursor.executemany("INSERT INTO sifrarnik_boja (boja) VALUES (%s)", [("Black",), ("Blue",), ("Red",), ("Gray",), ("White",), ("Beige",)])
    conn.commit()

kreiraj_tabele()

def konvertuj_u_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Magacin')
    return output.getvalue()

def pronadji_sliku_u_df(df, sifra):
    if df.empty: return ""
    filtrirano = df[(df["sifra"] == sifra) & (df["slika_putanja"] != "") & (df["slika_putanja"].notna())]
    return filtrirano.iloc[0]["slika_putanja"] if not filtrirano.empty else ""

# --- DIZAJN ---
st.set_page_config(page_title="Magacin", layout="wide")
st.markdown("""
    <style>
    .block-container { padding-top: 3.5rem !important; padding-bottom: 2rem !important; }
    h1 { font-size: 1.8rem !important; padding-bottom: 10px !important; margin: 0px !important; }
    h2 { font-size: 1.35rem !important; padding-bottom: 15px !important; margin: 0px !important; }
    h3 { font-size: 1.05rem !important; font-weight: bold !important; }
    [data-testid="stMetricValue"] { font-size: 1.05rem !important; }
    [data-testid="stMetricLabel"] { font-size: 0.75rem !important; }
    .stTextInput p, .stNumberInput p, .stSelectbox p, .stDateInput p, label p { font-size: 0.85rem !important; }
    div[data-testid="stHorizontalBlock"] { background: var(--secondary-background-color); padding: 15px; border-radius: 8px; margin-bottom: 12px; border: 1px solid var(--border-color); }
    div.stButton > button { width: 100% !important; padding: 2px 10px !important; font-size: 0.9rem !important; height: 38px !important; }
    .indikator-stranice { background-color: var(--secondary-background-color); padding: 6px 12px; border-radius: 4px; font-weight: bold; color: #26a69a; font-size: 0.95rem; border: 1px solid var(--border-color); border-left: 4px solid #26a69a; margin-bottom: 15px; display: inline-block; }
    </style>
""", unsafe_allow_html=True)

st.title("📦 Višekorisnički sistem za praćenje stanja u magacinu")
izabrana_sezona = st.sidebar.radio("🌸 IZABERI KATEGORIJU / SEZONU:", ["Proleće-Leto", "Jesen-Zima", "Torbe"])
st.sidebar.markdown("---")
meni = st.sidebar.selectbox("Izaberi opciju:", ["Trenutno stanje", "Unos nove robe", "Evidencija izlaza (Po danima)", "Korekcija stanja zaliha"])
st.sidebar.info(f"Trenutno radite u sekciji:\n**{izabrana_sezona}**")

if "trenutna_stranica" not in st.session_state: st.session_state["trenutna_stranica"] = 1
if "prethodna_sezona" not in st.session_state: st.session_state["prethodna_sezona"] = izabrana_sezona
if "prethodni_meni" not in st.session_state: st.session_state["prethodni_meni"] = meni
if "reset_brojac" not in st.session_state: st.session_state["reset_brojac"] = 0

if izabrana_sezona != st.session_state["prethodna_sezona"] or meni != st.session_state["prethodni_meni"]:
    st.session_state["trenutna_stranica"] = 1
    st.session_state["prethodna_sezona"] = izabrana_sezona
    st.session_state["prethodni_meni"] = meni

# --- 1. UNOS NOVE ROBE ---
if meni == "Unos nove robe":
    st.header(f"➕ Unos novog artikla ({izabrana_sezona})")
    lista_boja = ucitaj_boje()
    
    sifra = st.text_input("Šifra modela:").strip().upper()
    boja = st.selectbox("Boja modela:", lista_boja)
    
    labela_kol = "Količina (komada/pari):" if izabrana_sezona == "Torbe" else "Količina pari:"
    labela_kut = "Broj komada u jednoj kutiji/pakovanju:" if izabrana_sezona == "Torbe" else "Broj pari u jednoj kutiji:"
    
    col1, col2 = st.columns(2)
    with col1:
        broj_pari = st.number_input(labela_kol, min_value=0, step=1, value=None, key=f"u_pari_{st.session_state['reset_brojac']}")
        pari_u_kutiji = st.number_input(labela_kut, min_value=1, step=1, value=None, key=f"u_kut_{st.session_state['reset_brojac']}")
    with col2:
        prodajna_cena = st.number_input("Prodajna cena (RSD):", min_value=0.0, step=50.0, value=None, key=f"u_pc_{st.session_state['reset_brojac']}")
        internet_cena = st.number_input("Internet cena (RSD):", min_value=0.0, step=50.0, value=None, key=f"u_ic_{st.session_state['reset_brojac']}")
        slika = st.file_uploader("Ubaci sliku modela:", type=["jpg", "jpeg", "png"], key=f"u_img_{st.session_state['reset_brojac']}")
        
    podaci_nedostaju = (sifra == "" or boja is None or broj_pari is None or pari_u_kutiji is None or prodajna_cena is None or internet_cena is None)
    
    if st.button("Sačuvaj artikal u bazu", type="primary", disabled=podaci_nedostaju):
        df_trenutni = ucitaj_artikle_za_sezonu(izabrana_sezona)
        url_slike = ""
        if slika is not None:
            with st.spinner("Slanje slike..."):
                try:
                    rezultat_slike = cloudinary.uploader.upload(slika, folder="magacin/", public_id=f"{sifra}_{boja}", transformation=[{"width": 800, "crop": "limit"}, {"quality": "auto", "fetch_format": "auto"}])
                    url_slike = rezultat_slike["secure_url"]
                except Exception as e: st.error(f"Greška slanja slike: {e}")
        else:
            url_slike = pronadji_sliku_u_df(df_trenutni, sifra)
            
        try:
            conn = uzmi_vezu_sa_bazom()
            cursor = conn.cursor()
            cursor.execute("INSERT INTO artikli (sifra, boja, sezona, broj_pari, pari_u_kutiji, prodajna_cena, internet_cena, slika_putanja) VALUES (%s, %s, %s, %s, %s, %s, %s, %s)", (sifra, boja, izabrana_sezona, broj_pari, pari_u_kutiji, prodajna_cena, internet_cena, url_slike))
            conn.commit()
            st.cache_data.clear()
            st.session_state["reset_brojac"] += 1
            st.success("Uspešno sačuvano!")
            st.rerun()
        except psycopg2.IntegrityError: st.error("Model u ovoj boji već postoji!")

    st.markdown("---")
    st.subheader("🎨 Upravljanje listom boja")
    c_boja, c_dugme = st.columns([3, 1])
    with c_boja: nova_boja = st.text_input("Naziv nove boje:").strip().capitalize()
    with c_dugme:
        st.write(""); st.write("")
        if st.button("➕ Dodaj boju") and nova_boja:
            try:
                conn = uzmi_vezu_sa_bazom()
                cursor = conn.cursor()
                cursor.execute("INSERT INTO sifrarnik_boja (boja) VALUES (%s)", (nova_boja,))
                conn.commit()
                st.cache_data.clear()
                st.success("Boja dodata!")
                st.rerun()
            except psycopg2.IntegrityError: st.warning("Boja već postoji.")

# --- 2. TRENUTNO STANJE ---
elif meni == "Trenutno stanje":
    st.header(f"📋 Stanje robe - Sekcija: {izabrana_sezona}")
    lista_boja = ucitaj_boje()
    df = ucitaj_artikle_za_sezonu(izabrana_sezona)
    
    if df.empty: st.info("Trenutno nema unete robe.")
    else:
        df_excel = df.copy()
        df_excel["Broj kutija"] = df_excel["broj_pari"] // df_excel["pari_u_kutiji"]
        df_excel["Ostatak"] = df_excel["broj_pari"] % df_excel["pari_u_kutiji"]
        lbl_kom = "Ukupno komada" if izabrana_sezona == "Torbe" else "Ukupno pari"
        lbl_kut = "Kutija/Pakovanja" if izabrana_sezona == "Torbe" else "Pari u kutiji"
        
        df_excel = df_excel.rename(columns={"sifra": "Šifra modela", "boja": "Boja proizvoda", "sezona": "Kategorija", "broj_pari": lbl_kom, "pari_u_kutiji": lbl_kut, "prodajna_cena": "Prodajna cena (RSD)", "internet_cena": "Internet cena (RSD)"}).drop(columns=["slika_putanja"], errors="ignore")
        st.download_button("🟢 Preuzmi kao Excel (.xlsx)", data=konvertuj_u_excel(df_excel), file_name=f"stanje_{izabrana_sezona}.xlsx")
        
        pretraga = st.text_input("🔍 Pretraži po šifri modela:", "").strip().upper()
        df_prikaz = df[df["sifra"].str.contains(pretraga, na=False)] if pretraga else df.copy()
        
        if df_prikaz.empty: st.warning("Nema rezultata.")
        else:
            PAGINACIJA = 20
            ukupno = len(df_prikaz)
            stranice = (ukupno // PAGINACIJA) + (1 if ukupno % PAGINACIJA > 0 else 0)
            if pretraga or st.session_state["trenutna_stranica"] > stranice: st.session_state["trenutna_stranica"] = 1
            
            if stranice > 1:
                st.markdown(f'<div class="indikator-stranice">📄 Stranica: {st.session_state["trenutna_stranica"]} od {stranice}</div>', unsafe_allow_html=True)
                st.session_state["trenutna_stranica"] = st.selectbox("Idi na stranicu:", list(range(1, stranice + 1)), index=list(range(1, stranice + 1)).index(st.session_state["trenutna_stranica"]))
                
            prikaz_ok = df_prikaz.iloc[(st.session_state["trenutna_stranica"]-1)*PAGINACIJA : st.session_state["trenutna_stranica"]*PAGINACIJA]
            
            for idx, red in prikaz_ok.iterrows():
                sif, boj, kljuc_id = red['sifra'], red['boja'], f"{red['sifra']}_{red['boja']}"
                img_url = red["slika_putanja"] if red["slika_putanja"] else pronadji_sliku_u_df(df, sif)
                br_kut, ost_par = red["broj_pari"] // red["pari_u_kutiji"], red["broj_pari"] % red["pari_u_kutiji"]
                
                with st.container():
                    c_img, c_det, c_act = st.columns([1.2, 3, 1.5])
                    with c_img:
                        if img_url:
                            st.image(img_url.replace("/upload/", "/upload/w_150,c_limit,q_auto,f_auto/"), width=120)
                            with st.expander("🔍 Veća slika"): st.image(img_url, use_container_width=True)
                        else: st.write("❌ Nema slike")
                    with c_det:
                        st.subheader(f"Šifra: {sif} | Boja: {boj}")
                        cx1, cx2, cx3, cx4 = st.columns(4)
                        lbl_m_kol = "Ukupno komada" if izabrana_sezona == "Torbe" else "Ukupno pari"
                        lbl_m_pak = f"{br_kut} pak. + {ost_par} kom" if izabrana_sezona == "Torbe" else f"{br_kut} kut. + {ost_par} par"
                        cx1.metric(lbl_m_kol, f"{red['broj_pari']} kom")
                        cx2.metric("Pakovanje", lbl_m_pak)
                        cx3.metric("Prodajna", f"{int(red['prodajna_cena'])} din")
                        cx4.metric("Internet", f"{int(red['internet_cena'])} din")
                    with c_act:
                        with st.expander("🛠️ Uredi"):
                            n_sifra = st.text_input("Šifra:", value=sif, key=f"s_{kljuc_id}").strip().upper()
                            n_boja = st.selectbox("Boja:", lista_boja, index=lista_boja.index(boj) if boj in lista_boja else 0, key=f"b_{kljuc_id}")
                            n_kol = st.number_input("Količina:", min_value=0, value=int(red['broj_pari']), key=f"k_{kljuc_id}")
                            n_pc = st.number_input("Prodajna:", value=float(red['prodajna_cena']), key=f"pc_{kljuc_id}")
                            n_ic = st.number_input("Internet:", value=float(red['internet_cena']), key=f"ic_{kljuc_id}")
                            n_img = st.file_uploader("Zameni sliku:", type=["jpg", "jpeg", "png"], key=f"img_{kljuc_id}")
                            
                            cb1, cb2 = st.columns(2)
                            with cb1:
                                if st.button("💾 Snimi", key=f"sv_{kljuc_id}") and n_sifra:
                                    final_img = red["slika_putanja"]
                                    if n_img:
                                        try:
                                            rez = cloudinary.uploader.upload(n_img, folder="magacin/", public_id=f"{n_sifra}_{n_boja}", transformation=[{"width": 800, "crop": "limit"}, {"quality": "auto", "fetch_format": "auto"}])
                                            final_img = rez["secure_url"]
                                        except: pass
                                    try:
                                        conn = uzmi_vezu_sa_bazom()
                                        cursor = conn.cursor()
                                        if n_sifra != sif:
                                            cursor.execute('UPDATE artikli SET sifra = %s WHERE sifra = %s AND sezona = %s', (n_sifra, sif, izabrana_sezona))
                                            cursor.execute('UPDATE izlaz_robe SET sifra_artikla = %s WHERE sifra_artikla = %s', (n_sifra, sif))
                                        cursor.execute('UPDATE artikli SET prodajna_cena = %s, internet_cena = %s WHERE sifra = %s AND sezona = %s', (n_pc, n_ic, n_sifra, izabrana_sezona))
                                        cursor.execute('UPDATE artikli SET boja = %s, broj_pari = %s, slika_putanja = %s WHERE sifra = %s AND boja = %s AND sezona = %s', (n_boja, n_kol, final_img, n_sifra, boj, izabrana_sezona))
                                        conn.commit()
                                        st.cache_data.clear()
                                        st.success("Sačuvano!")
                                        st.rerun()
                                    except psycopg2.IntegrityError: st.error("Već postoji!")
                            with cb2:
                                if st.button("🗑️ Obriši", key=f"del_{kljuc_id}"):
                                    conn = uzmi_vezu_sa_bazom()
                                    cursor = conn.cursor()
                                    cursor.execute("DELETE FROM artikli WHERE sifra = %s AND boja = %s AND sezona = %s", (sif, boj, izabrana_sezona))
                                    conn.commit()
                                    st.cache_data.clear()
                                    st.rerun()
                st.markdown("---")

# --- 3. EVIDENCIJA IZLAZA ---
elif meni == "Evidencija izlaza (Po danima)":
    st.header(f"📆 Dnevni izlaz robe - Sekcija: {izabrana_sezona}")
    df_artikli = ucitaj_artikle_za_sezonu(izabrana_sezona)
    sve_sifre = sorted(df_artikli["sifra"].unique().tolist()) if not df_artikli.empty else []
    
    if not sve_sifre: st.info("Trenutno nema robe na stanju.")
    else:
        gradovi = ["Internet", "Mladenovac Gore", "Mladenovac Dole", "Smederevska Palanka", "Zaječar", "Subotica", "Aleksinac", "Loznica", "Sremska Mitrovica", "Pančevo", "Vršac", "Bečej", "Prokuplje"]
        col1, col2 = st.columns(2)
        with col1:
            izabrani_datum = st.date_input("Izaberi datum izlaza:", datetime.now())
            izabrana_sifra = st.selectbox("Izaberi šifru modela:", sve_sifre, key=f"izl_sif_{izabrana_sezona}")
            boje_za_sifru = sorted(df_artikli[df_artikli["sifra"] == izabrana_sifra]["boja"].unique().tolist())
            izabrana_boja = st.selectbox("Izaberi boju modela:", boje_za_sifru, key=f"izl_boj_{izabrana_sezona}")
            izabrani_grad = st.selectbox("Izaberi grad:", gradovi)
            
        artikal = df_artikli[(df_artikli["sifra"] == izabrana_sifra) & (df_artikli["boja"] == izabrana_boja)]
        f_cena = float(artikal.iloc[0]["prodajna_cena"]) if not artikal.empty else 0.0
        zaliha = int(artikal.iloc[0]["broj_pari"]) if not artikal.empty else 0
        
        with col2:
            kolicina = st.number_input("Količina za izlaz:", min_value=1, step=1, value=None)
            p_cena_p = st.number_input("Prodajna cena (RSD):", min_value=0.0, step=50.0, value=f_cena)
            n_cena_p = st.number_input("Nabavna cena (Opciono):", min_value=0.0, step=50.0, value=None)
            
            lbl_suf = "komada" if izabrana_sezona == "Torbe" else "pari"
            st.info(f"📊 Trenutno stanje za **{izabrana_sifra} ({izabrana_boja})**: **{zaliha} {lbl_suf}**")
            
        btn_dis = (p_cena_par is None or p_cena_p <= 0.0 or kolicina is None)
        if st.button("Zapiši izlaz robe", type="primary", disabled=btn_dis):
            if zaliha < kolicina: st.error(f"Nema dovoljno robe! Na stanju: {zaliha}")
            else:
                conn = uzmi_vezu_sa_bazom()
                cursor = conn.cursor()
                cursor.execute("INSERT INTO izlaz_robe (datum, sifra_artikla, boja_artikla, kolicina_izlaz, grad, prodajna_cena, nabavna_cena) VALUES (%s, %s, %s, %s, %s, %s, %s)", (izabrani_datum.strftime("%Y-%m-%d"), izabrana_sifra, izabrana_boja, kolicina, izabrani_grad, p_cena_p, n_cena_p))
                cursor.execute('UPDATE artikli SET broj_pari = broj_pari - %s WHERE sifra = %s AND boja = %s AND sezona = %s', (kolicina, izabrana_sifra, izabrana_boja, izabrana_sezona))
                conn.commit()
                st.cache_data.clear()
                st.success("Izlaz proknjižen!")
                st.rerun()
                
    st.markdown("---")
    st.subheader(f"📋 Istorija izlaza robe za sekciju: {izabrana_sezona}")
    try:
        df_izlazi = ucitaj_istoriju_izlaza_za_sezonu(izabrana_sezona)
        if not df_izlazi.empty:
            cf1, cf2, cf3 = st.columns(3)
            s_dat = datetime.now()
            try: s_dat = datetime.strptime(df_izlazi['Datum'].min(), "%Y-%m-%d")
            except: pass
            with cf1: od_d = st.date_input("Od datuma:", s_dat)
            with cf2: do_d = st.date_input("Do datuma:", datetime.now())
            with cf3: f_grad = st.selectbox("Izaberi grad za tabelu:", ["SVI GRADOVI"] + gradovi)
            
            df_f = df_izlazi[(df_izlazi['Datum'] >= od_d.strftime("%Y-%m-%d")) & (df_izlazi['Datum'] <= do_d.strftime("%Y-%m-%d"))]
            if f_grad != "SVI GRADOVI": df_f = df_f[df_f['Grad'] == f_grad]
            
            st.download_button("🟢 Preuzmi istoriju kao Excel", data=konvertuj_u_excel(df_f), file_name="izlazi.xlsx")
            if not df_f.empty: st.dataframe(df_f, use_container_width=True)
            else: st.info("Nema zapisa za izabrani filter.")
        else: st.write("Nema zabeleženih izlaza.")
    except Exception as e: st.warning(f"Istorijski prikaz trenutno nedostupan. ({e})")
    
    st.markdown("---")
    st.write("### 🚨 Storniranje (Brisanje) zapisa")
    try:
        conn = uzmi_vezu_sa_bazom()
        cursor = conn.cursor(cursor_factory=RealDictCursor)
        cursor.execute("SELECT ir.id, ir.datum, ir.sifra_artikla, ir.boja_artikla, ir.grad, ir.kolicina_izlaz FROM izlaz_robe ir LEFT JOIN artikli a ON ir.sifra_artikla = a.sifra AND ir.boja_artikla = a.boja WHERE a.sezona = %s OR a.sezona IS NULL ORDER BY ir.id DESC LIMIT 50", (izabrana_sezona,))
        sirovi = cursor.fetchall()
        
        if not sirovi: st.info("Nema zapisa za storniranje.")
        else:
            opcije, mapa = [], {}
            for r in sirovi:
                txt = f"ID: {r['id']} | Model: {r['sifra_artikla']} | Boja: {r['boja_artikla']} | Količina: {r['kolicina_izlaz']} kom | Datum: {r['datum']}"
                opcije.append(txt)
                mapa[txt] = r
            izbor_storno = st.selectbox("Izaberi zapis za brisanje:", ["--- Izaberi zapis ---"] + opcije)
            if izbor_storno != "--- Izaberi zapis ---":
                z = mapa[izbor_storno]
                st.error(f"Upozorenje: Brisanjem ID {z['id']} vraćate {z['kolicina_izlaz']} kom na stanje modela {z['sifra_artikla']}.")
                if st.button("❌ POTVRDI BRISANJE"):
                    conn = uzmi_vezu_sa_bazom()
                    cursor = conn.cursor()
                    cursor.execute("DELETE FROM izlaz_robe WHERE id = %s", (int(z["id"]),))
                    cursor.execute("SELECT * FROM artikli WHERE sifra = %s AND boja = %s AND sezona = %s", (z["sifra_artikla"], z["boja_artikla"], izabrana_sezona))
                    if cursor.fetchone() is not None:
                        cursor.execute('UPDATE artikli SET broj_pari = broj_pari + %s WHERE sifra = %s AND boja = %s AND sezona = %s', (int(z["kolicina_izlaz"]), z["sifra_artikla"], z["boja_artikla"], izabrana_sezona))
                    conn.commit()
                    st.cache_data.clear()
                    st.success("Obrisano!")
                    st.rerun()
    except Exception as ex: st.error(f"Greška storniranja: {ex}")

# --- 4. KOREKCIJA STANJA ZALIHA ---
elif meni == "Korekcija stanja zaliha":
    st.header(f"🔧 Direktna korekcija stanja artikala ({izabrana_sezona})")
    df_artikli = ucitaj_artikle_za_sezonu(izabrana_sezona)
    sve_sifre_k = sorted(df_artikli["sifra"].unique().tolist()) if not df_artikli.empty else []
    lista_boja = ucitaj_boje()
    
    if not sve_sifre_k: st.warning("Nema artikala čije stanje možete menjati.")
    else:
        with st.form("forma_korekcija"):
            iz_sif = st.selectbox("Izaberi šifru:", sve_sifre_k)
            iz_boj = st.selectbox("Izaberi boju:", lista_boja)
            t_red = df_artikli[(df_artikli["sifra"] == iz_sif) & (df_artikli["boja"] == iz_boj)]
            
            if not t_red.empty: st.info(f"Trenutno u bazi: **{int(t_red.iloc[0]['broj_pari'])} kom.**")
            else: st.caption("Ovaj model trenutno ne postoji u izabranoj boji.")
            
            operacija = st.radio("Operacija:", ["DODAJ (Povećaj)", "ODUZMI (Smanji)"])
            br_kor = st.number_input("Broj komada:", min_value=1, step=1, value=None)
            potvrda = st.form_submit_button("Izvrši brzu korekciju", type="primary")
            
        if potvrda:
            if br_kor is None: st.error("Unesite broj komada!")
            elif t_red.empty: st.error("Model/boja ne postoje na stanju.")
            else:
                staro = int(t_red.iloc[0]['broj_pari'])
                novo = (staro + br_kor) if operacija == "DODAJ (Povećaj)" else (staro - br_kor)
                
                if novo < 0: st.error("Nemoguća operacija! Zalihe ne mogu ići u minus.")
                else:
                    conn = uzmi_vezu_sa_bazom()
                    cursor = conn.cursor()
                    cursor.execute('UPDATE artikli SET broj_pari = %s WHERE sifra = %s AND boja = %s AND sezona = %s', (novo, iz_sif, iz_boj, izabrana_sezona))
                    conn.commit()
                    st.cache_data.clear()
                    st.success(f"✅ Uspešno izmenjeno! Novo stanje: {novo} kom.")
                    st.rerun()