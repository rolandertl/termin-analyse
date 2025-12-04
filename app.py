import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="EDELWEISS Termin-Analyse", layout="wide")

# Titel
st.title("📊 EDELWEISS Termin-Analyse")
st.markdown("---")

# File Upload
uploaded_file = st.file_uploader("Excel-Datei hochladen", type=['xlsx'])

if uploaded_file is not None:
    
    # Daten einlesen
    with st.spinner('Daten werden verarbeitet...'):
        df = pd.read_excel(uploaded_file)
        
        # Datum/Uhrzeit zu datetime konvertieren
        df['Datum/Uhrzeit'] = pd.to_datetime(df['Datum/Uhrzeit'])
        
        # Alle "Termin vereinbart" Varianten finden
        termin_mask = df['Kontaktart'].str.contains('Termin vereinbart', na=False, case=False)
        termine_df = df[termin_mask].copy()
        
        # Nach Kunde und Datum sortieren
        termine_df = termine_df.sort_values(['Kontakt', 'Datum/Uhrzeit'])
        
        # Nur ersten Termin pro Kunde behalten
        erste_termine = termine_df.groupby('Kontakt').first().reset_index()
        
        st.success(f"✅ {len(erste_termine)} Kunden mit erstem 'Termin vereinbart' gefunden")
        
        # Für jeden ersten Termin die Customer Journey bauen
        results = []
        
        for _, termin_row in erste_termine.iterrows():
            kunde = termin_row['Kontakt']
            mitarbeiterin = termin_row['Mitarbeiter']
            termin_datum = termin_row['Datum/Uhrzeit']
            termin_art = termin_row['Kontaktart']
            
            # Alle Kontakte zu diesem Kunden NACH dem Termin
            kunde_kontakte = df[
                (df['Kontakt'] == kunde) & 
                (df['Datum/Uhrzeit'] > termin_datum)
            ].sort_values('Datum/Uhrzeit')
            
            # Folge-Kontaktarten sammeln
            folge_kontakte = kunde_kontakte['Kontaktart'].tolist()
            
            # Wenn keine Folge-Kontakte, dann bleibt es bei "Termin vereinbart"
            if len(folge_kontakte) == 0:
                folge_kontakte = ['Kein weiterer Kontakt']
            
            results.append({
                'Mitarbeiterin': mitarbeiterin,
                'Kunde': kunde,
                'Termin Datum': termin_datum,
                'Termin Art': termin_art,
                'Folge-Kontakte': ' → '.join(folge_kontakte),
                'Anzahl Folge-Kontakte': len(folge_kontakte),
                'Letzter Status': folge_kontakte[-1]
            })
        
        results_df = pd.DataFrame(results)
        
        # --- STATISTIKEN ---
        st.markdown("## 📈 Übersicht")
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("Gesamt Termine", len(results_df))
        
        with col2:
            unique_ma = results_df['Mitarbeiterin'].nunique()
            st.metric("Mitarbeiterinnen", unique_ma)
        
        with col3:
            mit_folge = len(results_df[results_df['Letzter Status'] != 'Kein weiterer Kontakt'])
            st.metric("Mit Folge-Kontakt", mit_folge)
        
        with col4:
            auftraege = len(results_df[results_df['Letzter Status'].str.contains('Auftrag', na=False, case=False)])
            st.metric("Endeten mit Auftrag", auftraege)
        
        st.markdown("---")
        
        # --- TOP PERFORMER ---
        st.markdown("## 🏆 Top Performer")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("### Meiste Termine vereinbart")
            top_termine = results_df['Mitarbeiterin'].value_counts().head(10)
            st.dataframe(top_termine.reset_index().rename(columns={'index': 'Mitarbeiterin', 'Mitarbeiterin': 'Anzahl'}), hide_index=True)
        
        with col2:
            st.markdown("### Meiste Aufträge (nach Termin)")
            auftraege_df = results_df[results_df['Letzter Status'].str.contains('Auftrag', na=False, case=False)]
            if len(auftraege_df) > 0:
                top_auftraege = auftraege_df['Mitarbeiterin'].value_counts().head(10)
                st.dataframe(top_auftraege.reset_index().rename(columns={'index': 'Mitarbeiterin', 'Mitarbeiterin': 'Anzahl'}), hide_index=True)
            else:
                st.info("Keine Aufträge gefunden")
        
        st.markdown("---")
        
        # --- DETAIL TABELLEN PRO MITARBEITERIN ---
        st.markdown("## 📋 Detail-Tabellen pro Mitarbeiterin")
        
        # Download Button für alle Daten ganz oben
        all_data_csv = results_df.copy()
        all_data_csv['Termin Datum'] = all_data_csv['Termin Datum'].dt.strftime('%Y-%m-%d %H:%M')
        
        st.download_button(
            label="📥 Gesamte Tabelle als CSV downloaden",
            data=all_data_csv.to_csv(index=False).encode('utf-8'),
            file_name=f'termin_analyse_gesamt_{datetime.now().strftime("%Y%m%d")}.csv',
            mime='text/csv'
        )
        
        st.markdown("---")
        
        # Sortiert nach Anzahl Termine (Top Performer zuerst)
        mitarbeiterinnen_sorted = results_df['Mitarbeiterin'].value_counts().index.tolist()
        
        for mitarbeiterin in mitarbeiterinnen_sorted:
            ma_df = results_df[results_df['Mitarbeiterin'] == mitarbeiterin].copy()
            
            # Statistiken für diese Mitarbeiterin
            anzahl_termine = len(ma_df)
            mit_folge = len(ma_df[ma_df['Letzter Status'] != 'Kein weiterer Kontakt'])
            auftraege = len(ma_df[ma_df['Letzter Status'].str.contains('Auftrag', na=False, case=False)])
            
            # Expander für jede Mitarbeiterin
            with st.expander(f"**{mitarbeiterin}** — {anzahl_termine} Termine | {mit_folge} mit Folge-Kontakt | {auftraege} mit Auftrag", expanded=False):
                
                # Datum formatieren
                ma_df_display = ma_df.copy()
                ma_df_display['Termin Datum'] = ma_df_display['Termin Datum'].dt.strftime('%Y-%m-%d %H:%M')
                
                # Tabelle
                st.dataframe(
                    ma_df_display[['Kunde', 'Termin Datum', 'Termin Art', 'Folge-Kontakte', 'Letzter Status']],
                    hide_index=True,
                    use_container_width=True
                )
                
                # Download für diese Mitarbeiterin
                st.download_button(
                    label=f"📥 {mitarbeiterin} als CSV downloaden",
                    data=ma_df_display.to_csv(index=False).encode('utf-8'),
                    file_name=f'termin_analyse_{mitarbeiterin}_{datetime.now().strftime("%Y%m%d")}.csv',
                    mime='text/csv',
                    key=f'download_{mitarbeiterin}'  # Unique key für jeden Button
                )

else:
    st.info("👆 Bitte Excel-Datei hochladen um zu starten")
    
    st.markdown("""
    ### Anleitung
    
    1. **Excel-Datei hochladen** mit folgenden Spalten:
       - `Mitarbeiter`
       - `Kontaktart`
       - `Kontakt` (Kundenname)
       - `Datum/Uhrzeit`
    
    2. Die App findet automatisch alle **"Termin vereinbart"** Einträge
    
    3. Auswertung zeigt:
       - Welche Mitarbeiterin den ersten Termin vereinbart hat
       - Was danach mit dem Kunden passiert ist
       - Customer Journey als Sankey-Diagramm
       - Detail-Tabelle mit allen Daten
    """)
