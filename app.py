import streamlit as st
import pandas as pd
import plotly.graph_objects as go
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
        
        # --- SANKEY DIAGRAM ---
        st.markdown("## 🌊 Customer Journey (Sankey)")
        
        # Sankey-Daten vorbereiten
        # Von "Termin vereinbart" zu allen möglichen ersten Folge-Kontakten
        sankey_data = []
        
        for _, row in results_df.iterrows():
            folge_liste = row['Folge-Kontakte'].split(' → ')
            
            # Erste Verbindung: Termin Art → Erster Folge-Kontakt
            if folge_liste[0] != 'Kein weiterer Kontakt':
                sankey_data.append({
                    'von': row['Termin Art'],
                    'zu': folge_liste[0],
                    'wert': 1
                })
                
                # Weitere Verbindungen in der Chain
                for i in range(len(folge_liste) - 1):
                    sankey_data.append({
                        'von': folge_liste[i],
                        'zu': folge_liste[i+1],
                        'wert': 1
                    })
            else:
                # Keine Folge-Kontakte
                sankey_data.append({
                    'von': row['Termin Art'],
                    'zu': 'Kein weiterer Kontakt',
                    'wert': 1
                })
        
        # Aggregieren
        sankey_df = pd.DataFrame(sankey_data)
        sankey_agg = sankey_df.groupby(['von', 'zu'])['wert'].sum().reset_index()
        
        # Nur die Top-Flows zeigen (sonst wird's zu komplex)
        top_flows = sankey_agg.nlargest(30, 'wert')
        
        # Nodes erstellen
        alle_nodes = pd.concat([top_flows['von'], top_flows['zu']]).unique().tolist()
        node_dict = {node: idx for idx, node in enumerate(alle_nodes)}
        
        # Sankey Figure
        fig = go.Figure(data=[go.Sankey(
            node=dict(
                pad=15,
                thickness=20,
                line=dict(color="black", width=0.5),
                label=alle_nodes,
                color=['#1f77b4' if 'Termin' in n else '#ff7f0e' if 'Auftrag' in n else '#2ca02c' if 'Verkaufsgespräch' in n else '#d62728' for n in alle_nodes]
            ),
            link=dict(
                source=[node_dict[x] for x in top_flows['von']],
                target=[node_dict[x] for x in top_flows['zu']],
                value=top_flows['wert'].tolist(),
                color='rgba(0,0,0,0.2)'
            )
        )])
        
        fig.update_layout(
            title="Customer Journey Flow (Top 30 Verbindungen)",
            font=dict(size=10),
            height=600
        )
        
        st.plotly_chart(fig, use_container_width=True)
        
        st.markdown("---")
        
        # --- DETAIL TABELLE ---
        st.markdown("## 📋 Detail-Tabelle")
        
        # Filter nach Mitarbeiterin
        mitarbeiterinnen = ['Alle'] + sorted(results_df['Mitarbeiterin'].unique().tolist())
        selected_ma = st.selectbox("Mitarbeiterin filtern", mitarbeiterinnen)
        
        if selected_ma != 'Alle':
            filtered_df = results_df[results_df['Mitarbeiterin'] == selected_ma]
        else:
            filtered_df = results_df
        
        # Sortierbar machen
        filtered_df_display = filtered_df.copy()
        filtered_df_display['Termin Datum'] = filtered_df_display['Termin Datum'].dt.strftime('%Y-%m-%d %H:%M')
        
        st.dataframe(
            filtered_df_display[['Mitarbeiterin', 'Kunde', 'Termin Datum', 'Termin Art', 'Folge-Kontakte', 'Letzter Status']],
            hide_index=True,
            use_container_width=True
        )
        
        # Download
        st.download_button(
            label="📥 Tabelle als CSV downloaden",
            data=filtered_df.to_csv(index=False).encode('utf-8'),
            file_name=f'termin_analyse_{datetime.now().strftime("%Y%m%d")}.csv',
            mime='text/csv'
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
