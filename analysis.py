import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import io
import msoffcrypto
import folium
import requests
import plotly.express as px
from branca.element import Template, MacroElement
import matplotlib as mpl
import matplotlib.patches as mpatches
import matplotlib.lines as mlines
from sklearn.preprocessing import StandardScaler, MinMaxScaler
from sklearn.cluster import KMeans
from sklearn.decomposition import PCA
from sklearn.metrics.pairwise import cosine_similarity

# --- CONFIGURATION VISUELLE GLOBALE ---
sns.set_theme(style="whitegrid")
mpl.rcParams['figure.dpi'] = 300
mpl.rcParams['savefig.dpi'] = 300
mpl.rcParams['figure.facecolor'] = 'none' 
mpl.rcParams['axes.facecolor'] = 'none'   
mpl.rcParams['axes.spines.top'] = False   
mpl.rcParams['axes.spines.right'] = False 
mpl.rcParams['axes.spines.left'] = False  
mpl.rcParams['text.color'] = '#2c3e50'    
mpl.rcParams['axes.labelcolor'] = '#7f8c8d'
mpl.rcParams['xtick.color'] = '#7f8c8d'
mpl.rcParams['ytick.color'] = '#7f8c8d'
# ══════════════════════════════════════════════════════════════
# CHARGEMENT & NETTOYAGE
# ══════════════════════════════════════════════════════════════

def load_and_clean(uploaded_file, password):
    """Déchiffre, charge et nettoie le fichier Excel."""
    import tempfile
    import io
    import msoffcrypto
    import pandas as pd
    import numpy as np

    # Déchiffrement
    temp_dec = io.BytesIO()
    excel_file = msoffcrypto.OfficeFile(uploaded_file)
    excel_file.load_key(password=password)
    excel_file.decrypt(temp_dec)
    
    temp_dec.seek(0)
    df_deals = pd.read_excel(temp_dec, sheet_name='Deals 2019-2026')
    temp_dec.seek(0)
    df_investisseurs = pd.read_excel(temp_dec, sheet_name='Investors 2019-2026')

    # ── Nettoyage df_deals ──
    # Suppression des lignes vides et doublons (corrige le +1)
    df_deals = df_deals.dropna(subset=['Start-up name', 'Deal Date'], how='all').drop_duplicates()
    
    df_deals['Year'] = pd.to_datetime(df_deals['Deal Date']).dt.year

    def clean_amount(val):
        if pd.isna(val): return np.nan
        val = str(val).strip().lower()
        # Correction du bug "0 disclosed" (ne pas utiliser '' dans le list comprehension)
        if val in ['', 'nan']: return np.nan
        if any(x in val for x in ['undisclosed', 'unknown', 'n/a']): return np.nan
        
        val = val.replace('$', '').replace('m', '').replace(',', '').strip()
        try:
            return float(val)
        except ValueError:
            return np.nan

    df_deals['Amount_numeric'] = df_deals['Amount raised $M'].apply(clean_amount)

    bracket_map = {
        'n.a':         np.nan,
        '$100K-$500K': 0.3,
        '$500K-$1M':   0.75,
        '$1M-$2M':     1.5,
        '$2M-$5M':     3.5,
        '$5M-$10M':    7.5,
        '$10M-$50M':   30.0,
        '$50M-$100M':  75.0,
        '$100M+':      150.0,
        'Unknown':     np.nan,
    }
    df_deals['Amount_clean'] = df_deals['Amount_numeric'].fillna(
        df_deals['Bracket'].map(bracket_map)
    )

    df_deals['Sector']  = df_deals['Sector'].astype(str).str.strip()
    df_deals['Country'] = df_deals['Country'].astype(str).str.strip()
    df_deals['Region']  = df_deals['Region'].astype(str).str.strip()

    def simplify_type(t):
        t = str(t).strip()
        if 'Series D' in t or 'Series E' in t or 'Series F' in t: return 'Series D+'
        if 'Series C' in t: return 'Series C'
        if 'Series B' in t: return 'Series B'
        if 'Series A' in t: return 'Series A'
        if 'Seed' in t or 'Pre-Seed' in t: return 'Seed'
        if 'Grant' in t: return 'Grants'
        if 'Debt' in t or 'Bond' in t: return 'Debt/Bond'
        return t

    df_deals['Type_Simple'] = df_deals['Type'].apply(simplify_type)

    def identify_yc(val):
        v = str(val).strip().upper()
        if v.startswith('YC') or 'Y COMBINATOR' in v: return 1
        return 0

    df_deals['is_yc']    = df_deals['Y Combinator'].apply(identify_yc)
    df_deals['YC_Label'] = df_deals['is_yc'].map({1: 'YC Alumni', 0: 'Non-YC'})
    df_deals['Founders_Num'] = (
        df_deals['# of Founders'].astype(str)
        .str.extract(r'(\d+)')[0].astype(float)
    )
    df_deals['is_disclosed'] = df_deals['Amount_numeric'].notna().astype(int)

    def categorize_market(country):
        big4 = ['Nigeria','Kenya','South Africa','Egypt']
        if country in big4:       return 'Big 4'
        elif country == 'Africa': return 'Pan-African'
        else:                     return 'Emerging Markets'

    df_deals['Market_Type'] = df_deals['Country'].apply(categorize_market)

    # ── Nettoyage df_investisseurs ──
    years_cols = [
        '2019 deals ($1M+)', '2020 deals ($500K+)', '2021 deals ($100K+)',
        '2022 deals ($100K+)', '2023 deals ($100K+)', '2024 deals ($100K+)',
        '2025 deals ($100K+)', '2026 deals ($100K+)'
    ]
    df_investisseurs[years_cols] = df_investisseurs[years_cols].fillna(0)

    # ── df_inv_final (exploded) ──
    df_inv_clean = df_deals.copy()
    df_inv_clean['Investors_List'] = df_inv_clean['Investors'].str.split(',')
    df_inv_clean = df_inv_clean.explode('Investors_List')
    df_inv_clean['Investors_List'] = df_inv_clean['Investors_List'].str.strip()
    df_inv_clean = df_inv_clean[
        df_inv_clean['Investors_List'].notna() &
        (df_inv_clean['Investors_List'] != '')
    ]
    
    df_inv_final = df_inv_clean.copy()
    df_inv_final['Investors_List'] = df_inv_final['Investors_List'].str.title().str.strip()
    black_list = ['N.A','Unknown','Nan','None','','Confidential',
                  'Individual Investors','Undisclosed Investors']
    df_inv_final = df_inv_final[~df_inv_final['Investors_List'].isin(black_list)]

    return df_deals, df_investisseurs, df_inv_final, years_cols


# ══════════════════════════════════════════════════════════════
# SECTION 1 — MARKET OVERVIEW
# ══════════════════════════════════════════════════════════════

def plot_deal_landscape(df_deals):
    fig, axes = plt.subplots(2, 3, figsize=(20, 12))

    sns.countplot(data=df_deals, y='Type',
                  order=df_deals['Type'].value_counts().index[:12],
                  palette='mako', ax=axes[0,0],
                  hue='Type', legend=False)
    axes[0,0].set_title('Deal Types')

    bracket_order = ['$100K-$500K','$500K-$1M','$1M-$2M','$2M-$5M',
                     '$5M-$10M','$10M-$50M','$50M-$100M','$100M+']
    sns.countplot(data=df_deals, x='Bracket',
                  order=[b for b in bracket_order if b in df_deals['Bracket'].values],
                  palette='rocket', ax=axes[0,1],
                  hue='Bracket', legend=False)
    axes[0,1].set_title('Funding Brackets')
    axes[0,1].tick_params(axis='x', rotation=45)

    sns.countplot(data=df_deals, y='Sector',
                  order=df_deals['Sector'].value_counts().index[:10],
                  palette='viridis', ax=axes[0,2],
                  hue='Sector', legend=False)
    axes[0,2].set_title('Top 10 Sectors')

    sns.countplot(data=df_deals, y='Country',
                  order=df_deals['Country'].value_counts().index[:10],
                  palette='Blues_r', ax=axes[1,0],
                  hue='Country', legend=False)
    axes[1,0].set_title('Top 10 Countries')

    sns.countplot(data=df_deals, x='Region',
                  order=df_deals['Region'].value_counts().index,
                  palette='Oranges_r', ax=axes[1,1],
                  hue='Region', legend=False)
    axes[1,1].set_title('Deals by Region')
    axes[1,1].tick_params(axis='x', rotation=45)

    axes[1,2].axis('off')
    stats_text = (
        f"Total deals : {len(df_deals):,}\n"
        f"Period : {df_deals['Year'].min()}–{df_deals['Year'].max()}\n"
        f"Countries : {df_deals['Country'].nunique()}\n"
        f"Sectors : {df_deals['Sector'].nunique()}\n"
        f"Total raised : ${df_deals['Amount_clean'].sum():,.0f}M\n"
        f"Deals with disclosed amount : {df_deals['Amount_numeric'].notna().sum():,}"
    )
    axes[1,2].text(0.1, 0.5, stats_text, transform=axes[1,2].transAxes,
                   fontsize=13, verticalalignment='center',
                   bbox=dict(boxstyle='round', facecolor='#f0f4ff', alpha=0.8))
    axes[1,2].set_title('Dataset Summary', fontsize=14)

    plt.suptitle('African Startup Funding — Deal Landscape', fontsize=18, fontweight='bold')
    plt.tight_layout()
    return fig


def plot_market_trajectory(df_deals):
    evolution = df_deals.groupby('Year').agg(
        Count=('Start-up name', 'count'),
        Total_USD_M=('Amount_clean', 'sum')
    )
    years  = list(evolution.index)
    x_pos  = range(len(years))

    fig, ax1 = plt.subplots(figsize=(14, 6))
    ax2 = ax1.twinx()

    ax1.bar(x_pos, evolution['Count'], color='lightgray', alpha=0.6, label='Number of deals')
    ax1.set_xticks(list(x_pos))
    ax1.set_xticklabels(years, rotation=45)
    ax1.set_ylabel('Number of Deals')

    ax2.plot(x_pos, evolution['Total_USD_M'], marker='o', color='#e74c3c',
             linewidth=3, label='Total Capital ($M)')
    ax2.set_ylabel('Total Capital ($M)')

    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2, loc='upper left')

    plt.title('African Market Trajectory (2019–2026)', fontsize=16)
    plt.tight_layout()
    return fig


def plot_heatmaps(df_deals):
    top_sectors = df_deals['Sector'].value_counts().index[:10]
    top_regions = df_deals['Region'].value_counts().index[:6]
    df_heat = df_deals[
        df_deals['Sector'].isin(top_sectors) &
        df_deals['Region'].isin(top_regions)
    ]
    pivot_total  = df_heat.pivot_table(values='Amount_clean', index='Sector',
                                       columns='Region', aggfunc='sum').fillna(0)
    pivot_median = df_heat.pivot_table(values='Amount_clean', index='Sector',
                                       columns='Region', aggfunc='median').fillna(0)

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(20, 8))
    sns.heatmap(pivot_total,  annot=True, fmt=".1f", cmap="YlOrRd", ax=ax1)
    ax1.set_title('Total Capital by Sector & Region ($M)')
    sns.heatmap(pivot_median, annot=True, fmt=".1f", cmap="YlGnBu", ax=ax2)
    ax2.set_title('Median Deal Size by Sector & Region ($M)')
    plt.tight_layout()
    return fig


def plot_transparency(df_deals):
    climate_trend = df_deals.groupby('Year')['Climate Tech'].apply(
        lambda x: x.astype(str).str.strip().str.lower().eq('climate tech').mean() * 100
    )
    climate_abs = df_deals.groupby('Year')['Climate Tech'].apply(
        lambda x: x.astype(str).str.strip().str.lower().eq('climate tech').sum()
    )
    disclosure_rate = df_deals.groupby('Year')['is_disclosed'].mean() * 100

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 6))

    ax1.plot(climate_trend.index, climate_trend.values, marker='o',
             color='#27ae60', linewidth=3)
    for yr, pct, n in zip(climate_trend.index, climate_trend.values, climate_abs.values):
        ax1.annotate(f'n={int(n)}', xy=(yr, pct), xytext=(0, 8),
                     textcoords='offset points', ha='center', fontsize=8)
    ax1.set_title('Climate Tech Share (%)')
    ax1.set_ylabel('%')
    ax1.grid(alpha=0.3)

    ax2.plot(disclosure_rate.index, disclosure_rate.values, marker='o',
             color='#3498db', linewidth=3)
    ax2.set_title('Amount Disclosure Rate (%)')
    ax2.set_ylabel('%')
    ax2.grid(alpha=0.3)

    plt.suptitle('Market Quality and Transparency (2019–2026)', fontsize=15)
    plt.tight_layout()
    return fig

import folium
import requests
from branca.element import Template, MacroElement

def create_interactive_map(df_deals, df_inv_final):
    top_sector_series = df_deals.groupby('Country')['Sector'].apply(
        lambda x: x.value_counts().index[0] if not x.empty else "N/A"
    )
    top_sector_dict = top_sector_series.to_dict()

    mask = ~df_inv_final['Investors_List'].str.contains('Angel|nan|Unknown', case=False, na=False)
    df_clean_inv = df_inv_final[mask]

    def get_clean_popup_html(country_df):
        top_5 = country_df['Investors_List'].value_counts().head(5)
        if top_5.empty:
            return "<p style='color:gray;'>No institutional data available</p>"
        
        html = "<ul style='margin: 0; padding: 0; list-style: none;'>"
        for i, (inv, count) in enumerate(top_5.items()):
            bg_color = "#f9f9f9" if i % 2 == 0 else "#ffffff"
            html += f"""
            <li style='background: {bg_color}; padding: 4px 8px; border-radius: 4px; margin-bottom: 2px; display: flex; justify-content: space-between; font-size: 0.9em;'>
                <span style='font-weight: 600; color: #2c3e50;'>{inv}</span>
                <span style='color: #27ae60;'>{int(count)}</span>
            </li>"""
        html += "</ul>"
        return html

    top_5_html_dict = df_clean_inv.groupby('Country').apply(get_clean_popup_html, include_groups=False).to_dict()

    stats_map = df_deals.groupby('Country').agg(
        Nb_Deals=('Amount_clean', 'count'),
        Total_Vol=('Amount_clean', 'sum')
    ).reset_index()

    url = 'https://raw.githubusercontent.com/python-visualization/folium/master/examples/data/world-countries.json'
    geo_data = requests.get(url).json()

    for feature in geo_data['features']:
        name = feature['properties']['name']
        row = stats_map[stats_map['Country'] == name]
        
        deals = int(row['Nb_Deals'].values[0]) if not row.empty else 0
        vol = f"{row['Total_Vol'].values[0]:.1f}" if not row.empty else "0"
        sector = top_sector_dict.get(name, "N/A")
        investors_list = top_5_html_dict.get(name, "No data")

        feature['properties']['popup_content'] = f"""
        <div style="font-family: 'Arial'; width: 280px; padding: 5px;">
            <h3 style="margin: 0 0 10px 0; color: #2c3e50; border-bottom: 2px solid #27ae60; padding-bottom: 5px;">{name}</h3>
            
            <div style="display: flex; justify-content: space-between; margin-bottom: 10px; background: #f4f7f6; padding: 8px; border-radius: 5px;">
                <div style="text-align: center; flex: 1;">
                    <span style="display: block; font-size: 0.7em; color: #7f8c8d; text-transform: uppercase;">Deals</span>
                    <span style="font-size: 1.1em; font-weight: bold; color: #2c3e50;">{deals}</span>
                </div>
                <div style="text-align: center; flex: 1; border-left: 1px solid #ddd; border-right: 1px solid #ddd;">
                    <span style="display: block; font-size: 0.7em; color: #7f8c8d; text-transform: uppercase;">Volume</span>
                    <span style="font-size: 1.1em; font-weight: bold; color: #27ae60;">${vol}M</span>
                </div>
                <div style="text-align: center; flex: 1.5;">
                    <span style="display: block; font-size: 0.7em; color: #7f8c8d; text-transform: uppercase;">Top Sector</span>
                    <span style="font-size: 0.9em; font-weight: bold; color: #2c3e50;">{sector}</span>
                </div>
            </div>

            <h4 style="margin: 0 0 8px 0; font-size: 0.8em; color: #7f8c8d; text-transform: uppercase; letter-spacing: 1px;">Top 5 Institutional Investors</h4>
            {investors_list}
        </div>
        """

    m = folium.Map(location=[0, 20], zoom_start=3, tiles='CartoDB positron')

    folium.Choropleth(
        geo_data=geo_data,
        data=stats_map,
        columns=['Country', 'Nb_Deals'],
        key_on='feature.properties.name',
        fill_color='YlGnBu',
        fill_opacity=0.8,
        line_opacity=0.3,
        nan_fill_color='#ecf0f1',
        legend_name='Market Activity (Number of Deals)',
        bins=8
    ).add_to(m)

    folium.GeoJson(
        geo_data,
        style_function=lambda x: {'fillColor': 'transparent', 'color': 'transparent'},
        popup=folium.features.GeoJsonPopup(fields=['popup_content'], labels=False),
        tooltip=folium.features.GeoJsonTooltip(fields=['name'], aliases=['Country:'])
    ).add_to(m)

    return m

# ══════════════════════════════════════════════════════════════
# SECTION 2 — GEOGRAPHY
# ══════════════════════════════════════════════════════════════

def plot_top_countries(df_deals):
    country_stats = df_deals.groupby('Country').agg(
        Nb_Deals=('Start-up name', 'count'),
        Total_USD_M=('Amount_clean', 'sum')
    ).reset_index()

    top_vol = country_stats.nlargest(10, 'Nb_Deals').set_index('Country')
    top_val = country_stats.nlargest(10, 'Total_USD_M').set_index('Country')

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 7))
    sns.barplot(x=top_vol['Nb_Deals'], y=top_vol.index,
                palette='mako', ax=ax1, hue=top_vol.index, legend=False)
    ax1.set_title('Top 10 Countries by Deal Count')

    sns.barplot(x=top_val['Total_USD_M'], y=top_val.index,
                palette='rocket', ax=ax2, hue=top_val.index, legend=False)
    ax2.set_title('Top 10 Countries by Capital ($M)')
    plt.tight_layout()
    return fig


def plot_market_concentration(df_deals):
    market_dist = df_deals['Market_Type'].value_counts()

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 6))

    ax1.pie(market_dist.values, labels=market_dist.index,
            autopct='%1.1f%%', colors=['#2c3e50','#e74c3c','#95a5a6'],
            startangle=140)
    ax1.set_title('Market Concentration: Big 4 vs Rest\n(cumulative since 2019)')

    emerging_trend = df_deals.groupby('Year')['Market_Type'].apply(
        lambda x: (x == 'Emerging Markets').mean() * 100
    )
    avg = emerging_trend.mean()
    ax2.plot(emerging_trend.index, emerging_trend.values,
             marker='o', color='#27ae60', linewidth=3)
    ax2.axhline(avg, color='gray', linestyle='--', alpha=0.6)
    ax2.text(emerging_trend.index.min() + 0.1, avg + 0.5,
             f'Average: {avg:.1f}%', fontsize=9, color='gray')
    ax2.set_title('Emerging Markets Share per Year (%)')
    ax2.set_ylabel('%')
    ax2.grid(alpha=0.3)

    max_val = emerging_trend.values.max()
    ax2.set_ylim(0, min(max_val * 1.2, 100))
    plt.tight_layout()
    return fig


# ══════════════════════════════════════════════════════════════
# SECTION 3 — SECTORS
# ══════════════════════════════════════════════════════════════

def plot_leading_sectors(df_deals):
    sector_summary = df_deals.groupby('Sector').agg(
        Nb_Deals=('Start-up name', 'count'),
        Total_USD_M=('Amount_clean', 'sum')
    ).reset_index()

    top_vol = sector_summary.nlargest(10, 'Nb_Deals').set_index('Sector')
    top_val = sector_summary.nlargest(10, 'Total_USD_M').set_index('Sector')

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 7))
    sns.barplot(x=top_vol['Nb_Deals'], y=top_vol.index,
                palette='mako', ax=ax1, hue=top_vol.index, legend=False)
    ax1.set_title('Top 10 Sectors by Deal Count')

    sns.barplot(x=top_val['Total_USD_M'], y=top_val.index,
                palette='rocket', ax=ax2, hue=top_val.index, legend=False)
    ax2.set_title('Top 10 Sectors by Capital ($M)')
    plt.tight_layout()
    return fig


def plot_sector_trends(df_deals):
    major_sectors = df_deals['Sector'].value_counts().index[:5].tolist()
    area_data = (df_deals[df_deals['Sector'].isin(major_sectors)]
                 .groupby(['Year','Sector']).size().unstack().fillna(0))
    area_perc = area_data.div(area_data.sum(axis=1), axis=0) * 100

    fig, ax = plt.subplots(figsize=(14, 7))
    area_perc.plot(kind='area', stacked=True, alpha=0.75, ax=ax,
                   colormap='tab10')
    ax.set_title('Relative Share of Leading Sectors (2019–2026)', fontsize=15)
    ax.set_ylabel('%')
    ax.legend(bbox_to_anchor=(1.01, 1))
    plt.tight_layout()
    return fig


def plot_sector_stage(df_deals):
    main_types  = ['Seed','Series A','Series B','Series C','Series D+','Grants']
    top_sectors = df_deals['Sector'].value_counts().index[:8].tolist()
    df_cross = df_deals[
        df_deals['Type_Simple'].isin(main_types) &
        df_deals['Sector'].isin(top_sectors)
    ]
    pivot_corr = (df_cross.groupby(['Sector','Type_Simple']).size()
                  .unstack(fill_value=0)
                  .reindex(columns=main_types, fill_value=0))
    pivot_corr = pivot_corr.div(pivot_corr.sum(axis=1), axis=0) * 100

    fig, ax = plt.subplots(figsize=(14, 7))
    sns.heatmap(pivot_corr, annot=True, fmt=".1f", cmap="Purples", ax=ax)
    ax.set_title('Funding Stage Distribution by Sector (%)', fontsize=14)
    plt.tight_layout()
    return fig

def plot_climate_tech(df_deals):
    climate = df_deals[df_deals['Climate Tech'].astype(str).str.lower() == 'climate tech']
    
    if climate.empty:
        fig, ax = plt.subplots()
        ax.text(0.5, 0.5, 'No Climate Tech deals found', ha='center', va='center')
        return fig

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 6))

    # Évolution
    climate_counts = climate.groupby('Year').size()
    sns.barplot(x=climate_counts.index, y=climate_counts.values, color='#27ae60', ax=ax1)
    ax1.set_title('Climate Tech: Number of Deals per Year', fontsize=13)
    ax1.set_ylabel('Number of deals')

    # Top secteurs d'application
    top_sectors = climate['Sector'].value_counts().head(6)
    sns.barplot(x=top_sectors.values, y=top_sectors.index, palette='Greens_r', ax=ax2)
    ax2.set_title('Top Sub-Sectors within Climate Tech', fontsize=13)
    ax2.set_xlabel('Number of deals')

    plt.tight_layout()
    return fig

# ══════════════════════════════════════════════════════════════
# SECTION 4 — FUNDING DYNAMICS
# ══════════════════════════════════════════════════════════════

def plot_funding_funnel(df_deals):
    funnel_order = ['Grants','Pre-Seed','Seed','Series A','Series B','Series C','Series D+']
    funnel_data  = (df_deals[df_deals['Type_Simple'].isin(funnel_order)]
                    .groupby('Type_Simple')['Start-up name']
                    .nunique().reindex(funnel_order).dropna())

    # Nouveau calcul : Taux de déperdition global (Macro Conversion)
    conversion_text = ""
    transitions = [('Seed','Series A'),('Series A','Series B'),('Series B','Series C')]
    for f, t in transitions:
        if f in funnel_data.index and t in funnel_data.index:
            vol_f = funnel_data[f]
            vol_t = funnel_data[t]
            if vol_f > 0:
                conversion_text += f"{f}→{t}: {(vol_t/vol_f)*100:.0f}%  "

    fig, ax = plt.subplots(figsize=(12, 7))
    sns.barplot(x=funnel_data.values, y=funnel_data.index,
                palette='Blues_r', hue=funnel_data.index, legend=False, ax=ax)
    
    for i, v in enumerate(funnel_data.values):
        ax.text(v + 5, i, str(int(v)), va='center', fontsize=10)

    ax.set_title(f'Startup Funding Funnel (Ecosystem Drop-off)\nMacro conversion rates: {conversion_text}',
                 fontsize=13)
    ax.set_xlabel('Number of unique startups')
    plt.tight_layout()
    return fig


def plot_time_between_rounds(df_deals):
    df_s = df_deals.sort_values(['Start-up name','Deal Date'])
    df_s['Prev_Date']  = df_s.groupby('Start-up name')['Deal Date'].shift(1)
    df_s['Prev_Stage'] = df_s.groupby('Start-up name')['Type_Simple'].shift(1)
    df_gap = df_s[df_s['Prev_Date'].notna()].copy()
    df_gap['Days_Between_Rounds'] = (df_gap['Deal Date'] - df_gap['Prev_Date']).dt.days
    df_gap = df_gap[df_gap['Days_Between_Rounds'] > 0]

    # Global
    gap_by_sector = (df_gap.groupby('Sector')['Days_Between_Rounds']
                     .agg(['median','count'])
                     .rename(columns={'median':'Median_Days','count':'N_Rounds'})
                     .query('N_Rounds >= 5')
                     .assign(Median_Months=lambda d: d['Median_Days']/30)
                     .sort_values('Median_Months'))

    # Par transition
    transitions = [('Seed','Series A'),('Series A','Series B'),('Series B','Series C')]
    results = []
    for f, t in transitions:
        sub = df_gap[(df_gap['Prev_Stage']==f) & (df_gap['Type_Simple']==t)]
        by_s = (sub.groupby('Sector')['Days_Between_Rounds']
                .agg(['median','count'])
                .rename(columns={'median':'Median_Days','count':'N'})
                .query('N >= 3')
                .assign(Median_Months=lambda d: d['Median_Days']/30,
                        Transition=f'{f} → {t}')
                .reset_index())
        results.append(by_s)
    df_trans = pd.concat(results, ignore_index=True)

    fig = plt.figure(figsize=(22, 14))
    gs  = fig.add_gridspec(2, 3)

    ax_global = fig.add_subplot(gs[0, :])
    sns.barplot(x=gap_by_sector['Median_Months'], y=gap_by_sector.index,
                hue=gap_by_sector.index, palette='coolwarm_r',
                legend=False, ax=ax_global)
    for i, (_, row) in enumerate(gap_by_sector.iterrows()):
        ax_global.text(row['Median_Months']+0.2, i,
                       f"n={int(row['N_Rounds'])}", va='center', fontsize=9)
    ax_global.set_title('Median Interval Between Rounds by Sector — All Transitions',
                        fontsize=13)

    palette = {'Seed → Series A':'#3498db',
               'Series A → Series B':'#27ae60',
               'Series B → Series C':'#e67e22'}

    for idx, (ax_pos, (f, t)) in enumerate(
            zip([gs[1,0], gs[1,1], gs[1,2]], transitions)):
        ax = fig.add_subplot(ax_pos)
        label  = f'{f} → {t}'
        subset = df_trans[df_trans['Transition']==label].sort_values('Median_Months')
        if subset.empty:
            ax.axis('off'); continue
        sns.barplot(x='Median_Months', y='Sector', data=subset,
                    color=palette[label], ax=ax, legend=False)
        for i, (_, row) in enumerate(subset.iterrows()):
            ax.text(row['Median_Months']+0.3, i,
                    f"n={int(row['N'])}", va='center', fontsize=9)
        ax.set_title(label, fontsize=12, fontweight='bold', color=palette[label])
        ax.set_xlabel('Median months')
        ax.grid(axis='x', alpha=0.3)

    plt.suptitle('Time Between Rounds — Global & By Stage Transition',
                 fontsize=16, fontweight='bold')
    plt.tight_layout()
    return fig


# ══════════════════════════════════════════════════════════════
# SECTION 5 — FOUNDERS
# ══════════════════════════════════════════════════════════════

def plot_gender_diversity(df_deals):
    def fix_gender(val, label):
        v = str(val).strip()
        if v == label or v == 'Yes': return 'Yes'
        if v == 'No': return 'No'
        return np.nan

    df_g = df_deals.copy()
    df_g['Woman_CF'] = df_g['Woman co-founder'].apply(lambda x: fix_gender(x, 'Woman co-founder'))
    df_g['Female_CEO'] = df_g['CEO Gender'].apply(lambda x: 'Yes' if str(x).strip() == 'Female' else 'No')

    gender_analysis = df_g.groupby('Year').agg(
        Female_CEO_pct=('Female_CEO', lambda x: (x=='Yes').mean()*100),
        Woman_CF_pct=('Woman_CF',   lambda x: (x=='Yes').mean()*100),
    )

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(18, 6))
    gender_analysis.plot(marker='o', linewidth=3, ax=ax1)
    ax1.set_title('Gender Diversity Trends (2019–2026)', fontsize=14)
    ax1.set_ylabel('%')
    ax1.legend(['Female CEO %', 'Woman Co-Founder %'])

    sector_stats = df_g.groupby('Sector')['Female_CEO'].agg(
        ['count', lambda x: (x=='Yes').mean()*100]
    ).rename(columns={'count':'n','<lambda_0>':'pct'})
    reliable = sector_stats[sector_stats['n'] >= 10].sort_values('pct')
    sns.barplot(x=reliable['pct'], y=reliable.index,
                palette='RdPu', hue=reliable.index, legend=False, ax=ax2)
    ax2.set_title('Female CEOs by Sector (sectors with >10 deals)')
    ax2.set_xlabel('%')
    plt.tight_layout()
    return fig


def plot_yc_effect(df_deals):
    yc_sectors = df_deals[df_deals['is_yc']==1]['Sector'].unique()
    df_plot_yc = df_deals[df_deals['Sector'].isin(yc_sectors)]

    fig, ax = plt.subplots(figsize=(14, 7))
    if not df_plot_yc.empty:
        sns.barplot(data=df_plot_yc, x='Sector', y='Amount_clean',
                    hue='YC_Label', estimator='median', errorbar=None,
                    palette={'YC Alumni':'#ff6600','Non-YC':'#2c3e50'}, ax=ax)
        ax.set_title('YC Effect on Median Amount Raised by Sector ($M)', fontsize=15)
        ax.tick_params(axis='x', rotation=45)
    plt.tight_layout()
    return fig


# ══════════════════════════════════════════════════════════════
# SECTION 6 — INVESTORS
# ══════════════════════════════════════════════════════════════

def plot_top_investors(df_inv_final):
    top15 = df_inv_final['Investors_List'].value_counts().head(15)
    fig, ax = plt.subplots(figsize=(12, 8))
    sns.barplot(x=top15.values, y=top15.index,
                palette='viridis', hue=top15.index, legend=False, ax=ax)
    ax.set_title('Top 15 Most Active Investors (2019–2026)', fontsize=15)
    ax.set_xlabel('Number of transactions')
    plt.tight_layout()
    return fig


def plot_investor_clustering(df_investisseurs, years_cols):
    X        = df_investisseurs[years_cols].fillna(0)
    X_scaled = StandardScaler().fit_transform(X)
    kmeans   = KMeans(n_clusters=4, random_state=42, n_init=10)
    df_investisseurs['Cluster'] = kmeans.fit_predict(X_scaled)

    trends = df_investisseurs.groupby('Cluster')[years_cols].mean().T
    trends.index = [2019,2020,2021,2022,2023,2024,2025,2026]

    cluster_labels = {
        0: 'Consistent Mid-Tier',
        1: 'Occasional / One-Shot',
        2: 'Boom & Retreat (2021 peak)',
        3: 'Dominant Late Surge'
    }

    fig, ax = plt.subplots(figsize=(12, 6))
    for cluster in trends.columns:
        ax.plot(trends.index, trends[cluster], marker='o', linewidth=3,
                label=cluster_labels.get(cluster, f'Cluster {cluster}'))
    ax.set_title('Investor Behavioral Archetypes (2019–2026)', fontsize=15)
    ax.set_ylabel('Average number of deals per year')
    ax.legend(title='Investor Profile')
    ax.grid(alpha=0.3)
    plt.tight_layout()
    return fig


def plot_south_africa_investors(df_inv_final):
    df_sa = df_inv_final[df_inv_final['Country'] == 'South Africa'].copy()
    sa_stats = (df_sa.groupby('Investors_List')['Amount_clean']
                .agg(['median','mean','count','sum'])
                .rename(columns={'median':'Median_Ticket','mean':'Avg_Ticket',
                                 'count':'Nb_Deals','sum':'Total_Deployed'})
                .reset_index())

    def assign_bracket(m):
        if pd.isna(m) or m == 0: return 'Undisclosed'
        elif m < 1:    return '< $1M'
        elif m < 5:    return '$1M – $5M'
        elif m < 15:   return '$5M – $15M'
        elif m < 50:   return '$15M – $50M'
        else:          return '$50M+'

    sa_stats['Ticket_Bracket'] = sa_stats['Median_Ticket'].apply(assign_bracket)

    bracket_order  = ['< $1M','$1M – $5M','$5M – $15M','$15M – $50M','$50M+']
    bracket_colors = {'< $1M':'#95a5a6','$1M – $5M':'#3498db',
                      '$5M – $15M':'#27ae60','$15M – $50M':'#e67e22','$50M+':'#e74c3c'}
    bracket_counts = sa_stats['Ticket_Bracket'].value_counts()

    figs = []

    # Plot 1 — par nombre de deals
    fig1, ax = plt.subplots(figsize=(12, 8))
    top20 = sa_stats.sort_values('Nb_Deals', ascending=False).head(20)
    bars  = ax.barh(top20['Investors_List'], top20['Nb_Deals'],
                    color='#2c3e50', alpha=0.85)
    for bar, (_, row) in zip(bars, top20.iterrows()):
        ax.text(bar.get_width()+0.05, bar.get_y()+bar.get_height()/2,
                f"avg ${row['Avg_Ticket']:.1f}M", va='center', fontsize=8.5)
    ax.invert_yaxis()
    ax.set_title('Top 20 Most Active Investors — South Africa\n(by deal count)', fontsize=13)
    ax.set_xlabel('Number of deals')
    ax.grid(axis='x', alpha=0.3)
    plt.tight_layout()
    figs.append(fig1)

    # Plot 2 — par montant total
    fig2, ax = plt.subplots(figsize=(12, 8))
    top20a = sa_stats[sa_stats['Total_Deployed']>0].sort_values('Total_Deployed', ascending=False).head(20)
    bars   = ax.barh(top20a['Investors_List'], top20a['Total_Deployed'],
                     color='#e74c3c', alpha=0.85)
    for bar, (_, row) in zip(bars, top20a.iterrows()):
        ax.text(bar.get_width()+0.5, bar.get_y()+bar.get_height()/2,
                f"{int(row['Nb_Deals'])} deals", va='center', fontsize=8.5)
    ax.invert_yaxis()
    ax.set_title('Top 20 Largest Investors — South Africa\n(by total capital deployed)', fontsize=13)
    ax.set_xlabel('Total deployed ($M)')
    ax.grid(axis='x', alpha=0.3)
    plt.tight_layout()
    figs.append(fig2)

    # Plot 3 — par tranche de ticket
    brackets_to_plot = [b for b in bracket_order if bracket_counts.get(b,0) > 0]
    n = len(brackets_to_plot)
    fig3, axes = plt.subplots(1, n, figsize=(6*n, 8))
    if n == 1: axes = [axes]
    for ax, bracket in zip(axes, brackets_to_plot):
        subset = (sa_stats[sa_stats['Ticket_Bracket']==bracket]
                  .sort_values('Nb_Deals', ascending=False).head(10))
        bars = ax.barh(subset['Investors_List'], subset['Nb_Deals'],
                       color=bracket_colors[bracket], alpha=0.85)
        for bar, (_, row) in zip(bars, subset.iterrows()):
            ax.text(bar.get_width()+0.05, bar.get_y()+bar.get_height()/2,
                    f"avg ${row['Avg_Ticket']:.1f}M", va='center', fontsize=8.5)
        ax.set_title(f'{bracket}\n({bracket_counts.get(bracket,0)} investors)',
                     fontsize=11, fontweight='bold', color=bracket_colors[bracket])
        ax.invert_yaxis()
        ax.grid(axis='x', alpha=0.3)
        m = sa_stats[sa_stats['Ticket_Bracket']==bracket]['Nb_Deals'].median()
        ax.axvline(m, color='gray', linestyle=':', linewidth=1.2,
                   label=f'Median ({m:.0f})')
        ax.legend(fontsize=8)
    plt.suptitle('South Africa — Investors by Ticket Bracket', fontsize=14, fontweight='bold')
    plt.tight_layout()
    figs.append(fig3)

    return figs


# ══════════════════════════════════════════════════════════════
# SECTION 7 — INVESTMENT INTELLIGENCE
# ══════════════════════════════════════════════════════════════

def compute_investment_scores(df_deals):
    startup_profile = df_deals.groupby('Start-up name').agg(
        Sector=('Sector','last'),
        Country=('Country','last'),
        Total_Raised=('Amount_clean','sum'),
        Nb_Rounds=('Amount_clean','count'),
        Last_Round_Size=('Amount_clean','last'),
        Last_Round_Date=('Deal Date','max'),
        First_Round_Date=('Deal Date','min'),
        Avg_Round_Growth=('Amount_clean', lambda x: x.pct_change()
                          .replace([np.inf,-np.inf],np.nan).mean()),
        Is_YC=('is_yc','max'),
        Last_Stage=('Type_Simple','last'),
    ).reset_index()

    startup_profile['Days_Since_Last'] = (
        pd.Timestamp('today') - startup_profile['Last_Round_Date']
    ).dt.days
    startup_profile['Months_Since_Last'] = startup_profile['Days_Since_Last'] / 30
    startup_profile['Startup_Age_Days']  = (
        pd.Timestamp('today') - startup_profile['First_Round_Date']
    ).dt.days
    startup_profile['Raise_Velocity'] = (
        startup_profile['Total_Raised'] / (startup_profile['Startup_Age_Days']/365)
    )

    fundraising_window = {
        'Seed':(10,28),'Series A':(12,30),
        'Series B':(14,32),'Series C':(16,36),
        'Venture Round':(10,28)
    }

    def in_window(row):
        stage  = row['Last_Stage']
        months = row['Months_Since_Last']
        if stage not in fundraising_window: return False
        mn, mx = fundraising_window[stage]
        return mn <= months <= mx

    startup_profile['In_Window'] = startup_profile.apply(in_window, axis=1)

    target_stages = ['Seed','Series A','Series B','Venture Round']
    filtered = startup_profile[
        startup_profile['Last_Stage'].isin(target_stages) &
        (startup_profile['Total_Raised'] < 200) &
        (startup_profile['Total_Raised'] > 0.5) &
        (startup_profile['Nb_Rounds'] >= 1) &
        (startup_profile['In_Window'] == True)
    ].copy()

    features = ['Raise_Velocity','Nb_Rounds','Last_Round_Size','Avg_Round_Growth','Is_YC']
    clean = filtered.dropna(subset=features).copy()
    clean[features] = clean[features].replace([np.inf,-np.inf],np.nan).fillna(0)
    scaler = MinMaxScaler()
    clean[features] = scaler.fit_transform(clean[features])

    clean['Raw_Score'] = (
        clean['Raise_Velocity']    * 0.25 +
        clean['Nb_Rounds']         * 0.10 +
        clean['Last_Round_Size']   * 0.20 +
        clean['Avg_Round_Growth']  * 0.30 +
        clean['Is_YC']             * 0.15
    )
    clean['Stage_Confidence'] = clean['Last_Stage'].map({
        'Seed':0.50,'Series A':0.65,
        'Series B':0.70,'Venture Round':0.55
    }).fillna(0.5)
    clean['Investment_Score'] = clean['Raw_Score'] * clean['Stage_Confidence']

    return clean




def plot_fundraising_window(df_deals):
    df_s = df_deals.sort_values(['Start-up name','Deal Date'])
    df_s['Prev_Date']  = df_s.groupby('Start-up name')['Deal Date'].shift(1)
    df_s['Prev_Stage'] = df_s.groupby('Start-up name')['Type_Simple'].shift(1)
    df_gap = df_s[df_s['Prev_Date'].notna()].copy()
    df_gap['Days_To_Next'] = (df_gap['Deal Date'] - df_gap['Prev_Date']).dt.days
    df_gap = df_gap[df_gap['Days_To_Next'] > 0]

    valid_transitions = [
        ('Seed','Series A'),('Series A','Series B'),
        ('Series B','Series C'),('Series C','Series D+')
    ]
    transition_stats = {}
    for f, t in valid_transitions:
        mask = ((df_gap['Prev_Stage']==f) &
                (df_gap['Type_Simple']==t) &
                (df_gap['Days_To_Next']>0))
        data = df_gap[mask]['Days_To_Next']
        if len(data) < 5: continue
        transition_stats[f] = {
            'to_stage': t,
            'median_m': data.median()/30,
            'p25_m':    data.quantile(0.25)/30,
            'p75_m':    data.quantile(0.75)/30,
            'danger_m': data.quantile(0.90)/30,
            'n':        len(data)
        }

    last_round = df_deals.sort_values('Deal Date').groupby('Start-up name').last().reset_index()
    last_round['Months_Since_Last'] = (
        pd.Timestamp('today') - last_round['Deal Date']
    ).dt.days / 30

    def classify(row):
        stage  = row['Type_Simple']
        months = row['Months_Since_Last']
        if stage not in transition_stats: return 'Unknown',None,None,None
        s = transition_stats[stage]
        if months < s['p25_m']:   status = 'Too Early'
        elif months <= s['p75_m']: status = 'In Window'
        elif months <= s['danger_m']: status = 'Late'
        else: status = 'At Risk'
        return status, s['to_stage'], s['median_m'], s['p75_m']

    last_round[['Window_Status','Next_Stage','Median_Target','Normal_Max']] = (
        last_round.apply(classify, axis=1, result_type='expand')
    )

    multi = df_deals.groupby('Start-up name').size()
    multi = multi[multi >= 2].index

    in_win = last_round[
        (last_round['Window_Status']=='In Window') &
        (last_round['Start-up name'].isin(multi)) &
        (last_round['Type_Simple'].isin(transition_stats.keys()))
    ].copy()

    in_win['Distance_To_Median'] = abs(
        in_win['Months_Since_Last'] - in_win['Median_Target']
    )
    top10 = in_win.sort_values('Distance_To_Median').head(10).reset_index(drop=True)

    stage_colors = {'Seed':'#3498db','Series A':'#2ecc71',
                    'Series B':'#e67e22','Series C':'#9b59b6'}

    fig, axes = plt.subplots(1, 2, figsize=(22, 8))

    # Gauche : fenêtres par stade
    stages = list(transition_stats.keys())
    y = range(len(stages))
    labels  = [f"{s} → {transition_stats[s]['to_stage']}" for s in stages]
    medians = [transition_stats[s]['median_m'] for s in stages]
    p25s    = [transition_stats[s]['p25_m']    for s in stages]
    p75s    = [transition_stats[s]['p75_m']    for s in stages]
    dangers = [transition_stats[s]['danger_m'] for s in stages]

    axes[0].barh(list(y), [p75s[i]-p25s[i] for i in range(len(p25s))],
                 left=p25s, height=0.4, color='#27ae60', alpha=0.5,
                 label='Normal window (P25–P75)')
    axes[0].scatter(medians, list(y), color='#27ae60', zorder=5, s=100, label='Median')
    axes[0].scatter(dangers, list(y), color='#e74c3c', zorder=5,
                    s=100, marker='x', linewidths=2, label='Danger (P90)')
    axes[0].set_yticks(list(y))
    axes[0].set_yticklabels(labels, fontsize=11)
    axes[0].set_xlabel('Months between rounds')
    axes[0].set_title('Real Transition Windows\n(African Ecosystem)', fontsize=13)
    axes[0].legend()
    axes[0].grid(axis='x', alpha=0.3)

    # Droite : top 10 in window
    for i, row in top10.iterrows():
        color = stage_colors.get(row['Type_Simple'], '#95a5a6')
        axes[1].barh(i, row['Months_Since_Last'], color=color, alpha=0.85, height=0.6)
        axes[1].vlines(row['Median_Target'], i-0.35, i+0.35,
                       color='#27ae60', linestyle='--', linewidth=1.5)
        axes[1].vlines(row['Normal_Max'], i-0.35, i+0.35,
                       color='#e74c3c', linestyle='--', linewidth=1.5)
        axes[1].text(row['Months_Since_Last']+0.3, i,
                     f"→ {row['Next_Stage']}  ({row['Months_Since_Last']:.0f}m)",
                     va='center', fontsize=9)

    axes[1].set_yticks(range(len(top10)))
    axes[1].set_yticklabels(
        [f"{r['Start-up name']}  [{r['Type_Simple']}]" for _, r in top10.iterrows()],
        fontsize=10
    )
    axes[1].set_xlabel('Months since last round')
    axes[1].set_title('Top 10 Startups in Fundraising Window\n(Optimal timing to approach now)',
                      fontsize=13)

    legend_e = [
        mpatches.Patch(facecolor='#3498db', alpha=0.85, label='Seed'),
        mpatches.Patch(facecolor='#2ecc71', alpha=0.85, label='Series A'),
        mpatches.Patch(facecolor='#e67e22', alpha=0.85, label='Series B'),
        mlines.Line2D([0],[0], color='#27ae60', linestyle='--', label='Ecosystem median'),
        mlines.Line2D([0],[0], color='#e74c3c', linestyle='--', label='Normal max (P75)'),
    ]
    axes[1].legend(handles=legend_e, loc='lower right')
    axes[1].grid(axis='x', alpha=0.3)

    plt.suptitle('Startup Fundraising Window Analysis', fontsize=15, fontweight='bold')
    plt.tight_layout()
    return fig, top10

# ══════════════════════════════════════════════════════════════
# SECTION 2.3 — PAN-AFRICAN VS LOCAL
# ══════════════════════════════════════════════════════════════

def plot_pan_african(df_deals):
    pan_african_comp = (
        df_deals
        .assign(Scope=lambda d: d['Country'].eq('Africa')
                .map({True: 'Pan-African', False: 'Country-Specific'}))
        .groupby('Scope')
        .agg(Nb_Deals=('Start-up name','count'),
             Median_Amount=('Amount_clean','median'))
        .reset_index()
    )
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6))
    sns.barplot(x='Scope', y='Nb_Deals', data=pan_african_comp,
                palette='Paired', hue='Scope', legend=False, ax=ax1)
    ax1.set_title('Number of Deals: Pan-African vs Country-Specific')
    ax1.set_ylabel('Number of deals')

    sns.barplot(x='Scope', y='Median_Amount', data=pan_african_comp,
                palette='Paired', hue='Scope', legend=False, ax=ax2)
    ax2.set_title('Median Amount Raised: Pan-African vs Country-Specific ($M)')
    ax2.set_ylabel('Median amount ($M)')
    plt.tight_layout()
    return fig


# ══════════════════════════════════════════════════════════════
# SECTION 3.3 — DEAL SIZE BY SECTOR
# ══════════════════════════════════════════════════════════════

def plot_deal_size_by_sector(df_deals):
    median_size = df_deals.groupby('Sector')['Amount_clean'].median().sort_values(ascending=False).head(10).reset_index()
    
    fig = px.bar(median_size, x='Amount_clean', y='Sector', 
                 orientation='h', 
                 title='Median Deal Size by Sector ($M)',
                 color='Amount_clean', 
                 color_continuous_scale='Viridis',
                 labels={'Amount_clean': 'Median Size ($M)', 'Sector': ''})
    
    fig.update_layout(
        plot_bgcolor='rgba(0,0,0,0)', 
        paper_bgcolor='rgba(0,0,0,0)',
        yaxis={'categoryorder':'total ascending'}, 
        font=dict(color='#2c3e50')
    )
    return fig


# ══════════════════════════════════════════════════════════════
# SECTION 4.3 — FIRST-TIME FUNDRAISERS
# ══════════════════════════════════════════════════════════════

def plot_first_time_fundraisers(df_deals):
    first_deals = (df_deals.sort_values('Deal Date')
                   .groupby('Start-up name').first().reset_index())

    top_sectors   = first_deals['Sector'].value_counts().head(10)
    top_countries = first_deals['Country'].value_counts().head(10)

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 7))
    sns.barplot(x=top_sectors.values, y=top_sectors.index,
                palette='mako', hue=top_sectors.index, legend=False, ax=ax1)
    ax1.set_title('Key Sectors for First Fundraising Rounds')
    ax1.set_xlabel('Number of first deals')

    sns.barplot(x=top_countries.values, y=top_countries.index,
                palette='rocket', hue=top_countries.index, legend=False, ax=ax2)
    ax2.set_title('Most Dynamic Countries for First Fundraising Rounds')
    ax2.set_xlabel('Number of first deals')
    plt.tight_layout()
    return fig


# ══════════════════════════════════════════════════════════════
# SECTION 4.4 — NON-DILUTIVE FINANCING
# ══════════════════════════════════════════════════════════════

def plot_non_dilutive(df_deals):
    non_equity_types = ['Debt','Grant','Bonds','Bond','Green Bonds',
                        'Revenue-based financing','Convertible Note']
    non_equity = df_deals[
        df_deals['Type'].str.contains('|'.join(non_equity_types), case=False, na=False)
    ].copy()
    non_equity = non_equity[non_equity['Sector'].notna()]

    if non_equity.empty:
        fig, ax = plt.subplots()
        ax.text(0.5, 0.5, 'No non-dilutive deals found',
                ha='center', va='center')
        return fig

    fig, ax = plt.subplots(figsize=(12, 7))
    sns.countplot(data=non_equity, y='Sector',
                  order=non_equity['Sector'].value_counts().index[:10],
                  palette='magma', hue='Sector', legend=False, ax=ax)
    ax.set_title('Sectors Financing Growth Without Dilution\n(Debt, Grants, Bonds)',
                 fontsize=14)
    ax.set_xlabel('Number of deals')
    plt.tight_layout()
    return fig


# ══════════════════════════════════════════════════════════════
# SECTION 4.5 — STARTUP SURVIVAL SIGNALS
# ══════════════════════════════════════════════════════════════

def plot_survival_signals(df_deals):
    last_deal = df_deals.groupby('Start-up name')['Deal Date'].max().reset_index()
    last_deal['Days_Since'] = (pd.Timestamp('today') - last_deal['Deal Date']).dt.days
    last_deal['Years_Since'] = last_deal['Days_Since'] / 365
    last_deal['Status'] = pd.cut(
        last_deal['Years_Since'],
        bins=[0, 1.5, 3, 100],
        labels=['Active (<18m)', 'Slowing (18m–3y)', 'Dormant (3y+)']
    )

    status_counts = last_deal['Status'].value_counts()

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6))

    colors = {'Active (<18m)':'#27ae60','Slowing (18m–3y)':'#e67e22','Dormant (3y+)':'#e74c3c'}
    ax1.pie(status_counts.values,
            labels=status_counts.index,
            colors=[colors[s] for s in status_counts.index],
            autopct='%1.1f%%', startangle=140)
    ax1.set_title('Startup Activity Status\n(based on time since last deal)', fontsize=13)

    last_deal_sector = df_deals.sort_values('Deal Date').groupby('Start-up name').last().reset_index()
    last_deal_sector['Days_Since'] = (pd.Timestamp('today') - last_deal_sector['Deal Date']).dt.days
    last_deal_sector['Years_Since'] = last_deal_sector['Days_Since'] / 365
    last_deal_sector['Status'] = pd.cut(
        last_deal_sector['Years_Since'],
        bins=[0, 1.5, 3, 100],
        labels=['Active (<18m)', 'Slowing (18m–3y)', 'Dormant (3y+)']
    )
    pivot = (last_deal_sector.groupby(['Sector','Status']).size()
             .unstack(fill_value=0)
             .apply(lambda x: x/x.sum()*100, axis=1))
    top_sectors = last_deal_sector['Sector'].value_counts().index[:8]
    pivot = pivot.reindex(top_sectors)
    pivot.plot(kind='barh', stacked=True, ax=ax2,
               color=[colors[c] for c in pivot.columns if c in colors])
    ax2.set_title('Startup Status by Sector (%)', fontsize=13)
    ax2.set_xlabel('%')
    ax2.legend(loc='lower right')
    plt.tight_layout()
    return fig


# ══════════════════════════════════════════════════════════════
# SECTION 5.2 — ACADEMIC BACKGROUND
# ══════════════════════════════════════════════════════════════

def plot_academic_background(df_deals):
    edu_continent = (df_deals['CEO - University/School - Continent']
                     .value_counts().dropna())
    edu_continent = edu_continent[edu_continent.index != 'nan']

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(18, 7))

    bars = ax1.barh(edu_continent.index, edu_continent.values,
                    color=sns.color_palette('pastel', len(edu_continent)))
    total = edu_continent.sum()
    for bar, val in zip(bars, edu_continent.values):
        ax1.text(bar.get_width() + 5, bar.get_y() + bar.get_height()/2,
                 f'{val/total*100:.1f}%', va='center', fontsize=10)
    ax1.set_title('Educational Background of CEOs\n(Continent of graduation)', fontsize=13)
    ax1.set_xlabel('Number of CEOs')

    df_edu = df_deals[df_deals['Amount_clean'].notna() &
                      df_deals['CEO - University/School - Continent'].notna()].copy()
    sns.boxplot(data=df_edu, x='CEO - University/School - Continent',
                y='Amount_clean', palette='Set3',
                hue='CEO - University/School - Continent', legend=False, ax=ax2)
    ax2.set_yscale('log')
    ax2.set_title('Amount Raised vs CEO Education Continent\n(log scale)', fontsize=13)
    ax2.set_xlabel('')
    ax2.tick_params(axis='x', rotation=45)
    plt.tight_layout()
    return fig


# ══════════════════════════════════════════════════════════════
# SECTION 5.3 — TEAM SIZE
# ══════════════════════════════════════════════════════════════

def plot_team_size(df_deals):
    df_t = df_deals.copy()
    df_t['Founders_Num'] = (df_t['# of Founders'].astype(str)
                            .str.extract(r'(\d+)')[0].astype(float))
    df_t['Founders_Group'] = df_t['Founders_Num'].apply(
        lambda x: '1' if x == 1 else '2' if x == 2 else '3' if x == 3
        else '4+' if x >= 4 else np.nan
    )
    df_t = df_t[df_t['Founders_Group'].notna()]

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6))
    sns.countplot(data=df_t, x='Founders_Group', order=['1','2','3','4+'],
                  palette='Blues', hue='Founders_Group', legend=False, ax=ax1)
    ax1.set_title('Number of Deals by Founding Team Size')
    ax1.set_xlabel('Number of founders')

    sns.barplot(data=df_t, x='Founders_Group', y='Amount_clean',
                order=['1','2','3','4+'], estimator='mean',
                palette='Oranges', hue='Founders_Group', legend=False, ax=ax2)
    ax2.set_title('Average Amount Raised by Team Size ($M)')
    ax2.set_xlabel('Number of founders')
    ax2.set_ylabel('Average amount ($M)')
    plt.tight_layout()
    return fig


# ══════════════════════════════════════════════════════════════
# SECTION 6.2 — INVESTOR ACTIVITY OVER TIME
# ══════════════════════════════════════════════════════════════

def plot_investor_activity_over_time(df_investisseurs, years_cols):
    top10 = df_investisseurs.nlargest(10, '2019-25 deals')
    activity = top10.set_index('Investor')[years_cols].T
    activity.index = [2019,2020,2021,2022,2023,2024,2025,2026]

    fig, ax = plt.subplots(figsize=(14, 7))
    for col in activity.columns:
        ax.plot(activity.index, activity[col], marker='o', linewidth=2, label=col)
    ax.set_title('Activity Trends of the Top 10 Investors (2019–2026)', fontsize=14)
    ax.set_ylabel('Number of deals per year')
    ax.legend(bbox_to_anchor=(1.01, 1), fontsize=8)
    ax.grid(alpha=0.3)
    plt.tight_layout()
    return fig


# ══════════════════════════════════════════════════════════════
# SECTION 6.4 — CAPITAL ORIGINS
# ══════════════════════════════════════════════════════════════

def plot_capital_origins(df_investisseurs):
    region_flow = (df_investisseurs.groupby('HQ Region')['2019-25 deals']
                   .sum().sort_values(ascending=False))
    fig, ax = plt.subplots(figsize=(12, 6))
    sns.barplot(x=region_flow.values, y=region_flow.index,
                hue=region_flow.index, palette='magma', legend=False, ax=ax)
    ax.set_title('Origin of Capital: Total Deals by Investor HQ Region', fontsize=14)
    ax.set_xlabel('Total deals')
    plt.tight_layout()
    return fig


# ══════════════════════════════════════════════════════════════
# SECTION 7.2 — KNIFE CAPITAL
# ══════════════════════════════════════════════════════════════

def plot_knife_portfolio(df_inv_final):
    knife = df_inv_final[
        df_inv_final['Investors_List'].str.contains('Knife', case=False, na=False)
    ].copy()

    if knife.empty:
        fig, ax = plt.subplots()
        ax.text(0.5, 0.5, 'No Knife Capital deals found', ha='center', va='center')
        return fig, None

    fig, axes = plt.subplots(2, 3, figsize=(20, 12))

    knife.groupby('Year').size().plot(kind='bar', ax=axes[0,0],
                                      color='#2c3e50', edgecolor='white')
    axes[0,0].set_title('Deals per Year')

    knife.groupby('Year')['Amount_clean'].sum().plot(kind='bar', ax=axes[0,1],
                                                      color='#e74c3c', edgecolor='white')
    axes[0,1].set_title('Total Capital Deployed per Year ($M)')

    knife['Sector'].value_counts().plot(kind='barh', ax=axes[0,2], color='#2980b9')
    axes[0,2].set_title('Deals by Sector')

    knife['Type_Simple'].value_counts().plot(kind='barh', ax=axes[1,0], color='#27ae60')
    axes[1,0].set_title('Deals by Stage')

    axes[1,1].hist(knife['Amount_clean'].dropna(), bins=10,
                   color='#8e44ad', edgecolor='white')
    axes[1,1].set_title('Ticket Size Distribution ($M)')

    knife['Country'].value_counts().plot(kind='barh', ax=axes[1,2], color='#e67e22')
    axes[1,2].set_title('Deals by Country')

    plt.suptitle('Knife Capital — Portfolio Overview', fontsize=18, fontweight='bold')
    plt.tight_layout()
    return fig, knife


def plot_knife_similar_deals(df_inv_final, df_deals):
    knife = df_inv_final[
        df_inv_final['Investors_List'].str.contains('Knife', case=False, na=False)
    ].copy()

    if knife.empty:
        fig, ax = plt.subplots()
        ax.text(0.5, 0.5, 'No Knife Capital deals found', ha='center', va='center')
        return fig, None

    knife_startup_names = knife['Start-up name'].unique()
    stage_order = {'Grants':0,'Pre-Seed':1,'Seed':2,'Series A':3,
                   'Series B':4,'Series C':5,'Series D+':6,'Venture Round':2.5}

    knife['Stage_Num']   = knife['Type_Simple'].map(stage_order).fillna(2)
    knife_sectors        = knife['Sector'].value_counts(normalize=True).to_dict()
    knife_countries      = knife['Country'].value_counts(normalize=True).to_dict()

    candidates = df_inv_final[
        ~df_inv_final['Start-up name'].isin(knife_startup_names)
    ].copy()
    candidates['Sector_Score']  = candidates['Sector'].map(knife_sectors).fillna(0)
    candidates['Country_Score'] = candidates['Country'].map(knife_countries).fillna(0)
    candidates['Stage_Num']     = candidates['Type_Simple'].map(stage_order).fillna(2)
    candidates['Amount_norm']   = candidates['Amount_clean'].fillna(0)

    knife['Sector_Score']  = knife['Sector'].map(knife_sectors).fillna(0)
    knife['Country_Score'] = knife['Country'].map(knife_countries).fillna(0)

    knife_vector = np.array([[
        knife['Sector_Score'].mean(), knife['Country_Score'].mean(),
        knife['Stage_Num'].mean(),    knife['Amount_clean'].mean()
    ]])

    feature_cols     = ['Sector_Score','Country_Score','Stage_Num','Amount_norm']
    candidate_matrix = candidates[feature_cols].fillna(0).values

    scaler   = MinMaxScaler()
    all_data = np.vstack([knife_vector, candidate_matrix])
    all_scaled = scaler.fit_transform(all_data)

    similarity = cosine_similarity(all_scaled[1:], all_scaled[0:1]).flatten()
    candidates['Similarity_Score'] = similarity

    startup_sim = candidates.groupby('Start-up name').agg(
        Sector=('Sector','last'),
        Country=('Country','last'),
        Last_Stage=('Type_Simple','last'),
        Total_Raised=('Amount_clean','sum'),
        Nb_Rounds=('Amount_clean','count'),
        Last_Round_Size=('Amount_clean','last'),
        Last_Round_Date=('Deal Date','max'),
        Is_YC=('is_yc','max'),
        Similarity_Score=('Similarity_Score','mean'),
    ).reset_index()

    startup_sim['Months_Since_Last'] = (
        pd.Timestamp('today') - startup_sim['Last_Round_Date']
    ).dt.days / 30

    fundraising_window = {
        'Seed':(10,28),'Series A':(12,30),
        'Series B':(14,32),'Venture Round':(10,28)
    }

    def in_window(row):
        stage = row['Last_Stage']; months = row['Months_Since_Last']
        if stage not in fundraising_window: return False
        mn, mx = fundraising_window[stage]
        return mn <= months <= mx

    startup_sim['In_Window'] = startup_sim.apply(in_window, axis=1)

    target_stages = ['Seed','Series A','Series B','Venture Round']
    filtered = startup_sim[
        startup_sim['Last_Stage'].isin(target_stages) &
        (startup_sim['Total_Raised'] < 200) &
        (startup_sim['Total_Raised'] > 0.5) &
        (startup_sim['In_Window'] == True)
    ].copy()

    momentum_cols = ['Total_Raised','Nb_Rounds','Last_Round_Size']
    clean = filtered.dropna(subset=momentum_cols).copy()
    clean[momentum_cols] = (clean[momentum_cols]
                            .replace([np.inf,-np.inf],np.nan).fillna(0))
    scaler2 = MinMaxScaler()
    clean[momentum_cols] = scaler2.fit_transform(clean[momentum_cols])

    clean['Momentum_Score'] = (
        clean['Total_Raised']     * 0.25 +
        clean['Nb_Rounds']        * 0.15 +
        clean['Last_Round_Size']  * 0.25 +
        clean['Is_YC']            * 0.10 +
        clean['Similarity_Score'] * 0.25
    )
    clean['Stage_Confidence'] = clean['Last_Stage'].map({
        'Seed':0.50,'Series A':0.65,
        'Series B':0.70,'Venture Round':0.55
    }).fillna(0.5)
    clean['Final_Score'] = clean['Momentum_Score'] * clean['Stage_Confidence']

    top15 = clean.nlargest(15,'Final_Score')[
        ['Start-up name','Sector','Country','Last_Stage',
         'Total_Raised','Months_Since_Last','Similarity_Score','Final_Score']
    ]

    fig, ax = plt.subplots(figsize=(14, 8))
    sns.barplot(x='Final_Score', y='Start-up name', data=top15,
                palette='viridis', hue='Last_Stage', dodge=False, ax=ax)
    ax.set_title("Top 15 Deals Most Similar to Knife Capital's Investment DNA",
                 fontsize=14)
    ax.set_xlabel('Final Score (Similarity × Momentum × Stage Confidence)')
    ax.legend(title='Stage', bbox_to_anchor=(1,1))
    plt.tight_layout()
    return fig, top15

def compute_summary_stats(df_deals):
    """Retourne un dict de stats clés pour l'IA."""
    return {
        'total_deals':      len(df_deals),
        'total_capital':    df_deals['Amount_clean'].sum(),
        'period':           f"{df_deals['Year'].min()}–{df_deals['Year'].max()}",
        'countries':        df_deals['Country'].nunique(),
        'top_sector':       df_deals['Sector'].value_counts().index[0],
        'top_country':      df_deals['Country'].value_counts().index[0],
        'yoy_growth':       df_deals.groupby('Year').size().pct_change().iloc[-1]*100,
        'seed_pct':         (df_deals['Type_Simple']=='Seed').mean()*100,
        'big4_pct':         (df_deals['Market_Type']=='Big 4').mean()*100,
        'climate_pct':      df_deals['Climate Tech'].astype(str)
                            .str.lower().eq('climate tech').mean()*100,
        'female_ceo_pct':   (df_deals['CEO Gender']=='Female').mean()*100,
        'yc_count':         df_deals['is_yc'].sum(),
        'median_deal':      df_deals['Amount_clean'].median(),
    }

# ══════════════════════════════════════════════════════════════
# SOUTH AFRICA — COMPLETE INVESTOR ANALYSIS
# ══════════════════════════════════════════════════════════════

import networkx as nx

def get_sa_data(df_inv_final):
    """Filtre les deals Sud-Africains et calcule les stats par investisseur."""
    df_sa = df_inv_final[df_inv_final['Country'] == 'South Africa'].copy()

    sa_stats = (
        df_sa.groupby('Investors_List')
        .agg(
            Nb_Deals=('Start-up name', 'count'),
            Total_Deployed=('Amount_clean', 'sum'),
            Median_Ticket=('Amount_clean', 'median'),
            Avg_Ticket=('Amount_clean', 'mean'),
            Max_Ticket=('Amount_clean', 'max'),
            First_Deal=('Deal Date', 'min'),
            Last_Deal=('Deal Date', 'max'),
        )
        .reset_index()
        .rename(columns={'Investors_List': 'Investor'})
    )
    sa_stats['Years_Active'] = (
        (sa_stats['Last_Deal'] - sa_stats['First_Deal']).dt.days / 365
    ).round(1)

    def assign_bracket(m):
        if pd.isna(m) or m == 0: return 'Undisclosed'
        elif m < 1:    return '< $1M'
        elif m < 5:    return '$1M–$5M'
        elif m < 15:   return '$5M–$15M'
        elif m < 50:   return '$15M–$50M'
        else:          return '$50M+'
    sa_stats['Ticket_Bracket'] = sa_stats['Median_Ticket'].apply(assign_bracket)

    return df_sa, sa_stats


# ── BLOC 1 : Vue d'ensemble ───────────────────────────────────

def plot_sa_overview(df_sa, sa_stats, df_inv_final):
    fig, axes = plt.subplots(2, 3, figsize=(22, 14))

    # 1. Évolution du nombre d'investisseurs actifs par année
    inv_by_year = (df_sa.groupby('Year')['Investors_List']
                   .nunique().reset_index()
                   .rename(columns={'Investors_List': 'Nb_Investors'}))
    axes[0,0].bar(inv_by_year['Year'], inv_by_year['Nb_Investors'],
                  color='#2c3e50', edgecolor='white')
    axes[0,0].set_title('Active Investors per Year — South Africa', fontsize=12)
    axes[0,0].set_ylabel('Number of unique investors')
    axes[0,0].grid(axis='y', alpha=0.3)

    # 2. Capital déployé par année
    capital_year = df_sa.groupby('Year')['Amount_clean'].sum()
    axes[0,1].bar(capital_year.index, capital_year.values,
                  color='#e74c3c', edgecolor='white')
    axes[0,1].set_title('Total Capital Deployed per Year ($M) — SA', fontsize=12)
    axes[0,1].set_ylabel('Total ($M)')
    axes[0,1].grid(axis='y', alpha=0.3)

    # 3. Local vs International
    if 'HQ Region' in df_inv_final.columns:
        df_sa_hq = df_sa.merge(
            df_inv_final[['Investors_List','HQ Region']].drop_duplicates(),
            on='Investors_List', how='left'
        )
        df_sa_hq['Origin'] = df_sa_hq['HQ Region'].apply(
            lambda x: 'Local (SA)' if str(x).strip() == 'Southern Africa'
            else 'International'
        )
        origin_counts = df_sa_hq['Origin'].value_counts()
        axes[0,2].pie(origin_counts.values, labels=origin_counts.index,
                      autopct='%1.1f%%', colors=['#27ae60','#3498db'],
                      startangle=140)
        axes[0,2].set_title('Local vs International Investors — SA', fontsize=12)
    else:
        axes[0,2].axis('off')

    # 4. Distribution des tickets
    valid_tickets = sa_stats[sa_stats['Median_Ticket'].notna() &
                             (sa_stats['Median_Ticket'] > 0)]['Median_Ticket']
    axes[1,0].hist(valid_tickets, bins=20, color='#8e44ad', edgecolor='white')
    axes[1,0].set_title('Median Ticket Size Distribution — SA ($M)', fontsize=12)
    axes[1,0].set_xlabel('Median ticket ($M)')
    axes[1,0].set_ylabel('Number of investors')
    axes[1,0].grid(axis='y', alpha=0.3)

    # 5. Taux de co-investissement
    deals_sa = df_sa.groupby(['Start-up name','Year'])['Investors_List'].count().reset_index()
    deals_sa['Is_Syndicated'] = deals_sa['Investors_List'] > 1
    syndication_rate = deals_sa.groupby('Year')['Is_Syndicated'].mean() * 100
    axes[1,1].plot(syndication_rate.index, syndication_rate.values,
                   marker='o', color='#e67e22', linewidth=3)
    axes[1,1].set_title('Co-Investment Rate per Year (%) — SA', fontsize=12)
    axes[1,1].set_ylabel('% of deals with 2+ investors')
    axes[1,1].grid(alpha=0.3)
    axes[1,1].fill_between(syndication_rate.index, syndication_rate.values,
                            alpha=0.15, color='#e67e22')

    # 6. Stats résumées
    axes[1,2].axis('off')
    total_inv   = sa_stats['Investor'].nunique()
    total_cap   = df_sa['Amount_clean'].sum()
    median_tick = sa_stats['Median_Ticket'].median()
    avg_deals   = sa_stats['Nb_Deals'].mean()
    synd_avg    = deals_sa['Is_Syndicated'].mean() * 100

    stats_text = (
        f"Total active investors : {total_inv:,}\n"
        f"Total capital deployed : ${total_cap:,.0f}M\n"
        f"Ecosystem median ticket : ${median_tick:.1f}M\n"
        f"Avg deals per investor : {avg_deals:.1f}\n"
        f"Avg co-investment rate : {synd_avg:.0f}%\n"
        f"Period covered : {df_sa['Year'].min()}–{df_sa['Year'].max()}"
    )
    axes[1,2].text(0.1, 0.5, stats_text, transform=axes[1,2].transAxes,
                   fontsize=13, verticalalignment='center',
                   bbox=dict(boxstyle='round', facecolor='#f0f4ff', alpha=0.8))
    axes[1,2].set_title('SA Ecosystem Summary', fontsize=12)

    plt.suptitle('South Africa — Investor Ecosystem Overview',
                 fontsize=18, fontweight='bold')
    plt.tight_layout()
    return fig


# ── BLOC 2 : Analyse sectorielle SA ──────────────────────────

def plot_sa_sectors(df_sa):
    fig, axes = plt.subplots(2, 2, figsize=(20, 14))

    # Top secteurs par deals
    top_s_vol = df_sa['Sector'].value_counts().head(10)
    sns.barplot(x=top_s_vol.values, y=top_s_vol.index,
                hue=top_s_vol.index, palette='mako', legend=False, ax=axes[0,0])
    axes[0,0].set_title('Top 10 Sectors by Deal Count — SA', fontsize=13)
    axes[0,0].set_xlabel('Number of deals')

    # Top secteurs par capital
    top_s_cap = df_sa.groupby('Sector')['Amount_clean'].sum().nlargest(10)
    sns.barplot(x=top_s_cap.values, y=top_s_cap.index,
                hue=top_s_cap.index, palette='rocket', legend=False, ax=axes[0,1])
    axes[0,1].set_title('Top 10 Sectors by Capital Deployed ($M) — SA', fontsize=13)
    axes[0,1].set_xlabel('Total capital ($M)')

    # Évolution sectorielle
    major_sectors = df_sa['Sector'].value_counts().index[:5].tolist()
    area_data = (df_sa[df_sa['Sector'].isin(major_sectors)]
                 .groupby(['Year','Sector']).size().unstack().fillna(0))
    area_perc = area_data.div(area_data.sum(axis=1), axis=0) * 100
    area_perc.plot(kind='area', stacked=True, alpha=0.75, ax=axes[1,0], colormap='tab10')
    axes[1,0].set_title('Sector Share Evolution — SA (2019–2026)', fontsize=13)
    axes[1,0].set_ylabel('%')
    axes[1,0].legend(bbox_to_anchor=(1.01, 1), fontsize=8)

    # Heatmap investisseur × secteur (top 15 investisseurs)
    top15_inv = df_sa['Investors_List'].value_counts().head(15).index
    df_heat = df_sa[df_sa['Investors_List'].isin(top15_inv)]
    pivot = (df_heat.groupby(['Investors_List','Sector']).size()
             .unstack(fill_value=0))
    sns.heatmap(pivot, cmap='YlOrRd', annot=True, fmt='d',
                linewidths=0.5, ax=axes[1,1])
    axes[1,1].set_title('Investor × Sector Matrix (Top 15 Investors) — SA', fontsize=13)
    axes[1,1].tick_params(axis='x', rotation=45)
    axes[1,1].tick_params(axis='y', rotation=0)

    plt.suptitle('South Africa — Sector Analysis', fontsize=18, fontweight='bold')
    plt.tight_layout()
    return fig


# ── BLOC 3 : Tickets SA ──────────────────────────────────────

def plot_sa_tickets(df_sa, sa_stats):
    fig, axes = plt.subplots(2, 2, figsize=(18, 12))

    # Évolution ticket médian par année
    tick_year = df_sa.groupby('Year')['Amount_clean'].median()
    axes[0,0].plot(tick_year.index, tick_year.values, marker='o',
                   color='#2c3e50', linewidth=3)
    axes[0,0].set_title('Median Ticket Size per Year — SA ($M)', fontsize=13)
    axes[0,0].set_ylabel('Median ($M)')
    axes[0,0].grid(alpha=0.3)
    axes[0,0].fill_between(tick_year.index, tick_year.values, alpha=0.15, color='#2c3e50')

    # Ticket médian par secteur
    tick_sector = (df_sa.groupby('Sector')['Amount_clean']
                   .median().sort_values(ascending=False).head(10))
    sns.barplot(x=tick_sector.values, y=tick_sector.index,
                hue=tick_sector.index, palette='viridis', legend=False, ax=axes[0,1])
    axes[0,1].set_title('Median Ticket by Sector — SA ($M)', fontsize=13)
    axes[0,1].set_xlabel('Median ($M)')

    # Ticket médian par stade
    stage_order = ['Seed','Series A','Series B','Series C','Series D+','Venture Round']
    tick_stage = (df_sa[df_sa['Type_Simple'].isin(stage_order)]
                  .groupby('Type_Simple')['Amount_clean']
                  .median().reindex(stage_order).dropna())
    axes[1,0].bar(tick_stage.index, tick_stage.values,
                  color='#e67e22', edgecolor='white')
    axes[1,0].set_title('Median Ticket by Stage — SA ($M)', fontsize=13)
    axes[1,0].set_ylabel('Median ($M)')
    axes[1,0].tick_params(axis='x', rotation=45)
    axes[1,0].grid(axis='y', alpha=0.3)

    # Comparaison SA vs Nigeria vs Kenya
    countries = ['South Africa', 'Nigeria', 'Kenya']
    comp_data = []
    for c in countries:
        med = df_sa['Amount_clean'].median() if c == 'South Africa' \
              else None
    df_comp_all = df_sa.copy()
    df_comp_all['Country_Group'] = 'South Africa'

    comp_rows = []
    for country in countries:
        subset = df_sa if country == 'South Africa' else None
        if country != 'South Africa':
            comp_rows.append({'Country': country,
                              'Median_Ticket': float('nan'),
                              'Nb_Deals': 0})
        else:
            comp_rows.append({'Country': country,
                              'Median_Ticket': df_sa['Amount_clean'].median(),
                              'Nb_Deals': len(df_sa)})

    # Version correcte : utiliser df_inv_final global
    # Cette fonction reçoit df_sa mais on annote pour que l'utilisateur
    # passe df_inv_final à la place si besoin
    tick_sa = df_sa['Amount_clean'].median()
    axes[1,1].text(0.5, 0.5,
                   f"SA Ecosystem Median Ticket:\n${tick_sa:.1f}M\n\n"
                   f"Total investors tracked: {sa_stats['Investor'].nunique()}\n"
                   f"Avg ticket: ${sa_stats['Avg_Ticket'].mean():.1f}M\n"
                   f"Max ticket: ${sa_stats['Max_Ticket'].max():.1f}M",
                   transform=axes[1,1].transAxes,
                   ha='center', va='center', fontsize=14,
                   bbox=dict(boxstyle='round', facecolor='#f0f4ff', alpha=0.8))
    axes[1,1].axis('off')
    axes[1,1].set_title('SA Ticket Summary', fontsize=13)

    plt.suptitle('South Africa — Ticket Size Analysis', fontsize=18, fontweight='bold')
    plt.tight_layout()
    return fig


# ── BLOC 4 : Profil individuel ────────────────────────────────

def get_investor_profile(investor_name, df_sa, sa_stats):
    """Retourne toutes les stats d'un investisseur SA."""
    inv_deals = df_sa[df_sa['Investors_List'] == investor_name].copy()
    inv_stats = sa_stats[sa_stats['Investor'] == investor_name]

    if inv_deals.empty:
        return None, None

    # Co-investisseurs fréquents
    all_deals_startup = df_sa[df_sa['Start-up name'].isin(inv_deals['Start-up name'].unique())]
    co_inv = (all_deals_startup[all_deals_startup['Investors_List'] != investor_name]
              ['Investors_List'].value_counts().head(10))

    return inv_deals, co_inv


def plot_investor_profile(investor_name, df_sa, sa_stats):
    inv_deals, co_inv = get_investor_profile(investor_name, df_sa, sa_stats)
    if inv_deals is None:
        fig, ax = plt.subplots()
        ax.text(0.5, 0.5, f'No data found for {investor_name}',
                ha='center', va='center')
        return fig

    fig, axes = plt.subplots(2, 3, figsize=(22, 14))

    # Timeline deals par année
    by_year = inv_deals.groupby('Year').size()
    axes[0,0].bar(by_year.index, by_year.values, color='#2c3e50', edgecolor='white')
    axes[0,0].set_title(f'Deals per Year — {investor_name}', fontsize=12)
    axes[0,0].set_ylabel('Number of deals')
    axes[0,0].grid(axis='y', alpha=0.3)

    # Répartition sectorielle
    sector_dist = inv_deals['Sector'].value_counts()
    axes[0,1].barh(sector_dist.index, sector_dist.values,
                   color=sns.color_palette('mako', len(sector_dist)))
    axes[0,1].set_title(f'Sector Distribution — {investor_name}', fontsize=12)
    axes[0,1].invert_yaxis()

    # Répartition par stade
    stage_dist = inv_deals['Type_Simple'].value_counts()
    axes[0,2].bar(stage_dist.index, stage_dist.values,
                  color=sns.color_palette('Set2', len(stage_dist)))
    axes[0,2].set_title(f'Stage Distribution — {investor_name}', fontsize=12)
    axes[0,2].tick_params(axis='x', rotation=45)

    # Évolution du ticket moyen
    tick_ev = inv_deals.groupby('Year')['Amount_clean'].mean()
    if not tick_ev.empty and tick_ev.notna().sum() > 0:
        axes[1,0].plot(tick_ev.index, tick_ev.values, marker='o',
                       color='#e74c3c', linewidth=3)
        axes[1,0].set_title(f'Average Ticket per Year ($M) — {investor_name}', fontsize=12)
        axes[1,0].set_ylabel('Avg ticket ($M)')
        axes[1,0].grid(alpha=0.3)
    else:
        axes[1,0].axis('off')
        axes[1,0].text(0.5, 0.5, 'No amount data', ha='center', va='center')

    # Co-investisseurs fréquents
    if not co_inv.empty:
        axes[1,1].barh(co_inv.index, co_inv.values,
                       color='#8e44ad', edgecolor='white')
        axes[1,1].set_title(f'Most Frequent Co-Investors — {investor_name}', fontsize=12)
        axes[1,1].invert_yaxis()
        axes[1,1].set_xlabel('Number of shared deals')
    else:
        axes[1,1].axis('off')
        axes[1,1].text(0.5, 0.5, 'No co-investments found', ha='center', va='center')

    # Stats résumées
    axes[1,2].axis('off')
    inv_row = sa_stats[sa_stats['Investor'] == investor_name]
    if not inv_row.empty:
        r = inv_row.iloc[0]
        stats_text = (
            f"Total deals : {int(r['Nb_Deals'])}\n"
            f"Total deployed : ${r['Total_Deployed']:.1f}M\n"
            f"Median ticket : ${r['Median_Ticket']:.1f}M\n"
            f"Avg ticket : ${r['Avg_Ticket']:.1f}M\n"
            f"Max ticket : ${r['Max_Ticket']:.1f}M\n"
            f"Ticket bracket : {r['Ticket_Bracket']}\n"
            f"Years active : {r['Years_Active']}\n"
            f"First deal : {str(r['First_Deal'])[:10]}\n"
            f"Last deal : {str(r['Last_Deal'])[:10]}"
        )
        axes[1,2].text(0.05, 0.5, stats_text, transform=axes[1,2].transAxes,
                       fontsize=12, verticalalignment='center',
                       bbox=dict(boxstyle='round', facecolor='#f0f4ff', alpha=0.8))
    axes[1,2].set_title(f'Key Stats — {investor_name}', fontsize=12)

    plt.suptitle(f'Investor Profile: {investor_name}',
                 fontsize=18, fontweight='bold')
    plt.tight_layout()
    return fig


def get_investor_portfolio_table(investor_name, df_sa):
    """Retourne le portfolio complet de l'investisseur sous forme de DataFrame."""
    inv_deals = df_sa[df_sa['Investors_List'] == investor_name].copy()
    if inv_deals.empty:
        return pd.DataFrame()
    return (inv_deals[['Start-up name','Sector','Country','Year',
                        'Type_Simple','Amount_clean','Deal Date']]
            .rename(columns={'Type_Simple':'Stage','Amount_clean':'Amount ($M)'})
            .sort_values('Deal Date', ascending=False)
            .reset_index(drop=True))


# ── BLOC 5 : Galaxie co-investissements ──────────────────────

def build_coinvestment_network(df_sa, min_shared=2):
    """Construit le graphe de co-investissements."""
    # Grouper par deal pour trouver les co-investisseurs
    deal_groups = df_sa.groupby(['Start-up name','Year'])['Investors_List'].apply(list)

    G = nx.Graph()

    for (startup, year), investors in deal_groups.items():
        investors = [i for i in investors if isinstance(i, str) and i.strip()]
        if len(investors) < 2:
            continue
        for i in range(len(investors)):
            for j in range(i+1, len(investors)):
                inv_a, inv_b = investors[i], investors[j]
                if G.has_edge(inv_a, inv_b):
                    G[inv_a][inv_b]['weight'] += 1
                    G[inv_a][inv_b]['deals'].append(startup)
                else:
                    G.add_edge(inv_a, inv_b, weight=1, deals=[startup])

    # Supprimer les liens faibles
    edges_to_remove = [(u,v) for u,v,d in G.edges(data=True)
                       if d['weight'] < min_shared]
    G.remove_edges_from(edges_to_remove)

    # Supprimer les nœuds isolés
    G.remove_nodes_from(list(nx.isolates(G)))

    # Ajouter les attributs des nœuds
    deal_counts = df_sa['Investors_List'].value_counts().to_dict()
    for node in G.nodes():
        G.nodes[node]['deals_count'] = deal_counts.get(node, 1)

    return G


def plot_coinvestment_network(df_sa, min_shared=2):
    """Visualise le réseau de co-investissements avec matplotlib."""
    G = build_coinvestment_network(df_sa, min_shared)

    if len(G.nodes()) == 0:
        fig, ax = plt.subplots()
        ax.text(0.5, 0.5, f'No co-investments with {min_shared}+ shared deals',
                ha='center', va='center')
        return fig

    fig, ax = plt.subplots(figsize=(16, 12))

    # Layout
    pos = nx.spring_layout(G, k=2, seed=42)

    # Taille des nœuds proportionnelle au nombre de deals
    node_sizes = [G.nodes[n]['deals_count'] * 80 for n in G.nodes()]

    # Épaisseur des liens proportionnelle au co-investissement
    edge_weights = [G[u][v]['weight'] for u,v in G.edges()]
    max_w = max(edge_weights) if edge_weights else 1
    edge_widths = [1 + (w/max_w) * 6 for w in edge_weights]

    # Couleurs par communauté
    communities = list(nx.community.greedy_modularity_communities(G))
    color_map = {}
    colors = ['#3498db','#e74c3c','#27ae60','#f39c12','#9b59b6',
              '#1abc9c','#e67e22','#2c3e50']
    for i, community in enumerate(communities):
        for node in community:
            color_map[node] = colors[i % len(colors)]
    node_colors = [color_map.get(n, '#95a5a6') for n in G.nodes()]

    # Dessiner
    nx.draw_networkx_edges(G, pos, ax=ax, width=edge_widths,
                           alpha=0.4, edge_color='#7f8c8d')
    nx.draw_networkx_nodes(G, pos, ax=ax, node_size=node_sizes,
                           node_color=node_colors, alpha=0.85)

    # Labels uniquement pour les nœuds importants
    important_nodes = {n: n for n in G.nodes()
                       if G.nodes[n]['deals_count'] >= 3}
    nx.draw_networkx_labels(G, pos, labels=important_nodes, ax=ax,
                            font_size=8, font_weight='bold')

    ax.set_title(
        f'Co-Investment Galaxy — South Africa\n'
        f'({len(G.nodes())} investors, {len(G.edges())} connections, '
        f'min. {min_shared} shared deals)\n'
        f'Node size = deal count | Edge thickness = co-investment frequency | '
        f'Color = community cluster',
        fontsize=13
    )
    ax.axis('off')

    # Légende communautés
    from matplotlib.patches import Patch
    legend_elements = [
        Patch(facecolor=colors[i], label=f'Cluster {i+1} ({len(c)} investors)')
        for i, c in enumerate(communities[:6])
    ]
    ax.legend(handles=legend_elements, loc='lower left', fontsize=9)

    plt.tight_layout()
    return fig


# ── BLOC 6 : Dynamiques temporelles SA ───────────────────────

def plot_sa_temporal(df_sa):
    fig, axes = plt.subplots(2, 2, figsize=(20, 12))

    # Top 5 investisseurs par année
    yearly_top = {}
    for year in sorted(df_sa['Year'].unique()):
        top = df_sa[df_sa['Year']==year]['Investors_List'].value_counts().head(5)
        yearly_top[year] = top

    # Heatmap activité top 20 investisseurs
    top20 = df_sa['Investors_List'].value_counts().head(20).index
    pivot = (df_sa[df_sa['Investors_List'].isin(top20)]
             .groupby(['Investors_List','Year']).size()
             .unstack(fill_value=0))
    sns.heatmap(pivot, cmap='Blues', annot=True, fmt='d',
                linewidths=0.5, ax=axes[0,0])
    axes[0,0].set_title('Activity Heatmap — Top 20 SA Investors', fontsize=12)
    axes[0,0].tick_params(axis='y', rotation=0)

    # Nouveaux investisseurs par année
    first_year = df_sa.groupby('Investors_List')['Year'].min()
    new_inv = first_year.value_counts().sort_index()
    axes[0,1].bar(new_inv.index, new_inv.values, color='#27ae60', edgecolor='white')
    axes[0,1].set_title('New Investors Entering SA per Year', fontsize=12)
    axes[0,1].set_ylabel('Number of new investors')
    axes[0,1].grid(axis='y', alpha=0.3)

    # Saisonnalité
    df_sa['Month'] = pd.to_datetime(df_sa['Deal Date']).dt.month
    monthly = df_sa.groupby('Month').size()
    month_names = ['Jan','Feb','Mar','Apr','May','Jun',
                   'Jul','Aug','Sep','Oct','Nov','Dec']
    axes[1,0].bar(range(1,13), [monthly.get(m,0) for m in range(1,13)],
                  color='#9b59b6', edgecolor='white')
    axes[1,0].set_xticks(range(1,13))
    axes[1,0].set_xticklabels(month_names)
    axes[1,0].set_title('Deal Seasonality — South Africa', fontsize=12)
    axes[1,0].set_ylabel('Number of deals')
    axes[1,0].grid(axis='y', alpha=0.3)

    # Investisseurs disparus (plus actifs depuis 2+ ans)
    last_year = df_sa.groupby('Investors_List')['Year'].max()
    current_year = df_sa['Year'].max()
    dormant = last_year[last_year <= current_year - 2].value_counts()
    if not dormant.empty:
        axes[1,1].bar(dormant.index, dormant.values, color='#e74c3c', edgecolor='white')
        axes[1,1].set_title('Last Active Year of Dormant Investors (2+ years inactive)',
                            fontsize=12)
        axes[1,1].set_ylabel('Number of investors')
        axes[1,1].grid(axis='y', alpha=0.3)
    else:
        axes[1,1].axis('off')
        axes[1,1].text(0.5, 0.5, 'No dormant investors found',
                       ha='center', va='center')

    plt.suptitle('South Africa — Investor Temporal Dynamics',
                 fontsize=18, fontweight='bold')
    plt.tight_layout()
    return fig


# ── BLOC 7 : Syndicats et co-investissements ─────────────────

def plot_sa_syndicates(df_sa):
    fig, axes = plt.subplots(1, 3, figsize=(22, 8))

    deal_groups = df_sa.groupby(['Start-up name','Year'])['Investors_List'].apply(list)

    # Top paires
    pair_counts = {}
    trio_counts = {}
    solo_investors = set()
    syndicated_investors = set()

    for (startup, year), investors in deal_groups.items():
        investors = [i for i in investors if isinstance(i, str) and i.strip()]
        if len(investors) == 1:
            solo_investors.add(investors[0])
        elif len(investors) > 1:
            for inv in investors:
                syndicated_investors.add(inv)
            for i in range(len(investors)):
                for j in range(i+1, len(investors)):
                    pair = tuple(sorted([investors[i], investors[j]]))
                    pair_counts[pair] = pair_counts.get(pair, 0) + 1
            if len(investors) >= 3:
                for i in range(len(investors)):
                    for j in range(i+1, len(investors)):
                        for k in range(j+1, len(investors)):
                            trio = tuple(sorted([investors[i],investors[j],investors[k]]))
                            trio_counts[trio] = trio_counts.get(trio, 0) + 1

    # Top 10 paires
    top_pairs = sorted(pair_counts.items(), key=lambda x: x[1], reverse=True)[:10]
    if top_pairs:
        pair_labels = [f"{p[0][:15]}…\n+\n{p[1][:15]}…" for p,_ in top_pairs]
        pair_vals   = [v for _,v in top_pairs]
        axes[0].barh(range(len(pair_labels)), pair_vals, color='#3498db')
        axes[0].set_yticks(range(len(pair_labels)))
        axes[0].set_yticklabels(pair_labels, fontsize=8)
        axes[0].invert_yaxis()
        axes[0].set_title('Top 10 Co-Investment Pairs — SA', fontsize=12)
        axes[0].set_xlabel('Number of shared deals')
    else:
        axes[0].axis('off')
        axes[0].text(0.5, 0.5, 'No pairs found', ha='center', va='center')

    # Top 10 triades
    top_trios = sorted(trio_counts.items(), key=lambda x: x[1], reverse=True)[:10]
    if top_trios:
        trio_labels = [f"{t[0][:10]}, {t[1][:10]}, {t[2][:10]}" for t,_ in top_trios]
        trio_vals   = [v for _,v in top_trios]
        axes[1].barh(range(len(trio_labels)), trio_vals, color='#27ae60')
        axes[1].set_yticks(range(len(trio_labels)))
        axes[1].set_yticklabels(trio_labels, fontsize=8)
        axes[1].invert_yaxis()
        axes[1].set_title('Top 10 Co-Investment Triads — SA', fontsize=12)
        axes[1].set_xlabel('Number of shared deals')
    else:
        axes[1].axis('off')
        axes[1].text(0.5, 0.5, 'Not enough triads found', ha='center', va='center')

    # Solo vs Syndicated
    only_solo = solo_investors - syndicated_investors
    only_synd = syndicated_investors - solo_investors
    both      = solo_investors & syndicated_investors

    cats = ['Always Solo', 'Always Syndicated', 'Both Styles']
    vals = [len(only_solo), len(only_synd), len(both)]
    axes[2].bar(cats, vals, color=['#e74c3c','#27ae60','#f39c12'], edgecolor='white')
    axes[2].set_title('Investment Style: Solo vs Syndicated — SA', fontsize=12)
    axes[2].set_ylabel('Number of investors')
    for i, v in enumerate(vals):
        axes[2].text(i, v + 0.3, str(v), ha='center', fontsize=11)
    axes[2].grid(axis='y', alpha=0.3)

    plt.suptitle('South Africa — Syndication & Co-Investment Patterns',
                 fontsize=18, fontweight='bold')
    plt.tight_layout()
    return fig


# ── BLOC 8 : Matrice de positionnement ───────────────────────

def plot_sa_positioning_matrix(df_sa, sa_stats, highlight_investor=None):
    """Bubble chart : ticket médian × nombre de deals, taille = capital total."""
    df_plot = sa_stats[
        sa_stats['Median_Ticket'].notna() &
        (sa_stats['Median_Ticket'] > 0) &
        (sa_stats['Nb_Deals'] >= 1)
    ].copy()

    fig, ax = plt.subplots(figsize=(16, 10))

    # Normaliser la taille des bulles
    max_cap = df_plot['Total_Deployed'].max()
    sizes   = ((df_plot['Total_Deployed'] / max_cap) * 2000 + 50).fillna(50)

    scatter = ax.scatter(
        df_plot['Median_Ticket'],
        df_plot['Nb_Deals'],
        s=sizes,
        alpha=0.6,
        c=df_plot['Total_Deployed'],
        cmap='viridis',
        edgecolors='white',
        linewidths=0.5
    )
    plt.colorbar(scatter, ax=ax, label='Total Capital Deployed ($M)')

    # Labels pour les top investisseurs
    top_by_deals  = df_plot.nlargest(10, 'Nb_Deals')['Investor']
    top_by_ticket = df_plot.nlargest(8, 'Median_Ticket')['Investor']
    to_label = set(top_by_deals) | set(top_by_ticket)

    for _, row in df_plot[df_plot['Investor'].isin(to_label)].iterrows():
        ax.annotate(
            row['Investor'][:20],
            (row['Median_Ticket'], row['Nb_Deals']),
            xytext=(5, 5), textcoords='offset points',
            fontsize=8, alpha=0.9
        )

    # Surligner un investisseur spécifique
    if highlight_investor:
        hi = df_plot[df_plot['Investor'] == highlight_investor]
        if not hi.empty:
            ax.scatter(hi['Median_Ticket'], hi['Nb_Deals'],
                       s=500, color='#e74c3c', zorder=5,
                       edgecolors='black', linewidths=2)
            ax.annotate(
                f"◀ {highlight_investor}",
                (hi['Median_Ticket'].iloc[0], hi['Nb_Deals'].iloc[0]),
                xytext=(10, 0), textcoords='offset points',
                fontsize=10, color='#e74c3c', fontweight='bold'
            )

    # Lignes de quadrant
    med_ticket = df_plot['Median_Ticket'].median()
    med_deals  = df_plot['Nb_Deals'].median()
    ax.axvline(med_ticket, color='gray', linestyle='--', alpha=0.5, linewidth=1)
    ax.axhline(med_deals,  color='gray', linestyle='--', alpha=0.5, linewidth=1)

    # Labels quadrants
    ax.text(ax.get_xlim()[1]*0.7, med_deals*1.1,
            'High Volume\nLarge Tickets', fontsize=9, color='gray', alpha=0.7)
    ax.text(ax.get_xlim()[0], med_deals*1.1,
            'High Volume\nSmall Tickets', fontsize=9, color='gray', alpha=0.7)
    ax.text(ax.get_xlim()[1]*0.7, med_deals*0.3,
            'Selective\nLarge Tickets', fontsize=9, color='gray', alpha=0.7)
    ax.text(ax.get_xlim()[0], med_deals*0.3,
            'Occasional\nSmall Tickets', fontsize=9, color='gray', alpha=0.7)

    ax.set_xlabel('Median Ticket Size ($M)', fontsize=12)
    ax.set_ylabel('Number of Deals', fontsize=12)
    ax.set_title(
        'Investor Positioning Matrix — South Africa\n'
        '(Bubble size = total capital deployed)',
        fontsize=14
    )
    ax.grid(alpha=0.2)
    plt.tight_layout()
    return fig
