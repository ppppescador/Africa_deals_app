import streamlit as st
import pandas as pd
import io
from analysis import (
    load_and_clean, compute_summary_stats,
    plot_deal_landscape, plot_market_trajectory,
    plot_heatmaps, plot_transparency,
    plot_top_countries, plot_market_concentration,
    plot_pan_african,
    plot_leading_sectors, plot_sector_trends,
    plot_deal_size_by_sector, plot_sector_stage,
    plot_climate_tech,
    plot_funding_funnel, plot_time_between_rounds,
    plot_first_time_fundraisers, plot_non_dilutive,
    plot_survival_signals,
    plot_gender_diversity, plot_academic_background,
    plot_team_size, plot_yc_effect,
    plot_top_investors, plot_investor_activity_over_time,
    plot_investor_clustering, plot_capital_origins,
    plot_south_africa_investors,
    plot_fundraising_window,
    plot_knife_portfolio, plot_knife_similar_deals,
    get_sa_data,
    plot_sa_overview, plot_sa_sectors, plot_sa_tickets,
    plot_investor_profile, get_investor_portfolio_table,
    plot_coinvestment_network, plot_sa_temporal,
    plot_sa_syndicates, plot_sa_positioning_matrix,
    create_interactive_map,
)
from streamlit_folium import st_folium
from groq import Groq

# ── CONFIG ────────────────────────────────────────────────────
st.set_page_config(
    page_title="Africa Deals Dashboard",
    page_icon="🌍",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    [data-testid="metric-container"] {
        background-color: #ffffff;
        border: 1px solid #e0e6ed;
        border-radius: 12px;
        padding: 20px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05);
    }
    [data-testid="metric-container"] label {
        color: #7f8c8d;
        font-size: 0.85rem !important;
        text-transform: uppercase;
        letter-spacing: 1px;
    }
    [data-testid="metric-container"] div[data-testid="stMetricValue"] {
        color: #1F3864;
        font-size: 1.8rem !important;
        font-weight: 700;
    }
    .section-header {
        font-size: 22px;
        font-weight: 700;
        color: #27ae60;
        border-bottom: 2px solid #f0f2f6;
        padding-bottom: 10px;
        margin: 30px 0 20px 0;
    }
    .block-container { padding-top: 2rem; padding-bottom: 2rem; }
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
</style>
""", unsafe_allow_html=True)


# ── FONCTION DOWNLOAD ─────────────────────────────────────────
def download_chart(fig, filename):
    """Boutons de téléchargement PNG et PDF pour un graphique matplotlib."""
    col1, col2 = st.columns(2)

    buf_png = io.BytesIO()
    fig.savefig(buf_png, format='png', dpi=300,
                bbox_inches='tight', facecolor='white')
    buf_png.seek(0)
    col1.download_button(
        label="⬇️ PNG",
        data=buf_png,
        file_name=f"{filename}.png",
        mime="image/png",
        key=f"png_{filename}"
    )

    buf_pdf = io.BytesIO()
    fig.savefig(buf_pdf, format='pdf',
                bbox_inches='tight', facecolor='white')
    buf_pdf.seek(0)
    col2.download_button(
        label="⬇️ PDF",
        data=buf_pdf,
        file_name=f"{filename}.pdf",
        mime="application/pdf",
        key=f"pdf_{filename}"
    )


def download_df(df, filename, label="⬇️ Download as Excel"):
    """Bouton de téléchargement Excel pour un DataFrame."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    buf.seek(0)
    st.download_button(
        label=label,
        data=buf,
        file_name=f"{filename}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=f"xlsx_{filename}"
    )


# ── SIDEBAR ───────────────────────────────────────────────────
with st.sidebar:
    st.image(
        "https://upload.wikimedia.org/wikipedia/commons/thumb/8/8d/"
        "Africa_%28orthographic_projection%29.svg/240px-Africa_%28orthographic_projection%29.svg.png",
        width=80
    )
    st.title("🌍 Africa Deals")
    st.divider()

    st.subheader("📁 Upload Data")
    uploaded_file = st.file_uploader("Upload Excel database", type=["xlsx"])
    password = st.text_input("File password", type="password")
    st.divider()
    st.caption("Upload a new Excel file each month to refresh all analyses automatically. Password July: B2$wF8^qL5#t")

# ── HEADER ────────────────────────────────────────────────────
st.markdown(
    '<h1 style="text-align:center; color:#1F3864;">🌍 African Startup Funding Report</h1>',
    unsafe_allow_html=True
)

if not uploaded_file or not password:
    st.info("👈 Upload your Excel file and enter the password in the sidebar to get started.")
    st.stop()


# ── CHARGEMENT ────────────────────────────────────────────────
@st.cache_data(show_spinner="Loading and processing data...")
def load_data(file_bytes, pwd):
    return load_and_clean(io.BytesIO(file_bytes), pwd)


try:
    file_bytes = uploaded_file.read()
    df_deals, df_investisseurs, df_inv_final, years_cols = load_data(file_bytes, password)
    stats = compute_summary_stats(df_deals)
    st.success(
        f"✅ {len(df_deals):,} deals loaded — "
        f"{df_deals['Year'].min()} to {df_deals['Year'].max()}"
    )
except Exception as e:
    st.error(f"❌ Error loading file: {e}")
    st.stop()


# ── CHATBOT SIDEBAR ───────────────────────────────────────────
with st.sidebar:
    st.markdown("---")
    client = Groq(api_key="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")

    with st.popover("💬 AI Assistant", use_container_width=True):
        st.markdown("**Ask your question**")

        if "chat_history" not in st.session_state:
            st.session_state.chat_history = []

        for msg in st.session_state.chat_history:
            st.chat_message(msg["role"]).write(msg["content"])

        if user_question := st.chat_input("What do you want to know?"):
            st.chat_message("user").write(user_question)
            st.session_state.chat_history.append(
                {"role": "user", "content": user_question}
            )
            context = (
                f"You are an expert in African venture capital. You are assisting a user "
                f"looking at a dashboard containing {len(df_deals)} deals with a total value "
                f"of ${df_deals['Amount_clean'].sum():.1f}M. "
                f"Answer concisely, professionally and analytically."
            )
            with st.spinner("Thinking..."):
                try:
                    completion = client.chat.completions.create(
                        messages=[
                            {"role": "system", "content": context},
                            {"role": "user", "content": user_question}
                        ],
                        model="llama-3.3-70b-versatile",
                    )
                    response = completion.choices[0].message.content
                    st.chat_message("assistant").write(response)
                    st.session_state.chat_history.append(
                        {"role": "assistant", "content": response}
                    )
                except Exception as e:
                    st.error(f"Error: {e}")


# ── KPI BAR ───────────────────────────────────────────────────
c1, c2, c3, c4, c5, c6 = st.columns(6)
c1.metric("Total Deals",    f"{stats['total_deals']:,}")
c2.metric("Total Capital",  f"${stats['total_capital']:,.0f}M")
c3.metric("Countries",      stats['countries'])
c4.metric("Top Sector",     stats['top_sector'])
c5.metric("Female CEO %",   f"{stats['female_ceo_pct']:.1f}%")
c6.metric("YC Alumni",      stats['yc_count'])
st.divider()


# ── TABS ──────────────────────────────────────────────────────
tabs = st.tabs([
    "📊 Market Overview",
    "🌍 Geography",
    "🏢 Sectors",
    "💰 Funding Dynamics",
    "👤 Founders",
    "🏦 Investors",
    "🇿🇦 South Africa Deep Dive",
    "🎯 Bonus — Predictions",
])


# ════════════════════════════════════════════════════════════
# TAB 1 — MARKET OVERVIEW
# ════════════════════════════════════════════════════════════
with tabs[0]:

    st.markdown('<div class="section-header">1.1 Deal Landscape</div>',
                unsafe_allow_html=True)
    st.markdown(
        "_An overview of the volume, size, and type of deals recorded across the study period. "
        "This dashboard provides a high-level snapshot of the dataset across five dimensions: "
        "round type, funding bracket, top sectors, top countries, and regional distribution._"
    )
    fig = plot_deal_landscape(df_deals)
    st.pyplot(fig, use_container_width=True)
    download_chart(fig, "1_1_deal_landscape")

    st.markdown('<div class="section-header">Interactive Map</div>',
                unsafe_allow_html=True)
    m = create_interactive_map(df_deals, df_inv_final)
    st_folium(m, width=1200, height=600, returned_objects=[])

    st.markdown('<div class="section-header">1.2 Market Trajectory (2019–2026)</div>',
                unsafe_allow_html=True)
    st.markdown(
        f"_Deal activity and total capital are tracked year over year. "
        f"The market peaked in 2021–2022, with the latest year showing a "
        f"**{stats['yoy_growth']:+.1f}%** change in deal count._"
    )
    # plot_market_trajectory retourne 2 figures
    fig_annual, fig_monthly = plot_market_trajectory(df_deals)
    st.markdown("**Annual view**")
    st.pyplot(fig_annual, use_container_width=True)
    download_chart(fig_annual, "1_2a_market_trajectory_annual")
    st.markdown("**Monthly breakdown — current dynamics**")
    st.pyplot(fig_monthly, use_container_width=True)
    download_chart(fig_monthly, "1_2b_market_trajectory_monthly")

    st.markdown('<div class="section-header">1.3 Capital Distribution by Sector & Region</div>',
                unsafe_allow_html=True)
    st.markdown(
        "_Two heatmaps cross-reference the top sectors with African regions. "
        "The left shows total capital deployed, the right shows median deal size._"
    )
    fig = plot_heatmaps(df_deals)
    st.pyplot(fig, use_container_width=True)
    download_chart(fig, "1_3_capital_distribution")

    st.markdown('<div class="section-header">1.4 Data Transparency & Climate Tech</div>',
                unsafe_allow_html=True)
    st.markdown(
        f"_Only a fraction of deals disclose their funding amount. "
        f"Climate Tech represents **{stats['climate_pct']:.1f}%** of all tagged deals._"
    )
    fig = plot_transparency(df_deals)
    st.pyplot(fig, use_container_width=True)
    download_chart(fig, "1_4_transparency")


# ════════════════════════════════════════════════════════════
# TAB 2 — GEOGRAPHY
# ════════════════════════════════════════════════════════════
with tabs[1]:

    st.markdown('<div class="section-header">2.1 Top Countries</div>',
                unsafe_allow_html=True)
    st.markdown(
        f"_**{stats['top_country']}** leads both in deal count and capital raised. "
        f"Rankings by volume and by financial value do not always align._"
    )
    fig = plot_top_countries(df_deals)
    st.pyplot(fig, use_container_width=True)
    download_chart(fig, "2_1_top_countries")

    st.markdown('<div class="section-header">2.2 Market Concentration</div>',
                unsafe_allow_html=True)
    st.markdown(
        f"_The Big 4 account for **{stats['big4_pct']:.0f}%** of all deals cumulatively. "
        f"Note: the pie chart reflects the cumulative distribution, "
        f"while the line chart shows year-by-year evolution._"
    )
    fig = plot_market_concentration(df_deals)
    st.pyplot(fig, use_container_width=True)
    download_chart(fig, "2_2_market_concentration")

    st.markdown('<div class="section-header">2.3 Pan-African vs Country-Specific</div>',
                unsafe_allow_html=True)
    st.markdown(
        "_Pan-African strategies tend to attract larger rounds, reflecting "
        "the broader addressable market narrative._"
    )
    fig = plot_pan_african(df_deals)
    st.pyplot(fig, use_container_width=True)
    download_chart(fig, "2_3_pan_african")


# ════════════════════════════════════════════════════════════
# TAB 3 — SECTORS
# ════════════════════════════════════════════════════════════
with tabs[2]:

    st.markdown('<div class="section-header">3.1 Leading Sectors</div>',
                unsafe_allow_html=True)
    st.markdown(
        f"_**{stats['top_sector']}** leads by a wide margin in deal count. "
        f"Capital-intensive sectors attract fewer but larger rounds._"
    )
    fig = plot_leading_sectors(df_deals)
    st.pyplot(fig, use_container_width=True)
    download_chart(fig, "3_1_leading_sectors")

    st.markdown('<div class="section-header">3.2 Sector Trends Over Time</div>',
                unsafe_allow_html=True)
    st.markdown(
        "_A stacked area chart tracks the relative share of the top 5 sectors from 2019 to 2026._"
    )
    fig = plot_sector_trends(df_deals)
    st.pyplot(fig, use_container_width=True)
    download_chart(fig, "3_2_sector_trends")

    st.markdown('<div class="section-header">3.3 Deal Size by Sector</div>',
                unsafe_allow_html=True)
    st.markdown(
        "_Median deal size reveals which sectors attract larger institutional capital._"
    )
    fig = plot_deal_size_by_sector(df_deals)
    st.pyplot(fig, use_container_width=True)
    download_chart(fig, "3_3_deal_size_sector")

    st.markdown('<div class="section-header">3.4 Funding Stage by Sector</div>',
                unsafe_allow_html=True)
    st.markdown(
        "_A normalized heatmap shows how funding stages distribute within each sector._"
    )
    fig = plot_sector_stage(df_deals)
    st.pyplot(fig, use_container_width=True)
    download_chart(fig, "3_4_sector_stage")

    st.markdown('<div class="section-header">3.5 Climate Tech Momentum</div>',
                unsafe_allow_html=True)
    st.markdown(
        "_Climate Tech has emerged as a major narrative in African VC. "
        "This section isolates startups tagged as Climate Tech to observe "
        "volume trajectory and which sectors are absorbing green capital._"
    )
    fig = plot_climate_tech(df_deals)
    st.pyplot(fig, use_container_width=True)
    download_chart(fig, "3_5_climate_tech")


# ════════════════════════════════════════════════════════════
# TAB 4 — FUNDING DYNAMICS
# ════════════════════════════════════════════════════════════
with tabs[3]:

    st.markdown('<div class="section-header">4.1 The Funding Funnel</div>',
                unsafe_allow_html=True)
    st.markdown(
        f"_**{stats['seed_pct']:.0f}%** of all deals are Seed-stage. "
        f"Real conversion rates between consecutive stages are computed separately._"
    )
    fig = plot_funding_funnel(df_deals)
    st.pyplot(fig, use_container_width=True)
    download_chart(fig, "4_1_funding_funnel")

    st.markdown('<div class="section-header">4.2 Time Between Rounds</div>',
                unsafe_allow_html=True)
    st.markdown(
        "_Global view and breakdown by stage transition "
        "(Seed→A, A→B, B→C)._"
    )
    fig = plot_time_between_rounds(df_deals)
    st.pyplot(fig, use_container_width=True)
    download_chart(fig, "4_2_time_between_rounds")

    st.markdown('<div class="section-header">4.3 First-Time Fundraisers</div>',
                unsafe_allow_html=True)
    st.markdown(
        "_First deals identify where new ventures are entering the funding ecosystem._"
    )
    fig = plot_first_time_fundraisers(df_deals)
    st.pyplot(fig, use_container_width=True)
    download_chart(fig, "4_3_first_fundraisers")

    st.markdown('<div class="section-header">4.4 Non-Dilutive Financing</div>',
                unsafe_allow_html=True)
    st.markdown(
        "_Grants, debt, and bonds — particularly active in Energy, Healthcare, and Agriculture._"
    )
    fig = plot_non_dilutive(df_deals)
    st.pyplot(fig, use_container_width=True)
    download_chart(fig, "4_4_non_dilutive")

    st.markdown('<div class="section-header">4.5 Startup Survival Signals</div>',
                unsafe_allow_html=True)
    st.markdown(
        "_Startups classified as Active, Slowing, or Dormant based on time since last deal._"
    )
    fig = plot_survival_signals(df_deals)
    st.pyplot(fig, use_container_width=True)
    download_chart(fig, "4_5_survival_signals")


# ════════════════════════════════════════════════════════════
# TAB 5 — FOUNDERS
# ════════════════════════════════════════════════════════════
with tabs[4]:

    st.markdown('<div class="section-header">5.1 Gender Diversity</div>',
                unsafe_allow_html=True)
    st.markdown(
        f"_Female CEO representation stands at **{stats['female_ceo_pct']:.1f}%**. "
        f"Despite an upward trend, representation remains structurally low._"
    )
    fig = plot_gender_diversity(df_deals)
    st.pyplot(fig, use_container_width=True)
    download_chart(fig, "5_1_gender_diversity")

    st.markdown('<div class="section-header">5.2 Academic Background</div>',
                unsafe_allow_html=True)
    st.markdown(
        "_Continent of CEO's university education as a proxy for founder background._"
    )
    fig = plot_academic_background(df_deals)
    st.pyplot(fig, use_container_width=True)
    download_chart(fig, "5_2_academic_background")

    st.markdown('<div class="section-header">5.3 Team Size</div>',
                unsafe_allow_html=True)
    st.markdown(
        "_Solo founders vs co-founding teams on deal volume and average capital raised._"
    )
    fig = plot_team_size(df_deals)
    st.pyplot(fig, use_container_width=True)
    download_chart(fig, "5_3_team_size")

    st.markdown('<div class="section-header">5.4 Y Combinator Effect</div>',
                unsafe_allow_html=True)
    st.markdown(
        f"_**{stats['yc_count']} startups** are YC alumni. "
        f"Their median funding is consistently higher than non-YC peers._"
    )
    fig = plot_yc_effect(df_deals)
    st.pyplot(fig, use_container_width=True)
    download_chart(fig, "5_4_yc_effect")


# ════════════════════════════════════════════════════════════
# TAB 6 — INVESTORS
# ════════════════════════════════════════════════════════════
with tabs[5]:

    st.markdown('<div class="section-header">6.1 Most Active Investors</div>',
                unsafe_allow_html=True)
    st.markdown(
        "_Investors ranked by total deal participation across the full period._"
    )
    fig = plot_top_investors(df_inv_final)
    st.pyplot(fig, use_container_width=True)
    download_chart(fig, "6_1_top_investors")

    st.markdown('<div class="section-header">6.2 Investor Activity Over Time</div>',
                unsafe_allow_html=True)
    st.markdown(
        "_Year-by-year activity of the top 10 investors reveals structural strategy differences._"
    )
    fig = plot_investor_activity_over_time(df_investisseurs, years_cols)
    st.pyplot(fig, use_container_width=True)
    download_chart(fig, "6_2_investor_activity")

    st.markdown('<div class="section-header">6.3 Investor Behavioral Archetypes</div>',
                unsafe_allow_html=True)
    st.markdown(
        "_K-Means clustering (k=4) reveals four investor profiles based on activity patterns._"
    )
    fig = plot_investor_clustering(df_investisseurs, years_cols)
    st.pyplot(fig, use_container_width=True)
    download_chart(fig, "6_3_clustering")

    st.markdown('<div class="section-header">6.4 Capital Origins</div>',
                unsafe_allow_html=True)
    st.markdown(
        "_Investor HQ region as proxy for capital origin — most flows from North America and Europe._"
    )
    fig = plot_capital_origins(df_investisseurs)
    st.pyplot(fig, use_container_width=True)
    download_chart(fig, "6_4_capital_origins")

    st.markdown('<div class="section-header">6.5 Focus: South Africa</div>',
                unsafe_allow_html=True)
    st.markdown(
        "_SA investors profiled across deal volume, capital deployed, and ticket bracket._"
    )
    figs_sa = plot_south_africa_investors(df_inv_final)
    labels_sa = [
        "By Number of Deals",
        "By Total Capital Deployed",
        "By Ticket Bracket"
    ]
    for i, fig in enumerate(figs_sa):
        st.markdown(f"**{labels_sa[i] if i < len(labels_sa) else f'Chart {i+1}'}**")
        st.pyplot(fig, use_container_width=True)
        download_chart(fig, f"6_5_south_africa_{i+1}")


# ════════════════════════════════════════════════════════════
# TAB 7 — SOUTH AFRICA DEEP DIVE
# ════════════════════════════════════════════════════════════
with tabs[6]:

    df_sa, sa_stats = get_sa_data(df_inv_final)

    st.markdown(
        f"_South Africa hosts the most developed domestic VC market, "
        f"with **{sa_stats['Investor'].nunique()} active investors** tracked between "
        f"{df_sa['Year'].min()} and {df_sa['Year'].max()}, deploying "
        f"**${df_sa['Amount_clean'].sum():,.0f}M** across **{len(df_sa):,} deals**._"
    )

    st.markdown('<div class="section-header">1. Ecosystem Overview</div>',
                unsafe_allow_html=True)
    fig = plot_sa_overview(df_sa, sa_stats, df_inv_final)
    st.pyplot(fig, use_container_width=True)
    download_chart(fig, "sa_1_overview")

    st.markdown('<div class="section-header">2. Sector Analysis</div>',
                unsafe_allow_html=True)
    fig = plot_sa_sectors(df_sa)
    st.pyplot(fig, use_container_width=True)
    download_chart(fig, "sa_2_sectors")

    st.markdown('<div class="section-header">3. Ticket Size Analysis</div>',
                unsafe_allow_html=True)
    fig = plot_sa_tickets(df_sa, sa_stats)
    st.pyplot(fig, use_container_width=True)
    download_chart(fig, "sa_3_tickets")

    st.markdown('<div class="section-header">4. Individual Investor Profile</div>',
                unsafe_allow_html=True)
    st.markdown(
        "_Select an investor to see their complete profile, portfolio, and co-investment patterns._"
    )
    investor_list = sorted(sa_stats['Investor'].tolist())
    selected_investor = st.selectbox(
        "Select an investor", investor_list, key='sa_investor_select'
    )
    if selected_investor:
        col_left, col_right = st.columns([2, 1])
        with col_left:
            fig = plot_investor_profile(selected_investor, df_sa, sa_stats)
            st.pyplot(fig, use_container_width=True)
            download_chart(
                fig,
                f"sa_4_profile_{selected_investor.replace(' ', '_')}"
            )
        with col_right:
            st.markdown(f"**📋 Full Portfolio — {selected_investor}**")
            portfolio = get_investor_portfolio_table(selected_investor, df_sa)
            if not portfolio.empty:
                st.dataframe(portfolio, use_container_width=True, height=400)
                st.metric("Total deployed", f"${portfolio['Amount ($M)'].sum():.1f}M")
                st.metric("Total deals", len(portfolio))
                download_df(
                    portfolio,
                    f"portfolio_{selected_investor.replace(' ', '_')}",
                    "⬇️ Download Portfolio"
                )
            else:
                st.info("No portfolio data available.")

    st.markdown('<div class="section-header">5. Co-Investment Galaxy</div>',
                unsafe_allow_html=True)
    st.markdown(
        "_Node size = deal count. Edge thickness = co-investment frequency. Colors = communities._"
    )
    min_shared = st.slider(
        "Minimum shared deals to show a connection",
        min_value=1, max_value=5, value=2, key='sa_network_slider'
    )
    with st.spinner("Building co-investment network..."):
        fig = plot_coinvestment_network(df_sa, min_shared=min_shared)
        st.pyplot(fig, use_container_width=True)
        download_chart(fig, "sa_5_network")

    st.markdown('<div class="section-header">6. Temporal Dynamics</div>',
                unsafe_allow_html=True)
    st.markdown(
        "_New entrants, dormant players, and seasonal deal patterns._"
    )
    fig = plot_sa_temporal(df_sa)
    st.pyplot(fig, use_container_width=True)
    download_chart(fig, "sa_6_temporal")

    st.markdown('<div class="section-header">7. Syndication Patterns</div>',
                unsafe_allow_html=True)
    st.markdown(
        "_Top co-investment pairs and triads, solo vs syndicated styles._"
    )
    fig = plot_sa_syndicates(df_sa)
    st.pyplot(fig, use_container_width=True)
    download_chart(fig, "sa_7_syndicates")

    st.markdown('<div class="section-header">8. Investor Positioning Matrix</div>',
                unsafe_allow_html=True)
    st.markdown(
        "_X = median ticket. Y = deal count. Bubble size = total capital. "
        "Quadrant lines = ecosystem medians._"
    )
    highlight = st.selectbox(
        "Highlight a specific investor (optional)",
        ["None"] + investor_list,
        key='sa_highlight_select'
    )
    hi_inv = None if highlight == "None" else highlight
    fig = plot_sa_positioning_matrix(df_sa, sa_stats, highlight_investor=hi_inv)
    st.pyplot(fig, use_container_width=True)
    download_chart(fig, "sa_8_positioning")

    st.markdown('<div class="section-header">9. Complete Investor Rankings</div>',
                unsafe_allow_html=True)
    rank_tab1, rank_tab2, rank_tab3 = st.tabs([
        "By Deal Count", "By Capital Deployed", "By Ticket Bracket"
    ])
    with rank_tab1:
        df_rank1 = (
            sa_stats.sort_values('Nb_Deals', ascending=False)
            [['Investor','Nb_Deals','Total_Deployed','Median_Ticket','Ticket_Bracket']]
            .round(1).reset_index(drop=True)
        )
        st.dataframe(df_rank1, use_container_width=True)
        download_df(df_rank1, "sa_rank_by_deals")

    with rank_tab2:
        df_rank2 = (
            sa_stats[sa_stats['Total_Deployed'] > 0]
            .sort_values('Total_Deployed', ascending=False)
            [['Investor','Total_Deployed','Nb_Deals','Median_Ticket','Ticket_Bracket']]
            .round(1).reset_index(drop=True)
        )
        st.dataframe(df_rank2, use_container_width=True)
        download_df(df_rank2, "sa_rank_by_capital")

    with rank_tab3:
        bracket_order = [
            '< $1M', '$1M–$5M', '$5M–$15M', '$15M–$50M', '$50M+', 'Undisclosed'
        ]
        for bracket in bracket_order:
            subset = (
                sa_stats[sa_stats['Ticket_Bracket'] == bracket]
                .sort_values('Nb_Deals', ascending=False)
            )
            if not subset.empty:
                st.markdown(f"**{bracket}** — {len(subset)} investors")
                df_bracket = (
                    subset[['Investor','Nb_Deals','Total_Deployed',
                             'Median_Ticket','Avg_Ticket']]
                    .round(1).reset_index(drop=True)
                )
                st.dataframe(df_bracket, use_container_width=True)
                download_df(
                    df_bracket,
                    f"sa_bracket_{bracket.replace(' ', '_').replace('$', '').replace('–', '-')}"
                )


# ════════════════════════════════════════════════════════════
# TAB 8 — BONUS PREDICTIONS
# ════════════════════════════════════════════════════════════
with tabs[7]:

    st.markdown('<div class="section-header">7.1 Optimal Fundraising Window</div>',
                unsafe_allow_html=True)
    st.markdown(
        "_Startups at the statistically optimal moment to raise their next round, "
        "based on real African ecosystem transition timings (P25–P75 per stage)._"
    )
    fig_win, top10_df = plot_fundraising_window(df_deals)
    st.pyplot(fig_win, use_container_width=True)
    download_chart(fig_win, "7_1_fundraising_window")
    st.dataframe(top10_df.round(1), use_container_width=True)
    download_df(top10_df, "7_1_fundraising_window_data")

    st.divider()

    st.markdown('<div class="section-header">7.2 Knife Capital — Portfolio DNA & Similar Deals</div>',
                unsafe_allow_html=True)
    st.markdown(
        "_Knife Capital's portfolio analyzed across stage, sector, ticket, country, and timing. "
        "A cosine similarity model identifies deals most similar to Knife's investment DNA._"
    )
    fig_knife, knife_data = plot_knife_portfolio(df_inv_final)
    st.pyplot(fig_knife, use_container_width=True)
    download_chart(fig_knife, "7_2_knife_portfolio")

    if knife_data is not None:
        st.divider()
        fig_sim, top15_knife = plot_knife_similar_deals(df_inv_final, df_deals)
        st.pyplot(fig_sim, use_container_width=True)
        download_chart(fig_sim, "7_2_knife_similar_deals")
        if top15_knife is not None:
            st.dataframe(top15_knife.round(3), use_container_width=True)
            download_df(top15_knife, "7_2_knife_top15_similar")
st.set_page_config(page_title="Africa deals Dashboard", layout="wide")

st.markdown("""
<style>
    /* Transformation des cartes de métriques (KPIs) */
    [data-testid="metric-container"] {
        background-color: #ffffff;
        border: 1px solid #e0e6ed;
        border-radius: 12px;
        padding: 20px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05); /* Ombre légère */
    }
    [data-testid="metric-container"] label {
        color: #7f8c8d;
        font-size: 0.85rem !important;
        text-transform: uppercase;
        letter-spacing: 1px;
    }
    [data-testid="metric-container"] div[data-testid="stMetricValue"] {
        color: #1F3864;
        font-size: 1.8rem !important;
        font-weight: 700;
    }
    /* Style des en-têtes de section */
    .section-header {
        font-size: 22px;
        font-weight: 700;
        color: #27ae60;
        border-bottom: 2px solid #f0f2f6;
        padding-bottom: 10px;
        margin: 30px 0 20px 0;
    }
    /* Réduit l'espace blanc géant en haut de la page */
    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
    }
    /* Enlève le menu Streamlit et le footer "Made with Streamlit" si tu le souhaites */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

# ── SIDEBAR ───────────────────────────────────────────────────
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/8/8d/Africa_%28orthographic_projection%29.svg/240px-Africa_%28orthographic_projection%29.svg.png", width=80)
    st.title("🌍 Africa Deals")
    st.divider()
    st.subheader("📁 Upload Data")
    uploaded_file = st.file_uploader("Upload Excel database", type=["xlsx"])
    password      = st.text_input("File password", type="password")
    st.divider()
    st.caption("Upload a new Excel file each month to refresh all analyses automatically. Actual file password : B2$wF8^qL5#t")

st.markdown('<div class="main-header">🌍 African Startup Funding Report</div>',
            unsafe_allow_html=True)

if not uploaded_file or not password:
    st.info("👈 Upload your Excel file and enter the password in the sidebar to get started.")
    st.stop()

@st.cache_data(show_spinner="Loading and processing data...")
def load_data(file_bytes, pwd):
    return load_and_clean(io.BytesIO(file_bytes), pwd)

try:
    file_bytes = uploaded_file.read()
    df_deals, df_investisseurs, df_inv_final, years_cols = load_data(file_bytes, password)
    stats = compute_summary_stats(df_deals)
    st.success(f"✅ {len(df_deals):,} deals loaded — {df_deals['Year'].min()} to {df_deals['Year'].max()}")
except Exception as e:
    st.error(f"❌ Error loading file: {e}")
    st.stop()

# ════════════════════════════════════════════════════════════
# CHATBOT WIDGET (GROQ / LLAMA 3)
# ════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("---") 
    
    # Mets ta clé Groq ici
    client = Groq(api_key="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")
    
    with st.popover("💬 AI Assitant", use_container_width=True):
        st.markdown("**Ask your questioon**")
        
        if "chat_history" not in st.session_state:
            st.session_state.chat_history = []
        
        for msg in st.session_state.chat_history:
            st.chat_message(msg["role"]).write(msg["content"])
        
        if user_question := st.chat_input("What do you want ?"):
            st.chat_message("user").write(user_question)
            st.session_state.chat_history.append({"role": "user", "content": user_question})
            
            context = f"""
            You are an expert in African venture capital. You are assiting a user who is looking at a dashboard countaining {len(df_deals)} deals with a total value 
            of ${df_deals['Amount_clean'].sum():.1f}M. Answer this question doncisely, 
            profesionnaly and analytically.
            """
            
            with st.spinner("L'IA réfléchit..."):
                try:
                    chat_completion = client.chat.completions.create(
                        messages=[
                            {"role": "system", "content": context},
                            {"role": "user", "content": user_question}
                        ],
                        model="llama-3.3-70b-versatile", 
                    )
                    
                    response = chat_completion.choices[0].message.content
                    st.chat_message("assistant").write(response)
                    st.session_state.chat_history.append({"role": "assistant", "content": response})
                except Exception as e:
                    st.error(f"Erreur : {e}")
# ── KPI BAR ───────────────────────────────────────────────────
c1,c2,c3,c4,c5,c6 = st.columns(6)
c1.metric("Total Deals",   f"{stats['total_deals']:,}")
c2.metric("Total Capital", f"${stats['total_capital']:,.0f}M")
c3.metric("Countries",     stats['countries'])
c4.metric("Nombre de deals", len(df_deals))
st.divider()


# ── TABS ──────────────────────────────────────────────────────
tabs = st.tabs([
    "Market Overview",
    "Geography",
    "Sectors",
    "Funding Dynamics",
    "Founders",
    "Investors",
    "South Africa Deep Dive",
    "Bonus - Predictions",
])

# ════════════════════════════════════════════════════════════
# TAB 1 — MARKET OVERVIEW
# ════════════════════════════════════════════════════════════
with tabs[0]:
    st.markdown('<div class="section-header">1.1 Deal Landscape</div>',
                unsafe_allow_html=True)
    st.markdown(f"""
    _An overview of the volume, size, and type of deals recorded accross the study period._ 
    _This dashboard provides a high-level snapshot of the dataset across five dimensions: round type, funding bracket, top sectors, top countries, and regional distribution._
    
    """)
    st.pyplot(plot_deal_landscape(df_deals), use_container_width=True)
    download_chart(fig, "1_1_deal_landscape")
    st.markdown('<div class="section-header">Interactive Map</div>', unsafe_allow_html=True)
    m = create_interactive_map(df_deals, df_inv_final)
    st_folium(m, width=1200, height=600, returned_objects=[])

    st.markdown('<div class="section-header">1.2 Market Trajectory (2019–2026)</div>',
                unsafe_allow_html=True)
    st.markdown(f"""
    _Deal activity and total capital are tracked year over year._
    """)
    st.pyplot(plot_market_trajectory(df_deals), use_container_width=True)
    download_chart(fig, "1_2_market_trajectory")

    st.markdown('<div class="section-header">1.3 Capital Distribution by Sector & Region</div>',
                unsafe_allow_html=True)
    st.markdown(f"""
    _Two heatmaps cross-reference the top sectors with African regions.
    The left shows total capital deployed, the right shows median deal size._
    """)
    st.pyplot(plot_heatmaps(df_deals), use_container_width=True)
    download_chart(fig, "1_3_capital_distribution")

    st.markdown('<div class="section-header">1.4 Data Transparency & Climate Tech</div>',
                unsafe_allow_html=True)
    st.markdown(f"""
    _Not all deals disclose their funding amount. This section tracks disclosure rates alongside the growth of Climate Tech tagging. Two transparency metrics are tracked annually: the share of deals with a disclosed funding amount, and the proportion tagged as Climate Tech. Disclosure rates reveal the data quality of the ecosystem, while the Climate Tech trend captures the growing ESG narrative in African venture. Absolute deal counts are annotated to distinguish statistical trends from noise._
    """)
    
    st.pyplot(plot_transparency(df_deals), use_container_width=True)
    download_chart(fig, "1_4_transparency")

# ════════════════════════════════════════════════════════════
# TAB 2 — GEOGRAPHY
# ════════════════════════════════════════════════════════════
with tabs[1]:
    st.markdown('<div class="section-header">2.1 Top Countries</div>',
                unsafe_allow_html=True)
    st.markdown(f"""
    _**{stats['top_country']}** leads both in deal count and capital raised.
    Rankings by volume and by financial value do not always align —
    some countries attract fewer but significantly larger rounds,
    reflecting the presence of institutional rather than early-stage capital._
    """)
    st.pyplot(plot_top_countries(df_deals), use_container_width=True)
    download_chart(fig, "2_1_top_countries")

    st.markdown('<div class="section-header">2.2 Market Concentration</div>',
                unsafe_allow_html=True)
    st.markdown(f"""
    _The Big 4 markets (Nigeria, Kenya, South Africa, Egypt) account for
    **{stats['big4_pct']:.0f}%** of all deals cumulatively since 2019.
    However, the year-by-year trend shows their dominance declining steadily,
    with Emerging Markets growing their share every year since 2020. Note : the pie chart reflects the cumulative distribution over the full period , while the line chart shows year-by-year evolution._
    """)
    st.pyplot(plot_market_concentration(df_deals), use_container_width=True)
    download_chart(fig, "2_2_market_concentration")

    st.markdown('<div class="section-header">2.3 Pan-African vs Country-Specific</div>',
                unsafe_allow_html=True)
    st.markdown("""
    _Startups with an explicit pan-African mandate are compared to country-specific deals
    on both deal count and median capital raised. Pan-African strategies tend to attract
    larger rounds, reflecting the broader addressable market narrative
    and the institutional profile of investors backing continent-wide plays._
    """)
    st.pyplot(plot_pan_african(df_deals), use_container_width=True)
    download_chart(fig, "2_3_pan_african")

# ════════════════════════════════════════════════════════════
# TAB 3 — SECTORS
# ════════════════════════════════════════════════════════════
with tabs[2]:
    st.markdown('<div class="section-header">3.1 Leading Sectors</div>',
                unsafe_allow_html=True)
    st.markdown(f"""
    _**{stats['top_sector']}** leads by a wide margin in deal count.
    However, ranking by total capital reveals different winners —
    capital-intensive sectors like Energy and Infrastructure attract
    fewer but significantly larger individual rounds._
    """)
    st.pyplot(plot_leading_sectors(df_deals), use_container_width=True)
    download_chart(fig, "3_1_leading_sectors")

    st.markdown('<div class="section-header">3.2 Sector Trends Over Time</div>',
                unsafe_allow_html=True)
    st.markdown("""
    _A stacked area chart tracks the relative share of the top 5 sectors from 2019 to 2026.
    Shifts in area indicate structural changes in investor appetite.
    Percentages normalize for overall market growth to highlight real compositional shifts._
    """)
    st.plotly_chart(plot_deal_size_by_sector(df_deals), use_container_width=True)
    download_chart(fig, "3_2_sector_trends")

    st.markdown('<div class="section-header">3.3 Deal Size by Sector</div>',
                unsafe_allow_html=True)
    st.markdown("""
    _Median deal size reveals which sectors attract larger institutional capital,
    independent of deal frequency. Sectors with high medians but few deals
    are typically backed by DFIs or specialized funds rather than generalist VCs._
    """)
    st.plotly_chart(plot_deal_size_by_sector(df_deals), use_container_width=True, key="deal_size_sector_chart")
    download_chart(fig, "3_3_deal_size_sector")

    st.markdown('<div class="section-header">3.4 Funding Stage by Sector</div>',
                unsafe_allow_html=True)
    st.markdown("""
    _A normalized heatmap shows how funding stages distribute within each sector.
    Sectors skewed toward Seed reflect early-stage ecosystems still building foundations,
    while those with significant Series B and C activity signal more mature investment landscapes._
    """)
    st.pyplot(plot_sector_stage(df_deals), use_container_width=True)
    download_chart(fig, "3_4_sector_stage")
    
    st.markdown('<div class="section-header">3.5 Focus: Climate Tech Momentum</div>',
                unsafe_allow_html=True)
    st.markdown("""
    _Climate Tech has emerged as a major narrative in African VC. This section isolates startups explicitly tagged as Climate Tech to observe the volume trajectory and identify which traditional sectors (e.g., Energy, Agriculture, Logistics) are absorbing this green capital._
    """)
    st.pyplot(plot_climate_tech(df_deals), use_container_width=True)

# ════════════════════════════════════════════════════════════
# TAB 4 — FUNDING DYNAMICS
# ════════════════════════════════════════════════════════════
with tabs[3]:
    st.markdown('<div class="section-header">4.1 The Funding Funnel</div>',
                unsafe_allow_html=True)
    st.markdown(f"""
    _The funnel shows unique startups at each stage, from Grants to Series D+.
    **{stats['seed_pct']:.0f}%** of all deals are Seed-stage, highlighting how
    early the ecosystem skews. Real conversion rates between consecutive stages
    are computed separately — the funnel shape reflects structure, not guaranteed progression._
    """)
    st.pyplot(plot_funding_funnel(df_deals), use_container_width=True)
    download_chart(fig, "4_1_funding_funnel")

    st.markdown('<div class="section-header">4.2 Time Between Rounds</div>',
                unsafe_allow_html=True)
    st.markdown("""
    _The first chart shows the overall median time between any two consecutive rounds by sector.
    The three charts below break this down by specific stage transition
    (Seed→Series A, Series A→Series B, Series B→Series C),
    revealing that velocity varies significantly by both sector and stage._
    """)
    st.pyplot(plot_time_between_rounds(df_deals), use_container_width=True)
    download_chart(fig, "4_2_time_between_rounds")

    st.markdown('<div class="section-header">4.3 First-Time Fundraisers</div>',
                unsafe_allow_html=True)
    st.markdown("""
    _First deals identify where new ventures are entering the funding ecosystem.
    A country generating many first deals but few follow-ons signals an early-stage
    ecosystem that has yet to develop the infrastructure for growth-stage capital._
    """)
    st.pyplot(plot_first_time_fundraisers(df_deals), use_container_width=True)
    download_chart(fig, "4_3_first_fundraisers")

    st.markdown('<div class="section-header">4.4 Non-Dilutive Financing</div>',
                unsafe_allow_html=True)
    st.markdown("""
    _Grants, debt, and bonds represent a significant share of African startup funding —
    particularly in Energy, Healthcare, and Agriculture where DFIs are active.
    This analysis isolates non-dilutive instruments to reveal which sectors
    rely most heavily on non-equity capital structures._
    """)
    st.pyplot(plot_non_dilutive(df_deals), use_container_width=True)
    download_chart(fig, "4_4_non_dilutive")

    st.markdown('<div class="section-header">4.5 Startup Survival Signals</div>',
                unsafe_allow_html=True)
    st.markdown("""
    _Startups are classified by time since their last recorded deal:
    Active (under 18 months), Slowing (18 months to 3 years), and Dormant (over 3 years).
    A high proportion of dormant startups in a sector may signal funding gaps
    or structural difficulty in securing follow-on capital._
    """)
    st.pyplot(plot_survival_signals(df_deals), use_container_width=True)
    download_chart(fig, "4_5_survival_signals")

# ════════════════════════════════════════════════════════════
# TAB 5 — FOUNDERS
# ════════════════════════════════════════════════════════════
with tabs[4]:
    st.markdown('<div class="section-header">5.1 Gender Diversity</div>',
                unsafe_allow_html=True)
    st.markdown(f"""
    _Female CEO representation currently stands at **{stats['female_ceo_pct']:.1f}%** across
    the dataset. Despite a consistent upward trend, representation remains structurally low.
    The gap between female CEO rates and mixed-team rates suggests co-founding dynamics
    differ significantly from solo leadership patterns._
    """)
    st.pyplot(plot_gender_diversity(df_deals), use_container_width=True)
    download_chart(fig, "5_1_gender_diversity")

    st.markdown('<div class="section-header">5.2 Academic Background</div>',
                unsafe_allow_html=True)
    st.markdown("""
    _The continent of the CEO's university education is used as a proxy for founder background.
    A boxplot tests whether educational origin correlates with capital raised.
    Locally trained founders show competitive fundraising outcomes in several sectors,
    challenging the assumption that international education is a prerequisite for success._
    """)
    st.pyplot(plot_academic_background(df_deals), use_container_width=True)
    download_chart(fig, "5_2_academic_background")

    st.markdown('<div class="section-header">5.3 Team Size</div>',
                unsafe_allow_html=True)
    st.markdown("""
    _Solo founders are compared to co-founding teams of 2, 3, and 4+ on deal volume
    and average capital raised. Larger teams tend to raise more, likely reflecting
    both complementary skills and increased investor confidence in execution capacity._
    """)
    st.pyplot(plot_team_size(df_deals), use_container_width=True)
    download_chart(fig, "5_3_team_size")

    st.markdown('<div class="section-header">5.4 Y Combinator Effect</div>',
                unsafe_allow_html=True)
    st.markdown(f"""
    _**{stats['yc_count']} startups** in this dataset are Y Combinator alumni.
    Their median funding is consistently higher than non-YC peers within the same sectors.
    The gap is most pronounced in Fintech and Deeptech,
    where the YC network effect and brand recognition are strongest._
    """)
    st.pyplot(plot_yc_effect(df_deals), use_container_width=True)
    download_chart(fig, "5_4_yc_effect")

# ════════════════════════════════════════════════════════════
# TAB 6 — INVESTORS
# ════════════════════════════════════════════════════════════
with tabs[5]:
    st.markdown('<div class="section-header">6.1 Most Active Investors</div>',
                unsafe_allow_html=True)
    st.markdown("""
    _Investors are ranked by total deal participation across the full period,
    based on the exploded investor dataset where each co-investor is counted individually.
    This ranking measures network breadth and ecosystem presence
    rather than fund size or capital deployed._
    """)
    st.pyplot(plot_top_investors(df_inv_final), use_container_width=True)
    download_chart(fig, "6_1_top_investors")

    st.markdown('<div class="section-header">6.2 Investor Activity Over Time</div>',
                unsafe_allow_html=True)
    st.markdown("""
    _Year-by-year activity of the top 10 investors reveals structural differences
    in strategy. Some accelerated post-2020 and maintained consistent activity,
    while others peaked in 2021–2022 and pulled back sharply —
    consistent with global VC contraction patterns following the rate hiking cycle._
    """)
    st.pyplot(plot_investor_activity_over_time(df_investisseurs, years_cols), use_container_width=True)
    download_chart(fig, "6_2_investor_activity")

    st.markdown('<div class="section-header">6.3 Investor Behavioral Archetypes</div>',
                unsafe_allow_html=True)
    st.markdown("""
    _K-Means clustering (k=4) applied to annual deal activity reveals four archetypes:
    occasional one-shot participants (the vast majority), consistent mid-tier players,
    boom-and-retreat funds concentrated in 2021–2022,
    and one dominant late-surge actor accelerating through 2025._
    """)
    st.pyplot(plot_investor_clustering(df_investisseurs, years_cols), use_container_width=True)
    download_chart(fig, "6_3_clustering")

    st.markdown('<div class="section-header">6.4 Capital Origins</div>',
                unsafe_allow_html=True)
    st.markdown("""
    _Investor headquarters region serves as a proxy for capital origin.
    The majority of capital flows from North America and Europe,
    highlighting a structural dependency on foreign capital
    and the relative underdevelopment of domestic institutional investor capacity._
    """)
    st.pyplot(plot_capital_origins(df_investisseurs), use_container_width=True)
    download_chart(fig, "6_4_capital_origins")

    st.markdown('<div class="section-header">6.5 Focus: South Africa</div>',
                unsafe_allow_html=True)
    st.markdown("""
    _South Africa hosts the most developed domestic VC market on the continent.
    Investors are profiled across three dimensions: deal volume, total capital deployed,
    and ticket size bracket. The three rankings rarely overlap,
    revealing distinct archetypes from high-frequency micro-ticket angels
    to selective large-cap institutional funds._
    """)
    figs_sa = plot_south_africa_investors(df_inv_final)
    for fig in figs_sa:
        st.pyplot(fig, use_container_width=True)

# ════════════════════════════════════════════════════════════
# TAB 7 — SOUTH AFRICA DEEP DIVE
# ════════════════════════════════════════════════════════════
with tabs[6]:

    # Préparer les données SA une seule fois
    df_sa, sa_stats = get_sa_data(df_inv_final)

    st.markdown(f"""
    _South Africa hosts the most developed domestic venture capital market on the continent,
    with **{sa_stats['Investor'].nunique()} active investors** tracked between
    {df_sa['Year'].min()} and {df_sa['Year'].max()},
    deploying a total of **${df_sa['Amount_clean'].sum():,.0f}M** across
    **{len(df_sa):,} deals**. This section provides a complete statistical portrait
    of the SA investor ecosystem._
    """)

    # ── Bloc 1 : Overview ──
    st.markdown('<div class="section-header">1. Ecosystem Overview</div>',
                unsafe_allow_html=True)
    st.pyplot(plot_sa_overview(df_sa, sa_stats, df_inv_final))
    download_chart(fig, "sa_1_overview")

    # ── Bloc 2 : Secteurs ──
    st.markdown('<div class="section-header">2. Sector Analysis</div>',
                unsafe_allow_html=True)
    st.pyplot(plot_sa_sectors(df_sa))
    download_chart(fig, "sa_2_sectors")

    # ── Bloc 3 : Tickets ──
    st.markdown('<div class="section-header">3. Ticket Size Analysis</div>',
                unsafe_allow_html=True)
    st.pyplot(plot_sa_tickets(df_sa, sa_stats))
    download_chart(fig, "sa_3_tickets")

    # ── Bloc 4 : Profil individuel ──
    st.markdown('<div class="section-header">4. Individual Investor Profile</div>',
                unsafe_allow_html=True)
    st.markdown("_Select an investor to see their complete profile, portfolio, and co-investment patterns._")

    investor_list = sorted(sa_stats['Investor'].tolist())
    selected_investor = st.selectbox(
        "Select an investor", investor_list, key='sa_investor_select'
    )

    if selected_investor:
        col_left, col_right = st.columns([2, 1])

        with col_left:
            st.pyplot(plot_investor_profile(selected_investor, df_sa, sa_stats))
            download_chart(fig, f"sa_4_profile_{selected_investor.replace(' ', '_')}")

        with col_right:
            st.markdown(f"**📋 Full Portfolio — {selected_investor}**")
            portfolio = get_investor_portfolio_table(selected_investor, df_sa)
            if not portfolio.empty:
                st.dataframe(portfolio, use_container_width=True, height=400)
                total = portfolio['Amount ($M)'].sum()
                st.metric("Total deployed", f"${total:.1f}M")
                st.metric("Total deals", len(portfolio))
            else:
                st.info("No portfolio data available.")

    # ── Bloc 5 : Galaxie réseau ──
    st.markdown('<div class="section-header">5. Co-Investment Galaxy</div>',
                unsafe_allow_html=True)
    st.markdown("_Network of co-investments between SA investors. Node size = deal count. Edge thickness = co-investment frequency. Colors = detected communities._")

    min_shared = st.slider(
        "Minimum shared deals to show a connection",
        min_value=1, max_value=5, value=2, key='sa_network_slider'
    )
    with st.spinner("Building co-investment network..."):
        st.pyplot(plot_coinvestment_network(df_sa, min_shared=min_shared))
        download_chart(fig, "sa_5_network")

    # ── Bloc 6 : Dynamiques temporelles ──
    st.markdown('<div class="section-header">6. Temporal Dynamics</div>',
                unsafe_allow_html=True)
    st.markdown("_Evolution of investor activity over time: new entrants, dormant players, seasonal patterns._")
    st.pyplot(plot_sa_temporal(df_sa))
    download_chart(fig, "sa_6_temporal")

    # ── Bloc 7 : Syndicats ──
    st.markdown('<div class="section-header">7. Syndication Patterns</div>',
                unsafe_allow_html=True)
    st.markdown("_Top co-investment pairs and triads, and classification of investors by syndication style._")
    st.pyplot(plot_sa_syndicates(df_sa))
    download_chart(fig, "sa_7_syndicates")

    # ── Bloc 8 : Matrice positionnement ──
    st.markdown('<div class="section-header">8. Investor Positioning Matrix</div>',
                unsafe_allow_html=True)
    st.markdown("_Each bubble is an investor. X axis = median ticket size. Y axis = number of deals. Bubble size = total capital deployed. Quadrant lines show ecosystem medians._")

    highlight = st.selectbox(
        "Highlight a specific investor (optional)",
        ["None"] + investor_list,
        key='sa_highlight_select'
    )
    hi_inv = None if highlight == "None" else highlight
    st.pyplot(plot_sa_positioning_matrix(df_sa, sa_stats, highlight_investor=hi_inv))
    download_chart(fig, "sa_8_positioning")

    # ── Table récap globale ──
    st.markdown('<div class="section-header">9. Complete Investor Rankings</div>',
                unsafe_allow_html=True)

    rank_tab1, rank_tab2, rank_tab3 = st.tabs([
        "By Deal Count", "By Capital Deployed", "By Ticket Bracket"
    ])

    with rank_tab1:
        st.dataframe(
            sa_stats.sort_values('Nb_Deals', ascending=False)
            [['Investor','Nb_Deals','Total_Deployed','Median_Ticket','Ticket_Bracket']]
            .round(1).reset_index(drop=True),
            use_container_width=True
        )

    with rank_tab2:
        st.dataframe(
            sa_stats[sa_stats['Total_Deployed']>0]
            .sort_values('Total_Deployed', ascending=False)
            [['Investor','Total_Deployed','Nb_Deals','Median_Ticket','Ticket_Bracket']]
            .round(1).reset_index(drop=True),
            use_container_width=True
        )

    with rank_tab3:
        bracket_order = ['< $1M','$1M–$5M','$5M–$15M','$15M–$50M','$50M+','Undisclosed']
        for bracket in bracket_order:
            subset = sa_stats[sa_stats['Ticket_Bracket']==bracket].sort_values('Nb_Deals', ascending=False)
            if not subset.empty:
                st.markdown(f"**{bracket}** — {len(subset)} investors")
                st.dataframe(
                    subset[['Investor','Nb_Deals','Total_Deployed','Median_Ticket','Avg_Ticket']]
                    .round(1).reset_index(drop=True),
                    use_container_width=True
                )
# ════════════════════════════════════════════════════════════
# TAB 8 — INVESTMENT INTELLIGENCE
# ════════════════════════════════════════════════════════════
with tabs[7]:
    

    st.markdown('<div class="section-header">7.1 Optimal Fundraising Window</div>',
                unsafe_allow_html=True)
    st.markdown("""
    _Startups at the statistically optimal moment to raise their next round,
    based on real African ecosystem transition timings (P25–P75 window per stage).
    The left chart shows the reference windows; the right shows where each
    startup currently sits relative to its stage-specific window._
    """)
    fig_win, top10_df = plot_fundraising_window(df_deals)
    st.pyplot(fig_win, use_container_width=True)
    st.dataframe(top10_df.round(1), use_container_width=True)

    st.divider()

    st.markdown('<div class="section-header">7.2 Knife Capital — Portfolio DNA & Similar Deals</div>',
                unsafe_allow_html=True)
    st.markdown("""
    _Knife Capital's historical investment portfolio is analyzed across stage, sector,
    ticket size, country, and deployment timing. A cosine similarity model then scans
    the full dataset to identify deals that most closely match Knife's investment DNA —
    filtered to startups currently in their fundraising window._
    """)
    fig_knife, knife_data = plot_knife_portfolio(df_inv_final)
    st.pyplot(fig_knife, use_container_width=True)

    if knife_data is not None:
        st.divider()
        fig_sim, top15_knife = plot_knife_similar_deals(df_inv_final, df_deals)
        st.pyplot(fig_sim, use_container_width=True)
        if top15_knife is not None:
            st.dataframe(top15_knife.round(3), use_container_width=True)

