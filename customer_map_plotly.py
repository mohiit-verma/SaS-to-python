import plotly.graph_objects as go
import pandas as pd

# ── LOAD YOUR DATA ─────────────────────────────────────────────────────────────
df = pd.read_csv("your_file.csv")
# Expected columns: latitude, longitude, group  (group values: 1, 2, 3)

# ── GROUP SETTINGS: color, bubble size, label ──────────────────────────────────
group_config = {
    1: {"color": "#E24B4A", "size": 18, "label": "Group 1"},   # Red   — large
    2: {"color": "#EF9F27", "size": 12, "label": "Group 2"},   # Amber — medium
    3: {"color": "#85B7EB", "size": 7,  "label": "Group 3"},   # Light blue — small
}

# ── BUILD ONE TRACE PER GROUP (keeps legend clean) ────────────────────────────
fig = go.Figure()

for grp, cfg in group_config.items():
    subset = df[df["group"] == grp]
    fig.add_trace(go.Scattergeo(
        lat=subset["latitude"],
        lon=subset["longitude"],
        mode="markers",
        name=cfg["label"],
        marker=dict(
            size=cfg["size"],
            color=cfg["color"],
            opacity=0.85,
            line=dict(width=0.5, color="white"),
        ),
        hovertemplate=(
            f"<b>{cfg['label']}</b><br>"
            "Lat: %{lat:.4f}<br>"
            "Lon: %{lon:.4f}<br>"
            "<extra></extra>"
        ),
    ))

# ── MAP LAYOUT ────────────────────────────────────────────────────────────────
fig.update_layout(
    title="Customer Distribution Across the US",
    geo=dict(
        scope="usa",
        projection_type="albers usa",
        showland=True,
        landcolor="rgb(230, 230, 220)",
        showlakes=True,
        lakecolor="rgb(200, 220, 240)",
        showcoastlines=True,
        coastlinecolor="white",
        showframe=False,
    ),
    legend_title_text="Group",
    margin=dict(l=0, r=0, t=40, b=0),
    height=550,
)

fig.show()                      # opens in browser
# fig.write_html("customer_map.html")   # uncomment to save as HTML file
