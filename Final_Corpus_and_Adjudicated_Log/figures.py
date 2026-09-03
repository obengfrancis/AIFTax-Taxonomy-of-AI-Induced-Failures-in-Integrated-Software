import json, os
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.gridspec as gridspec
from matplotlib.patches import FancyArrowPatch
import numpy as np
from collections import Counter, OrderedDict
from matplotlib.colors import BoundaryNorm

HERE = os.path.dirname(os.path.abspath(__file__))
with open(os.path.join(HERE, 'clean_data.json')) as f:
    data = json.load(f)

FIGDIR = os.path.join(HERE, 'figures')
os.makedirs(FIGDIR, exist_ok=True)

# ── Shared style ──────────────────────────────────────────────────────────────
plt.rcParams.update({
    'font.family':      'DejaVu Sans',
    'font.size':        9,
    'axes.titlesize':   10,
    'axes.titleweight': 'bold',
    'axes.spines.top':  False,
    'axes.spines.right':False,
    'figure.dpi':       200,
    'savefig.dpi':      300,
    'savefig.bbox':     'tight',
    'savefig.pad_inches': 0.05,
})

N = len(data)

# ── Colour palettes ───────────────────────────────────────────────────────────
# Primary categories
CAT_COLORS = {
    'Operational':    '#2166AC',   # blue
    'Distributional': '#D6604D',   # red-orange
    'Adversarial':    '#4DAC26',   # green
    'Mixed/Hybrid':   '#8073AC',   # purple
}

# Risk severity  (ordered dark → light, critical is darkest)
RISK_COLORS = {
    'Critical': '#67000D',
    'Severe':   '#D73027',
    'High':     '#F46D43',
    'Moderate': '#FDAE61',
    'Low':      '#A6D96A',
}
RISK_ORDER = ['Critical','Severe', 'High', 'Moderate','Low']

# Sector colours
SECTOR_COLORS = {
    'Healthcare & Medicine':          '#1B7837',
    'Government & Public Services':   '#2166AC',
    'Social Media & Platforms':       '#8073AC',
    'Transportation & Automotive':    '#D6604D',
    'Technology & AI Services':       '#4393C3',
    'Security & Law Enforcement':     '#B2182B',
    'Finance, Retail & Commerce':     '#F46D43',
    'Education & Workforce':          '#74ADD1',
}

# ═══════════════════════════════════════════════════════════════════════════════
# FIGURE 1 — Failure Category Distribution  (primary + subcategories, 2-panel)
# ═══════════════════════════════════════════════════════════════════════════════
fig, axes = plt.subplots(1, 2, figsize=(10, 4.5))

# ── Panel (a): Primary categories ─────────────────────────────────────────────
ax = axes[0]
primary_order = ['Operational','Distributional','Adversarial','Mixed/Hybrid']
prim_counts = Counter(r['primary_category'] for r in data)
vals   = [prim_counts[k] for k in primary_order]
colors = [CAT_COLORS[k] for k in primary_order]
labels = [f"{k}\n({prim_counts[k]}, {prim_counts[k]/N*100:.0f}%)" for k in primary_order]

wedges, texts = ax.pie(vals, colors=colors, startangle=90,
                        wedgeprops=dict(width=0.62, edgecolor='white', linewidth=2),
                        pctdistance=0.75)
#ax.set_title('(a) Primary Failure Categories (n=100)', pad=12, fontsize=11)

# Custom legend outside
legend_handles = [mpatches.Patch(color=CAT_COLORS[k],
                  label=f"{k}  ({prim_counts[k]}, {prim_counts[k]/N*100:.0f}%)")
                  for k in primary_order]
ax.legend(handles=legend_handles, loc='lower center',
          bbox_to_anchor=(0.5, -0.18), ncol=1, fontsize=11,
          frameon=False)

# ── Panel (b): Subcategory horizontal bar ─────────────────────────────────────
ax2 = axes[1]

sub_counts = Counter((r['primary_category'], r['subcategory']) for r in data)
# Build ordered list: group by primary, sort by count within group
groups = []
for prim in primary_order:
    subs = [(s, sub_counts[(prim,s)]) for p,s in sub_counts if p == prim]
    subs.sort(key=lambda x: -x[1])
    for s, c in subs:
        groups.append((prim, s, c))

y_labels = [f"{s}" for (p,s,c) in groups]
y_vals   = [c for (p,s,c) in groups]
y_cols   = [CAT_COLORS[p] for (p,s,c) in groups]
y_pos    = list(range(len(groups)))

bars = ax2.barh(y_pos, y_vals, color=y_cols, edgecolor='white', linewidth=0.8, height=0.7)
for bar, val in zip(bars, y_vals):
    ax2.text(bar.get_width() + 0.3, bar.get_y() + bar.get_height()/2,
             str(val), va='center', ha='left', fontsize=11, color='#333333')

ax2.set_yticks(y_pos)
ax2.set_yticklabels(y_labels, fontsize=11)
ax2.set_xlim(0, max(y_vals) + 4)
ax2.set_xlabel('Number of Cases', fontsize=11)
#ax2.set_title('(b) Subcategory Breakdown', pad=12, fontsize=11)
ax2.invert_yaxis()

# # Add primary-category group labels on the right axis
# current_p, group_start = None, 0
# for i, (p, s, c) in enumerate(groups):
#     if p != current_p:
#         if current_p is not None:
#             mid = (group_start + i - 1) / 2
#             ax2.annotate(current_p, xy=(1.01, mid/(len(groups)-1)),
#                          xycoords=('axes fraction','axes fraction'),
#                          fontsize=10, color=CAT_COLORS[current_p],
#                          fontweight='bold', va='center',
#                          rotation=90 if len(current_p) > 8 else 0)
#         current_p, group_start = p, i
# # last group
# mid = (group_start + len(groups) - 1) / 2
# ax2.annotate(current_p, xy=(1.01, mid/(len(groups)-1)),
#              xycoords=('axes fraction','axes fraction'),
#              fontsize=10, color=CAT_COLORS[current_p],
#              fontweight='bold', va='center')

# Separator lines between primary groups
for i in range(1, len(groups)):
    if groups[i][0] != groups[i-1][0]:
        ax2.axhline(i - 0.5, color='#cccccc', linewidth=0.8, linestyle='--')

plt.tight_layout(w_pad=3)
plt.savefig(os.path.join(FIGDIR, 'fig1_failure_categories.pdf'))
plt.savefig(os.path.join(FIGDIR, 'fig1_failure_categories.png'))
plt.close()
print("✓ Figure 1 saved")

# ═══════════════════════════════════════════════════════════════════════════════
# FIGURE 2 — Risk Severity + Year Trend  (2-panel)
# ═══════════════════════════════════════════════════════════════════════════════
fig, axes = plt.subplots(1, 2, figsize=(11, 4.2))

# ── Panel (a): Risk severity donut ────────────────────────────────────────────
ax = axes[0]
risk_counts = Counter(r['risk'] for r in data)
vals   = [risk_counts[k] for k in RISK_ORDER]
colors = [RISK_COLORS[k] for k in RISK_ORDER]

wedges, texts = ax.pie(vals, colors=colors, startangle=90,
                        wedgeprops=dict(width=0.58, edgecolor='white', linewidth=2))

# Centre annotation: calculate the headline from the normalized data.
# "High or worse" combines Critical, Severe, and High while preserving
# the full five-level distribution in the legend.
high_or_worse = sum(
    risk_counts.get(level, 0)
    for level in ['Critical', 'Severe', 'High']
)
high_or_worse_pct = high_or_worse / N * 100 if N else 0

ax.text(0, 0, f"{high_or_worse_pct:.0f}%\nHigh or worse",
        ha='center', va='center',
        fontsize=11, fontweight='bold', color='#67000D')
#ax.set_title('(a) Risk Severity Distribution (n=100)', pad=12, fontsize=11, )

legend_handles = [mpatches.Patch(color=RISK_COLORS[k],
                  label=f"{k}  ({risk_counts[k]}, {risk_counts[k]/N*100:.0f}%)")
                  for k in RISK_ORDER]
ax.legend(handles=legend_handles, loc='lower center',
          bbox_to_anchor=(0.5, -0.18), ncol=2, fontsize=11, frameon=False)

# ── Panel (b): Stacked bar by year, coloured by primary category ──────────────
ax2 = axes[1]
years = sorted(set(r['year'] for r in data))
cats  = ['Operational','Distributional','Adversarial','Mixed/Hybrid']

year_cat = {y: Counter() for y in years}
for r in data:
    year_cat[r['year']][r['primary_category']] += 1

bottoms = np.zeros(len(years))
for cat in cats:
    vals = [year_cat[y][cat] for y in years]
    ax2.bar(years, vals, bottom=bottoms,
            color=CAT_COLORS[cat], edgecolor='white', linewidth=0.6,
            label=cat, width=0.65)
    bottoms += np.array(vals)

# Total labels on top of each bar
totals = [sum(year_cat[y].values()) for y in years]
for i, (year, tot) in enumerate(zip(years, totals)):
    ax2.text(i, tot + 0.3, str(tot), ha='center', va='bottom', fontsize=11)

ax2.set_xticks(range(len(years)))
ax2.set_xticklabels(years, rotation=0)
ax2.set_ylabel('Number of Cases', fontsize=11)
ax2.set_xlabel('Year', fontsize=11)
#ax2.set_title('(b) Year-over-Year Trend by Failure Category', pad=12, fontsize=11)
ax2.legend(loc='upper left', bbox_to_anchor=(0.0, 1.05), fontsize=11, frameon=False)

plt.tight_layout(w_pad=3)
plt.savefig(os.path.join(FIGDIR, 'fig2_risk_trend.pdf'))
plt.savefig(os.path.join(FIGDIR, 'fig2_risk_trend.png'))
plt.close()
print("✓ Figure 2 saved")

# ═══════════════════════════════════════════════════════════════════════════════
# FIGURE 3 — Sector Distribution with Risk Heatmap
# ═══════════════════════════════════════════════════════════════════════════════
fig, axes = plt.subplots(1, 2, figsize=(13, 4.5))

# ── Panel (a): Sector bar chart ───────────────────────────────────────────────
ax = axes[0]
sector_counts = Counter(r['sector'] for r in data)
sector_order  = [s for s,_ in sector_counts.most_common()]

vals   = [sector_counts[s] for s in sector_order]
colors = [SECTOR_COLORS.get(s,'#999999') for s in sector_order]
short  = [s.replace(' & ',' &\n') for s in sector_order]

bars = ax.barh(short, vals, color=colors, edgecolor='white', linewidth=0.8, height=0.7)
for bar, val in zip(bars, vals):
    ax.text(bar.get_width() + 0.2, bar.get_y() + bar.get_height()/2,
            str(val), va='center', ha='left', fontsize=11)

ax.set_xlim(0, max(vals) + 4)
ax.set_xlabel('Number of Cases', fontsize=11)
#ax.set_title('(a) Cases by Industry Sector', pad=12, fontsize=11)
ax.invert_yaxis()

# ── Panel (b): Risk heatmap: sector × severity ────────────────────────────────
ax2 = axes[1]
mat = np.zeros((len(sector_order), len(RISK_ORDER)), dtype=int)
for r in data:
    si = sector_order.index(r['sector'])
    ri = RISK_ORDER.index(r['risk'])
    mat[si, ri] += 1

im = ax2.imshow(mat, cmap='YlOrRd', aspect='auto')
ax2.set_xticks(range(len(RISK_ORDER)))
ax2.set_xticklabels(RISK_ORDER, fontsize=11)
ax2.set_yticks(range(len(sector_order)))
ax2.set_yticklabels([s.replace(' & ','\n& ') for s in sector_order], fontsize=11)
#ax2.set_title('(b) Risk Severity by Sector', pad=12, fontsize=11)

# Annotate cells
for i in range(len(sector_order)):
    for j in range(len(RISK_ORDER)):
        if mat[i,j] > 0:
            ax2.text(j, i, str(mat[i,j]), ha='center', va='center',
                     fontsize=10, color='black' if mat[i,j] < 5 else 'white',
                     fontweight='bold' if mat[i,j] >= 3 else 'normal')

plt.colorbar(im, ax=ax2, shrink=0.7, label='Cases')
plt.tight_layout(w_pad=3)
plt.savefig(os.path.join(FIGDIR, 'fig3_sector_risk.pdf'),
            dpi=300)
plt.savefig(os.path.join(FIGDIR, 'fig3_sector_risk.png')) 
          
plt.close()
print("✓ Figure 3 saved")

# ═══════════════════════════════════════════════════════════════════════════════
# FIGURE 4 — Recovery Complexity Distribution
# ═══════════════════════════════════════════════════════════════════════════════
fig, ax = plt.subplots(figsize=(6.5, 3.5))

rec_order  = ['Very High','Severe','High','Moderate','Low']
rec_colors = ['#67000D','#D73027','#F46D43','#FDAE61','#A6D96A']
rec_counts = Counter(r['recovery'] for r in data)

vals = [rec_counts.get(k,0) for k in rec_order]
bars = ax.bar(rec_order, vals, color=rec_colors, edgecolor='white', linewidth=1, width=0.6)

for bar, val in zip(bars, vals):
    pct = val/N*100
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.4,
            f"{val}\n({pct:.0f}%)", ha='center', va='bottom', fontsize=9)

# Bracket: Very High + Severe + High = "difficult recovery"
difficult = sum(rec_counts.get(k,0) for k in ['Very High','Severe','High'])
ax.annotate('', xy=(-0.4, max(vals)+5.5), xytext=(2.4, max(vals)+5.5),
            arrowprops=dict(arrowstyle='-', color='#444444', lw=1.2))
ax.text(1.0, max(vals)+6.2,
        f"High-or-worse recovery: {difficult} cases ({difficult/N*100:.0f}%)",
        ha='center', fontsize=10, color='#444444')

ax.set_ylabel('Number of Cases', fontsize=11)
ax.set_xlabel('Recovery Complexity Level', fontsize=11)
#ax.set_title('Recovery Complexity Distribution (n=100)', pad=10, fontweight='bold', fontsize=10)
ax.set_ylim(0, max(vals) + 10)
ax.set_yticks(range(0, max(vals)+2, 5))

plt.tight_layout()
plt.savefig(os.path.join(FIGDIR, 'fig4_recovery.pdf'))
plt.savefig(os.path.join(FIGDIR, 'fig4_recovery.png'))
plt.close()
print("✓ Figure 4 saved")


# ═══════════════════════════════════════════════════════════════════════════════
# FIGURE 5 — AI Failure Propagation Pattern Diagram
# ═══════════════════════════════════════════════════════════════════════════════
fig, ax = plt.subplots(figsize=(13, 6))
ax.set_xlim(-0.05, 11.09)          # tighter — content runs ~0.0 to 11.8
ax.set_ylim(1.0, 4.8)           # tighter — cuts dead space above and below
ax.axis('off')

stages = [
    ("Stage 1\n(Origin)",       "AI Component\nFailure",
     "Model drift, hallucination,\ncapability limitation, adversarial\ninput, or bias activation",
     1.2, '#2166AC'),
    ("Stage 2\n(Propagation)",  "Undetected\nPropagation",
     "Plausible or structurally valid\noutput masks error; downstream\nsoftware may not intercept\nAI failure",
     4.1, '#D73027'),
    ("Stage 3\n(Amplification)","System-Wide\nImpact",
     "Cascading failures across\nintegrated components;\nresource misallocation",
     7.0, '#F46D43'),
    ("Stage 4\n(Consequence)",  "User Harm or\nOperational Impact",
     "Safety incidents, financial loss,\ndiscrimination, service disruption,\nloss of public trust",
     9.9, '#67000D'),
]

box_w, box_h = 1.9, 1.2     #2.25, 1.55
for (stage_label, title, desc, x, col) in stages:
    # Main coloured box
    rect = mpatches.FancyBboxPatch((x - box_w/2, 2.775), box_w, box_h,
                                    boxstyle="round,pad=0.08",
                                    facecolor=col, edgecolor='white', linewidth=1.5,
                                    zorder=3)
    ax.add_patch(rect)
    ax.text(x, 3.37, title, ha='center', va='center',
            color='white', fontsize=13, fontweight='bold', zorder=4,
            linespacing=1.3)

    # Stage label above box
    ax.text(x, 4.20, stage_label, ha='center', va='center',
            color=col, fontsize=11, fontweight='bold',
            bbox=dict(boxstyle='round,pad=0.3', facecolor='white',
                      edgecolor=col, linewidth=1.2))

    # Description box below
    desc_box = mpatches.FancyBboxPatch((x - box_w/2, 1.55), box_w, 1.0,
                                        boxstyle="round,pad=0.06",
                                        facecolor='#F7F7F7', edgecolor=col,
                                        linewidth=1.0, zorder=2)
    ax.add_patch(desc_box)
    ax.text(x, 2.05, desc, ha='center', va='center',
            color="#070707", fontsize=10, zorder=3, linespacing=1.4)

# Arrows between stages
arrow_xs = [
    (2.20, 3.10),  # Stage 1 -> Stage 2
    (5.10, 6.00),  # Stage 2 -> Stage 3
    (8.00, 8.90),  # Stage 3 -> Stage 4
]

arrow_labels = [
    "silent boundary\ntransfer",
    "amplification",
    "consequence",
]

for (x1, x2), label in zip(arrow_xs, arrow_labels):
    ax.annotate(
        '',
        xy=(x2, 3.38),
        xytext=(x1, 3.38),
        arrowprops=dict(
            arrowstyle='->',
            color="#282626",
            lw=2.2,
            mutation_scale=20
        )
    )

    ax.text(
        (x1 + x2) / 2,
        3.62,
        label,
        ha='center',
        va='bottom',
        fontsize=9,
        color="#080808",
        style='italic'
    )
# # Arrows between stages
# arrow_xs =   [(2.20, 3.10), (5.10, 6.00), (8.00, 8.90)] #[(2.41, 2.95), (5.31, 5.90), (8.20, 8.75)]
# # arrow_xs = [(2.41, 3.00), (5.41, 6.00), (8.41, 9.00)]
# for (x1, x2) in arrow_xs:
#     ax.annotate('', xy=(x2, 3.38), xytext=(x1, 3.38),
#                 arrowprops=dict(arrowstyle='->', color="#282626",
#                                 lw=2.2, mutation_scale=20))
#     ax.text((x1+x2)/2, 3.62, "silent\ntransfer", ha='center', va='bottom',
#             fontsize=9, color="#080808", style='italic')
    

# Detection gap callout
ax.annotate('', xy=(5.0, 2.9), xytext=(2.9, 2.5),
            arrowprops=dict(arrowstyle='->', color='#D73027', lw=1.4,
                            connectionstyle='arc3,rad=-0.2'))
ax.text(2.5, 2.4,
        "94%  silent transfer mode",
        fontsize=11, color='#D73027', ha='center', style='italic',
        bbox=dict(boxstyle='round,pad=0.3', facecolor='#FFF5F0',
                  edgecolor='#D73027', linewidth=0.9))

# Title
#ax.text(5.9, 5.72,
#        "Recurring AI Failure Propagation Pattern (observed across 100 cases)",
#        ha='center', va='center', fontsize=13, fontweight='bold', color='#222222')

plt.tight_layout(pad=0.3)
plt.savefig(os.path.join(FIGDIR, 'fig5_propagation.pdf'),
            dpi=300, bbox_inches='tight', pad_inches=0.1)
plt.savefig(os.path.join(FIGDIR, 'fig5_propagation.png'),
            dpi=300, bbox_inches='tight', pad_inches=0.1)
plt.close()
print("✓ Figure 5 saved")

print(f"\n✓ All 5 figures generated in {FIGDIR}")