"""
MOVITEC S.A. Detección de Problemas e Ineficiencias

Aplica z-score multivariado, IQR, y K-Means para segmentar técnicos
y equipos problemáticos.
"""

import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.gridspec import GridSpec
from matplotlib.ticker import FuncFormatter
from sklearn.preprocessing import StandardScaler
from sklearn.cluster import KMeans
from sklearn.decomposition import PCA
from scipy import stats
import warnings, os
warnings.filterwarnings('ignore')

# Paleta
C = {
    'azul_osc': '#1A3A5C', 'azul_med': '#2E6DA4', 'azul_clar': '#D6E4F0',
    'rojo': '#C0392B', 'verde': '#1E8449', 'naranja': '#E67E22',
    'gris': '#7F8C8D', 'amarillo': '#F39C12', 'morado': '#7D3C98',
}
TIPO_COL = {'Eléctrica':'#2E86AB','Mecánica':'#E84855','Preventiva':'#3BB273',
            'Correctiva':'#E67E22','Emergencia':'#8B1A1A'}

plt.rcParams.update({
    'font.family':'DejaVu Sans','axes.spines.top':False,'axes.spines.right':False,
    'axes.grid':True,'grid.alpha':0.3,'grid.linestyle':'--',
    'figure.facecolor':'#FAFAFA','axes.facecolor':'#FFFFFF',
    'axes.titlesize':11,'axes.titleweight':'bold','axes.titlepad':10,
})
miles = FuncFormatter(lambda x,_: f'${x/1_000:.0f}K')
os.makedirs('C:/Users/Axel/Downloads/movitec/graficos', exist_ok=True)

# Carga de datos
df = pd.read_excel(
    r'C:/Users/Axel/Downloads/Movitec S.A/ordenes_trabajo.xlsx')

# limpiar columnas
df.columns = (
    df.columns
    .str.strip()
    .str.lower()
)

print("COLUMNAS OT:")
print(df.columns.tolist())

# convertir fechas
df['fecha_creacion'] = pd.to_datetime(df['fecha_creacion'])
df['fecha_cierre'] = pd.to_datetime(df['fecha_cierre'])
df_tec = pd.read_excel(r'C:/Users/Axel/Downloads/Movitec S.A/tecnicos.xlsx')
df_c   = df[df['estado']=='Cerrada'].copy()


# SCORECARD EXTENDIDO DE TÉCNICOS

df_c = df_c.merge(df_tec[['id tecnico','nombre','Años de experiencia','especialidad']], on='id tecnico')

scorecard = df_c.groupby('nombre').agg(
    total_ot        =('id orden',                'count'),
    horas_prom      =('horas trabajadas',        'mean'),
    horas_std       =('horas trabajadas',        'std'),
    costo_prom      =('costo_total',             'mean'),
    costo_std       =('costo_total',             'std'),
    pct_retrabajo   =('retrabajo',               lambda x: (x=='Sí').mean()*100),
    mttr_prom       =('tiempo resolucion horas', 'mean'),
    mttr_std        =('tiempo resolucion horas', 'std'),
    exp             =('Años de experiencia',        'first'),
    especialidad    =('especialidad',            'first'),
).reset_index()

# Score compuesto normalizado
for col, w in [('horas_prom',0.30),('costo_prom',0.40),('pct_retrabajo',0.30)]:
    mx, mn = scorecard[col].max(), scorecard[col].min()
    scorecard[f'{col}_n'] = (scorecard[col]-mn)/(mx-mn+1e-9)
scorecard['inef_score'] = (scorecard['horas_prom_n']*0.30 +
                            scorecard['costo_prom_n']*0.40 +
                            scorecard['pct_retrabajo_n']*0.30)*100

# Índice costo-experiencia (alto exp + alto costo = anomalía)
scorecard['roi_experiencia'] = scorecard['costo_prom'] / (scorecard['exp'] + 1)


# CLUSTERING K-MEANS (TÉCNICOS)

features_tec = ['horas_prom','costo_prom','pct_retrabajo','mttr_prom','exp']
X_tec = scorecard[features_tec].fillna(0)
scaler_tec = StandardScaler()
X_scaled   = scaler_tec.fit_transform(X_tec)

kmeans = KMeans(n_clusters=3, random_state=42, n_init=20)
scorecard['cluster'] = kmeans.fit_predict(X_scaled)

# Etiquetar clusters por ineficiencia media
cluster_inef = scorecard.groupby('cluster')['inef_score'].mean().sort_values()
labels_map   = {cluster_inef.index[0]: 'Eficiente',
                cluster_inef.index[1]: 'Intermedio',
                cluster_inef.index[2]: 'Crítico'}
scorecard['cluster_label'] = scorecard['cluster'].map(labels_map)
CLUSTER_COL = {'Eficiente': C['verde'], 'Intermedio': C['naranja'], 'Crítico': C['rojo']}

# PCA para visualización 2D
pca = PCA(n_components=2, random_state=42)
coords = pca.fit_transform(X_scaled)
scorecard['pca1'] = coords[:,0]
scorecard['pca2'] = coords[:,1]


# ANÁLISIS DE EQUIPOS PROBLEMÁTICOS

equipo_stats = df.groupby('equipo').agg(
    total_ot        =('id orden',         'count'),
    costo_total     =('costo_total',      'sum'),
    costo_prom      =('costo_total',      'mean'),
    pct_retrabajo   =('retrabajo',        lambda x: (x=='Sí').mean()*100),
    horas_prom      =('horas trabajadas', 'mean'),
    n_emergencias   =('tipo falla',       lambda x: (x=='Emergencia').sum()),
    n_correctivas   =('tipo falla',       lambda x: (x=='Correctiva').sum()),
).reset_index()

equipo_stats['pct_no_planif'] = (equipo_stats['n_emergencias']+equipo_stats['n_correctivas'])/equipo_stats['total_ot']*100
equipo_stats['pct_costo']     = equipo_stats['costo_total']/equipo_stats['costo_total'].sum()*100
equipo_stats['costo_acum']    = equipo_stats.sort_values('costo_total',ascending=False)['costo_total'].cumsum()/equipo_stats['costo_total'].sum()*100

# Risk score equipos (frecuencia + costo + retrabajo + fallas no planif)
for col, w in [('total_ot',0.20),('costo_prom',0.30),('pct_retrabajo',0.25),('pct_no_planif',0.25)]:
    mx,mn = equipo_stats[col].max(), equipo_stats[col].min()
    equipo_stats[f'{col}_n'] = (equipo_stats[col]-mn)/(mx-mn+1e-9)
equipo_stats['risk_score'] = (equipo_stats['total_ot_n']*0.20 + equipo_stats['costo_prom_n']*0.30 +
                               equipo_stats['pct_retrabajo_n']*0.25 + equipo_stats['pct_no_planif_n']*0.25)*100
#
# OUTLIERS MULTIVARIADOS (Mahalanobis)

cols_mv = ['horas trabajadas','costo_total','tiempo resolucion horas']
df_mv   = df_c[cols_mv].dropna()
mu_vec  = df_mv.mean().values
cov_mat = np.cov(df_mv.values.T)
try:
    inv_cov = np.linalg.inv(cov_mat)
    diffs   = df_mv.values - mu_vec
    mah_sq  = np.array([d @ inv_cov @ d for d in diffs])
    df_c.loc[df_mv.index, 'mahal_sq'] = mah_sq
    umbral_mah = stats.chi2.ppf(0.975, df=3)
    df_c['outlier_mah'] = df_c['mahal_sq'] > umbral_mah
    n_mah = df_c['outlier_mah'].sum()
except:
    df_c['outlier_mah'] = False
    n_mah = 0


# DETECCIÓN DE PROBLEMAS (4 paneles)

fig6 = plt.figure(figsize=(18,14), facecolor='#FAFAFA')
fig6.suptitle('MOVITEC S.A. — Fase 5: Detección de Ineficiencias y Problemas Operativos',
              fontsize=15, fontweight='bold', color=C['azul_osc'], y=0.98)
gs6 = GridSpec(2,2, figure=fig6, hspace=0.50, wspace=0.40)

# Clustering PCA técnicos 
ax1 = fig6.add_subplot(gs6[0,0])
for label, grp in scorecard.groupby('cluster_label'):
    ax1.scatter(grp['pca1'], grp['pca2'],
                c=CLUSTER_COL[label], label=label,
                s=grp['total_ot']*3.5, alpha=0.85, edgecolors='white', linewidth=0.8)
    for _, r in grp.iterrows():
        ax1.annotate(r['nombre'].split()[0], (r['pca1'], r['pca2']),
                     textcoords='offset points', xytext=(5,3), fontsize=7.5,
                     color=CLUSTER_COL[label], fontweight='bold')
ax1.set_title('K-Means Clustering de Técnicos\n(PCA 2D · burbuja = volumen OT)')
ax1.set_xlabel(f'PC1 ({pca.explained_variance_ratio_[0]*100:.0f}% var)')
ax1.set_ylabel(f'PC2 ({pca.explained_variance_ratio_[1]*100:.0f}% var)')
leyenda_cl = [mpatches.Patch(color=CLUSTER_COL[l], label=l) for l in ['Eficiente','Intermedio','Crítico']]
ax1.legend(handles=leyenda_cl, fontsize=8, loc='best')

# Risk Score equipos (heatmap-style) 
ax2 = fig6.add_subplot(gs6[0,1])
eq_sorted = equipo_stats.sort_values('risk_score', ascending=True)
colores_risk = [C['rojo'] if s>65 else C['naranja'] if s>40 else C['verde'] for s in eq_sorted['risk_score']]
bars_risk = ax2.barh(eq_sorted['equipo'], eq_sorted['risk_score'], color=colores_risk, edgecolor='white')
ax2.axvline(40, color=C['naranja'], linestyle='--', linewidth=1.3, alpha=0.8, label='Alerta (40)')
ax2.axvline(65, color=C['rojo'],    linestyle='--', linewidth=1.3, alpha=0.8, label='Crítico (65)')
ax2.set_title('Risk Score por Equipo\n(frecuencia 20% · costo 30% · retrabajo 25% · no planif. 25%)')
ax2.set_xlabel('Risk Score (0–100)')
ax2.legend(fontsize=8)
for bar in bars_risk:
    ax2.text(bar.get_width()+0.5, bar.get_y()+bar.get_height()/2,
             f'{bar.get_width():.0f}', va='center', fontsize=7.5)

# Outliers Mahalanobis
ax3 = fig6.add_subplot(gs6[1,0])
normal  = df_c[~df_c['outlier_mah']]
outlier = df_c[ df_c['outlier_mah']]
ax3.scatter(normal['horas trabajadas'],  normal['costo_total']/1000,
            c=C['azul_clar'], s=8, alpha=0.35, edgecolors='none', label='Normal')
ax3.scatter(outlier['horas trabajadas'], outlier['costo_total']/1000,
            c=C['rojo'],      s=40, alpha=0.80, edgecolors=C['azul_osc'], linewidth=0.5,
            label=f'Outlier Mahalanobis (n={n_mah})', zorder=5)
ax3.set_title(f'Outliers Multivariados — Mahalanobis Distance\n'
              f'χ²(df=3, p=0.975) · {n_mah} OT anómalas detectadas')
ax3.set_xlabel('Horas trabajadas')
ax3.set_ylabel('Costo total (K$)')
ax3.yaxis.set_major_formatter(miles)
ax3.legend(fontsize=8)

# Matriz de ineficiencias (heatmap técnico × dimensión) 
ax4 = fig6.add_subplot(gs6[1,1])
dims = ['horas_prom','costo_prom','pct_retrabajo','mttr_prom']
dims_labels = ['Horas\nProm','Costo\nProm','% Retrabajo','MTTR\nProm']
heat_data = scorecard.sort_values('inef_score')[['nombre']+dims].set_index('nombre')
for col in dims:
    mn,mx = heat_data[col].min(), heat_data[col].max()
    heat_data[col] = (heat_data[col]-mn)/(mx-mn+1e-9)
import seaborn as sns
sns.heatmap(heat_data, ax=ax4, cmap='RdYlGn_r', vmin=0, vmax=1,
            linewidths=0.5, linecolor='white', cbar_kws={'label':'Nivel (0=óptimo, 1=peor)'},
            annot=True, fmt='.2f', annot_kws={'size':7.5})
ax4.set_title('Matriz de Ineficiencias por Técnico\n(ordenado de mejor a peor)')
ax4.set_xticklabels(dims_labels, fontsize=8)
ax4.set_yticklabels(ax4.get_yticklabels(), fontsize=7.5, rotation=0)

plt.savefig('C:/Users/Axel/Downloads/Movitec S.A/fig6_fase5_problemas.png', dpi=150, bbox_inches='tight')
plt.close()
print("✓ Figura 6: Detección de problemas guardada")

#  Exportar scorecard y equipos a excel para reporte 
scorecard.to_excel('C:/Users/Axel/Downloads/Movitec S.A/scorecard_tecnicos.xlsx', index=False)
equipo_stats.to_excel('C:/Users/Axel/Downloads/Movitec S.A/risk_equipos.xlsx', index=False)
print("✓ scorecard_tecnicos.xlsx y risk_equipos.xlsx exportados")

# ─ Resumen de hallazgos Fase 5 
print("\n── HALLAZGOS FASE 5")
criticos  = scorecard[scorecard['cluster_label']=='Crítico']
alerta_eq = equipo_stats[equipo_stats['risk_score']>40].sort_values('risk_score',ascending=False)
print(f"  Técnicos en cluster CRÍTICO : {len(criticos)}")
for _,r in criticos.iterrows():
    print(f"    → {r['nombre']:<22} inef={r['inef_score']:.0f}  ret={r['pct_retrabajo']:.1f}%  costo=${r['costo_prom']:,.0f}")
print(f"\n  Equipos en ALERTA/CRÍTICO   : {len(alerta_eq)}")
for _,r in alerta_eq.head(5).iterrows():
    print(f"    → {r['equipo']:<22} risk={r['risk_score']:.0f}  ret={r['pct_retrabajo']:.1f}%  noplanif={r['pct_no_planif']:.0f}%")
print(f"\n  Outliers Mahalanobis        : {n_mah} OT")
print("──────────────────────────────────────────────────────────────")
