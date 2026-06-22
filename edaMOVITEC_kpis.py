"""
MOVITEC S.A. — EDA Completo + KPIs Operacionales

Análisis exploratorio y cálculo
de indicadores clave de mantenimiento industrial.
"""

import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.gridspec import GridSpec
from matplotlib.ticker import FuncFormatter
import seaborn as sns
from scipy import stats
import warnings
import os

warnings.filterwarnings('ignore')

# Configuración visual
PALETTE = {
    'azul_osc':  '#1A3A5C',
    'azul_med':  '#2E6DA4',
    'azul_clar': '#D6E4F0',
    'rojo':      '#C0392B',
    'verde':     '#1E8449',
    'naranja':   '#E67E22',
    'gris':      '#7F8C8D',
    'amarillo':  '#F39C12',
}

COLORES_TIPO = {
    'Electrica':  '#2E86AB',
    'Mecanica':   '#E84855',
    'Preventiva': '#3BB273',
    'Correctiva': '#E67E22',
    'Emergencia': '#8B1A1A',
}

COLORES_ESTADO = {
    'Cerrada':    '#1E8449',
    'En Proceso': '#E67E22',
    'Pendiente':  '#C0392B',
}

plt.rcParams.update({
    'font.family':       'DejaVu Sans',
    'axes.spines.top':   False,
    'axes.spines.right': False,
    'axes.grid':         True,
    'grid.alpha':        0.3,
    'grid.linestyle':    '--',
    'figure.facecolor':  '#FAFAFA',
    'axes.facecolor':    '#FFFFFF',
    'axes.titlesize':    12,
    'axes.titleweight':  'bold',
    'axes.titlepad':     10,
    'axes.labelsize':    10,
    'xtick.labelsize':   9,
    'ytick.labelsize':   9,
})

millones = FuncFormatter(lambda x, _: f'${x/1_000_000:.1f}M')
miles    = FuncFormatter(lambda x, _: f'${x/1_000:.0f}K')

os.makedirs('C:/Users/Axel/Downloads/Movitec S.A/graficos', exist_ok=True)


# CARGA Y PREPARACIÓN

df_raw = pd.read_excel(
    r'C:\Users\Axel\Downloads\Movitec S.A\ordenes_trabajo.xlsx')

# limpiar columnas
df_raw.columns = (
    df_raw.columns
    .str.strip()
    .str.lower()
)

print("COLUMNAS OT:")
print(df_raw.columns.tolist())

# convertir fechas
df_raw['fecha_creacion'] = pd.to_datetime(df_raw['fecha_creacion'])
df_raw['fecha_cierre'] = pd.to_datetime(df_raw['fecha_cierre'])


# TÉCNICOS
df_tec = pd.read_excel(
    r'C:\Users\Axel\Downloads\Movitec S.A\tecnicos.xlsx')

# limpiar columnas
df_tec.columns = (
    df_tec.columns
    .str.strip()
    .str.lower()
)

print("COLUMNAS TEC:")
print(df_tec.columns.tolist())

# COPIA DATASET

df = df_raw.copy()

# variables temporales
df['mes'] = df['fecha_creacion'].dt.to_period('M')
df['año'] = df['fecha_creacion'].dt.year
df['trimestre'] = df['fecha_creacion'].dt.to_period('Q')

# cerradas
df_cerradas = df[df['estado'] == 'Cerrada'].copy()

print(
    f"Dataset: {len(df):,} órdenes | "
    f"{len(df_cerradas):,} cerradas | "
    f"{df['id tecnico'].nunique()} técnicos"
)


# VISIÓN PANORÁMICA DE LOS DATOS

fig1 = plt.figure(figsize=(18, 12), facecolor='#FAFAFA')
fig1.suptitle('MOVITEC S.A. — Visión Panorámica del Dataset', fontsize=16, fontweight='bold',
               color=PALETTE['azul_osc'], y=0.98)
gs1 = GridSpec(2, 3, figure=fig1, hspace=0.45, wspace=0.35)

# Distribución por tipo de falla
ax1 = fig1.add_subplot(gs1[0, 0])
conteo_tipo = df['tipo falla'].value_counts()
bars = ax1.bar(conteo_tipo.index, conteo_tipo.values,
               color=[COLORES_TIPO[t] for t in conteo_tipo.index],
               edgecolor='white', linewidth=0.8)
ax1.set_title('Órdenes por Tipo de Falla')
ax1.set_ylabel('Cantidad')
ax1.set_xticklabels(conteo_tipo.index, rotation=30, ha='right')
for bar in bars:
    ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 10,
             f'{bar.get_height():,.0f}', ha='center', va='bottom', fontsize=8, fontweight='bold')

# Distribución de estados (donut)
ax2 = fig1.add_subplot(gs1[0, 1])
conteo_estado = df['estado'].value_counts()
wedges, texts, autotexts = ax2.pie(
    conteo_estado.values, labels=conteo_estado.index, autopct='%1.1f%%',
    colors=[COLORES_ESTADO[e] for e in conteo_estado.index],
    startangle=90, pctdistance=0.75, wedgeprops=dict(width=0.5, edgecolor='white'))
for t in autotexts:
    t.set_fontsize(9); t.set_fontweight('bold')
ax2.set_title('Distribución de Estados')

# Distribución de costos (histograma + kde)
ax3 = fig1.add_subplot(gs1[0, 2])
datos_costo = df_cerradas['costo_total'].clip(upper=df_cerradas['costo_total'].quantile(0.99))
ax3.hist(datos_costo, bins=50, color=PALETTE['azul_med'], alpha=0.7, edgecolor='white', density=True)
datos_costo.plot.kde(ax=ax3, color=PALETTE['rojo'], linewidth=2)
ax3.set_title('Distribución de Costo Total (OT Cerradas)')
ax3.set_xlabel('Costo CLP')
ax3.xaxis.set_major_formatter(miles)
ax3.set_ylabel('Densidad')
ax3.axvline(df_cerradas['costo_total'].median(), color=PALETTE['naranja'],
            linestyle='--', linewidth=1.5, label=f"Mediana: ${df_cerradas['costo_total'].median()/1000:.0f}K")
ax3.legend(fontsize=8)

# Volumen mensual de órdenes
ax4 = fig1.add_subplot(gs1[1, :2])
vol_mensual = df.groupby('mes').size().reset_index(name='n')
vol_mensual['mes_dt'] = vol_mensual['mes'].dt.to_timestamp()
ax4.fill_between(vol_mensual['mes_dt'], vol_mensual['n'], alpha=0.25, color=PALETTE['azul_med'])
ax4.plot(vol_mensual['mes_dt'], vol_mensual['n'], color=PALETTE['azul_med'], linewidth=2, marker='o', markersize=4)
ax4.set_title('Volumen Mensual de Órdenes Creadas')
ax4.set_ylabel('Órdenes / mes')
ax4.set_xlabel('')
# Anotación tendencia
z = np.polyfit(range(len(vol_mensual)), vol_mensual['n'], 1)
tendencia = np.polyval(z, range(len(vol_mensual)))
ax4.plot(vol_mensual['mes_dt'], tendencia, 'r--', linewidth=1, alpha=0.6, label='Tendencia')
ax4.legend(fontsize=8)

# Costo acumulado por tipo de falla
ax5 = fig1.add_subplot(gs1[1, 2])
costo_tipo = df.groupby('tipo falla')['costo_total'].sum().sort_values(ascending=True)
bars5 = ax5.barh(costo_tipo.index, costo_tipo.values,
                 color=[COLORES_TIPO[t] for t in costo_tipo.index],
                 edgecolor='white')
ax5.set_title('Costo Acumulado por Tipo')
ax5.xaxis.set_major_formatter(millones)
for bar in bars5:
    ax5.text(bar.get_width() + costo_tipo.max()*0.01, bar.get_y() + bar.get_height()/2,
             f'${bar.get_width()/1e6:.1f}M', va='center', fontsize=8)

plt.savefig('C:/Users/Axel/Downloads/Movitec S.A/graficos/fig1_panoramica.png', dpi=150, bbox_inches='tight')
plt.close()
print("✓ Figura 1: Panorámica guardada")


# ANÁLISIS DE TÉCNICOS (EDA)

fig2 = plt.figure(figsize=(18, 14), facecolor='#FAFAFA')
fig2.suptitle('MOVITEC S.A. — Análisis de Desempeño por Técnico', fontsize=16, fontweight='bold',
               color=PALETTE['azul_osc'], y=0.98)
gs2 = GridSpec(2, 2, figure=fig2, hspace=0.5, wspace=0.4)

# Scorecard de técnicos
df_score = df_cerradas.merge(df_tec[['id tecnico','nombre','años de experiencia']], on='id tecnico')
scorecard = df_score.groupby('nombre').agg(
    total_ot        = ('id orden',          'count'),
    horas_prom      = ('horas trabajadas',   'mean'),
    costo_prom      = ('costo_total',        'mean'),
    pct_retrabajo   = ('retrabajo',          lambda x: (x=='Sí').mean() * 100),
    mttr_prom       = ('tiempo resolucion horas', 'mean'),
    exp             = ('años de experiencia',   'first'),
).reset_index()

# Score de ineficiencia compuesto (menor = más eficiente)

for col, w in [('horas_prom', 0.30), ('costo_prom', 0.40), ('pct_retrabajo', 0.30)]:
    mx, mn = scorecard[col].max(), scorecard[col].min()
    scorecard[f'{col}_norm'] = (scorecard[col] - mn) / (mx - mn + 1e-9)
scorecard['ineficiencia_score'] = (
    scorecard['horas_prom_norm']    * 0.30 +
    scorecard['costo_prom_norm']    * 0.40 +
    scorecard['pct_retrabajo_norm'] * 0.30
) * 100
scorecard = scorecard.sort_values('ineficiencia_score')

# Ranking de eficiencia
ax1 = fig2.add_subplot(gs2[0, :])
colores_rank = []
for _, row in scorecard.iterrows():
    if row['ineficiencia_score'] <= 30:
        colores_rank.append(PALETTE['verde'])
    elif row['ineficiencia_score'] <= 60:
        colores_rank.append(PALETTE['naranja'])
    else:
        colores_rank.append(PALETTE['rojo'])

bars_r = ax1.barh(scorecard['nombre'], scorecard['ineficiencia_score'],
                  color=colores_rank, edgecolor='white', linewidth=0.8)
ax1.axvline(30, color=PALETTE['verde'],  linestyle=':', linewidth=1.5, alpha=0.7, label='Zona eficiente (≤30)')
ax1.axvline(60, color=PALETTE['naranja'],linestyle=':', linewidth=1.5, alpha=0.7, label='Zona de alerta (≤60)')
ax1.set_title('Score de Ineficiencia por Técnico (compuesto: costo 40% · horas 30% · retrabajo 30%)')
ax1.set_xlabel('Score de Ineficiencia (0=óptimo, 100=peor)')
for bar in bars_r:
    ax1.text(bar.get_width() + 0.5, bar.get_y() + bar.get_height()/2,
             f'{bar.get_width():.1f}', va='center', fontsize=8)
leyenda = [
    mpatches.Patch(color=PALETTE['verde'],   label='Eficiente  (≤30)'),
    mpatches.Patch(color=PALETTE['naranja'], label='Alerta     (31–60)'),
    mpatches.Patch(color=PALETTE['rojo'],    label='Crítico    (>60)'),
]
ax1.legend(handles=leyenda, fontsize=9, loc='lower right')

# % Retrabajo vs Experiencia 
ax2 = fig2.add_subplot(gs2[1, 0])
scatter = ax2.scatter(scorecard['exp'], scorecard['pct_retrabajo'],
                      c=scorecard['ineficiencia_score'], cmap='RdYlGn_r',
                      s=scorecard['total_ot']*2.5, alpha=0.85, edgecolors='white', linewidth=0.8)
plt.colorbar(scatter, ax=ax2, label='Score ineficiencia')
# Etiquetar técnico problemático
for _, row in scorecard[scorecard['pct_retrabajo'] > 35].iterrows():
    ax2.annotate(row['nombre'].split()[0], (row['exp'], row['pct_retrabajo']),
                 textcoords='offset points', xytext=(6, 4), fontsize=7.5,
                 color=PALETTE['rojo'], fontweight='bold')
# Línea de tendencia
z = np.polyfit(scorecard['exp'], scorecard['pct_retrabajo'], 1)
x_line = np.linspace(scorecard['exp'].min(), scorecard['exp'].max(), 100)
ax2.plot(x_line, np.polyval(z, x_line), '--', color=PALETTE['azul_med'], linewidth=1.5, alpha=0.7)
ax2.axhline(10, color=PALETTE['verde'], linestyle='--', linewidth=1.2, alpha=0.6, label='Benchmark 10%')
ax2.set_xlabel('Años de Experiencia')
ax2.set_ylabel('% Retrabajo')
ax2.set_title('Retrabajo vs Experiencia\n(burbuja = volumen de OT)')
ax2.legend(fontsize=8)

#  Costo promedio por técnico 
ax3 = fig2.add_subplot(gs2[1, 1])
sc_sorted = scorecard.sort_values('costo_prom', ascending=False)
colores_costo = [PALETTE['rojo'] if c > df_cerradas['costo_total'].mean()*1.3
                 else PALETTE['naranja'] if c > df_cerradas['costo_total'].mean()
                 else PALETTE['verde'] for c in sc_sorted['costo_prom']]
bars3 = ax3.bar(range(len(sc_sorted)), sc_sorted['costo_prom'],
                color=colores_costo, edgecolor='white')
ax3.axhline(df_cerradas['costo_total'].mean(), color=PALETTE['azul_med'],
            linestyle='--', linewidth=2, label=f'Promedio ${df_cerradas["costo_total"].mean()/1000:.0f}K')
ax3.set_title('Costo Promedio por OT — por Técnico')
ax3.set_ylabel('CLP promedio / OT')
ax3.yaxis.set_major_formatter(miles)
ax3.set_xticks(range(len(sc_sorted)))
ax3.set_xticklabels([n.split()[0] for n in sc_sorted['nombre']], rotation=45, ha='right', fontsize=8)
ax3.legend(fontsize=8)

plt.savefig('C:/Users/Axel/Downloads/Movitec S.A/graficos/fig2_tecnicos.png', dpi=150, bbox_inches='tight')
plt.close()
print("✓ Figura 2: Análisis técnicos guardada")

# ANÁLISIS DE TIEMPOS Y COSTOS
 
fig3 = plt.figure(figsize=(18, 12), facecolor='#FAFAFA')
fig3.suptitle('MOVITEC S.A. — Análisis de Tiempos y Costos', fontsize=16, fontweight='bold',
               color=PALETTE['azul_osc'], y=0.98)
gs3 = GridSpec(2, 3, figure=fig3, hspace=0.5, wspace=0.4)

# Boxplot de horas por tipo de falla
ax1 = fig3.add_subplot(gs3[0, :2])
orden_tipos = df_cerradas.groupby('tipo falla')['horas trabajadas'].median().sort_values(ascending=False).index
bp = ax1.boxplot(
    [df_cerradas[df_cerradas['tipo falla']==t]['horas trabajadas']
     .clip(upper=df_cerradas['horas trabajadas'].quantile(0.95)) for t in orden_tipos],
    labels=orden_tipos, patch_artist=True, notch=False,
    medianprops=dict(color='white', linewidth=2),
    flierprops=dict(marker='o', markersize=2, alpha=0.3),
)
for patch, tipo in zip(bp['boxes'], orden_tipos):
    patch.set_facecolor(COLORES_TIPO[tipo])
    patch.set_alpha(0.75)
ax1.set_title('Distribución de Horas Trabajadas por Tipo de Falla (p95 recortado)')
ax1.set_ylabel('Horas trabajadas')
ax1.axhline(df_cerradas['horas trabajadas'].mean(), color=PALETTE['gris'],
            linestyle='--', linewidth=1.2, label=f'Media global: {df_cerradas["horas trabajadas"].mean():.1f}h')
ax1.legend(fontsize=8)

# MTTR por tipo de falla (barras horizontales con IC 95%)
ax2 = fig3.add_subplot(gs3[0, 2])
mttr_tipo = df_cerradas.groupby('tipo falla')['tiempo resolucion horas'].agg(['mean','sem']).sort_values('mean', ascending=True)
mttr_tipo['ic95'] = mttr_tipo['sem'] * 1.96
bars2 = ax2.barh(mttr_tipo.index, mttr_tipo['mean'],
                 xerr=mttr_tipo['ic95'], capsize=4,
                 color=[COLORES_TIPO[t] for t in mttr_tipo.index],
                 edgecolor='white', linewidth=0.8, error_kw={'elinewidth':1.5, 'ecolor':'gray'})
ax2.set_title('MTTR por Tipo de Falla\n(horas ± IC 95%)')
ax2.set_xlabel('Horas promedio resolución')
for i, (_, row) in enumerate(mttr_tipo.iterrows()):
    ax2.text(row['mean'] + row['ic95'] + 1, i, f'{row["mean"]:.1f}h', va='center', fontsize=8)

# Evolución mensual de costos (apilado por tipo)
ax3 = fig3.add_subplot(gs3[1, :])
costo_mensual_tipo = (df.groupby([df['fecha_creacion'].dt.to_period('M'), 'tipo falla'])['costo_total']
                      .sum().unstack(fill_value=0))
costo_mensual_tipo.index = costo_mensual_tipo.index.to_timestamp()
tipos_orden = df.groupby('tipo falla')['costo_total'].sum().sort_values(ascending=False).index
bottom = np.zeros(len(costo_mensual_tipo))
for tipo in tipos_orden:
    if tipo in costo_mensual_tipo.columns:
        vals = costo_mensual_tipo[tipo].values / 1e6
        ax3.bar(costo_mensual_tipo.index, vals, bottom=bottom,
                label=tipo, color=COLORES_TIPO[tipo], edgecolor='white', linewidth=0.3, alpha=0.85)
        bottom += vals
ax3.set_title('Evolución Mensual de Costos por Tipo de Falla (Millones CLP)')
ax3.set_ylabel('Costo Total (M$)')
ax3.legend(loc='upper left', fontsize=8, ncol=5)
ax3.set_xlabel('')

plt.savefig('C:/Users/Axel/Downloads/Movitec S.A/graficos/fig3_tiempos_costos.png', dpi=150, bbox_inches='tight')
plt.close()
print("✓ Figura 3: Tiempos y costos guardada")

# FIGURA 4 — DETECCIÓN DE OUTLIERS Y EQUIPOS

fig4 = plt.figure(figsize=(18, 12), facecolor='#FAFAFA')
fig4.suptitle('MOVITEC S.A. — Outliers y Análisis de Equipos', fontsize=16, fontweight='bold',
               color=PALETTE['azul_osc'], y=0.98)
gs4 = GridSpec(2, 2, figure=fig4, hspace=0.5, wspace=0.35)

#Z-score de costos (scatter)
ax1 = fig4.add_subplot(gs4[0, :])
mu, sigma = df_cerradas['costo_total'].mean(), df_cerradas['costo_total'].std()
df_cerradas = df_cerradas.copy()
df_cerradas['z_score'] = (df_cerradas['costo_total'] - mu) / sigma
colores_z = df_cerradas['z_score'].apply(
    lambda z: PALETTE['rojo'] if z > 3.0 else
              PALETTE['naranja'] if z > 2.5 else
              PALETTE['azul_clar'])
ax1.scatter(range(len(df_cerradas)), df_cerradas['z_score'].values,
            c=colores_z, s=6, alpha=0.5)
ax1.axhline(2.5,  color=PALETTE['naranja'], linestyle='--', linewidth=1.5, label='Z=2.5 (outlier)')
ax1.axhline(3.0,  color=PALETTE['rojo'],    linestyle='--', linewidth=1.5, label='Z=3.0 (crítico)')
ax1.axhline(-2.5, color=PALETTE['naranja'], linestyle='--', linewidth=1.5)
ax1.set_title(f'Detección de Outliers por Z-Score en Costo Total · '
              f'{(df_cerradas["z_score"].abs()>2.5).sum()} outliers detectados')
ax1.set_ylabel('Z-Score')
ax1.set_xlabel('Índice de orden')
ax1.legend(fontsize=8, loc='upper left')
n_outliers = (df_cerradas['z_score'].abs() > 2.5).sum()
ax1.text(0.98, 0.95, f'{n_outliers} outliers\n({n_outliers/len(df_cerradas)*100:.1f}%)',
         transform=ax1.transAxes, ha='right', va='top', fontsize=9,
         bbox=dict(boxstyle='round,pad=0.4', fc=PALETTE['rojo'], alpha=0.85, ec='none'), color='white')

#Pareto de equipos (costo acumulado)
ax2 = fig4.add_subplot(gs4[1, 0])
pareto = (df.groupby('equipo')['costo_total'].sum()
            .sort_values(ascending=False)
            .reset_index())
pareto['pct_acum'] = pareto['costo_total'].cumsum() / pareto['costo_total'].sum() * 100
ax2b = ax2.twinx()
bars_p = ax2.bar(range(len(pareto)), pareto['costo_total']/1e6,
                 color=PALETTE['azul_med'], alpha=0.75, edgecolor='white')
ax2b.plot(range(len(pareto)), pareto['pct_acum'], color=PALETTE['rojo'],
          linewidth=2, marker='o', markersize=4)
ax2b.axhline(80, color=PALETTE['gris'], linestyle=':', linewidth=1.5, alpha=0.7)
ax2.set_title('Pareto de Equipos por Costo Acumulado')
ax2.set_ylabel('Costo (M$)')
ax2b.set_ylabel('% Acumulado', color=PALETTE['rojo'])
ax2.set_xticks(range(len(pareto)))
ax2.set_xticklabels([e.replace(' ', '\n') for e in pareto['equipo']], fontsize=6.5, rotation=45, ha='right')
# Marcar el corte 80%
corte80 = (pareto['pct_acum'] <= 80).sum()
ax2.axvline(corte80 - 0.5, color=PALETTE['naranja'], linestyle='--', linewidth=1.5,
            label=f'80% del costo → {corte80} equipos')
ax2.legend(fontsize=7)

# Tasa retrabajo por equipo
ax3 = fig4.add_subplot(gs4[1, 1])
retrabajo_equipo = (df.groupby('equipo')
                     .agg(total=('id orden','count'), retrabajos=('retrabajo', lambda x: (x=='Sí').sum()))
                     .assign(pct_retrabajo=lambda d: d['retrabajos']/d['total']*100)
                     .sort_values('pct_retrabajo', ascending=True))
colores_ret = [PALETTE['rojo'] if p > 30 else PALETTE['naranja'] if p > 20 else PALETTE['verde']
               for p in retrabajo_equipo['pct_retrabajo']]
bars_r = ax3.barh(retrabajo_equipo.index, retrabajo_equipo['pct_retrabajo'],
                  color=colores_ret, edgecolor='white')
ax3.axvline(10, color=PALETTE['verde'],   linestyle='--', linewidth=1.5, alpha=0.7, label='Benchmark 10%')
ax3.axvline(21.7, color=PALETTE['azul_med'], linestyle=':', linewidth=1.5, alpha=0.7, label='Promedio 21.7%')
ax3.set_title('Tasa de Retrabajo por Equipo')
ax3.set_xlabel('% Retrabajo')
ax3.legend(fontsize=8)
for bar in bars_r:
    ax3.text(bar.get_width() + 0.2, bar.get_y() + bar.get_height()/2,
             f'{bar.get_width():.1f}%', va='center', fontsize=7)

plt.savefig('C:/Users/Axel/Downloads/Movitec S.A/graficos/fig4_outliers_equipos.png', dpi=150, bbox_inches='tight')
plt.close()
print("✓ Figura 4: Outliers y equipos guardada")

# CORRELACIONES Y MAPA DE CALOR

fig5 = plt.figure(figsize=(18, 12), facecolor='#FAFAFA')
fig5.suptitle('MOVITEC S.A. — Correlaciones y KPIs por Dimensión', fontsize=16, fontweight='bold',
               color=PALETTE['azul_osc'], y=0.98)
gs5 = GridSpec(2, 3, figure=fig5, hspace=0.5, wspace=0.45)

# 5A: Heatmap de correlaciones numéricas
ax1 = fig5.add_subplot(gs5[0, :2])
vars_num = ['horas trabajadas','costo mano obra','costo repuestos','costo_total','tiempo resolucion horas']
corr_mat = df_cerradas[vars_num].corr()
mask = np.triu(np.ones_like(corr_mat, dtype=bool))
sns.heatmap(corr_mat, ax=ax1, mask=mask, annot=True, fmt='.2f',
            cmap='RdYlGn', center=0, vmin=-1, vmax=1,
            linewidths=0.5, linecolor='white',
            annot_kws={'size': 10, 'weight': 'bold'})
ax1.set_title('Mapa de Correlaciones — Variables Métricas Clave')
ax1.set_xticklabels([v.replace('_',' ').title() for v in vars_num], rotation=30, ha='right', fontsize=9)
ax1.set_yticklabels([v.replace('_',' ').title() for v in vars_num], rotation=0, fontsize=9)

# Correlación horas vs costo (scatter)
ax2 = fig5.add_subplot(gs5[0, 2])
sample = df_cerradas.sample(min(1500, len(df_cerradas)), random_state=42)
ax2.scatter(sample['horas trabajadas'], sample['costo_total']/1000,
            c=[COLORES_TIPO[t] for t in sample['tipo falla']],
            alpha=0.4, s=18, edgecolors='none')
z_fit = np.polyfit(sample['horas trabajadas'], sample['costo_total'], 1)
x_r = np.linspace(sample['horas trabajadas'].min(), sample['horas trabajadas'].quantile(0.97), 100)
ax2.plot(x_r, np.polyval(z_fit, x_r)/1000, 'k--', linewidth=1.5, alpha=0.7, label='Tendencia lineal')
legend_el = [mpatches.Patch(color=COLORES_TIPO[t], label=t) for t in COLORES_TIPO]
ax2.legend(handles=legend_el, fontsize=7, loc='upper left')
ax2.set_xlabel('Horas trabajadas')
ax2.set_ylabel('Costo total (K$)')
ax2.set_title('Horas Trabajadas vs Costo Total')

# % SLA por prioridad
ax3 = fig5.add_subplot(gs5[1, 0])
sla_limites = {'Alta': 24, 'Media': 72, 'Baja': 168}
sla_rows = []
for prio, limite in sla_limites.items():
    sub = df_cerradas[(df_cerradas['prioridad']==prio) & df_cerradas['tiempo resolucion horas'].notna()]
    cumple = (sub['tiempo resolucion horas'] <= limite).mean() * 100
    sla_rows.append({'prioridad': prio, 'pct_sla': cumple, 'limite_h': limite, 'total': len(sub)})
df_sla = pd.DataFrame(sla_rows)
colores_sla = [PALETTE['rojo'] if p < 70 else PALETTE['naranja'] if p < 85 else PALETTE['verde']
               for p in df_sla['pct_sla']]
bars_sla = ax3.bar(df_sla['prioridad'], df_sla['pct_sla'], color=colores_sla, edgecolor='white', width=0.55)
ax3.axhline(90, color=PALETTE['verde'], linestyle='--', linewidth=1.5, label='Objetivo 90%')
ax3.set_title('% Cumplimiento SLA por Prioridad')
ax3.set_ylabel('% Órdenes dentro del SLA')
ax3.set_ylim(0, 105)
ax3.legend(fontsize=8)
for bar, row in zip(bars_sla, df_sla.itertuples()):
    ax3.text(bar.get_x()+bar.get_width()/2, bar.get_height()+1,
             f'{row.pct_sla:.1f}%\n(SLA ≤{row.limite_h}h)', ha='center', fontsize=8, fontweight='bold')

# KPI — MTTR y Costo por prioridad
ax4 = fig5.add_subplot(gs5[1, 1])
kpi_prio = df_cerradas.groupby('prioridad').agg(
    mttr_h   = ('tiempo resolucion horas','mean'),
    costo_k  = ('costo_total', lambda x: x.mean()/1000),
).reindex(['Alta','Media','Baja'])
x = np.arange(3)
w = 0.35
ax4b = ax4.twinx()
b1 = ax4.bar(x - w/2, kpi_prio['mttr_h'], w, label='MTTR (horas)', color=PALETTE['azul_med'], alpha=0.8)
b2 = ax4b.bar(x + w/2, kpi_prio['costo_k'], w, label='Costo promedio (K$)', color=PALETTE['naranja'], alpha=0.8)
ax4.set_title('MTTR y Costo Promedio por Prioridad')
ax4.set_ylabel('MTTR (horas)', color=PALETTE['azul_med'])
ax4b.set_ylabel('Costo promedio (K$)', color=PALETTE['naranja'])
ax4.set_xticks(x)
ax4.set_xticklabels(['Alta','Media','Baja'])
lines = [mpatches.Patch(color=PALETTE['azul_med'], label='MTTR (h)'),
         mpatches.Patch(color=PALETTE['naranja'],   label='Costo prom (K$)')]
ax4.legend(handles=lines, fontsize=8)

# % retrabajo por tipo de falla
ax5 = fig5.add_subplot(gs5[1, 2])
ret_tipo = df.groupby('tipo falla').apply(
    lambda g: (g['retrabajo']=='Sí').mean()*100
).sort_values(ascending=True)
colores_ret_t = [PALETTE['rojo'] if p > 25 else PALETTE['naranja'] if p > 15 else PALETTE['verde']
                 for p in ret_tipo.values]
bars_rt = ax5.barh(ret_tipo.index, ret_tipo.values, color=colores_ret_t, edgecolor='white')
ax5.axvline(10, color=PALETTE['verde'], linestyle='--', linewidth=1.5, label='Benchmark 10%')
ax5.set_title('% Retrabajo por Tipo de Falla')
ax5.set_xlabel('% Retrabajo')
ax5.legend(fontsize=8)
for bar in bars_rt:
    ax5.text(bar.get_width()+0.2, bar.get_y()+bar.get_height()/2,
             f'{bar.get_width():.1f}%', va='center', fontsize=8, fontweight='bold')

plt.savefig('C:/Users/Axel/Downloads/Movitec S.A/graficos/fig5_correlaciones_kpis.png', dpi=150, bbox_inches='tight')
plt.close()
print("✓ Figura 5: Correlaciones y KPIs guardada")

# CÁLCULO FORMAL DE KPIs 
print("\n" + "═"*65)
print("  MOVITEC S.A. — KPIs OPERACIONALES  (Fase 4)")
print("═"*65)

# KPI: MTTR
mttr_global = df_cerradas['tiempo resolucion horas'].mean()
mttr_por_tipo = df_cerradas.groupby('tipo falla')['tiempo resolucion horas'].mean().round(1)
print(f"\n▸ MTTR Global               : {mttr_global:.1f} horas")
print(f"  MTTR por tipo de falla    :")
for t, v in mttr_por_tipo.sort_values(ascending=False).items():
    flag = '⚠' if v > 50 else '✓'
    print(f"    {flag}  {t:<14}: {v:>6.1f} h")

# KPI: Costo promedio por OT
costo_prom_global  = df_cerradas['costo_total'].mean()
costo_mediana      = df_cerradas['costo_total'].median()
costo_p95          = df_cerradas['costo_total'].quantile(0.95)
costo_por_tipo     = df_cerradas.groupby('tipo falla')['costo_total'].mean()
print(f"\n▸ Costo Promedio / OT       : ${costo_prom_global:>12,.0f} CLP")
print(f"  Mediana                   : ${costo_mediana:>12,.0f} CLP")
print(f"  P95 (umbral outlier)      : ${costo_p95:>12,.0f} CLP")
print(f"  Costo total flota (3 años): ${df['costo_total'].sum():>12,.0f} CLP")

# KPI: % retrabajo
pct_retrabajo_global = (df['retrabajo']=='Sí').mean() * 100
costo_retrabajo      = df[df['retrabajo']=='Sí']['costo_total'].sum()
print(f"\n▸ Tasa de Retrabajo Global  : {pct_retrabajo_global:.1f}%  (benchmark: ≤10%)  ⚠ CRÍTICO")
print(f"  Costo atribuible retrabajo: ${costo_retrabajo:>12,.0f} CLP")
print(f"  → Exceso sobre benchmark  : {pct_retrabajo_global-10:.1f} pp  ← oportunidad de ahorro clave")

# KPI: Eficiencia por técnico (top y bottom)
top3    = scorecard.nsmallest(3,  'ineficiencia_score')[['nombre','ineficiencia_score','pct_retrabajo','costo_prom','mttr_prom']]
bottom3 = scorecard.nlargest(3,   'ineficiencia_score')[['nombre','ineficiencia_score','pct_retrabajo','costo_prom','mttr_prom']]
print(f"\n▸ Top 3 Técnicos Eficientes :")
for _, r in top3.iterrows():
    print(f"    ✓ {r['nombre']:<20} Score={r['ineficiencia_score']:.1f}  Retrabajo={r['pct_retrabajo']:.1f}%  CostoProm=${r['costo_prom']:,.0f}")
print(f"\n▸ Bottom 3 Técnicos (alerta):")
for _, r in bottom3.iterrows():
    print(f"    ⚠ {r['nombre']:<20} Score={r['ineficiencia_score']:.1f}  Retrabajo={r['pct_retrabajo']:.1f}%  CostoProm=${r['costo_prom']:,.0f}")

# KPI: SLA
print(f"\n▸ Cumplimiento SLA          :")
for _, row in df_sla.iterrows():
    flag = '✓' if row['pct_sla'] >= 90 else '⚠' if row['pct_sla'] >= 70 else '✗'
    print(f"    {flag}  Prioridad {row['prioridad']:<6}: {row['pct_sla']:.1f}%  (objetivo ≥90%, umbral ≤{row['limite_h']}h)")

# KPI: Outliers
n_out  = (df_cerradas['z_score'].abs() > 2.5).sum()
costo_out = df_cerradas[df_cerradas['z_score'].abs() > 2.5]['costo_total'].sum()
print(f"\n▸ Outliers detectados       : {n_out} OT  ({n_out/len(df_cerradas)*100:.1f}% del total cerradas)")
print(f"  Costo acumulado outliers  : ${costo_out:>12,.0f} CLP")
print(f"  Impacto sobre costo total : {costo_out/df['costo_total'].sum()*100:.1f}%")

# KPI: Tiempo promedio por tipo de falla (tabla limpia)
print(f"\n▸ Tiempo Prom. por Tipo     :")
for t, v in df_cerradas.groupby('tipo falla')['horas trabajadas'].mean().sort_values(ascending=False).items():
    print(f"    {t:<14}: {v:.1f} h/OT")

print("\n" + "═"*65)
print("  POTENCIAL DE AHORRO ESTIMADO")
print("═"*65)

# Ahorro 1: reducir retrabajo a 10%
ots_retrabajo_exceso    = df[df['retrabajo']=='Sí']
costo_prom_retrabajo    = ots_retrabajo_exceso['costo_total'].mean()
n_retrabajo_evitable    = len(ots_retrabajo_exceso) * (1 - 10/pct_retrabajo_global)
ahorro_retrabajo        = n_retrabajo_evitable * costo_prom_retrabajo * 0.70  # 70% evitable
print(f"\n  Reducir retrabajo de {pct_retrabajo_global:.1f}% → 10%")
print(f"  → Ahorro estimado anual   : ${ahorro_retrabajo/3/1e6:.1f}M CLP / año")

df_cerradas = df_cerradas.merge(
    df_tec[['id tecnico', 'nombre']],
    on='id tecnico',
    how='left'
)

# AHORRO 2: TECNICOS CRITICOS → EFICIENCIA MEDIA

# Cantidad de OT por técnico
ot_por_tecnico = (
    df_cerradas
    .groupby('nombre')
    .size()
    .reset_index(name='n')
)

# Diferencia de costo vs mediana general
diff_costo = (
    (
        bottom3['costo_prom'].mean()
        - scorecard['costo_prom'].median()
    )
    *
    (
        bottom3
        .merge(
            ot_por_tecnico,
            on='nombre',
            how='left'
        )['n']
        .mean()
    )
)

# Ahorro estimado
ahorro_tecnicos = max(0, diff_costo / 3)


print(f"\n  Mejorar eficiencia técnicos críticos → promedio")
print(f"  → Ahorro estimado anual   : ${ahorro_tecnicos/1e6:.1f}M CLP / año")

print(
    f"\n  AHORRO TOTAL ESTIMADO     : "
    f"${(ahorro_retrabajo/3 + ahorro_tecnicos)/1e6:.1f}M CLP / año"
)

print("═"*65)


