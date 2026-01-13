import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# Carga de los datos
archivo_excel = 'turismo.xlsx'
try:
    xls = pd.ExcelFile(archivo_excel)
    df = pd.read_excel(archivo_excel)
    # Limpieza
    if 'drop' in df.columns: df = df.drop(columns=['drop'])
    df['value'] = pd.to_numeric(df['value'], errors='coerce').fillna(0)
    print(f"Carga exitosa del archivo: {archivo_excel}")
except Exception as e:
    print(f"Error al cargar el archivo: {e}")
    import sys; sys.exit()

# Definición de la región de estudio (países de África Central con mercado de arte)
africa_central = ['Ghana', "Côte d’Ivoire", 'Cameroon', 'Gabon', 'Congo', 'Democratic Republic of the Congo']

# Definir colores fijos para mantener la coherencia en todo el proyecto
color_map = {
    'Ghana': '#006B3F',                           # Verde fuerte
    "Côte d’Ivoire": '#FF8200',                   # Dorado/Naranja
    'Cameroon': '#CE1126',                        # Rojo
    'Gabon': '#FCD116',                           # Amarillo
    'Congo': '#264653',                           # Azul oscuro
    'Democratic Republic of the Congo': '#007FFF' # Azul claro
}

# Título del dashboard
fig_titulo = go.Figure()
fig_titulo.add_annotation(
    text="<b>DASHBOARD: DESTINOS TURÍSTICOS<br>PARA EL MERCADO DEL ARTE AFRICANO</b><br><span style='font-size:18px; color:gray;'>Análisis de flujos turísticos en 6 países de África central y occidental</span>",
    showarrow=False,
    xref="paper", yref="paper",
    x=0.5, y=0.5,
    font=dict(size=35, color="#D4AF37"),
    align="center"
)
fig_titulo.update_layout(
    plot_bgcolor="black",
    paper_bgcolor="black",
    xaxis=dict(showgrid=False, zeroline=False, showticklabels=False),
    yaxis=dict(showgrid=False, zeroline=False, showticklabels=False),
    height=250
)
fig_titulo.show()

# Filtrar los datos del mundo y de nuestra región de estudio
df_world = df[df['partner_area_label'] == 'World'].groupby('year')['value'].sum().reset_index()
df_ac = df[df['reporter_area_label'].isin(africa_central)].groupby('year')['value'].sum().reset_index()

# Visualizar gráfico comparativo de la región vs el mundo para contexto
fig_contexto = go.Figure()
fig_contexto.add_trace(go.Scatter(x=df_world['year'], y=df_world['value'], name='Mundo', line=dict(color='silver', width=2, dash='dot')))
fig_contexto.add_trace(go.Scatter(x=df_ac['year'], y=df_ac['value'], name='África 6 (c/o)', line=dict(color='#D4AF37', width=5))) # Color Oro
fig_contexto.update_layout(
    title="<b>Contexto:</b> 6 países de África central y occidental en el turismo global",
    yaxis_type="log", template="plotly_dark",
    xaxis_title="Año", yaxis_title="Llegadas (miles) - escala logarítmica",
    hovermode="x unified"
)
fig_contexto.show()

# Filtrar datos sin propósito para evitar datos ausentes
df_estudio = df[
    (df['reporter_area_label'].isin(africa_central)) &
    (df['indicator_label'].str.contains('region', case=False)) &
    (df['partner_area_label'] != 'World') &
    (df['year'] >= 2017) # Recientes
].copy()

# Agrupar por país receptor y por región de origen (media de los últimos años disponibles)
df_final = df_estudio.groupby(['reporter_area_label', 'partner_area_label'])['value'].mean().reset_index()

# Visualizar gráfico de países receptores y regiones de origen
fig = px.bar(
    df_final,
    y="reporter_area_label",
    x="value",
    color="partner_area_label",
    title="<b>Análisis de los flujos turísticos:</b> ¿quién visita estos países? (2016-2022)",
    labels={'value': 'Promedio llegadas (miles)', 'reporter_area_label': 'País de destino', 'partner_area_label': 'Origen:'},
    barmode="stack",
    orientation='h',
    template="plotly_dark",
    color_discrete_sequence=px.colors.qualitative.Vivid
)
fig.update_layout(yaxis={'categoryorder':'total ascending'})
fig.show()

# Filtrar los datos de nuestra región de estudio y de llegadas totales
df_f = df[df['reporter_area_label'].isin(africa_central)].copy()
df_burbujas = df_f.groupby(['reporter_area_label', 'year'])['value'].max().reset_index()

# Crear una rejilla de años y países, uniéndola con los datos reales para quitar años sin datos
años = range(1995, 2023)
rejilla = pd.MultiIndex.from_product([africa_central, años], names=['reporter_area_label', 'year']).to_frame(index=False)
df_final = pd.merge(rejilla, df_burbujas, on=['reporter_area_label', 'year'], how='left')
df_final['size_value'] = df_final['value'].fillna(0)

# Visualizar gráfico de burbujas animado
fig = px.scatter(
    df_final,
    x="year",
    y="value",
    animation_frame="year",
    animation_group="reporter_area_label",
    size="size_value",
    color="reporter_area_label",
    color_discrete_map=color_map,
    category_orders={'reporter_area_label': africa_central},
    hover_name="reporter_area_label",
    title="<b>Evolución del mercado:</b> volumen de turistas (1995-2022)",
    labels={'value': 'Llegadas (miles) - escala logarítmica', 'year': 'Año', 'reporter_area_label': 'País:'},
    log_y=True,  # Escala logarítmica para comparar distintas magnitudes
    range_x=[1994, 2023],
    range_y=[df_final[df_final['value']>0]['value'].min() * 0.5, df_final['value'].max() * 2],
    size_max=60,
    template="plotly_dark"
)
fig.layout.updatemenus[0].buttons[0].args[1]['frame']['redraw'] = True
fig.show()

# Filtrar los datos para tener un solo valor total por año y país
df_burbujas = df_burbujas[df_burbujas['value'] > 0]

# Visualizar gráfico de burbujas
fig = px.scatter(
    df_burbujas,
    x="year",
    y="value",
    size="value",
    color="reporter_area_label",
    color_discrete_map=color_map,
    category_orders={"reporter_area_label": africa_central},
    hover_name="reporter_area_label",
    labels={'value': 'Total llegadas (miles)', 'year': 'Año', 'reporter_area_label': 'País'},
    log_y=True,  # Escala logarítmica para comparar distintas magnitudes
    size_max=45,
    template="plotly_dark"
)
fig.update_traces(marker=dict(line=dict(width=0.5, color='darkgrey'), opacity=0.8))
fig.update_layout(
    hovermode="closest",
    legend_title="Países seleccionados:",
    xaxis_title="Año",
    yaxis_title="Llegadas (miles) - escala logarítmica"
)
fig.show()


# Definir colores para propósito
color_map_proposito = {
    'Negocios': '#0E14E0',                  # Azul marino
    'Ocio / Cultura': '#33DEF0',            # Azul cielo
    'Dato general no desglosado': '#5B6061' # Gris neutro
}

# Filtrar datos por propósito
mapeo_labels = {
    'inbound - trips - by purpose - business - overnight visitors (tourists)': 'Negocios',
    'inbound - trips - by purpose - total - overnight visitors (tourists)': 'Ocio / Cultura'
}
df_f = df[df['reporter_area_label'].isin(africa_central)].copy()
df_f['Categoria'] = df_f['indicator_label'].map(mapeo_labels)
con_desglose = df_f.dropna(subset=['Categoria'])
sin_desglose = df_f[~df_f['reporter_area_label'].isin(con_desglose['reporter_area_label'].unique())]
totales_alternativos = sin_desglose[sin_desglose['indicator_label'].str.contains('total', case=False)]
totales_alternativos = totales_alternativos.copy()
totales_alternativos['Categoria'] = 'Dato general no desglosado'
df_final_prop = pd.concat([con_desglose, totales_alternativos])

# Filtrar último año disponible
df_ultimo = df_final_prop[df_final_prop['value'] > 0].groupby('reporter_area_label')['year'].max().reset_index()
df_resumen = pd.merge(df_final_prop, df_ultimo, on=['reporter_area_label', 'year'])

# Calcular procentaje
df_totales = df_resumen.groupby('reporter_area_label')['value'].transform('sum')
df_resumen['porcentaje'] = (df_resumen['value'] / df_totales) * 100

# Visualiza gráfico de barras apiladas horizontales
fig_proposito = px.bar(
    df_resumen,
    x="porcentaje",
    y="reporter_area_label",
    color="Categoria",
    orientation='h',
    title="<b>El principal tipo de turismo:</b> ¿ocio/cultura o negocios?",
    labels={'porcentaje': 'Distribución (%)', 'reporter_area_label': 'País'},
    color_discrete_map=color_map_proposito,
    category_orders={
        "reporter_area_label": africa_central,
        "Categoria": ["Ocio / Cultura", "Negocios", "Dato general no desglosado (datos de propósito sin especificar)"]
    },
    template="plotly_dark",
    text=df_resumen['porcentaje'].apply(lambda x: f'{x:.1f}%')
)
fig_proposito.update_layout(xaxis_title="Distribución del tipo de turismo en los países de la región", yaxis_title=None, margin=dict(l=150))
fig_proposito.show()

# Filtrar los datos de nuestra región de estudio
df_proposito_hist = df[df['reporter_area_label'].isin(africa_central)].copy()

# Filtrar datos por propósito
mapeo_labels = {
    'inbound - trips - by purpose - business - overnight visitors (tourists)': 'Negocios',
    'inbound - trips - by purpose - personal - overnight visitors (tourists)': 'Ocio / Cultura'
}
df_proposito_hist['indicator_label'] = df_proposito_hist['indicator_label'].map(mapeo_labels)
df_proposito_hist = df_proposito_hist.dropna(subset=['indicator_label'])

# Visualizar gráficos de líneas
fig_proposito_hist = px.line(
    df_proposito_hist,
    x='year',
    y='value',
    color='indicator_label',
    facet_col='reporter_area_label',
    facet_col_wrap=3,
    facet_row_spacing=0.15,
    facet_col_spacing=0.08,
    title="<b>Distribución histórica con los datos disponibles en los últimos años</b> ",
    labels={'value': 'Llegadas (miles)', 'year': 'Año', 'indicator_label': 'Propósito'},
    category_orders={"indicator_label": ["Ocio / Cultura", "Negocios"]},
    template="plotly_dark",
    color_discrete_map=color_map_proposito
)
fig_proposito_hist.update_xaxes(showticklabels=True, matches=None)
fig_proposito_hist.for_each_annotation(lambda a: a.update(text=f"<b>{a.text.split('=')[-1]}</b>"))
fig_proposito_hist.update_layout(
    margin=dict(t=100, b=80, l=50, r=50),
    showlegend=True,
    legend=dict(
        yanchor="bottom",
        y=0.01,
        xanchor="right",
        x=0.99,
        bgcolor="rgba(0,0,0,0)",
        borderwidth=0,
        title_text=""
    )
)
fig_proposito_hist.show()


# Filtrar los datos por origen del visitante
indicador_total = 'inbound - trips - by area of residence - total - overnight visitors (tourists)'
df_hubs = df[
    (df['reporter_area_label'].isin(africa_central)) &
    (df['indicator_label'] == indicador_total)
].copy()

# Calculamos los 3 principales destinos de los últimos 5 años
hubs_data = df_hubs[df_hubs['year'] >= 2015].groupby('reporter_area_label')['value'].sum().sort_values(ascending=False).head(3)
df_pie = hubs_data.reset_index()

# Visualizar gráfico de anillos
fig_hubs = px.pie(
    df_pie,
    names='reporter_area_label',
    values='value',
    hole=0.6,
    title="<b>Principales 3 hubs para inversión:</b> capacidad para atraer visitantes (cuota llegadas internacionales)",
    template="plotly_dark",
    color='reporter_area_label',
    color_discrete_map=color_map
)
fig_hubs.update_traces(
    textfont_size=14,
    textfont_color="white"
)
fig_hubs.show()

# Consolidar en un archivo HTML
with open('index.html', 'a') as f:
    f.write(fig_contexto.to_html(full_html=False, include_plotlyjs='cdn'))
    f.write(fig.to_html(full_html=False, include_plotlyjs='cdn'))
    f.write(fig_proposito.to_html(full_html=False, include_plotlyjs='cdn'))
    f.write(fig_proposito_hist.to_html(full_html=False, include_plotlyjs='cdn'))
    f.write(fig_hubs.to_html(full_html=False, include_plotlyjs='cdn'))
