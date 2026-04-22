# 🚗 Dashboard UBER — Ingreso / Egreso

Dashboard interactivo para análisis financiero de operación UBER (2025/2026).  
Construido con **Streamlit** + **Pandas** + **Plotly**.

---

## ✨ Funcionalidades

- KPIs de ingresos, gastos operativos e inversión separados
- Gráfica de conductores únicos por semana (sin sobreconteo)
- Separación automática de **INVERSIÓN** vs **gasto operativo**
- Análisis de amortización de flota (PAGO DE SUBASTA)
- Filtros por año, semana, socio y conductor
- Exportación a Excel del análisis de inversión

---

## 🚀 Deploy en Render (recomendado)

### Pasos desde Render Dashboard

1. Ve a [render.com](https://render.com) → **New +** → **Web Service**
2. Conecta tu repositorio de GitHub/GitLab
3. Completa los campos:

| Campo | Valor |
|-------|-------|
| **Name** | `uber-ingreso-egreso-dashboard` *(o el que prefieras)* |
| **Environment** | `Python 3` |
| **Region** | Oregon (o Frankfurt) |
| **Branch** | `main` |
| **Build Command** | `pip install -r requirements.txt` |
| **Start Command** | `streamlit run dashboard.py --server.port $PORT --server.address 0.0.0.0 --server.headless true` |

4. En **Environment Variables** agrega:

| Variable | Valor |
|----------|-------|
| `PYTHON_VERSION` | `3.12.0` |
| `STREAMLIT_SERVER_HEADLESS` | `true` |
| `PYTHONUNBUFFERED` | `1` |

5. Click en **Create Web Service**

> ℹ️ Render detectará el `render.yaml` automáticamente si conectas el repo con Blueprint.

---

### ⚡ Deploy rápido con Blueprint (render.yaml)

Si tu repo tiene el archivo `render.yaml` incluido, puedes usar el botón de Blueprint en Render:

1. En Render Dashboard → **New +** → **Blueprint**
2. Conecta el repo → Render leerá `render.yaml` y configurará todo automáticamente

---

## 📂 Uso de la app en producción

La app **no incluye datos** en el repo (los archivos Excel son privados).  
Al abrir el dashboard en producción:

1. En el panel lateral **"📥 Datos"**, haz clic en **"Browse files"**
2. Sube tu archivo Excel con las hojas:
   - `UBER 2025` / `UBER 2026` — datos de ingresos
   - `Gastos 2025` / `Gastos 2026` — datos de egresos
3. La app cargará y procesará automáticamente

---

## 💻 Desarrollo local

```bash
# 1. Clonar el repositorio
git clone <tu-repo-url>
cd UBER_INGRESO_EGRESO

# 2. Crear entorno virtual
python -m venv .venv
.\.venv\Scripts\activate   # Windows
# source .venv/bin/activate  # macOS/Linux

# 3. Instalar dependencias
pip install -r requirements.txt

# 4. Colocar tu archivo de datos
# Copia tu Excel como: prueba.xlsx  (en la raíz del proyecto)

# 5. Ejecutar
streamlit run dashboard.py
```

---

## 📁 Estructura del proyecto

```
UBER_INGRESO_EGRESO/
├── dashboard.py          # App Streamlit principal
├── business_rules.py     # Lógica de negocio y transformaciones
├── requirements.txt      # Dependencias Python
├── render.yaml           # Configuración de deploy en Render
├── runtime.txt           # Versión de Python para Render
├── .streamlit/
│   └── config.toml       # Configuración de Streamlit para producción
├── .gitignore            # Excluye Excel y archivos sensibles
├── .env.example          # Plantilla de variables de entorno
└── README.md             # Este archivo
```

---

## 🔐 Seguridad

- Los archivos `.xlsx` están excluidos del repo (`.gitignore`) porque contienen datos financieros
- No se exponen credenciales ni llaves API
- El dashboard no persiste datos entre sesiones (stateless)

---

## 📦 Dependencias principales

| Paquete | Versión | Uso |
|---------|---------|-----|
| `streamlit` | 1.52.1 | Framework de la app |
| `pandas` | 2.3.3 | Procesamiento de datos |
| `numpy` | 2.3.4 | Cálculos numéricos |
| `plotly` | 6.2.0 | Gráficas interactivas |
| `openpyxl` | 3.1.5 | Lectura de archivos Excel |

---

## ⚠️ Notas sobre el plan gratuito de Render

- Los servicios gratuitos se **duermen** después de 15 minutos de inactividad
- El primer request tras el sueño tarda ~30-60 segundos en despertar
- Si necesitas que esté siempre activo, usa el plan **Starter** ($7/mes)
- Los archivos subidos **no persisten** entre reinicios del servicio (sube el Excel en cada sesión)
