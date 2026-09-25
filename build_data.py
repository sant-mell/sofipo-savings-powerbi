"""Builds the CSVs, the scenario grid and the Excel workbook for the SOFIPO savings project.

Run: python3 build_data.py   (needs openpyxl for the .xlsx)
All figures were collected on 2026-09-24; see the source_url column of each file.
"""
import csv
import datetime as dt
import os

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "data")
ACC = "2026-09-24"

# ---------------- sources ----------------
S = {
    "banxico": "https://www.banxico.org.mx/publicaciones-y-prensa/anuncios-de-las-decisiones-de-politica-monetaria/anuncios-politica-monetaria-t.html",
    "isr": "https://www.siemprecontable.net/blog/retencion-isr-intereses-2026",
    "isr2": "https://russellbedford.mx/fiscal/aumento-en-la-tasa-de-retencion-por-intereses-en-el-ejercicio-2026/",
    "uma": "https://www.inegi.org.mx/contenidos/saladeprensa/boletines/2026/uma/uma2026.pdf",
    "exempt": "https://es-us.noticias.yahoo.com/sofipos-impuestos-2026-cu%C3%A1nto-te-170000926.html",
    "lisr": "https://mexico.justia.com/federales/leyes/ley-del-impuesto-sobre-la-renta/titulo-iv/capitulo-vi/",
    "udi": "https://calculadora-udi.com/",
    "cnbv": "https://www.gob.mx/cnbv/articulos/seguros-de-deposito-que-protegen-tus-ahorros-en-instituciones-reguladas-y-supervisadas-por-cnbv?idiom=es",
    "nu_bank": "https://www.nmas.com.mx/economia/finanzas-personales/nu-ya-es-un-banco-en-mexico-esto-significa-para-cuentahabientes-y-que-cambia-en-2026/",
    "tasas_nu": "https://www.tasas.mx/institucion/nu",
    "ef_feb": "https://www.elfinanciero.com.mx/mis-finanzas/2026/02/14/donde-invertir-mi-dinero-en-2026-estos-son-los-rendimientos-de-cetes-nu-revolut-y-otras-fintech/",
    "exp_jul": "https://expansion.mx/finanzas-personales/2026/07/24/nu-klar-mercado-pago-cetes-cual-da-mas-rendimiento",
    "tasas_mp": "https://www.tasas.mx/institucion/mercadopago",
    "tasas_klar": "https://www.tasas.mx/institucion/klar",
    "tasas_rev": "https://www.tasas.mx/institucion/revolut",
    "uala": "https://www.rendimientosmexico.com/instrumentos/uala",
    "decero": "https://deceroalinfinito.com/mejores-sofipos",
    "tasas": "https://www.tasas.mx/",
    "plata": "https://www.rendimientosmexico.com/instrumentos/plata",
    "plata_launch": "https://www.jornada.com.mx/noticia/2026/08/12/economia/plata-entra-a-las-cuentas-de-ahorro-ofrecera-rendimientos-de-hasta-15",
    "openbank": "https://www.tasas.mx/institucion/openbank",
    "openbank_alt": "https://www.cronista.com/mexico/finanzas-economia/openbank-desafia-a-las-fintechs-en-mexico-y-reedita-una-guerra-de-tasas-con-13-y-nuevo-tope-de-40-mil-pesos/",
    "sic_tax": "https://gbm.com/faqs/como-funcionan-los-impuestos-por-las-acciones-de-empresas-extranjeras-en-el-sic/",
    "gbm_fees": "https://gbm.com/faqs/que-comisiones-cobran-al-invertir-en-gbm/",
    "ivvpeso": "https://www.blackrock.com/mx/intermediarios/productos/268675/ishares-sp-500-peso-hedged-trac-fund",
    "spx_2025": "https://www.ftportfolios.com/Commentary/EconomicResearch/2026/1/8/the-sp-500-index-2025-recap",
    "spx_ytd": "https://chartrow.com/sp500/ytd",
    "fx_2025": "https://www.elfinanciero.com.mx/mercados/2025/12/31/peso-dolar-precio-hoy-31-de-diciembre-de-2025/",
    "fx_fix_2025": "https://dof.gob.mx/indicadores_detalle.php?cod_tipo_indicador=158&dfecha=31%2F12%2F2025&hfecha=31%2F12%2F2025",
    "fx_now": "https://tradingeconomics.com/mexico/currency",
    "art93": "https://www.contadigital.mx/posts/isr-sobre-intereses-personas-fisicas",
    "channel": "https://www.youtube.com/@miamigodinero",
}
# Germán Mi Amigo Dinero, monthly rate videos. Figures taken from spoken auto-captions only
# (the on-screen tables are not in the captions), so each row is a statement made in the video.
V = {
    "2025-10-04": ("7jNCqK43Fkc", "¿Quién PAGA MÁS en OCTUBRE 2025?"),
    "2025-11-04": ("FYmxm2lTYvQ", "¿Quién PAGA MÁS en NOVIEMBRE 2025?"),
    "2025-12-03": ("-rDq0PLHGs4", "¿Quién PAGA MÁS en DICIEMBRE 2025?"),
    "2026-01-07": ("zOLucpISr7c", "¿Quién PAGA MÁS en ENERO 2026?"),
    "2026-02-04": ("tqd4ME8JhSM", "¿Quién PAGA MÁS en FEBRERO 2026?"),
    "2026-03-04": ("ojcXlOxs5wI", "¿Quién PAGA MÁS en MARZO 2026?"),
    "2026-04-04": ("lSHPp29G8jo", "DÓNDE invertir en abril 2026"),
    "2026-05-07": ("DyNDCgcdDdM", "Mis inversiones FAVORITAS en mayo 2026"),
    "2026-06-04": ("4-ZB5rI_WnA", "Las Tasas Cambiaron... Junio"),
    "2026-07-05": ("YcVJv9_32r4", "Las Tasas Cambiaron... Julio"),
    "2026-08-07": ("xoA6vk1q0Mo", "Las Tasas Cambiaron... Agosto 2026"),
    "2026-09-05": ("2cCeCmegbkI", "Las Tasas Cambiaron... SEPTIEMBRE 2026"),
}
def vsrc(date):
    vid, title = V[date]
    return f"Germán Mi Amigo Dinero, {title}, {date}, https://www.youtube.com/watch?v={vid}"

# ---------------- tax and protection parameters ----------------
UDI = 8.823864                      # MXN per UDI on 2026-09-24
UMA_ANNUAL = 42794.64               # UMA 2026, annual
SOFIPO_EXEMPT = 5 * UMA_ANNUAL      # 213,973.20
WHT = 0.009                         # 0.90% annual provisional withholding, on capital
PROSOFIPO_MXN = round(25000 * UDI, 2)
IPAB_MXN = round(400000 * UDI, 2)

TAX = [
    # param, value, unit, applies_from, notes, source, verified
    ("isr_withholding_rate", WHT, "annual rate on invested capital", "2026-01-01",
     "Provisional withholding, credited in the annual return (final tax is on real interest). LIF 2026; article number differs between sources (17 vs 24).",
     S["isr"], "yes (secondary sources agree on 0.90%)"),
    ("isr_withholding_rate_2025", 0.005, "annual rate on invested capital", "2025-01-01",
     "Previous year's rate, for comparison. Also stated in the Dec 2025 video (-rDq0PLHGs4).", S["isr"], "yes (secondary)"),
    # earlier years, from the video series (history only; the model reads isr_withholding_rate)
    ("isr_withholding_rate_2020", 0.0145, "annual rate on invested capital", "2020-01-01",
     "History only. Matches the Ley de Ingresos rate; no video found for it.", "", "not verified in videos"),
    ("isr_withholding_rate_2021", 0.0097, "annual rate on invested capital", "2021-01-01",
     "History only.", "https://www.youtube.com/watch?v=RcEoeHDvND8", "video (auto-captions)"),
    ("isr_withholding_rate_2022", 0.0008, "annual rate on invested capital", "2022-01-01",
     "History only. Big cut from 2021.", "https://www.youtube.com/watch?v=RcEoeHDvND8", "video (auto-captions)"),
    ("isr_withholding_rate_2023", 0.0015, "annual rate on invested capital", "2023-01-01",
     "History only.", "https://www.youtube.com/watch?v=fn7DOZ3qKQ0", "video (auto-captions)"),
    ("isr_withholding_rate_2024", 0.005, "annual rate on invested capital", "2024-01-01",
     "History only. 1.48% was proposed first (https://www.youtube.com/watch?v=00F_xLy5eRk).",
     "https://www.youtube.com/watch?v=Y9y3h_eOuGE", "video (auto-captions)"),
    ("uma_daily_2026", 117.31, "MXN", "2026-02-01", "INEGI, DOF 2026-01-09.", S["uma"], "yes"),
    ("uma_annual_2026", UMA_ANNUAL, "MXN", "2026-02-01", "117.31 x 365 (rounded as published).", S["uma"], "yes"),
    ("sofipo_exempt_limit_mxn", round(SOFIPO_EXEMPT, 2), "MXN average daily balance", "2026-02-01",
     "LISR art. 93 fr. XX: SOFIPO interest exempt if the average daily balance across all SOFIPOs is at most 5 annual UMAs; above it, only the excess is taxed.",
     S["exempt"], "yes (secondary; statute not fetched directly)"),
    ("bank_savings_exemption", 1, "flag", "",
     "Art. 93 fr. XX also exempts interest from bank savings/deposit accounts up to the same 5 UMAs (not fixed terms). Whether a bank applies it varies; the Sep 2026 video says Nu's Cajita Turbo is not withheld while Nu fixed terms are. Other banks' on-demand accounts are marked 'unknown' and taxed in the model.",
     S["art93"], "secondary + video"),
    ("udi_value", UDI, "MXN per UDI", ACC, "Changes daily with inflation.", S["udi"], "yes"),
    ("prosofipo_limit_udis", 25000, "UDIs", "", "Per person, per SOFIPO.", S["cnbv"], "yes (secondary)"),
    ("prosofipo_limit_mxn", PROSOFIPO_MXN, "MXN", ACC, "25,000 x UDI value.", S["udi"], "computed"),
    ("ipab_limit_udis", 400000, "UDIs", "", "Per person, per bank.", S["cnbv"], "yes (secondary)"),
    ("ipab_limit_mxn", IPAB_MXN, "MXN", ACC, "400,000 x UDI value.", S["udi"], "computed"),
]

# ---------------- institutions ----------------
INST = [
    # name, type, protection, protection_mxn, notes, source, verified
    ("Nu", "Bank", "IPAB", IPAB_MXN,
     "SOFIPO until it started operating as a bank on 2026-08-06 (CNBV authorization 2026-07-10). Interest now subject to withholding like any bank.",
     S["nu_bank"], "yes"),
    ("Revolut", "Bank", "IPAB", IPAB_MXN, "Local banking license.", S["tasas_rev"], "yes"),
    ("Ualá", "Bank", "IPAB", IPAB_MXN, "Operates through ABC Capital, a bank it acquired. Reported a 563 mdp net loss in H1 2026.",
     S["uala"], "secondary"),
    ("Banco Plata", "Bank", "IPAB", IPAB_MXN,
     "Bank authorized by CNBV; launched savings in Aug 2026. Plata+ membership costs 99 MXN + IVA per month. 15% Ahorro Ultra is only for the first 60 days of a new account.",
     S["plata"], "secondary"),
    ("Openbank", "Bank", "IPAB", IPAB_MXN,
     "Santander's digital bank. 13% on the first 30,000 in Apartados Open. The cap was 40,000 until Aug 2026 (El Cronista, and the Aug 2026 video says it fell from 40,000 to 30,000).",
     S["openbank"], "yes (cap change confirmed by video)"),
    ("Mifel", "Bank", "IPAB", IPAB_MXN,
     "Cuenta Digital: about 10% on demand up to 500,000, no spending conditions (video, Mar and Sep 2026). Only source is the video series; check the official site.",
     vsrc("2026-09-05"), "video only"),
    ("Mercado Pago", "Fintech", "Not reported", "",
     "Not a SOFIPO or bank per tasas.mx; deposit protection not reported. The 25,000 cap is shared between the 12% and 15% options and balances above it earn nothing (Sep 2026 video). Tax treatment assumed equal to a bank (unverified).",
     S["tasas_mp"], "yes (type); tax treatment unverified"),
    ("Klar", "SOFIPO", "Prosofipo", PROSOFIPO_MXN,
     "Two products: Inversión Max pays 15% on up to 25,000 for Plus/Platino (2,500 monthly spend since Jun 2026), and the regular on-demand balance pays 8% (Plus) or 6%. That explains the 8% vs 15% gap between tasas.mx and other sources.",
     S["tasas_klar"], "resolved with video + Expansión"),
    ("DiDi", "SOFIPO", "Prosofipo", PROSOFIPO_MXN, "", S["tasas"], "yes"),
    ("Stori", "SOFIPO", "Prosofipo", PROSOFIPO_MXN, "", S["tasas"], "yes"),
    ("Finsus", "SOFIPO", "Prosofipo", PROSOFIPO_MXN, "", S["tasas"], "yes"),
    ("Kubo Financiero", "SOFIPO", "Prosofipo", PROSOFIPO_MXN,
     "Users reported blocked withdrawals in May-June 2026. Excluded from the default strategies.", S["decero"], "yes"),
    ("Supertasas", "SOFIPO", "Prosofipo", PROSOFIPO_MXN,
     "Covered only through the Germán Mi Amigo Dinero video series (rate history only, no current product). Mostly fixed terms.",
     S["channel"], "video only"),
    ("CETES (via GBM or Cetesdirecto)", "Government", "Federal government", "",
     "Government debt, no deposit insurance needed. Buying through GBM may add fees (not modeled). Withholding assumed at the 0.90% rate.",
     S["tasas"], "rates yes; GBM fees unverified"),
]
INST_TYPE = {r[0]: r[1] for r in INST}
INST_PROT = {r[0]: r[3] for r in INST}

# ---------------- rate history ----------------
# date, institution, product, term_days (0 = on demand), rate, cap, rate_on_excess, conditions, source, verified
H = []
def h(date, inst, prod, term, rate, cap="", excess="", cond="", src="", ver="yes"):
    H.append((date, inst, prod, term, rate, cap, excess, cond, src, ver))

# Feb 2026 snapshot (El Financiero, 2026-02-14)
f = "2026-02-14"
h(f, "Nu", "Cajita (on demand)", 0, 0.07, src=S["ef_feb"], ver="secondary", cond="Nu was a SOFIPO at this date")
h(f, "Nu", "Cajita Turbo", 0, 0.13, 25000, 0.07, "Monthly debit card purchase", S["ef_feb"], "secondary")
for t, r in [(7, 0.0705), (28, 0.071), (90, 0.072), (180, 0.073)]:
    h(f, "Nu", f"Fixed term {t}d", t, r, src=S["ef_feb"], ver="secondary")
h(f, "Revolut", "Savings (on demand)", 0, 0.15, 25000, 0.07, "0.07 to 0.075 on 25k-1M; 0.05 above 1M", S["ef_feb"], "secondary")
h(f, "Mercado Pago", "On demand", 0, 0.13, 25000, "", "At least 3,000 MXN monthly deposits", S["ef_feb"], "secondary")
h(f, "Ualá", "On demand", 0, 0.07, 30000, "", "Base; 12% with 3k monthly spend, 15% with 6k", S["ef_feb"], "secondary")
h(f, "Klar", "On demand (Plus)", 0, 0.15, "", "", "Klar Plus with 3,000 MXN monthly transactions", S["ef_feb"], "secondary")
for t, r in [(28, 0.0688), (91, 0.07), (182, 0.071), (364, 0.0737)]:
    h(f, "CETES (via GBM or Cetesdirecto)", f"CETES {t}d", t, r, src=S["ef_feb"], ver="secondary")

# Jul 2026 snapshot (Expansion, 2026-07-24)
j = "2026-07-24"
h(j, "Nu", "Cajita (on demand)", 0, 0.065, src=S["exp_jul"], ver="yes")
h(j, "Nu", "Cajita Turbo", 0, 0.13, 25000, 0.065, "Monthly debit card purchase", S["exp_jul"], "yes")
h(j, "Nu", "Fixed term 7d", 7, 0.0655, src=S["exp_jul"], ver="yes")
h(j, "Nu", "Fixed term 180d", 180, 0.068, src=S["exp_jul"], ver="yes")
h(j, "Mercado Pago", "On demand", 0, 0.12, 25000, 0.06, "3,000+ MXN monthly deposits, otherwise 6%", S["exp_jul"], "yes")
h(j, "Klar", "On demand (Plus/Platino)", 0, 0.15, "", "", "2,500 MXN monthly card spend; 6.5% for other users", S["exp_jul"], "yes")
h(j, "Kubo Financiero", "Fixed term 365d", 365, 0.12, src=S["exp_jul"], ver="yes")
for t, r in [(28, 0.0618), (91, 0.0649), (182, 0.0675), (364, 0.0693)]:
    h(j, "CETES (via GBM or Cetesdirecto)", f"CETES {t}d", t, r, src=S["exp_jul"], ver="yes")

# Aug-Sep 2026
h("2026-08-16", "Mercado Pago", "On demand", 0, 0.12, 25000, "", "", S["tasas_mp"], "yes")
h("2026-09-01", "Mercado Pago", "Apartados with Meli+", 0, 0.15, 25000, "", "Paid Meli+ subscription; 12% without it", S["tasas_mp"], "yes")
h("2026-08-29", "Revolut", "Savings (on demand)", 0, 0.15, 25000, 0.07,
  "4 purchases of at least 50 MXN every 30 days, or Premium/Metal plan; excess earns 7-7.5%", S["tasas_rev"], "yes")

h("2026-08-12", "Banco Plata", "Ahorro Ultra (new accounts)", 0, 0.15, 25000, 0.07, "First 60 days only", S["plata_launch"], "yes")
h("2026-08-12", "Banco Plata", "Ahorro Flexible (Plata+)", 0, 0.10, "", "", "Plata+ membership", S["plata"], "secondary")

# Germán Mi Amigo Dinero statements (auto-captions; ambiguous figures skipped)
def hv(date, inst, prod, term, rate, cap="", excess="", cond=""):
    h(date, inst, prod, term, rate, cap, excess, cond, vsrc(date), "video (auto-captions)")
hv("2025-10-04", "Klar", "Inversión Max (Plus/Platino)", 0, 0.15, 25000, "", "New product this month")
hv("2025-10-04", "Openbank", "Apartados Open", 0, 0.10, "", "", "No cap")
hv("2025-11-04", "DiDi", "On demand", 0, 0.16, 10000)
hv("2025-11-04", "Klar", "Inversión Max (Plus/Platino)", 0, 0.15, 25000, "", "Spend 3,000 monthly")
hv("2025-11-04", "Mercado Pago", "On demand", 0, 0.13, 25000, "", "Most users; some got 15% in Apartados")
hv("2025-11-04", "Openbank", "Apartados Open", 0, 0.10, "", "", "Guaranteed for November")
hv("2025-11-04", "Mifel", "Cuenta Digital", 0, 0.10, "", "", "Valid until 2026-04-30")
hv("2025-12-03", "DiDi", "On demand", 0, 0.16, 10000)
hv("2025-12-03", "Nu", "Cajita Turbo", 0, 0.15, 25000)
hv("2025-12-03", "Klar", "Inversión Max (Plus/Platino)", 0, 0.15, 25000, "", "Spend 3,000 monthly or Platino card")
hv("2025-12-03", "Openbank", "Apartados Open", 0, 0.10, "", "", "No cap")
hv("2026-01-07", "Mercado Pago", "On demand", 0, 0.13, 25000, "", "Deposit 3,000 monthly")
hv("2026-01-07", "Mifel", "Cuenta Digital", 0, 0.10)
hv("2026-02-04", "Openbank", "Apartados Open", 0, 0.09)
hv("2026-02-04", "Revolut", "Savings (on demand)", 0, 0.15, 25000, "", "Launch; no condition")
hv("2026-02-04", "Klar", "Inversión Max (Plus/Platino)", 0, 0.15, 25000, "", "Spend 3,000 monthly")
hv("2026-03-04", "Mifel", "Cuenta Digital", 0, 0.10, 500000, "", "No conditions")
hv("2026-03-04", "Openbank", "Apartados Open", 0, 0.09)
hv("2026-03-04", "Klar", "Fixed term 365d", 365, 0.10)
hv("2026-04-04", "Banco Plata", "Ahorro Flexible (Plata+)", 0, 0.12, "", "", "Launch")
hv("2026-04-04", "Banco Plata", "Fixed term 30d (Plata+)", 30, 0.09, "", "", "7% without membership")
hv("2026-05-07", "Mercado Pago", "On demand", 0, 0.10, "", "", "25,000 cap removed mid April")
hv("2026-05-07", "Openbank", "Apartados Open", 0, 0.13, 40000)
hv("2026-05-07", "Banco Plata", "Ahorro Flexible (Plata+)", 0, 0.10)
hv("2026-05-07", "Klar", "Fixed term 7d", 7, 0.081)
hv("2026-06-04", "Nu", "Cajita Turbo", 0, 0.13, 25000)
hv("2026-06-04", "Klar", "Inversión Max (Plus/Platino)", 0, 0.15, 25000, "", "Spend requirement cut to 2,500")
hv("2026-07-05", "Kubo Financiero", "Fixed term 365d", 365, 0.12, "", "", "No cap, no conditions")
hv("2026-08-07", "Openbank", "Apartados Open", 0, 0.13, 30000, "", "Cap cut from 40,000")
hv("2026-08-07", "Banco Plata", "Ahorro Ultra (new accounts)", 0, 0.15, 25000)
hv("2026-08-07", "Banco Plata", "Ahorro Flexible (Plata+)", 0, 0.09, "", "", "Was 10% with no cap")
hv("2026-09-05", "Mercado Pago", "Apartados with Meli+", 0, 0.15, 25000, 0.0, "Cap shared with the 12% option; nothing above it")
hv("2026-09-05", "Mercado Pago", "On demand (no Meli+)", 0, 0.12, 25000, 0.0, "Deposit 3,000 or spend 1,000-3,000 monthly")
hv("2026-09-05", "Kubo Financiero", "1-day term", 1, 0.10, "", "", "Cut from 13%")
hv("2026-09-05", "Finsus", "Cosecha 4 months", 120, 0.11, "", "", "No membership")
hv("2026-09-05", "DiDi", "On demand", 0, 0.15, 10000, "", "Only 15% with no condition or time limit")

# Same channel, earlier monthly roundups (Mar 2023 - Feb 2026) and 2026 topic videos, from the
# hand-checked tables of the 402-video review. Keyed by video id because some share a date.
# Term left blank ("") when the video does not state it; a range keeps its top value.
VX = {
    "gmy1pM3RSQY": ("2023-03-02", "¿Quién PAGA MÁS en cada plazo? (Marzo 2023)"),
    "qABJs5rFw7A": ("2023-04-02", "¿Quién PAGA MÁS en ABRIL 2023? CETES, SOFIPOS, Hey Banco, mercado pago, etc."),
    "mx8pE7Ohydk": ("2023-06-02", "¿Quién PAGA MÁS en JUNIO 2023? CETES, SOFIPOS, Hey Banco, KLAR, etc."),
    "0UVRneDYElU": ("2023-08-03", "CETES detienen su caída! - ¿Quién PAGA MÁS en Agosto 2023?"),
    "_JHrncIaD2k": ("2023-11-03", "¿Quién PAGA MÁS en noviembre 2023? - Opciones con 15% de rendimiento, 200% más ISR"),
    "JWT4v1Y1v9A": ("2024-01-03", "¿Quién PAGA MÁS en ENERO 2024? - CETES y FINAMEX bajan"),
    "4t6Kf4HouwU": ("2024-02-02", "¿Quién PAGA MÁS en FEBRERO 2024? - Cajitas en NU, KLAR sube mucho, DINN baja tasa"),
    "eFXqSq8hqS8": ("2024-03-02", "¿Quién PAGA MÁS en MARZO 2024? - CETES bajan, Finsus sube a 16%, ¿Banxico bajará?"),
    "axxIJ237U9s": ("2024-04-02", "¿Quién PAGA MÁS en ABRIL 2024? - CAÍDAS por todos lados"),
    "D92SpdExA9k": ("2024-05-02", "¿Quién PAGA MÁS en MAYO 2024? - SIGUEN LAS CAÍDAS"),
    "pNDU4jQxVac": ("2024-08-03", "¿Quién PAGA MÁS en AGOSTO 2024? - Los CETES quedan a deber"),
    "Ju5oYtKpAKo": ("2024-10-03", "¿Quién PAGA MÁS en OCTUBRE 2024? - NU baja tasas, meli dólar llega, stori tiene plazos"),
    "3nRL92xKGy0": ("2025-01-04", "¿Quién PAGA MÁS en ENERO 2025? - CETES recupera el 10%, NU baja tasas, KLAR cambia mucho"),
    "Om-l5Akl7to": ("2025-02-04", "¿Quién PAGA MÁS en FEBRERO 2025? - NU domina, ARANCELES, y constancias de impuestos"),
    "l7D2bz8tBtg": ("2025-04-03", "¿Quién PAGA MÁS en ABRIL 2025? - CETES CAEN, Now Bank es nueva opción, NU cambia de tasa pronto"),
    "AH2vQLe0dQQ": ("2025-05-03", "¿Quién PAGA MÁS en MAYO 2025? - SOFIPO quiebra, mercado pago baja, NU será banco"),
    "w1QWl47g0uc": ("2025-06-05", "¿Quién PAGA MÁS en JUNIO 2025? - DIDI dará 15%, Novedades en NU, KLAR baja tasas"),
    "3qbfW5NrZog": ("2025-07-03", "¿Quién PAGA MÁS en JULIO 2025? - Cashback en NU+, 16% en ualá, didi al 15%"),
    "w8--F9sqbVw": ("2025-08-05", "¿Quién PAGA MÁS en AGOSTO 2025? - DESPLOME en NU, Stori, KLAR, los CETES se recuperan"),
    "ivm1_uAzvuY": ("2026-04-09", "NU ya se RINDIÓ, aquí GANAS MÁS"),
    "XTDcKIUcak4": ("2026-06-24", "Adiós Revolut: Se nos Cayó un Grande!"),
    "UurM0cORmNI": ("2026-07-02", "Revolut se Rindió. Estas son las Alternativas."),
    "ABZytAfldQk": ("2026-07-11", "KUBO Lanza una \"Cajita Turbo\" SIN TOPE!!"),
    "U6m5NDabFec": ("2026-07-16", "Las Mejores Inversiones en 2026, Rankeadas"),
    "jzhrIrAru-E": ("2026-07-30", "Fue un ERROR Aceptar las Inversiones Topadas"),
    "pkGJ4boI6tk": ("2026-08-07", "Cambios en las Cajitas NU Explicados"),
    "FnZBsYYLf4k": ("2026-08-24", "La Didi Cuenta Debería Preocuparnos"),
    "hvP_dB63NT8": ("2026-09-17", "Así Ganas $6,297.04 al Mes Sin Hacer Nada"),
}
def hx(vid, inst, prod, term, rate, cap="", excess="", cond=""):
    date, title = VX[vid]
    src = f"Germán Mi Amigo Dinero, {title}, {date}, https://www.youtube.com/watch?v={vid}"
    h(date, inst, prod, term, rate, cap, excess, cond, src, "video (auto-captions)")
HR = "Headline rate (video)"
NT = "term not stated in video"
NS = "Nu was a SOFIPO at this date"
# 2023: Finsus fixed terms lead, Stori starts the 15% era
hx("gmy1pM3RSQY", "Finsus", HR, "", 0.13, cond=f"Fixed terms, 11-13% by term; {NT}")
hx("qABJs5rFw7A", "Finsus", HR, "", 0.11, cond=f"Fixed terms, 10-11%; {NT}")
hx("mx8pE7Ohydk", "Finsus", HR, "", 0.14, cond=NT)
hx("0UVRneDYElU", "Finsus", HR, "", 0.1455, cond=f"Fixed term; {NT}")
hx("_JHrncIaD2k", "Finsus", HR, "", 0.15, cond=f"Fixed term; {NT}")
hx("_JHrncIaD2k", "Stori", "On demand", 0, 0.15, cond="No cap, no conditions; start of the 15% rate war")
# 2024: the 15% era, then cuts from April
hx("JWT4v1Y1v9A", "Nu", HR, "", 0.15, cond=f"{NT}; {NS}")
hx("4t6Kf4HouwU", "Klar", "On demand", 0, 0.15, cond="Klar rises a lot this month")
hx("4t6Kf4HouwU", "Supertasas", HR, "", 0.09, cond=f"Mostly fixed terms; {NT}")
hx("eFXqSq8hqS8", "Nu", "Cajita (on demand)", 0, 0.15, cond=f"Cajitas launched Feb 2024; {NS}")
hx("eFXqSq8hqS8", "Stori", "On demand", 0, 0.15)
hx("eFXqSq8hqS8", "Finsus", HR, "", 0.16, cond=f"Long fixed terms; {NT}")
hx("axxIJ237U9s", "Nu", "Cajita (on demand)", 0, 0.15, cond=NS)
hx("axxIJ237U9s", "Finsus", HR, "", 0.15, cond=f"Fixed term; {NT}")
hx("D92SpdExA9k", "Klar", "On demand", 0, 0.15)
hx("pNDU4jQxVac", "Nu", HR, "", 0.15, cond=f"{NT}; {NS}")
hx("pNDU4jQxVac", "Klar", "On demand", 0, 0.10)
hx("Ju5oYtKpAKo", "Stori", HR, "", 0.15, cond=f"Stori adds fixed terms this month; {NT}")
# 2025: 15% only capped or conditional
hx("3nRL92xKGy0", "Klar", "On demand", 0, 0.10, cond="Depends on account level")
hx("Om-l5Akl7to", "Stori", HR, "", 0.15, cond=f"Fixed term; {NT}")
hx("l7D2bz8tBtg", "Nu", HR, "", 0.14, cond=f"{NT}; {NS}")
hx("AH2vQLe0dQQ", "Nu", HR, "", 0.15, cond=f"9-15% depending on product and cap; {NT}; {NS}")
hx("w1QWl47g0uc", "DiDi", "On demand", 0, 0.15, cond="Announced: DiDi to pay 15%")
hx("3qbfW5NrZog", "Nu", HR, "", 0.14, cond=f"12-14%, capped; {NT}; {NS}")
hx("3qbfW5NrZog", "Finsus", HR, "", 0.14, cond=f"Fixed term; {NT}")
hx("3qbfW5NrZog", "Stori", HR, "", 0.08, cond=NT)
hx("w8--F9sqbVw", "Nu", HR, "", 0.08, cond=f"Sharp drop; {NT}; {NS}")
hx("w8--F9sqbVw", "Klar", "On demand", 0, 0.06, cond="Sharp drop; base level")
# Oct 2025 - Feb 2026 roundups: figures not already recorded above
hv("2025-10-04", "Kubo Financiero", HR, "", 0.14, cond=NT)
hv("2025-10-04", "Stori", HR, "", 0.125, cond=f"5-12.5% across products; {NT}")
hv("2025-11-04", "Finsus", HR, "", 0.09, cond=f"8-9%; {NT}")
hv("2025-11-04", "Kubo Financiero", HR, "", 0.085, cond=NT)
hv("2025-12-03", "Stori", HR, "", 0.15, cond=f"Stori rises a lot; {NT}")
hv("2025-12-03", "Finsus", HR, "", 0.135, cond=NT)
hv("2026-01-07", "Nu", "Cajita (on demand)", 0, 0.07, cond=f"Nu cuts; {NS}")
hv("2026-01-07", "Finsus", HR, "", 0.07, cond=NT)
hv("2026-02-04", "Kubo Financiero", HR, "", 0.1325, cond=NT)
# 2026 topic videos: caps, requirements and what sits behind the headline
hx("ivm1_uAzvuY", "Supertasas", HR, "", 0.076, cond=f"7.0-7.6% mentioned; mostly fixed terms; {NT}")
hv("2026-06-04", "Stori", "Fixed term 90d", 90, 0.105, cond="No requirement")
hx("XTDcKIUcak4", "Nu", "Cajita Turbo", 0, 0.13, cond=f"Capped (amount not stated) with an activation condition; {NS}")
hx("UurM0cORmNI", "Nu", "Cajita Turbo", 0, 0.13, cond=f"Capped (amount not stated) with an activation condition; {NS}")
hv("2026-07-05", "Kubo Financiero", "On demand", 0, 0.12, 25000, "", "About 3,000 MXN monthly deposits; cut to 12% this month")
hx("ABZytAfldQk", "Kubo Financiero", "Cajita Turbo", 0, 0.13, cond="No cap")
hx("U6m5NDabFec", "DiDi", "On demand", 0, 0.15, 10000, 0.075, "Excess about 7.5%; about 12.3% blended at 15,000; paid daily")
hx("jzhrIrAru-E", "Stori", "Fixed term 90d", 90, 0.10, cond="No requirement")
hx("jzhrIrAru-E", "Stori", "Fixed term 180d", 180, 0.10, cond="No requirement")
hv("2026-08-07", "Finsus", "Fixed term 120d (membership)", 120, 0.115, cond="Paid membership, about 500 MXN a year")
hx("pkGJ4boI6tk", "Nu", "Cajita (on demand)", 0, 0.06, cond="Regular Cajitas; Nu a bank since 2026-08-06")
hx("FnZBsYYLf4k", "DiDi", "On demand", 0, 0.15, 10000, 0.075, "Excess about 7.5%")
hx("hvP_dB63NT8", "Klar", HR, 0, 0.13, cond="13% tier tied to a 2,500 MXN requirement; terms got worse this month")

# Sep 24 2026 snapshot
s = ACC
h(s, "Nu", "Cajita (on demand)", 0, 0.065, src=S["nu_bank"], ver="yes")
h(s, "Nu", "Cajita Turbo", 0, 0.13, 25000, 0.065, "Monthly debit card purchase", S["nu_bank"], "yes")
h(s, "Nu", "Fixed term 7d", 7, 0.0655, src=S["tasas_nu"], ver="yes")
h(s, "Nu", "Fixed term 180d", 180, 0.068, src=S["tasas_nu"], ver="yes")
h(s, "Ualá", "On demand (Plus)", 0, 0.15, 30000, 0.0675, "Plus level; base account about 6.75%", S["uala"], "secondary")
h(s, "Klar", "On demand (Plus/Platino)", 0, 0.08, "", "", "Regular balance; 6% standard. The 15% is the separate Inversión Max product", S["tasas_klar"], "yes")
h(s, "Klar", "Inversión Max (Plus/Platino)", 0, 0.15, 25000, 0.08, "2,500 monthly spend; cap from video (Dec 2025)", S["decero"], "yes (cap from video)")
h(s, "Klar", "Fixed term 365d (Plus/Platino)", 365, 0.085, src=S["tasas_klar"], ver="yes", cond="Plus/Platino")
h(s, "Banco Plata", "Ahorro Flexible (base)", 0, 0.07, src=S["plata"], ver="secondary")
h(s, "Banco Plata", "Ahorro Flexible (Plata+)", 0, 0.09, "", "", "99 MXN + IVA monthly; cut from 10% in Aug 2026", S["plata"], "secondary")
h(s, "Banco Plata", "Fixed term 360d (Plata+)", 365, 0.11, src=S["plata"], ver="secondary", cond="Plata+; base terms 7-8%")
h("2026-09-14", "Openbank", "Apartados Open", 0, 0.13, 30000, 0.07, "Open+ debit account; 6.5% above 1M", S["openbank"], "conflict (see institutions)")
h(s, "Mercado Pago", "On demand (no Meli+)", 0, 0.12, 25000, "", "Without subscription", S["tasas_mp"], "yes")
h(s, "DiDi", "On demand", 0, 0.15, 10000, "", "Rate on balance above 10,000 not found", S["tasas"], "yes (excess rate unverified)")
h(s, "Stori", "On demand", 0, 0.0725, src=S["decero"], ver="yes")
for t, r in [(30, 0.0705), (90, 0.11), (180, 0.10), (365, 0.0825)]:
    h(s, "Stori", f"Fixed term {t}d", t, r, src=S["tasas"], ver="yes")
h(s, "Finsus", "On demand", 0, 0.0701, src=S["decero"], ver="yes")
for t, r in [(30, 0.0719), (90, 0.075), (180, 0.0759), (365, 0.0869)]:
    h(s, "Finsus", f"Fixed term {t}d", t, r, src=S["tasas"], ver="yes")
for t, r in [(90, 0.11), (365, 0.12)]:
    h(s, "Kubo Financiero", f"Fixed term {t}d", t, r, src=S["tasas"], ver="yes", cond="Withdrawal problems reported May-Jun 2026")
for t, r in [(28, 0.0615), (91, 0.0659), (182, 0.0691), (364, 0.0724)]:
    h(s, "CETES (via GBM or Cetesdirecto)", f"CETES {t}d", t, r, src=S["tasas"], ver="yes")

# ---------------- Banxico ----------------
BX = [("2024-02-08", .1125, 0), ("2024-03-21", .11, -25), ("2024-05-09", .11, 0), ("2024-06-27", .11, 0),
      ("2024-08-08", .1075, -25), ("2024-09-26", .105, -25), ("2024-11-14", .1025, -25), ("2024-12-19", .10, -25),
      ("2025-02-06", .095, -50), ("2025-03-27", .09, -50), ("2025-05-15", .085, -50), ("2025-06-26", .08, -50),
      ("2025-08-07", .0775, -25), ("2025-09-25", .075, -25), ("2025-11-06", .0725, -25), ("2025-12-18", .07, -25),
      ("2026-02-05", .07, 0), ("2026-03-26", .0675, -25), ("2026-05-07", .065, -25), ("2026-06-25", .065, 0),
      ("2026-08-06", .065, 0), ("2026-09-24", .065, 0)]

def banxico_monthly():
    out, d = [], dt.date(2024, 2, 1)
    while d <= dt.date(2026, 9, 1):
        nxt = (d.replace(day=28) + dt.timedelta(days=4)).replace(day=1)
        month_end = nxt - dt.timedelta(days=1)
        rate = [r for (x, r, _) in BX if dt.date.fromisoformat(x) <= month_end][-1]
        out.append((month_end.isoformat(), rate))
        d = nxt
    return out

# ---------------- current products (buckets for the allocation model) ----------------
# institution, product, term_days, rate, cap, rate_on_excess, needs_condition, include_default, monthly_fee_mxn
P0 = [
    ("Revolut", "Savings (on demand)", 0, 0.15, 25000, 0.07, "yes", "yes"),
    ("Ualá", "On demand (Plus)", 0, 0.15, 30000, 0.0675, "yes", "yes"),
    ("Mercado Pago", "Apartados with Meli+", 0, 0.15, 25000, 0.0, "yes (paid subscription)", "no"),
    ("Nu", "Cajita Turbo", 0, 0.13, 25000, 0.065, "yes", "yes"),
    ("DiDi", "On demand", 0, 0.15, 10000, "", "no", "yes"),
    ("Nu", "Cajita (on demand)", 0, 0.065, "", "", "no", "yes"),
    ("Stori", "On demand", 0, 0.0725, "", "", "no", "yes"),
    ("Finsus", "On demand", 0, 0.0701, "", "", "no", "yes"),
    ("Stori", "Fixed term 30d", 30, 0.0705, "", "", "no", "yes"),
    ("Stori", "Fixed term 90d", 90, 0.11, "", "", "no", "yes"),
    ("Stori", "Fixed term 180d", 180, 0.10, "", "", "no", "yes"),
    ("Stori", "Fixed term 365d", 365, 0.0825, "", "", "no", "yes"),
    ("Finsus", "Fixed term 30d", 30, 0.0719, "", "", "no", "yes"),
    ("Finsus", "Fixed term 90d", 90, 0.075, "", "", "no", "yes"),
    ("Finsus", "Fixed term 180d", 180, 0.0759, "", "", "no", "yes"),
    ("Finsus", "Fixed term 365d", 365, 0.0869, "", "", "no", "yes"),
    ("Nu", "Fixed term 180d", 180, 0.068, "", "", "no", "yes"),
    ("Kubo Financiero", "Fixed term 90d", 90, 0.11, "", "", "no", "no (withdrawal risk)"),
    ("Kubo Financiero", "Fixed term 365d", 365, 0.12, "", "", "no", "no (withdrawal risk)"),
    ("CETES (via GBM or Cetesdirecto)", "CETES 28d", 30, 0.0615, "", "", "no", "yes"),
    ("CETES (via GBM or Cetesdirecto)", "CETES 91d", 90, 0.0659, "", "", "no", "yes"),
    ("CETES (via GBM or Cetesdirecto)", "CETES 182d", 180, 0.0691, "", "", "no", "yes"),
    ("CETES (via GBM or Cetesdirecto)", "CETES 364d", 365, 0.0724, "", "", "no", "yes"),
    ("Openbank", "Apartados Open", 0, 0.13, 30000, 0.07, "no", "yes"),
    ("Banco Plata", "Ahorro Flexible (base)", 0, 0.07, "", "", "no", "yes"),
    ("Banco Plata", "Ahorro Flexible (Plata+)", 0, 0.09, "", "", "yes (99 MXN + IVA monthly)", "yes"),
    ("Banco Plata", "Fixed term 360d (Plata+)", 365, 0.11, "", "", "yes (99 MXN + IVA monthly)", "yes"),
    ("Banco Plata", "Ahorro Ultra (new accounts)", 0, 0.15, 25000, 0.07, "yes (first 60 days only)", "no"),
    ("Klar", "Inversión Max (Plus/Platino)", 0, 0.15, 25000, 0.08, "yes (2,500 monthly spend)", "yes"),
    ("Klar", "On demand (Plus/Platino)", 0, 0.08, "", "", "yes", "yes"),
    ("Mifel", "Cuenta Digital", 0, 0.10, 500000, "", "no", "yes"),
    ("Klar", "Fixed term 365d (Plus/Platino)", 365, 0.085, "", "", "yes", "yes"),
    ("Mercado Pago", "On demand (no Meli+)", 0, 0.12, 25000, 0.0, "yes", "no"),
]
PLATA_FEE = round(99 * 1.16, 2)


def exempt_flag(p):
    t = {r[0]: r[1] for r in INST}[p[0]]
    if t == "SOFIPO":
        return "yes"
    if t == "Bank" and p[2] == 0:
        return "yes" if p[0] == "Nu" else "unknown"
    return "no"


P = [p + ((PLATA_FEE if "Plata+" in p[1] else 0.0),) for p in P0]
P = [p + (exempt_flag(p),) for p in P]
# CETES terms are 28/91/182/364 days; they are grouped with the 30/90/180/365 horizons for comparison.

AMOUNTS = [50000, 100000, 200000, 300000, 500000]
HORIZONS = [30, 90, 180, 365]
DEFAULT_AMOUNT = 200000
SHOCKS = [0, 50]  # bp drop in on-demand rates after day 90 (fixed terms keep their rate)


def eff_rate(rate, term, days, shock_bp):
    """Average annual rate over the horizon; on-demand rates drop by shock_bp after day 90."""
    if term != 0 or shock_bp == 0 or days <= 90:
        return rate
    return (rate * 90 + (rate - shock_bp / 10000) * (days - 90)) / days


def product_interest(amount, prod, days, shock):
    inst, name, term, rate, cap, rx, *_ = prod
    rate_e = eff_rate(rate, term, days, shock)
    promo = amount if cap == "" else min(amount, cap)
    excess = amount - promo
    rx_e = 0.0 if rx == "" else eff_rate(rx, term, days, shock)
    fee = prod[8] * days / 30.4167
    return (promo * rate_e + excess * rx_e) * days / 365 - fee


def withholding(allocs, days):
    """0.90% on capital; exempt-eligible balances (SOFIPOs, Nu on demand) share one 5-UMA exemption."""
    sof = sum(a for (p, a) in allocs if p[9] == "yes")
    other = sum(a for (p, a) in allocs if p[9] != "yes")
    taxable = max(0.0, sof - SOFIPO_EXEMPT) + other
    return taxable * WHT * days / 365, sof, taxable


def available(prod, days, include_all):
    inst, name, term = prod[0], prod[1], prod[2]
    default = prod[7]
    if not include_all and default != "yes":
        return False
    return term == 0 or term == days


def prot_cap(inst, rate, days):
    """Largest deposit whose end balance stays within the institution's protection limit."""
    lim = INST_PROT[inst]
    if lim == "":
        return float("inf")
    return lim / (1 + rate * days / 365)


def strategy_single(amount, prod):
    return [(prod, amount)]


def strategy_split(amount, days, shock, include_mp):
    """Fill promo tiers first, then the highest-rate protected products; respects caps and protection limits."""
    allocs, left, used = [], amount, {}
    cands = [p for p in P if available(p, days, False) or (include_mp and p[0] == "Mercado Pago")]
    promos = sorted([p for p in cands if p[4] != ""], key=lambda p: -p[3])
    for p in promos:
        if left <= 0:
            break
        room = min(p[4], prot_cap(p[0], p[3], days) - used.get(p[0], 0))
        a = max(0, min(left, room))
        if a > 0:
            allocs.append((p, a)); left -= a; used[p[0]] = used.get(p[0], 0) + a
    def net_rank(p):
        r = eff_rate(p[3], p[2], days, shock)
        if p[9] != "yes":
            r -= WHT
        if p[8] and left > 0:
            r -= p[8] * 12 / left
        return -r
    rest = sorted([p for p in cands if p[4] == ""], key=net_rank)
    for p in rest:
        if left <= 0:
            break
        room = prot_cap(p[0], p[3], days) - used.get(p[0], 0)
        a = max(0, min(left, room))
        if a > 0:
            allocs.append((p, round(a, 2))); left -= a; used[p[0]] = used.get(p[0], 0) + a
    return allocs


def find(inst, name):
    return next(p for p in P if p[0] == inst and p[1] == name)


def best_term(days, sofipo_only=True):
    c = [p for p in P if p[2] == days and p[7] == "yes" and (not sofipo_only or INST_TYPE[p[0]] == "SOFIPO")]
    return max(c, key=lambda p: p[3]) if c else None


def cetes(days):
    return next(p for p in P if p[0].startswith("CETES") and p[2] == days)


def build_scenarios():
    alloc_rows, summ_rows = [], []
    for amount in AMOUNTS:
        for days in HORIZONS:
            for shock in SHOCKS:
                strategies = {
                    "A. All in Stori on demand (SOFIPO)": strategy_single(amount, find("Stori", "On demand")),
                    "B. All in Nu Cajita (bank)": strategy_single(amount, find("Nu", "Cajita (on demand)")),
                    "C. All in CETES (GBM)": strategy_single(amount, cetes(days)),
                    "D. All in best SOFIPO fixed term": strategy_single(amount, best_term(days)),
                    "E. Split, protected products only": strategy_split(amount, days, shock, False),
                    "F. Split, including Mercado Pago": strategy_split(amount, days, shock, True),
                }
                for sname, allocs in strategies.items():
                    gross = 0.0
                    over = False
                    per_inst = {}
                    for prod, a in allocs:
                        g = product_interest(a, prod, days, shock)
                        gross += g
                        per_inst.setdefault(prod[0], 0)
                        per_inst[prod[0]] += a + g
                        alloc_rows.append((sname, amount, days, shock, prod[0], INST_TYPE[prod[0]], prod[1], prod[2],
                                           prod[3], round(a, 2), round(g, 2)))
                    for inst, bal in per_inst.items():
                        if INST_PROT[inst] == "" and INST_TYPE[inst] != "Government":
                            over = True
                        elif INST_PROT[inst] != "" and bal > INST_PROT[inst] + 0.01:
                            over = True
                    wht, sof, taxable = withholding(allocs, days)
                    net = gross - wht
                    summ_rows.append((sname, amount, days, shock, round(gross, 2), round(wht, 2), round(net, 2),
                                      round(net / amount * 365 / days, 5), round(sof, 2), round(taxable, 2),
                                      "no" if over else "yes",
                                      "yes" if any(p[6] != "no" for p, _ in allocs) else "no",
                                      "yes" if amount == DEFAULT_AMOUNT else "no"))
    return alloc_rows, summ_rows


# ---------------- S&P 500 module ----------------
# Two ways to hold the S&P 500 from Mexico through GBM:
#   SIC:     VOO/IVV bought in the Sistema Internacional de Cotizaciones. Priced in MXN but the
#            value follows the dollar, so the peso/dollar move is part of the return.
#   IVVPESO: iShares S&P 500 Peso Hedged TRAC on the BMV. Currency risk is hedged; the hedge
#            earns roughly the MXN minus USD rate difference.
EQ = [
    # param, value, unit, notes, source, verified
    ("gbm_commission", 0.0025, "per trade, on value", "Up to 1M MXN traded in 3 months; plus IVA.", S["gbm_fees"], "secondary"),
    ("iva", 0.16, "rate on commission", "", S["gbm_fees"], "yes"),
    ("sic_etf_expense_ratio", 0.0003, "annual", "VOO expense ratio, typical (not fetched).", "", "unverified"),
    ("ivvpeso_expense_ratio", 0.0049, "annual", "", S["ivvpeso"], "secondary"),
    ("hedge_carry_assumed", 0.025, "annual", "Assumed MXN minus USD short-rate gap (Banxico 6.50% minus an assumed ~4% US rate). Edit this.", "", "assumption"),
    ("equity_gain_isr", 0.10, "rate on realized gain", "LISR art. 129 applies to BMV and SIC listed shares/ETFs; paid in the annual return, no withholding. Gain assumed measured in MXN (includes FX).", S["sic_tax"], "secondary"),
    ("spx_tr_2025", 0.179, "total return, USD", "", S["spx_2025"], "yes"),
    ("spx_tr_2026_ytd", 0.1347, "total return, USD, to 2026-09-23", "", S["spx_ytd"], "secondary"),
    ("usdmxn_2024_close", 20.7862, "MXN per USD", "Derived: 2025 FIX close 18.0012 plus the reported 2.785 appreciation.", S["fx_2025"], "secondary (derived)"),
    ("usdmxn_2025_close", 18.0012, "MXN per USD", "FIX, El Financiero. DOF shows 17.9528 for the FIX published 2025-12-31.", S["fx_2025"], "secondary"),
    ("usdmxn_2026_09", 17.20, "MXN per USD", "Approximate spot, mid September 2026.", S["fx_now"], "approximate"),
]
EQP = {r[0]: r[1] for r in EQ}
TRADE_COST = EQP["gbm_commission"] * (1 + EQP["iva"])
SPX_SCEN = [("Bear", -0.15), ("Flat", 0.0), ("Base", 0.08), ("Bull", 0.20)]
FX_SCEN = [("Peso +10% (USD/MXN -10%)", -0.10), ("Peso flat", 0.0), ("Peso -10% (USD/MXN +10%)", 0.10)]
VEHICLES = ["S&P 500 via SIC (VOO/IVV, USD exposure)", "S&P 500 via IVVPESO (peso hedged)"]
EQUITY_SHARES = [0, 0.10, 0.25, 0.50]


def equity_result(amount, vehicle, spx, fx, years=1.0):
    """Net MXN result of buying at t0 and selling at the horizon, after GBM costs and 10% ISR on the gain."""
    invested = amount / (1 + TRADE_COST)
    if vehicle.startswith("S&P 500 via SIC"):
        growth = (1 + spx) * (1 + fx) * (1 - EQP["sic_etf_expense_ratio"]) ** years
    else:
        growth = (1 + spx + EQP["hedge_carry_assumed"] * years) * (1 - EQP["ivvpeso_expense_ratio"]) ** years
    gross_value = invested * growth
    proceeds = gross_value * (1 - TRADE_COST)
    costs = amount - invested + gross_value * TRADE_COST
    gain = proceeds - amount
    isr = max(0.0, gain) * EQP["equity_gain_isr"]
    return round(gross_value - invested, 2), round(costs, 2), round(isr, 2), round(gain - isr, 2)


def banxico_avg(start, end):
    d0, d1 = dt.date.fromisoformat(start), dt.date.fromisoformat(end)
    total, d = 0.0, d0
    while d < d1:
        total += [r for (x, r, _) in BX if dt.date.fromisoformat(x) <= d][-1] if any(dt.date.fromisoformat(x) <= d for x, _, _ in BX) else BX[0][1]
        d += dt.timedelta(days=1)
    return total / (d1 - d0).days


def build_equity():
    scen = []
    for amount in AMOUNTS:
        for v in VEHICLES:
            for sn, spx in SPX_SCEN:
                for fn, fx in FX_SCEN:
                    if v.endswith("(peso hedged)") and fx != 0.0:
                        continue
                    mkt, costs, isr, net = equity_result(amount, v, spx, fx)
                    scen.append((amount, v, sn, spx, "n/a (hedged)" if v.endswith("(peso hedged)") else fn,
                                 fx if not v.endswith("(peso hedged)") else 0.0, mkt, costs, isr, net,
                                 round(net / amount, 5), "yes" if amount == DEFAULT_AMOUNT else "no"))
    # historical: what 1 USD-exposed peso did vs the S&P in dollars
    bt = []
    for label, spx, fx0, fx1, start, end in [
        ("2025 calendar year", EQP["spx_tr_2025"], EQP["usdmxn_2024_close"], EQP["usdmxn_2025_close"], "2025-01-01", "2026-01-01"),
        ("2026 to Sep 23", EQP["spx_tr_2026_ytd"], EQP["usdmxn_2025_close"], EQP["usdmxn_2026_09"], "2026-01-01", "2026-09-23"),
    ]:
        fx_chg = fx1 / fx0 - 1
        mxn = (1 + spx) * (1 + fx_chg) - 1
        bt.append((label, spx, fx0, fx1, round(fx_chg, 5), round(mxn, 5), round(banxico_avg(start, end), 5),
                   "Peso appreciation cut the dollar return" if fx_chg < 0 else "Peso depreciation added to the return"))
    # mixes: part in equity, rest in split strategy E for 365 days
    mix = []
    for amount in AMOUNTS:
        for share in EQUITY_SHARES:
            eq_amt = amount * share
            sav_amt = amount - eq_amt
            if sav_amt > 0:
                allocs = strategy_split(sav_amt, 365, 0, False)
                gross = sum(product_interest(a, p, 365, 0) for p, a in allocs)
                sav_net = gross - withholding(allocs, 365)[0]
            else:
                sav_net = 0.0
            for v in VEHICLES:
                if share == 0 and v != VEHICLES[0]:
                    continue
                for sn, spx in SPX_SCEN:
                    for fn, fx in FX_SCEN:
                        if (v.endswith("(peso hedged)") or share == 0) and fx != 0.0:
                            continue
                        eq_net = equity_result(eq_amt, v, spx, fx)[3] if eq_amt else 0.0
                        tot = sav_net + eq_net
                        mix.append((amount, share, "Savings only" if share == 0 else v, sn,
                                    "n/a" if (share == 0 or v.endswith("(peso hedged)")) else fn,
                                    round(sav_amt, 2), round(sav_net, 2), round(eq_amt, 2), round(eq_net, 2),
                                    round(tot, 2), round(tot / amount, 5), "yes" if amount == DEFAULT_AMOUNT else "no"))
    return scen, bt, mix


def write_csv(name, header, rows):
    path = os.path.join(DATA, name)
    with open(path, "w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(header)
        w.writerows(rows)
    return path, header, rows


def main():
    os.makedirs(DATA, exist_ok=True)
    alloc, summ = build_scenarios()
    eq_scen, eq_bt, eq_mix = build_equity()
    tables = [
        write_csv("institutions.csv", ["institution", "type", "protection_scheme", "protection_limit_mxn", "notes", "source_url", "accessed", "verified"],
                  [r[:5] + (r[5], ACC, r[6]) for r in INST]),
        write_csv("rates_history.csv", ["date", "institution", "product", "term_days", "rate_annual", "balance_cap_mxn", "rate_on_excess",
                                        "conditions", "source_url", "accessed", "verified"],
                  [r[:8] + (r[8], ACC, r[9]) for r in H]),
        write_csv("banxico_rate.csv", ["decision_date", "target_rate", "change_bp", "source_url", "accessed"],
                  [b + (S["banxico"], ACC) for b in BX]),
        write_csv("banxico_monthly.csv", ["month_end", "target_rate"], banxico_monthly()),
        write_csv("tax_params.csv", ["param", "value", "unit", "applies_from", "notes", "source_url", "verified"], TAX),
        write_csv("products_current.csv", ["institution", "product", "term_days", "rate_annual", "balance_cap_mxn", "rate_on_excess",
                                           "needs_condition", "include_default", "monthly_fee_mxn", "exempt_eligible", "type", "protection_limit_mxn"],
                  [p + (INST_TYPE[p[0]], INST_PROT[p[0]]) for p in P]),
        write_csv("amounts.csv", ["amount_mxn", "is_default"], [(a, "yes" if a == DEFAULT_AMOUNT else "no") for a in AMOUNTS]),
        write_csv("horizons.csv", ["horizon_days", "label"], [(30, "1 month"), (90, "3 months"), (180, "6 months"), (365, "1 year")]),
        write_csv("scenario_allocations.csv", ["strategy", "amount_mxn", "horizon_days", "rate_shock_bp", "institution", "type", "product",
                                               "term_days", "rate_annual", "allocated_mxn", "gross_interest_mxn"], alloc),
        write_csv("scenario_summary.csv", ["strategy", "amount_mxn", "horizon_days", "rate_shock_bp", "gross_interest_mxn", "isr_withheld_mxn",
                                           "net_return_mxn", "effective_net_rate", "exempt_eligible_balance_mxn", "taxable_balance_mxn",
                                           "fully_protected", "needs_conditions", "is_default_amount"], summ),
        write_csv("equity_params.csv", ["param", "value", "unit", "notes", "source_url", "verified"], EQ),
        write_csv("sp500_backtest.csv", ["period", "spx_total_return_usd", "usdmxn_start", "usdmxn_end", "usdmxn_change",
                                         "return_in_mxn_sic", "banxico_avg_target_rate", "note"], eq_bt),
        write_csv("equity_scenarios.csv", ["amount_mxn", "vehicle", "spx_scenario", "spx_return_usd", "fx_scenario", "usdmxn_change",
                                           "market_gain_mxn", "trading_costs_mxn", "isr_mxn", "net_result_mxn", "net_return",
                                           "is_default_amount"], eq_scen),
        write_csv("portfolio_mix.csv", ["amount_mxn", "equity_share", "equity_vehicle", "spx_scenario", "fx_scenario", "savings_mxn",
                                        "savings_net_mxn", "equity_mxn", "equity_net_mxn", "total_net_mxn", "total_net_return",
                                        "is_default_amount"], eq_mix),
    ]
    try:
        from openpyxl import Workbook
        from openpyxl.worksheet.table import Table, TableStyleInfo
        wb = Workbook()
        wb.remove(wb.active)
        for path, header, rows in tables:
            name = os.path.splitext(os.path.basename(path))[0]
            ws = wb.create_sheet(name[:31])
            ws.append(list(header))
            for r in rows:
                ws.append([None if v == "" else v for v in r])
            ref = f"A1:{ws.cell(row=1, column=len(header)).column_letter}{len(rows) + 1}"
            t = Table(displayName=name, ref=ref)
            t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
            ws.add_table(t)
        wb.save(os.path.join(HERE, "sofipo_savings.xlsx"))
    except ImportError:
        print("openpyxl not installed; CSVs written, workbook skipped")
    print("done:", len(tables), "tables")


if __name__ == "__main__":
    main()
