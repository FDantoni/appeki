# 🍋 APPeki – Calcolatore Iperbilirubinemia Neonatale

Applicazione web sviluppata con **Streamlit** per il calcolo e la classificazione dell'iperbilirubinemia neonatale, basata sulle tabelle cliniche del file Excel `_APPeki FINAL.xlsx` e sul nomogramma di Bhutani.

---

## Funzionalità

### Input dati paziente
- **Data e ora di nascita** + **data e ora del prelievo** → calcolo automatico delle ore di vita
- **Età gestazionale (EG)** in settimane
- **BTS** – Bilirubinemia Totale Sierica (mg/dL)
- **BTC** – Bilirubinemia Transcutanea (mg/dL)
- **8 fattori di rischio** (Maisels et al. 2009):
  - Malattia emolitica iso-immune / deficit G6PD / sferocitosi
  - Allattamento esclusivo con scarsa alimentazione
  - Ittero clinico nelle prime 24 ore
  - Fratello/sorella con ittero grave trattato
  - Cefalematoma o ecchimosi significative
  - Origine asiatica orientale
  - Sepsi / acidosi
  - Albumina < 3.0 g/dL

### Risultati calcolati
| Calcolo | Descrizione |
|---|---|
| **Ore di vita** | Differenza tra data/ora nascita e data/ora prelievo, arrotondata all'ora |
| **Prematurità** | Flag automatico se EG < 37 settimane |
| **Percentile BTS** | Confronto con 50°, 75°, 90° dalla tabella di riferimento |
| **Percentile BTC** | Confronto con 50°, 75° dalla tabella di riferimento |
| **Soglia Fototerapia (FT)** | Valore soglia (mg/dL) in base a EG e ore di vita, con alert se BTS ≥ soglia |
| **Soglia Exsanguinotrasfusione (EXT)** | Valore soglia (mg/dL) in base a EG e ore di vita, con alert se BTS ≥ soglia |

### Nomogramma di Bhutani
Grafico interattivo con le **4 zone di rischio** secondo Bhutani et al. (Pediatrics 1999):

| Zona | Significato |
|---|---|
| 🟢 Low-risk zone | Bilirubin < 40° percentile |
| 🟡 Low-intermediate-risk zone | 40°–75° percentile |
| 🟠 High-intermediate-risk zone | 75°–95° percentile |
| 🔴 High-risk zone | > 95° percentile |

Il punto del paziente viene plottato sul grafico con la classe di rischio evidenziata in un badge colorato.

---

## Installazione

### Prerequisiti
- Python 3.9+

### Installa le dipendenze

```bash
pip install -r requirements.txt
```

### Avvio

```bash
python3 -m streamlit run app.py
```

L'app si apre automaticamente nel browser all'indirizzo `http://localhost:8501`.

---

## File richiesti

| File | Descrizione |
|---|---|
| `app.py` | Applicazione Streamlit principale |
| `_APPeki FINAL.xlsx` | Tabelle di riferimento clinico (BTS, BTC, FT, EXT) |
| `requirements.txt` | Dipendenze Python |

> ⚠️ Il file Excel deve trovarsi nella stessa cartella di `app.py`.

---

## Fonti

- Bhutani VK, Johnson L, Sivieri EM. *Predictive ability of a predischarge hour-specific serum bilirubin for subsequent significant hyperbilirubinemia in healthy term and near-term newborns.* Pediatrics. 1999;103:6-14.
- Maisels MJ, Bhutani VK, Bogen D, Newman TB, Stark AR, Watchko JF. *Hyperbilirubinemia in the Newborn Infant ≥35 Weeks' Gestation: An Update With Clarifications.* Pediatrics. 2009;124:1193-1198.
