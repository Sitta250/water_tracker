# Water Quality Tracker — Setup Instructions

## Step-by-step (takes about 5 minutes)

### 1. Create a new Google Sheet
Go to [sheets.new](https://sheets.new) — a blank spreadsheet opens.

---

### 2. Open the Apps Script editor
**Extensions → Apps Script**

A new tab opens with a code editor. You'll see a default file called `Code.gs`.

---

### 3. Paste the main script
- Delete ALL existing code in `Code.gs`
- Open `Code.gs` from this folder and copy everything
- Paste it into the editor

---

### 4. Add the HTML date picker file
- Click the **+** button next to "Files" in the left sidebar
- Choose **HTML**
- Name it exactly: `DatePicker`  ← important, no typo
- Delete all default content
- Open `DatePicker.html` from this folder and copy everything
- Paste it into the new HTML file

---

### 5. Save the project
Press **Ctrl+S** (or Cmd+S on Mac).
Name the project something like "Wasserqualitäts-Tracker".

---

### 6. Run the setup (ONE TIME ONLY)
- In the function dropdown at the top, select **`setupSpreadsheet`**
- Click the **▶ Run** button
- A permissions dialog will appear — click **Review permissions**
- Choose your Google account, then click **Advanced → Go to ... (unsafe)**
  *(This warning appears for all self-written scripts — it is safe)*
- Click **Allow**

The script will run for ~30–60 seconds and build the entire spreadsheet.

---

### 7. Go back to your spreadsheet
You'll see 4 tabs:
| Tab | Purpose |
|-----|---------|
| **Übersicht** | Dashboard — latest value per pond, red/green traffic light |
| **Messdaten** | Data entry — one row per pond per visit |
| **Diagramme** | 9 charts — one per parameter, all 7 ponds as lines |
| **Diagramm-Daten** | (hidden helper sheet — ignore) |

---

## Daily use

### Adding a new measurement
1. Tap **🎣 Wasserdaten → Neuen Eintrag hinzufügen** from the menu
2. Pick the date (today pre-selected, change for historical data)
3. Press **"7 Zeilen anlegen →"**
4. 7 rows appear (one per pond) — fill in the values

> **On mobile:** The menu is in the ⋮ (overflow) menu → *Extensions* or *Script*.
> You can also just scroll to the bottom of **Messdaten** and type directly.

### Auto-behaviours
- **Date auto-fills** with today's date as soon as you start typing in a row
- **SVB** auto-calculates when you enter Carbonathärte (SVB = Carb ÷ 2.8)
- **Red cells** appear instantly if a value is outside limits
- Tap the Date cell to get a **native calendar picker** (especially useful on phones)

### Sorting
**🎣 Wasserdaten → Älteste zuerst** or **Neueste zuerst**

### After entering a batch of data
Run **🎣 Wasserdaten → Diagramme aktualisieren** to refresh charts.
(Charts do *not* auto-refresh — this is intentional to keep the sheet fast.)

---

## Limit values (Sollwerte)

| Parameter | Limit |
|-----------|-------|
| Sauerstoff O² | ≥ 5.0 mg/l |
| pH-Wert | 6.5 – 8.0 |
| Ammonium NH4+ | ≤ 2.0 mg/l |
| Carbonathärte | ≥ 1.0 D° |
| Nitrat NO3⁻ | ≤ 50.0 mg/l |
| Nitrit NO2⁻ | ≤ 0.01 mg/l |
| Phosphat PO4²⁻ | ≤ 0.3 mg/l |
| SVB / Gesamthärte | no limit defined |

To change a limit: edit the `PARAMS` array at the top of `Code.gs`, re-run `setupSpreadsheet`.

---

## Chart details
- Each of the 9 charts shows **all 7 ponds as separate coloured lines**
- **Orange dashed line** = minimum limit (Mindestens)
- **Red dashed line** = maximum limit (Höchstens)
- **Red dots** = data points that violate the limit
- Normal points are small; violations appear as large red circles
