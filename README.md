# Daily Inventory Cycle Count for Google Sheets

Count a few items every day instead of the whole store once a month.
Find losses **the next day**, not 30 days later.

Free Google Apps Script. No app to install, no monthly fee. Works on the phone.

> 🇪🇸 [Resumen en español abajo](#en-español)

---

## The problem

Most small stores and warehouses count everything once a month (or once a year).

- The full count takes hours and people rush it.
- When a number is wrong, nobody knows **when** it happened or **who** was working.
- The items that get stolen the most are counted as often as the ones nobody touches.

## The solution

A Google Sheet that builds a short count list every morning:

1. **All high-risk items, every day.** The items that get stolen or lost the most.
2. **Plus the items that have gone the longest without a count.** Everything gets counted on a rotation.
3. Staff type only one column: **Counted Qty** (yellow).
4. **Submit count** saves the result, flags big differences and emails the manager a short report.

```
 7:00 AM  list is ready  ──►  staff count (10–15 min)  ──►  Submit  ──►  manager gets the report
                                                                  │
                                                                  └──►  History (every count, every day)
```

## Where this comes from

I built the first version of this process for my own retail stores (shapewear, 8 locations).
A difference between the shelf and the point of sale could be a counting mistake, a sale under the wrong code,
or a real loss, and we found out too late to know when or how it happened.
With a short daily count and a review every morning, a difference shows up the **next day**,
while it is still easy to explain or investigate.

More about that project: [aptocode.com/portafolio](https://www.aptocode.com/portafolio/)

## What you get

| Sheet | What it does |
|---|---|
| **Items** | Your product list: SKU, name, location, system qty, unit cost, risk (High / Medium / Low) |
| **Count Today** | The daily list. Staff fill the yellow column. Difference and $ calculate live |
| **History** | Every count, with date, difference, $ and who counted |
| **Settings** | Items per day, alert limits (units and $), report email |

The email report shows: items counted, % exact, net $ difference, and the list of items to recount today.

## Setup (5 minutes)

1. Create a new Google Sheet.
2. Go to **Extensions → Apps Script**.
3. Replace the code with [`src/Code.gs`](src/Code.gs) and save.
4. Reload the sheet. A new menu appears: **Cycle Count**.
5. Click **Cycle Count → Set up sheets with sample data** (accept the Google permissions).
6. Click **Cycle Count → Generate today's count**. Done.

Optional: **Cycle Count → Run daily at 7 AM** to build the list automatically every morning.

To use your own products, paste them in the **Items** sheet
(same columns as [`sample-data/items.csv`](sample-data/items.csv)).

## Settings

| Setting | Default | Meaning |
|---|---|---|
| Items per day | 15 | Size of the daily list (high-risk items are always included) |
| Alert difference (units) | 2 | Flag when the difference is 2 units or more |
| Alert difference ($) | 25 | Flag when the difference is worth $25 or more |
| Report email | (owner) | Who receives the daily report |
| Update system qty after count | NO | YES = replace the system qty with the counted qty |

## Need it connected to your POS?

This version uses the quantities you keep in the sheet.
I also build versions that pull stock automatically from **Square, Shopify or QuickBooks**,
add barcode scanning from the phone, and send alerts by WhatsApp.

📩 Contact me on [Upwork](https://www.upwork.com/freelancers/~0176e8afde5db388a0) or through [aptocode.com](https://www.aptocode.com).

---

## En español

**Conteo cíclico diario de inventario en Google Sheets.** En vez de contar toda la tienda una vez al mes,
cada mañana se arma una lista corta: todos los productos de alto riesgo (los que más se roban)
más los que llevan más tiempo sin contarse. El personal llena una sola columna, y al enviar el conteo
se guarda el historial y el encargado recibe un reporte por email con las diferencias a revisar.

Nació del sistema que usé en mis propias tiendas: las pérdidas se detectan **al día siguiente**,
no 30 días después. Instalación en 5 minutos (ver *Setup* arriba). Si lo necesitas conectado a
Square, Shopify o QuickBooks, escríbeme.

---

MIT License · Built by [AptoCode](https://www.aptocode.com)
