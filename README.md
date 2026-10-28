
# Retail Inventory Control and Stock Reconciliation for a Nigerian FMCG Distributor

An Excel-based inventory tracking system for a small or medium-sized distributor of fast-moving consumer goods (FMCG) such as beverages, noodles, cooking oil, detergent and packaged food. The workbook records daily stock movements, reconciles book stock against physical counts, raises reorder alerts, classifies fast- and slow-moving products, and produces a weekly inventory performance report.

> **Sample data notice:** All business names, products, suppliers, customers, prices, quantities and investigation notes in this repository are fictitious and for demonstration only. Prices are illustrative, not market quotes.

---

## Business Scenario

A distributor supplies provision shops, supermarkets and open-market retailers in Lagos. The business faces:

- Stock discrepancies between records and physical counts
- Delayed replenishment and stock-outs
- Difficulty identifying fast-moving and slow-moving products
- Frequent cash, transfer and credit sales, product returns, damaged cartons, and deliveries from multiple suppliers

## Project Objectives

1. Build an Excel inventory register with opening stock, purchases, sales, returns, damaged goods and closing stock.
2. Record daily sales and stock movements using sample transactions.
3. Calculate expected closing stock.
4. Compare physical counts with book balances and identify variances.
5. Set minimum stock levels and reorder alerts.
6. Analyse fast- and slow-moving products using sales quantities and stock turnover.
7. Prepare a weekly inventory report covering availability, discrepancies, reorder requirements and potential losses.

## Core Formula

```
Expected Closing Stock = Opening Stock + Stock Received + Sales Returns
                         - Units Sold - Supplier Returns - Damaged Stock

Variance (units)       = Physical Count - Expected Closing Stock
```

A negative variance is a shortage; a positive variance is an overage.

---

## Repository Contents

| File | Description |
|------|-------------|
| `FMCG_Inventory_Control_Workbook.xlsx` | The complete inventory tracking workbook |
| `README.md` | Project documentation (this file) |

## Workbook Structure

| Sheet | Purpose |
|-------|---------|
| **README** | In-workbook guide, colour code, how to add data, key assumptions |
| **Weekly Report** | Weekly inventory performance summary: sales and margin, stock position, variances and losses, top and slowest movers, auto-generated observations, charts |
| **Settings** | Week dates, trading days, safety stock days, alert buffers, variance tolerances, fast/slow thresholds |
| **Products** | Product master: code, name, category, supplier, pack unit, unit cost, selling price, supplier lead time (18 SKUs, 5 suppliers) |
| **Transactions** | Daily log of purchases, sales (cash / transfer / credit), sales returns, supplier returns and damaged stock (203 sample lines) |
| **Register** | Inventory register with expected closing stock, physical count and variance per product |
| **Reorder Dashboard** | Reorder alert dashboard: minimum levels, colour-coded alert status, stock level bar, suggested order quantities, supplier order summary |
| **Reconciliation** | Stock reconciliation report: book vs physical, variance in units and naira, priority flag, investigation notes |
| **Movers Analysis** | Fast / medium / slow classification using units sold and weekly stock turnover |
| **Sales Summary** | Pivot-style SUMIFS tables: sales by day and payment type, and by category |

## Deliverables Mapping

| Required deliverable | Where to find it |
|----------------------|------------------|
| Inventory tracking workbook | The whole `.xlsx` file (`Products`, `Transactions`, `Register`) |
| Stock reconciliation report | `Reconciliation` sheet |
| Reorder alert dashboard | `Reorder Dashboard` sheet |
| Weekly inventory performance summary | `Weekly Report` sheet |

---

## How It Works

### Reorder logic

- **Minimum stock** = average daily sales x (supplier lead time + safety days). A manual override column is available.
- **Alert status** (colour-coded, with a level bar):
  - `STOCK-OUT`: stock on hand is zero
  - `REORDER NOW`: stock at or below the minimum
  - `APPROACHING`: stock within a set buffer above the minimum (default 25%)
  - `OK`: stock comfortably above the minimum
- **Suggested order quantity** = (average daily sales x target cover days) - stock on hand, for STOCK-OUT and REORDER NOW items only.

### Variance investigation

Variances are flagged **HIGH** when the variance value reaches a naira threshold or the variance percentage reaches a tolerance (both set in `Settings`). Otherwise they are flagged **LOW**. The report also totals shortages, overages, damaged stock and potential loss.

### Fast and slow movers

- **Weekly turnover** = units sold / average of opening and closing stock
- Classification: Fast-moving, Medium or Slow-moving, using thresholds set in `Settings`
- Products are also ranked by units sold and by turnover

---

## Default Assumptions

All of these are editable in the `Settings` sheet.

| Parameter | Default |
|-----------|---------|
| Reporting week | Mon 19 Oct to Sat 24 Oct 2026 |
| Trading days | 6 |
| Safety stock | 2 days of sales |
| Target cover when reordering | 14 days |
| "Approaching" alert buffer | 25% above minimum |
| Variance tolerance | 5% of book balance |
| High-priority variance value | NGN 20,000 at cost |
| Fast-moving threshold | Weekly turnover of 1.0 or more |
| Slow-moving threshold | Weekly turnover below 0.5 |

Valuation: purchases, supplier returns, damaged stock and variances are valued at unit cost; sales and sales returns at list selling price. Potential loss is shortages plus damaged stock at cost; supplier returns are excluded because they are recoverable by credit note.

---

## Sample Results (Demonstration Week)

- 5 of 18 products need reordering now (1 stock-out, 4 at or below minimum), with 3 more approaching minimum
- 7 products show stock variances; stock record accuracy is about 61%
- Net stock shortage of about NGN 86,900
- Potential loss (shortages plus damaged stock) of about NGN 236,100, roughly 3.7% of net sales

---

## Getting Started

1. Download `FMCG_Inventory_Control_Workbook.xlsx` and open it in Microsoft Excel (or another spreadsheet application that supports standard formulas).
2. Read the `README` sheet for the colour code and sheet guide. Blue text on a yellow background marks input cells; black text marks formulas.
3. Replace the sample data:
   - Update `Settings` with your week start date and trading days.
   - Edit `Products` with your own SKUs, suppliers, costs, prices and lead times.
   - Paste or type your transactions into `Transactions` (rows 6 to 1000; copy the formulas in columns D, E, J and K down).
   - Enter opening stock and physical counts in `Register`.
4. Review `Reorder Dashboard`, `Reconciliation` and `Weekly Report`.

### Weekly routine

1. Set the new week start date in `Settings`.
2. Copy last week's physical count into the Opening Stock column on `Register`, then clear the Physical Count column.
3. Log the new week's transactions on `Transactions`.
4. Count stock, enter the figures on `Register`, and review the reports.

### Creating a PivotTable (optional)

The workbook uses SUMIFS tables for its pivot-style summaries. To add a native PivotTable, select `Transactions!A5:L<last row>`, then Insert > PivotTable. Suggested layout: Rows = Product Name, Columns = Txn Type, Values = Sum of Qty.

---

## Tools and Excel Features Used

- Microsoft Excel
- SUMIFS, COUNTIFS, SUMPRODUCT, IF, IFERROR, INDEX/MATCH, RANK, ROUNDUP
- Conditional formatting (status colours and level bars)
- Data validation drop-downs
- Charts (daily sales, category sales, stock vs minimum, variance by product, units sold, alert status)

## Skills Demonstrated

Inventory record keeping, stock reconciliation, stock control, sales tracking, variance investigation, reorder planning and Excel reporting.

## Limitations

- Sample data only; results do not represent any real business.
- Sales are valued at list price; customer discounts are not modelled.
- The weekly calculations assume a single reporting week at a time.
- Classification thresholds and alert parameters are practical starting points, not industry standards. Tune them to your business.
- Annualised turnover (weekly turnover x 52) is indicative only.

## Licence

Add your preferred licence here 
