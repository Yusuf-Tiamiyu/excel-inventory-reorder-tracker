# Lab inventory reorder tracker in Excel

> **In short:** Our labs used to reorder supplies only after they had run out. I built an Excel tracker for more than 1,500 items that flags anything with less than two months of stock left, so orders go out in time. Procurement now orders before items run out instead of after.

A stock tracker for a diagnostic lab's consumables and reagents (1,578 line items) that flags each item for reorder automatically, based on a two-month stock buffer. It replaced a process where orders went out after something ran out.

**Tools:** Excel (IF, SUMIFS, COUNTIFS, INDEX/MATCH, MAXIFS, data validation, conditional formatting). No macros.

## The rule it enforces

Management policy is a two-month buffer: reorder when two months of stock remain, which leaves time to order and receive delivery before running out. So for every item:

```excel
Reorder Level     = 2 × average monthly usage
Status            = IF(CurrentQty <= ReorderLevel, "REORDER NOW", "OK")
Months of Stock   = IF(AvgMonthlyUsage = 0, "N/A", ROUND(CurrentQty / AvgMonthlyUsage, 1))
```

The status column fills itself in. Nobody has to scan 1,500 rows to decide what to order.

## Sheets

| Sheet | Purpose |
|---|---|
| `Inventory Master` | One row per item: ID, description, current quantity, reorder level, status, months of stock left |
| `Issuance Log` | Every issue of stock: date, item, quantity, which department it went to, who issued it. The month fills in from the date |
| `Monthly Summary` | Type a month in one cell; total issued per item for that month is pulled from the log with SUMIFS |
| `HOW TO USE` | Instructions for the people who update it |

The issuance log is kept separate from the master list on purpose: it's the audit trail. Quantities can be checked against what was actually issued, and average usage can be recalculated from real consumption rather than estimates.

## What it changed

Procurement started ordering before items ran out, because the shortfall was visible early enough to act on.

## Notes

- This copy reflects the stock position on the day it was exported, so most items show REORDER NOW. In daily use the quantities are updated as stock comes in.
- Average monthly usage is only as good as the reorder levels set for each item; the `HOW TO USE` sheet explains how to review them when consumption changes.

## Files

`Inventory_Reorder_Tracker_Portfolio.xlsx`
