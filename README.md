# Data Science Job Market Salary Calculator (Excel)

An interactive Excel dashboard that turns a **32,000+ row global job-postings dataset** into a live salary calculator — pick a **job title**, **country**, and **employment type**, and every number, table, and chart on the sheet updates instantly.

No Power Query, no Power Pivot, no VBA — just Excel Tables, dynamic arrays, and named ranges doing the work.

---

## What it does

Type in (or pick from a dropdown-style list) three inputs on the `Salary_Calculator` sheet:

| Input | Example |
|---|---|
| Job Title | `Data Scientist` |
| Country | `United States` |
| Employment Type | `Full-time` |

...and the workbook returns:

- **Median annual salary** for that exact combination
- Median salary **broken out by job title**, **by country**, and **by employment type**, each ranked so you can see where your selection stands
- **Job posting counts by platform** (LinkedIn, Indeed, ZipRecruiter, etc.) for that same filtered slice
- A **geographic (filled map) chart** showing where the postings/salaries are concentrated

## How it works

- **`Data`** — the raw dataset (`jobs` Excel Table, ~32,672 rows × 16 columns): title, location, company, remote flag, degree requirement, health insurance flag, salary rate/amount, and a skills list per posting.
- **`Data_Validation`** — builds the live, de-duplicated pick-lists straight off the `jobs` table using `UNIQUE`, `SORT`, and `FILTER`, so new rows added to `Data` automatically flow through to the dropdowns — nothing is hardcoded.
- **`Salary_Calculator`** — the three input cells, wired to named ranges (`title`, `country`, `type`), plus the headline output and summary charts.
- **`title_median` / `country_median` / `type_median` / `platform`** — one sheet per dimension, each computing a **median via an array-entered `MEDIAN(IF(...))` formula** that stacks three simultaneous conditions (title × country × type) against the full dataset, then ranks the results with `SORT`/`FILTER` so you can see how the selection compares to every other option.

### Notable Excel techniques used

- **Dynamic arrays end-to-end**: `UNIQUE`, `SORT`, `FILTER`, and `ANCHORARRAY` chained together so lookup lists and rankings regenerate automatically — no manual re-sorting or copy-paste.
- **Multi-condition array-formula medians** — Excel has no native `MEDIANIFS`, so each median is computed with `MEDIAN(IF(cond1 * cond2 * cond3, values))`, entered as an array formula across three simultaneous filters.
- **Named ranges as the wiring layer** (`title`, `country`, `type`, `median_salary`, `platform`, `job_count`) — every formula and chart reads from these names instead of hardcoded cell references, so the whole model reads like a spec.
- **`XLOOKUP`** for the final "get me this one number" step once the ranked list exists.
- **A Filled Map (geo) chart** — Excel's Bing-powered regional map, plotting the filtered results geographically.
- Everything is built on a proper **Excel Table** (`jobs`) with structured references (`jobs[salary_year_avg]`, etc.) rather than raw `A2:A32673`-style ranges.

## Design decisions

**Median, not average.** Salary data is heavily skewed by outliers — a handful of extreme executive or contractor salaries can pull an average far above what a typical job seeker would actually be offered. Median isn't affected by that skew and reflects the real middle of the market, which is why every summary in this workbook (by title, country, and employment type) is computed as a median rather than a mean.

## Dataset

~32,672 data analytics / data science job postings, including title, location, remote-work flag, degree requirement, salary (annual/hourly), company, and extracted skill list.

## Tech

Microsoft Excel — Tables, Dynamic Arrays (`UNIQUE`, `SORT`, `FILTER`, `XLOOKUP`), array formulas, named ranges, PivotChart-free live charting (Bar + Filled Map).

## File

- [`Data_Science_Calculator_Dashboard.xlsx`](./Data_Science_Calculator_Dashboard.xlsx)

<img width="960" height="540" alt="dashboard" src="https://github.com/user-attachments/assets/215b413f-17bc-4e53-be65-7af94967a575" />
