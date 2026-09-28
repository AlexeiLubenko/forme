# Qlik Sense: структура приложения мониторинга резервов

## 1. Executive Dashboard
Фильтры: report_date, risk_segment, business_segment, subportfolio, industry, currency.

KPI:
- Reserve latest: `Sum({<report_date={"=$(=Date(Max(report_date)))"}>} reserve_total_mln)`
- EAD latest: `Sum({<report_date={"=$(=Date(Max(report_date)))"}>} ead_mln)`
- Reserve rate latest: `Sum({<report_date={"=$(=Date(Max(report_date)))"}>} reserve_total_mln) / Sum({<report_date={"=$(=Date(Max(report_date)))"}>} ead_mln)`
- Delta reserve DoD: `Sum({<report_date={"=$(=Date(Max(report_date)))"}>} delta_reserve_dod_mln)`
- Delta reserve MTD: `Sum({<report_date={"=$(=Date(Max(report_date)))"}>} delta_reserve_mtd_mln)`
- Stage 2+3 share: `(Sum({<report_date={"=$(=Date(Max(report_date)))"}>} reserve_stage2_mln)+Sum({<report_date={"=$(=Date(Max(report_date)))"}>} reserve_stage3_mln))/Sum({<report_date={"=$(=Date(Max(report_date)))"}>} reserve_total_mln)`

Charts:
1. Line: dimension `report_date`, measures `reserve_total_mln`, `ead_mln` (лучше отдельными графиками из-за масштаба).
2. Line: dimension `report_date`, measure `reserve_rate`.
3. Waterfall/bar: dimension `movement_type`, measure `Sum(delta_reserve_mln)`, filter latest report_date.
4. Waterfall/bar: dimension `primary_driver`, measure `Sum(delta_reserve_mln)`, filter latest report_date.
5. Table TOP movements: client_name, contract_number, delta_reserve_mln, stage_migration, rating_migration, primary_driver, driver_flags.

## 2. Stage & Migration
- Pivot/heatmap: rows `stage_prev`, columns `stage_curr`, measure `Count(DISTINCT contract_key)`.
- Second measure: `Sum(delta_reserve_mln)`.
- Trend by current stage from DailyDrivers: filter `dimension='Стадия'`, dimension report_date + dimension_value, measure reserve_mln.
- KPI count deterioration: `Count({<flag_stage_deterioration={1}>} DISTINCT contract_key)`.

## 3. Risk Drivers
Bar charts by `primary_driver`:
- `Sum(delta_reserve_mln)`
- `Count(DISTINCT contract_key)`

Additional tables/filters:
- rating_migration
- overdue_prev -> overdue_curr
- flag_new_default
- flag_overdue_deterioration / flag_overdue_improvement
- flag_pd_change
- flag_lgd_change
- flag_ead_change
- multi_factor

## 4. Portfolio Segments
Use `DailyDrivers` and select one dimension at a time:
- Риск-сегмент
- Бизнес-сегмент
- Субпортфель
- Отрасль клиента
- Рейтинг на отчетную дату
- Бакет просрочки по договору
- Валюта

Measures: reserve_mln, ead_mln, reserve_rate, delta_reserve_mln where available.

## 5. Client / Contract Drill-down
From `LatestContracts` show:
- client, contract, tranche
- current Stage/rating/overdue bucket
- EAD, reserve, reserve rate
- PD, LGD, macro addon, CCF

From `Movements` show history by selected `contract_key`:
- delta_reserve_mln
- delta_ead_mln
- stage_migration
- rating_migration
- primary_driver
- driver_flags

## 6. Data Quality
KPI current `dq_status`; table of DQ metrics by report_date.
Red conditional formatting when status='FAIL', amber when 'WARN'.

## 7. Daily operating filters
The default sheet bookmark should select `report_date = Max(report_date)`.
For MTD analysis remove this selection and select MonthYear.
