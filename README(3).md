# Reserve Monitoring Package

Комплект для ежедневного мониторинга резервов по полным дневным срезам кредитного портфеля.

## Состав

- `python/reserve_monitor.py` — основной расчетный pipeline.
- `python/config.yaml` — настройка путей, ключа, колонок, порогов и рейтингов.
- `python/requirements.txt` — Python-зависимости.
- `python/run_monitoring.sh` / `run_monitoring.bat` — запуск.
- `qlik/ReserveMonitoring.qvs` — Qlik Sense load script.
- `qlik/Qlik_Dashboard_Spec.md` — страницы, KPI и выражения Qlik.
- `docs/Functional_Requirements_Reserve_Monitoring.docx` — функциональные требования.
- `sample/sample_raw_2026-09-27.csv`, `sample/sample_raw_2026-09-28.csv` — синтетический пример двух дневных срезов.
- `sample/Daily_Monitoring_Example.xlsx` — пример ежедневного мониторинга.

## Быстрый запуск

1. Создать окружение и установить зависимости:

```bash
python -m venv .venv
source .venv/bin/activate     # Linux/macOS
# .venv\\Scripts\\activate   # Windows
pip install -r requirements.txt
```

2. Скопировать очередной дневной файл в `python/data/raw/`.

3. Проверить `config.yaml`:
- `key_columns` — стабильный ключ договора/транша;
- `monetary_values_in_byn` — уже ли денежные поля в BYN;
- `rating_order` — порядок рейтингов от лучшего к худшему;
- пороги существенных изменений.

4. Запустить:

```bash
cd python
python reserve_monitor.py --config config.yaml --file data/raw/portfolio_2026-09-28.xlsx
```

Или положить файл в `data/raw` и выполнить без `--file`: будет выбран последний измененный файл.

## Что создает Python

`data/processed/daily/`
- полный нормализованный дневной срез (исходные поля + расчетные признаки) на уровне договора;

`data/processed/movements/`
- сравнение текущего дня с предыдущим;

`data/qlik/`
- `daily_summary.csv`
- `daily_drivers.csv`
- `movements.csv`
- `top_movements.csv`
- `latest_contracts.csv`
- `dq_history.csv`

`data/reports/`
- ежедневный Excel-отчет.

## Главная сверка

Для каждого дня должно выполняться:

`Delta Reserve = NEW + EXIT + EXISTING`

`bridge_check_mln` должен быть равен 0 с учетом машинной точности.

`primary_driver` — эвристическая классификация для аддитивного waterfall. Все исходные флаги изменений (`driver_flags`) сохраняются, поэтому результат можно аудировать.

## Qlik

1. Создать Folder/Data connection на папку `python/data`.
2. Назвать подключение `ReserveMonitoring` либо изменить `vDataRoot` в `.qvs`.
3. Вставить `ReserveMonitoring.qvs` в Data Load Editor.
4. Выполнить reload.
5. Создать страницы и master measures по `Qlik_Dashboard_Spec.md`.

## Проверка на тестовых данных

```bash
cd python
rm -rf data/processed data/qlik data/reports data/logs
mkdir -p data/raw
cp ../sample/sample_raw_2026-09-27.csv data/raw/
cp ../sample/sample_raw_2026-09-28.csv data/raw/
python reserve_monitor.py --config config.yaml --file data/raw/sample_raw_2026-09-27.csv
python reserve_monitor.py --config config.yaml --file data/raw/sample_raw_2026-09-28.csv
```

На втором дне появятся NEW, EXIT, Existing, миграции Stage и изменения EAD/PD/LGD.
