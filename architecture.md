# Architecture Overview

## High-Level Flow
1. **Configuration**: `load_config()` ingests `enhanced_config.json` and configures logging.
2. **Orchestration**: `screen_nasdaq()` coordinates the asynchronous workflow. It fetches the NASDAQ symbol list, enriches each entry with profile data via `fetch_stock_profiles()`, filters symbols with `filter_stocks_by_initial_criteria()`, and then analyses qualifying stocks concurrently.
3. **Data Acquisition**: For each candidate symbol `get_comprehensive_financial_data()` gathers income statements, cash-flow statements, ratios, and key metrics through the Financial Modeling Prep API with rate-limited requests handled by `fetch()`.
4. **Metric Preparation**: `prepare_financial_metrics()` aligns statement payloads by date, extracts relevant series (revenue, EPS, FCF, ROE, margins, working capital, R&D, capex, PER, PBR), and returns a `FinancialMetrics` dataclass instance for downstream analysis.
5. **Quality Analysis**: `QualityScorer.calculate_final_score()` composes sub-scores from the growth, risk, and valuation modules, applies weighting adjustments, and multiplies by a coherence bonus to produce a final quality score per stock.
6. **Output**: The workflow writes both text and Excel reports via `write_enhanced_output()` and `write_excel_output()` and logs failures for later inspection.

## Core Analytical Components

### Metric Normalization
`MetricNormalizer` winsorizes each metric sequence at a configurable percentile to reduce outlier impact and stores normalization ranges. Its `normalize()` method linearly scales metrics to the `[0, 1]` interval for later comparison.

### Growth Analysis
`GrowthQualityAnalyzer` evaluates growth across three pillars:
- **Magnitude**: `calculate_magnitude_score()` compares CAGR measurements against sector or target benchmarks. Revenue CAGR is first normalized relative to sector medians and standard deviations before scoring.
- **Consistency**: `calculate_growth_consistency()` derives a coefficient of variation on period-over-period changes to reward stable growth paths.
- **Sustainability**: `assess_growth_sustainability()` inspects supporting metrics (R&D intensity, capex efficiency, operating margin stability, and FCF conversion), each scored via `calculate_trend_score()` or `calculate_growth_consistency()` and averaged with equal weights.

`calculate_growth_scores()` combines these aspects with weights (35% magnitude, 35% consistency, 30% sustainability) to produce a composite growth score along with individual dimension scores returned to the caller.

### Risk Assessment
`RiskAssessmentModule` estimates operational resilience by examining margin behaviour and working capital efficiency:
- `calculate_margin_risk()` blends stability and trend metrics for gross and operating margins.
- `calculate_working_capital_efficiency()` normalizes revenue-to-working-capital turnover and uses trend analysis to flag improving or deteriorating efficiency.

### Valuation Analysis
`ValuationAnalyzer` considers relative valuation metrics:
- `calculate_valuation_score()` averages PER and PBR scores after normalizing each within configured bounds.
- `calculate_fcf_yield()` evaluates the attractiveness of trailing FCF relative to market capitalization.
- `calculate_growth_adjusted_valuation()` computes a PEG-style adjustment when growth rates are available.

### Quality Aggregation
`QualityScorer` integrates the modules above:
1. Computes growth, valuation, margin quality, working capital efficiency, and FCF yield components.
2. Applies sector-specific base weights (`growth` is 60% in general and 65% for technology).
3. Derives a `coherence_multiplier` via `calculate_coherence_bonus()`, which averages growth alignment, margin risk, working capital efficiency, and sustainability to reward consistent narratives.
4. Multiplies the weighted base score by the coherence multiplier to obtain the final quality score and returns detailed component breakdowns for reporting.

### Reporting
`write_enhanced_output()` and `write_excel_output()` render the metrics, sub-scores, and company metadata to human-readable text and spreadsheet formats for further analysis.

## Concurrency and Rate Limiting
- `fetch()` enforces a delay between requests and exponentially backs off when rate-limited or encountering transient errors.
- `screen_nasdaq()` uses two semaphores: a global request limiter (`MAX_CONCURRENT_REQUESTS`) and a dedicated analysis semaphore (`config['concurrency']['max_workers']`) to balance API throughput with CPU-bound processing.

## GUI Entry Point
`create_gui()` builds a Tkinter interface that allows users to override key thresholds before launching the asynchronous workflow through `main()`, which merges user overrides via `deep_update()` and then awaits `screen_nasdaq()`.
