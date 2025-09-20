# 아키텍처 개요

## 상위 흐름
1. **설정**: `load_config()`가 `enhanced_config.json`을 읽어들여 로깅을 설정합니다.
2. **오케스트레이션**: `screen_nasdaq()`이 비동기 워크플로우를 조율합니다. NASDAQ 심볼 목록을 가져오고, 각 항목을 `fetch_stock_profiles()`를 통해 프로필 데이터로 보강하며, 필터를 적용해 심볼을 선별합니다.
3. **데이터 획득**: 각 후보 심볼에 대해 `get_comprehensive_financial_data()`가 손익계산서, 현금흐름표, 재무비율, 주요 지표를 Financial Modeling Prep API를 통해 수집합니다.
4. **지표 준비**: `prepare_financial_metrics()`가 재무제표 데이터를 날짜별로 정렬하고, 관련 시계열(매출, EPS, FCF, ROE, 마진, 운전자본, R&D, 설비투자, PER, PBR)을 추출해 반환합니다.
5. **퀄리티 분석**: `QualityScorer.calculate_final_score()`가 성장, 리스크, 밸류에이션 모듈의 서브스코어를 결합하고 가중치를 조정한 뒤, 일관성 보너스를 곱해 최종 점수를 계산합니다.
6. **출력**: 워크플로우가 `write_enhanced_output()`, `write_excel_output()`을 통해 텍스트 및 엑셀 리포트를 생성하고, 실패 내역을 로깅합니다.

## 핵심 분석 컴포넌트

### 지표 정규화
`MetricNormalizer`는 각 지표 시퀀스를 설정 가능한 분위수로 윈저라이즈하여 이상치 영향을 줄이고, 정규화 범위를 저장합니다. `normalize()` 메서드는 지표를 `[0, 1]` 범위로 선형 변환합니다.

### 성장 분석
`GrowthQualityAnalyzer`는 성장성을 세 가지 축에서 평가합니다:
- **크기**: `calculate_magnitude_score()`가 CAGR(연평균 성장률)을 섹터 또는 목표 벤치마크와 비교합니다. 매출 CAGR은 먼저 섹터 중앙값, 표준편차를 기준으로 정규화된 후 점수가 산출됩니다.
- **일관성**: `calculate_growth_consistency()`가 기간별 변화에 대한 변동계수(coefficient of variation)를 계산해 성장의 안정성을 평가합니다.
- **지속가능성**: `assess_growth_sustainability()`가 R&D 투자 비율, 설비투자 효율, 영업이익률 안정성, FCF 전환율 등 보조 지표를 점수화하여 평가합니다.

`calculate_growth_scores()`는 이 세 요소를 각각 35%, 35%, 30% 가중치로 합산하여 성장 점수를 산출하고, 각 축별 점수도 반환합니다.

### 리스크 평가
`RiskAssessmentModule`은 마진 변동성과 운전자본 효율성을 분석해 사업 탄력성을 평가합니다:
- `calculate_margin_risk()`는 총마진 및 영업마진의 안정성과 추세 지표를 결합하여 점수를 냅니다.
- `calculate_working_capital_efficiency()`는 매출 대비 운전자본 회전율을 정규화하고, 개선 또는 악화 추세를 분석해 플래그를 설정합니다.

### 밸류에이션 분석
`ValuationAnalyzer`는 상대적 밸류에이션 지표를 사용합니다:
- `calculate_valuation_score()`는 PER, PBR을 정규화 후 평균하여 점수를 산출합니다.
- `calculate_fcf_yield()`는 최근 FCF를 시가총액 대비로 평가해 매력도를 측정합니다.
- `calculate_growth_adjusted_valuation()`는 성장률이 있을 경우 PEG(성장 대비 PER) 방식의 조정 점수를 계산합니다.

### 퀄리티 통합
`QualityScorer`는 위 모듈들을 통합합니다:
1. 성장, 밸류에이션, 마진 퀄리티, 운전자본 효율, FCF 수익률 컴포넌트를 계산합니다.
2. 섹터별 기본 가중치 적용(일반적으로 성장 60%, 기술 섹터는 65%).
3. `calculate_coherence_bonus()`로 성장, 마진 리스크, 운전자본 효율, 지속가능성의 평균을 내 일관성 보너스(coherence multiplier)를 산출합니다.
4. 가중 합산 점수에 일관성 보너스를 곱해 최종 퀄리티 점수를 만들고, 각 컴포넌트별 상세 점수도 반환합니다.

### 리포팅
`write_enhanced_output()`, `write_excel_output()`이 지표, 서브스코어, 기업 메타데이터를 사람이 읽을 수 있는 텍스트 및 스프레드시트 형식으로 출력합니다.

## 동시성 및 속도 제한
- `fetch()`는 요청 간 딜레이를 두고, 속도 제한이나 일시적 에러 발생 시 지수 백오프를 적용합니다.
- `screen_nasdaq()`은 전역 요청 제한자(`MAX_CONCURRENT_REQUESTS`)와 분석 전용 세마포어(`config['concurrency']['max_workers']`) 두 가지를 사용해 API 처리량과 리소스 사용을 조절합니다.

## GUI 진입점
`create_gui()`는 사용자가 주요 임계값을 오버라이드할 수 있는 Tkinter 기반 인터페이스를 생성합니다. 이후 비동기 워크플로우를 실행하는 `main()`을 통해 사용자 입력을 병합(`deep_update()`)하고 프로세스를 시작합니다.
