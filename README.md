# Daytime and nighttime temperature analysis

This repository estimates mortality risks across nine daytime-nighttime temperature
states, allocates temperature-related burden to four components, and tests differences
within and between cardiocerebrovascular outcomes. It starts from prepared
time-stratified case-crossover data. Individual records are not distributed.
Reproducing population estimates requires authorized access to the datasets in
[INPUT_SCHEMA.md](INPUT_SCHEMA.md).

## Workflow

| Script | Inputs | Outputs |
|---|---|---|
| `01_main_risk_models.R` | Matched temperature/humidity records | Continuous models, shared reference intervals and lag windows, nine-state risks, covariance, counts and comparison-support diagnostics |
| `02_main_burden.R` | Current risks, covariance, state death counts and metadata | Annual and mean annual state burden, Shapley allocation, uncertainty intervals and diagnostic figures |
| `03_main_heterogeneity.R` | Primary risks, covariance, counts, design and burden summaries | Global risk tests, between-disease omnibus and gated pairwise tests, within-outcome AN tests and profile intervals |
| `04_sensitivity_analysis.R` | Matched records, primary calibration and pollutant attachments | Sensitivity fits, burden, profile intervals, between-disease omnibus tests, sample flow and status |
| `05_collect_results.R` | Primary and sensitivity result directories | One workbook combining each result type across outcomes and specifications |

Run from the repository root:

```sh
Rscript 01_main_risk_models.R
Rscript 02_main_burden.R
Rscript 03_main_heterogeneity.R
Rscript 04_sensitivity_analysis.R
Rscript 05_collect_results.R
```

Script 04 performs sensitivity fitting and all associated burden calculations.
Script 05 collects stored results without additional statistical calculations.
Each header documents an option for sourcing functions without executing the
analysis. Population models require substantial memory and runtime; sequential
execution is recommended.

## Software and input locations

Use R with `data.table`, `survival`, `splines`, `ggplot2` and `openxlsx`.
`splines` is included with R. RDS is the default serialization format. The optional
`qs` package is required for QS inputs or when selecting QS output.

All analyses reported in the manuscript were performed using R version 3.6.3
on Ubuntu 18.04.5 LTS. The recorded study environment included `data.table`
1.14.2, `survival` 3.2-7, `ggplot2` 3.3.6, `qs` 0.25.4 and
`openxlsx` 4.2.8.1.

```r
install.packages(c("data.table", "survival", "ggplot2", "openxlsx"))
```

Edit `config/input_files.csv`, preserving the `kind` and `code` identifiers.
File locations are relative to `data/` unless absolute locations are supplied
locally. Default output roots are `results/main/` and `results/sensitivity/`.
Environment variables `DAYNIGHT_PROJECT_ROOT`, `DAYNIGHT_DATA_ROOT` and
`DAYNIGHT_RESULTS_ROOT` override these locations. Machine-specific locations
and study records are not embedded in the repository. `.gitignore` excludes
input data, generated results and local session files.

## Outcomes and selection

| Code | Outcome | Default |
|---|---|---|
| `I00_I52_I60_I69` | Overall cardiocerebrovascular mortality | Yes |
| `I10_I15` | Hypertensive diseases | Yes |
| `I21_I22` | Acute myocardial infarction | Yes |
| `I25` | Chronic ischemic heart disease | Yes |
| `I60_I62` | Hemorrhagic stroke | Yes |
| `I63` | Ischemic stroke | Yes |
| `I20_I25` | Complete ischemic heart disease | Optional |

Overall mortality combines the I00-I52 and I60-I69 source files; it is not
reconstructed from the selected specific causes. AMI and chronic IHD do not
exhaust the I20-I25 category.

Scripts 01 and 02 use individual `RUN_*` switches. Script 03 uses `RUN_OUTCOMES`.
Overall calibration is required even when fitting one specific cause. Set
`RUN_I20_I25 = TRUE` to additionally estimate complete IHD in scripts 01 and 02.
In script 03, complete IHD can replace AMI and chronic IHD; it cannot enter the
same disease comparison as either overlapping subtype. Overall is excluded from
between-disease tests. Scripts 04 and 05 use `INCLUDE_UNSPLIT_IHD` for optional
complete IHD. Sensitivity omnibus tests use the five default specific causes.

Script 04 supports `RUN_MODE = "all"`, `"selected"` or `"custom"`. Selected
mode crosses `SELECTED_OUTCOMES` with `SELECTED_SENSITIVITIES`; custom mode uses
the explicit pairs in `CUSTOM_RUNS`. For all non-pollution specifications:

```r
RUN_MODE <- "selected"
SELECTED_OUTCOMES <- PRIMARY_OUTCOMES
SELECTED_SENSITIVITIES <- c("S2_mmt_ci", "S3_outcome_lags", "S4_fixed_lags")
```

| Specification | Definition |
|---|---|
| `S1_pollution` | Separate adjustment for PM2.5, NO2 and O3_8H, plus joint adjustment for all three |
| `S2_mmt_ci` | Overall-outcome daytime and nighttime MMT 95% intervals define reference intervals; primary shared windows are used |
| `S3_outcome_lags` | Outcome-specific significant harmful lag patterns determine windows; primary reference intervals are used |
| `S4_fixed_lags` | Cold lags 0-20 and heat lags 0-2 in both periods; primary reference intervals are used |

`POLLUTION_MODELS_TO_RUN` selects adjustment sets within S1. Defaults fit 42
models: four pollutant-adjusted and three temperature-specification models for
each of six outcomes. Enabling complete IHD requests 49 models. Pollutant
attachments are required only for selected S1 models. Each adjustment set uses
complete pollutant measurements, followed by valid matched-set filtering.

## Sensitivity completion and resumption

`RUN_VERSION = "analysis_1"` identifies a data and analysis round.
`SKIP_COMPLETED_ANALYSES = TRUE` skips a model only when its completion marker
matches that version and all required output files exist and are nonempty.
Incomplete models are fitted. Completion is recorded after risk, burden and
required sample-flow outputs are written. Summaries and inference include all
completed models in that version, including earlier invocations.

Change `RUN_VERSION` or set `SKIP_COMPLETED_ANALYSES = FALSE` after changing data
or fitting settings. Completion checks do not detect content changes inside
existing files. Set `SENSITIVITY_RUN_VERSION` in script 05 to match script 04.
Collection uses matching completion metadata and reports missing coverage.
Inspect `_summary/tables/run_status.csv` and inference status. With
`CONTINUE_ON_ERROR = TRUE`, independent jobs continue after a failure; a finished
process does not establish that all requested models succeeded.

## Statistical definitions

Separate daytime and nighttime distributed lag nonlinear models use natural
cubic temperature splines with four degrees of freedom and lag splines with
three internal knots equally spaced on the log scale, including an intercept.
The lag domain is 0-20. Bases are built with `splines::ns` and matrix operations.
Conditional logistic regression uses `survival::clogit` with the Efron method,
holiday, and a three-degree-of-freedom natural spline for relative humidity
averaged across daytime and nighttime lags 0-2.

Exponentiated conditional-logistic coefficients are matched odds ratios (ORs).
The internal `rr` column and legacy RR filenames store these ORs. The burden
transformation uses the short-term rate-ratio approximation described in the
manuscript and Supplementary Note S4.

For each period, the primary reference interval is the connected range containing
the overall-outcome MMT with cumulative risk no more than 1% above the minimum.
The search is restricted to the first through 99th temperature percentiles.
A truncated interval stops by default. Automatic windows extend from lag 0
through the last lag with lower 95% RR bound above 1 for the 2.5th-percentile
cold or 97.5th-percentile heat contrast against MMT. Intervening nonsignificant
lags are included. With no qualifying lag, the default fallback is 21 days for
cold and 3 days for heat. Manual lengths include lag 0.

For each period, a cold-window mean below the lower reference bound indicates
cold; a heat-window mean above the upper bound indicates heat. Neither indicates
reference. Records satisfying both conditions are excluded, followed by matched
sets without one case and at least one referent. The periods form nine mutually
exclusive states. All observed estimable states contribute to allocation,
including discordant cold/heat states.

For state log RR `beta_s` and deaths `n_sy` in year `y`,
`AN_sy = n_sy * (1 - exp(-beta_s))`. Reference-state AN is zero.
For a compound state with components `a` and `b`, let `v_s = 1 - exp(-beta_s)`.
Allocations are `phi_a = (v_a + v_ab - v_b)/2` and
`phi_b = (v_b + v_ab - v_a)/2`. A single component receives its state value.
Summing death-weighted allocations produces four component ANs summing to total
AN. Negative contributions are retained.

AN denotes attributable number; AF denotes attributable fraction. Annual AF
divides AN by final included deaths that year. Mean annual AN and AF are
arithmetic means across observed years. Component shares divide mean annual
component AN by mean annual total AN. This allocation describes burden under
the fitted joint-state model and does not identify separate intervention effects.

Correlated coefficient draws propagate uncertainty conditional on the supplied
covariance, reference intervals, windows and counts. Primary burden uses 1,000
draws with seed 20260523; MMT intervals use 1,000 draws with seed 20260612.
Primary inference uses 10,000 draws with base seed 20260925 and fixed outcome
offsets. Sensitivity calculations use base seed 20260926. Empirical interval
endpoints use R quantile type 8. Settings are recorded in outputs.

Within-outcome tests assess four-component AN equality (df = 3) and daytime
versus nighttime AN separately for cold and heat (df = 1 each). In component
order `D_C, N_C, D_H, N_H`, the two contrasts are `(1,-1,0,0)` and `(0,0,1,-1)`.
Their variance is `Var(AN_day) + Var(AN_night) - 2*Cov(AN_day,AN_night)`.
Wald tests are two-sided and require full contrast rank. Separate Holm families
contain global risk tests, four-component tests, and combined cold/heat contrasts
across selected outcomes. With six outcomes, the sizes are 6, 6 and 12.
Unavailable tests keep their planned slots. Selected subsets define their own
explicitly recorded families.

Between-disease tests compare four-component share profiles through three
independent coordinates. For five mutually exclusive diseases the omnibus has
12 degrees of freedom; each disease-pair comparison has 3. A significant primary
omnibus gates pairwise tests with Holm adjustment across ten pairs. Cross-disease
coefficient errors are treated as approximately independent; shared calibration
and counts are fixed.

Profile inference requires a nonzero total-AN interval, limited simulated sign
switching, finite draws, stable split-half covariance and adequate rank. Primary
omnibus inference requires every selected disease profile to be eligible.
Sensitivity inference reports one omnibus per specification and model variant.
`EXCLUDE_UNSTABLE_PROFILES = TRUE` permits an eligible subset and reports included
and excluded diseases with reasons. Missing planned diseases prevent a test.
S3 is unavailable for disease comparison because its exposure windows differ.
AN contrasts do not require profile-ratio eligibility when their own moments
and covariance rank are valid.

## Outputs and result collection

Primary model directories are
`results/main/<outcome>/joint_analysis_<configuration>/`, containing `tables/`,
`figures/`, `models/` and `plot_data/`. Scripts 02 and 03 discover these directories.
Use `ANALYSIS_TAG` or `ANALYSIS_DIRS` to select among multiple runs. The pipeline
index locates directories and does not lock files. Both scripts default to CSV
covariance, independently of file dates. Script 02's default CSV mode rejects
a covariance diagonal that disagrees with the risk-table SEs. Its explicit
`rescale` mode and script 03 rescale the covariance to the risk-table SEs while
preserving correlations; script 03 records whether rescaling occurred.
A refitted model should supply its full covariance with its risks.

| Table | Content |
|---|---|
| `table_04_joint_9_state_rr.csv` | Nine-state risks |
| `table_04a_joint_logrr_covariance.csv` | Full coefficient covariance |
| `table_00c_annual_state_death_counts.csv` | Counts used for burden |
| `table_09_exact_9_state_burden_annual.csv`, `table_10_exact_9_state_burden_mean_annual.csv` | State burden |
| `table_11_shapley_4_component_burden_annual.csv`, `table_12_shapley_4_component_burden_mean_annual.csv` | Component burden |
| `table_03c_matched_state_comparison_support.csv`, `table_03d_case_referent_state_comparisons.csv` | Comparison support |
| `table_00b_joint_analysis_sample_flow.csv`, `table_00d_overlap_exclusion_distribution.csv` | Sample flow and exclusions |

Primary inference is in `results/main/_core_heterogeneity_inference/<tag>/`.
Tables 01-10 contain tests, estimates, diagnostics and settings. Table 11 contains
cold/heat daytime-nighttime contrasts; table 12 combines all within-outcome AN
tests; table 13 consolidates all tests; table 14 contains component AN covariance.
The inference workbook includes `H_Within_all`, with 18 rows by default.

Sensitivity models are in `results/sensitivity/<specification>/<outcome>/<variant>/`.
Inference is in `_heterogeneity/`; status, inventories and
`Sensitivity_analysis_results.xlsx` are in `_summary/`. That workbook has headers
on row 3.

Script 05 writes `_summary/tables/Main_and_sensitivity_all_results.xlsx` under
the sensitivity root. Headers are on row 4. Each result type occupies a sheet;
keys are `analysis_id`, `outcome_code` and `model_variant`. Primary rows use
`analysis_id = "Main"`. `Joint_state_RR`, `Component_mean_annual`,
`Component_profile_inference` and `Between_disease_global` contain central results.
`Within_outcome_comparisons` and `Between_disease_pairwise` contain primary tests.
`Availability` and `Source_index` document coverage. Global tests retain the
upstream disease set when collection is filtered. Recalculate upstream before
collecting after a model change.

## Synthetic example and checks

```sh
Rscript tests/configuration_tests.R
Rscript tests/within_outcome_tests.R
Rscript tests/smoke_test.R
```

The focused test checks analytical contrast variances, P values, Holm families,
selected outcomes and exports. The integration test fits six synthetic primary
models and all 42 sensitivity models, propagates burden uncertainty, runs
inference, checks resumption and creates the combined workbook. Tests use
temporary data and reduced simulation counts. They use no study records.
