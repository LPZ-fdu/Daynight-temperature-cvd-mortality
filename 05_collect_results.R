#!/usr/bin/env Rscript
# =============================================================================
# Collect primary and sensitivity results in one workbook
#
# Purpose: combine each type of current result into one worksheet across outcomes,
# primary analysis, sensitivity specifications, and pollution model variants.
# This script reads result tables and does not fit models or recalculate estimates.
#
# Inputs: CSV tables under results/main and results/sensitivity, including their
# heterogeneity output folders. Each input retains its original numeric precision.
# MAIN_ANALYSIS_TAG or MAIN_ANALYSIS_DIRS selects a primary run when necessary.
#
# Output: results/sensitivity/_summary/tables/Main_and_sensitivity_all_results.xlsx.
# Headers are on row 4. Main identifies the primary rows. Model identifiers are
# analysis_id, outcome_code, and model_variant. Availability and Source_index
# identify absent results and source coverage. No missing estimate is set to zero.
# Within_outcome_comparisons contains three primary comparisons per outcome,
# retaining the selected-outcome Holm families from script 03.
#
# Keep INCLUDE_ADDITIONAL_MODEL_TABLES=TRUE to include coefficient covariances,
# annual counts, exposure definitions, and allocation diagnostics for verification.
# This export replaces the workbook with current directory results. Run it after
# the primary and sensitivity calculations have completed.
# Functions only: options(daynight.collect.run = FALSE); source(this_file).
# =============================================================================

if (.Platform$OS.type == "windows")
  invisible(suppressWarnings(try(Sys.setlocale("LC_CTYPE", "English_United States.utf8"), silent = TRUE)))

# Local paths are configured through environment variables; defaults are relative.
PROJECT_ROOT <- normalizePath(Sys.getenv("DAYNIGHT_PROJECT_ROOT", unset = getwd()),
                              winslash = "/", mustWork = TRUE)
DATA_ROOT <- Sys.getenv("DAYNIGHT_DATA_ROOT", unset = file.path(PROJECT_ROOT, "data"))
RESULTS_BASE <- Sys.getenv("DAYNIGHT_RESULTS_ROOT", unset = file.path(PROJECT_ROOT, "results"))
WORK_DIR <- PROJECT_ROOT

PRIMARY_RESULTS_ROOT <- file.path(RESULTS_BASE, "main")
SENSITIVITY_RESULTS_ROOT <- file.path(RESULTS_BASE, "sensitivity")
OUTPUT_FILE <- file.path(SENSITIVITY_RESULTS_ROOT, "_summary", "tables",
                         "Main_and_sensitivity_all_results.xlsx")

MAIN_ANALYSIS_TAG <- NULL # Set an exact configuration tag when several runs exist.
MAIN_ANALYSIS_DIRS <- NULL # Optional named character vector: outcome code = model directory.
MAIN_HETEROGENEITY_DIR <- NULL # Optional directory containing the primary inference tables/.

OUTCOME_NAMES <- c(
  I00_I52_I60_I69 = "Overall cardiocerebrovascular mortality",
  I10_I15 = "Hypertensive diseases",
  I21_I22 = "Acute myocardial infarction",
  I25 = "Chronic ischemic heart disease",
  I60_I62 = "Hemorrhagic stroke",
  I63 = "Ischemic stroke",
  I20_I25 = "Ischemic heart disease")
PRIMARY_OUTCOMES <- c("I00_I52_I60_I69", "I10_I15", "I21_I22", "I25", "I60_I62", "I63")
INCLUDE_UNSPLIT_IHD <- FALSE
SENSITIVITY_RUN_VERSION <- "analysis_1"
SENSITIVITY_NAMES <- c(S1_pollution = "Air-pollutant adjustment",
                       S2_mmt_ci = "Overall MMT confidence-interval reference",
                       S3_outcome_lags = "Outcome-specific lag windows",
                       S4_fixed_lags = "Fixed cold and heat lag windows")

OUTCOMES_TO_INCLUDE <- c(PRIMARY_OUTCOMES, if (INCLUDE_UNSPLIT_IHD) "I20_I25")
SENSITIVITIES_TO_INCLUDE <- names(SENSITIVITY_NAMES)
POLLUTION_MODELS_TO_INCLUDE <- "all" # Or PM25, NO2, O3_8H, PM25_NO2_O3_8H.
INCLUDE_ADDITIONAL_MODEL_TABLES <- TRUE
STOP_ON_READ_ERROR <- TRUE

required_packages <- c("data.table", "openxlsx")
missing_packages <- required_packages[!vapply(required_packages, requireNamespace,
                                              logical(1), quietly = TRUE)]
if (length(missing_packages)) stop("Install required packages: ", paste(missing_packages, collapse = ", "))
suppressPackageStartupMessages(library(data.table))

# 2. Result types and workbook definitions ---------------------------------------
result_catalog <- function() {
  z <- data.table(scope = character(), sheet = character(), primary_file = character(),
                  title = character(), sensitivity_file = character())
  add <- function(scope, sheet, file, title, sensitivity_file = file) {
    z <<- rbind(z, data.table(scope = scope, sheet = sheet, primary_file = file,
                             title = title, sensitivity_file = sensitivity_file))
  }
  add("model", "Joint_state_RR", "table_04_joint_9_state_rr.csv", "Nine-state relative risks")
  add("model", "Component_mean_annual", "table_12_shapley_4_component_burden_mean_annual.csv", "Mean annual four-component burden")
  add("model", "State_mean_annual", "table_10_exact_9_state_burden_mean_annual.csv", "Mean annual nine-state burden")
  add("model", "Component_annual", "table_11_shapley_4_component_burden_annual.csv", "Annual four-component burden")
  add("model", "State_annual", "table_09_exact_9_state_burden_annual.csv", "Annual nine-state burden")
  add("model", "State_Shapley_values", "table_07_state_specific_shapley_values.csv", "State-specific Shapley allocations")
  add("model", "Reference_intervals", "table_02a_shared_reference_intervals.csv", "Temperature reference intervals")
  add("model", "Selected_lag_windows", "table_03a_auto_selected_lag_windows.csv", "Selected moving-average lag windows")
  add("model", "State_counts", "table_03_joint_state_counts.csv", "Joint-state sample counts")
  add("model", "Sample_flow", "table_00b_joint_analysis_sample_flow.csv", "Joint-analysis sample flow")
  add("model", "Analysis_sample", "table_00_analysis_sample.csv", "Analysis sample")
  add("model", "Sample_before_filter", "table_00_basic_analysis_sample_before_joint_state_filter.csv", "Sample before joint-state filtering")
  add("model", "Annual_state_counts", "table_00c_annual_state_death_counts.csv", "Annual state-specific death counts")
  add("model", "Overlap_exclusions", "table_00a_cold_heat_overlap_diagnostics.csv", "Cold-heat overlap exclusions")
  add("model", "Exclusion_distribution", "table_00d_overlap_exclusion_distribution.csv", "Distribution of overlap exclusions")
  add("model", "Matched_state_support", "table_03c_matched_state_comparison_support.csv", "Within-set support for joint-state comparisons")
  add("model", "Case_referent_comparisons", "table_03d_case_referent_state_comparisons.csv", "Case-referent joint-state comparisons")
  add("model", "Cumulative_RR", "table_01_cumulative_rr_p025_p975_vs_mmt_single_logknots.csv", "Cumulative exposure-response contrasts")
  add("model", "Lag_specific_RR", "table_02_lag_specific_rr_p025_p975_vs_mmt_single_logknots.csv", "Lag-specific exposure-response contrasts")
  add("model", "Reference_MMT_summary", "table_01_mmt_summary.csv", "MMT information used for reference definition")
  add("model", "Outcome_MMT_summary", "table_03_mmt_summary_single_logknots.csv", "Outcome-specific minimum mortality temperatures")
  add("model", "State_thresholds", "table_02_joint_state_thresholds.csv", "State classification thresholds")
  add("model", "Lag_selection_diagnostics", "table_03b_auto_lag_selection_diagnostics.csv", "Lag-selection diagnostics")
  add("model", "Reference_diagnostics", "table_00_reference_state_diagnostics.csv", "Reference-state diagnostics")
  add("model", "Pollution_sample_flow", "table_pollution_sample_flow.csv", "Pollutant model sample flow")
  add("model", "Pollution_coefficients", "table_pollution_adjustment_coefficients.csv", "Pollutant adjustment coefficients")
  add("model", "Pollution_covariates", "table_pollution_covariate_summary.csv", "Pollutant covariate summaries")
  add("model", "Sensitivity_configuration", "table_sensitivity_configuration.csv", "Sensitivity model settings")
  add("inference", "Global_risk_tests", "table_01_joint_state_risk_global.csv", "Primary global joint-state risk tests", NA_character_)
  add("inference", "Within_outcome_tests", "table_04_within_outcome_equal_AN.csv", "Primary four-component AN equality tests", NA_character_)
  add("inference", "Within_outcome_comparisons", "table_12_within_outcome_all_burden_tests.csv", "Primary within-outcome burden comparisons", NA_character_)
  add("inference", "Daytime_nighttime_tests", "table_11_within_outcome_daytime_nighttime_AN.csv", "Primary daytime-nighttime AN contrasts", NA_character_)
  add("inference", "All_primary_tests", "table_13_all_hypothesis_tests.csv", "All primary hypothesis tests", NA_character_)
  add("inference", "Component_AN_covariance", "table_14_component_AN_covariance.csv", "Primary mean annual component AN covariance", NA_character_)
  add("inference", "Component_profile_inference", "table_05_component_burden_and_profile.csv", "Component shares and inferential eligibility", "heterogeneity_components.csv")
  add("inference", "Between_disease_global", "table_02_between_disease_profile_global.csv", "Between-disease burden-profile omnibus tests", "heterogeneity_between.csv")
  add("inference", "Between_disease_pairwise", "table_03_between_disease_profile_pairwise.csv", "Primary disease-pair profile comparisons", NA_character_)
  add("inference", "Inference_diagnostics", "table_06_monte_carlo_diagnostics.csv", "Monte Carlo and profile stability diagnostics", "heterogeneity_diagnostics.csv")
  add("inference", "Reconstruction_checks", "table_07_stage2_reconstruction_checks.csv", "Burden reconstruction checks", "heterogeneity_validation.csv")
  add("inference", "Inference_settings", "table_09_analysis_settings.csv", "Primary inference settings", NA_character_)
  z[, core := sheet %in% c("Joint_state_RR", "Component_mean_annual", "State_mean_annual",
                           "Component_annual", "State_annual", "Global_risk_tests", "Within_outcome_tests",
                           "Component_profile_inference", "Between_disease_global", "Between_disease_pairwise",
                           "Inference_diagnostics", "Reconstruction_checks", "Within_outcome_comparisons")]
  z[, note := "Source values are copied without statistical recalculation. Unavailable columns are blank."]
  z[sheet == "Joint_state_RR", note := "Nine states including the D0_N0 reference. RR and confidence limits are copied as stored."]
  z[sheet %in% c("Component_mean_annual", "State_mean_annual"), note := "AN denotes mean annual attributable number; AF_percent is the mean of annual attributable fractions. Component shares use mean annual AN. Signed contributions are retained."]
  z[sheet %in% c("Component_annual", "State_annual"), note := "AN and AF_percent refer to the indicated year. Percent columns use the 0-100 scale."]
  z[sheet == "Component_profile_inference", note := "Confidence intervals and eligibility are copied from profile inference; consult the stability diagnostics."]
  z[sheet %in% c("Within_outcome_comparisons", "Daytime_nighttime_tests"), note := "Daytime minus nighttime mean annual AN is tested separately for cold and heat. Its variance uses the full component covariance. Four-component and cold/heat tests form separate Holm families across selected outcomes; family sizes are recorded."]
  z[sheet == "Between_disease_global", note := "Primary inference compares all selected mutually exclusive diseases. Sensitivity inference reports eligible and excluded diseases; missing planned diseases prevent a test. Outcome-specific lag definitions prevent cross-disease comparison. Collection does not recompute omnibus tests when filtering outcomes."]
  z[sheet == "Between_disease_pairwise", note := "Primary disease-pair tests follow a significant omnibus test; Holm adjustment applies across the planned pairs. Share differences are percentage points."]
  z[]
}

log_message <- function(...) message(format(Sys.time(), "%Y-%m-%d %H:%M:%S"),
                                     " | ", paste(..., collapse = " "))

validate_selection <- function(x, allowed, label, allow_empty = FALSE) {
  if (!is.character(x) || anyNA(x) || anyDuplicated(x) || any(!x %in% allowed) ||
      (!allow_empty && !length(x))) stop("Invalid ", label, ".")
  invisible(x)
}

pollution_models <- function() {
  allowed <- c("PM25", "NO2", "O3_8H", "PM25_NO2_O3_8H")
  take <- if (identical(POLLUTION_MODELS_TO_INCLUDE, "all")) allowed else POLLUTION_MODELS_TO_INCLUDE
  validate_selection(take, allowed, "POLLUTION_MODELS_TO_INCLUDE")
  take
}

resolve_primary_directory <- function(code) {
  if (!is.null(MAIN_ANALYSIS_DIRS) && code %in% names(MAIN_ANALYSIS_DIRS))
    return(unname(MAIN_ANALYSIS_DIRS[[code]]))
  base <- file.path(PRIMARY_RESULTS_ROOT, code)
  if (!dir.exists(base)) return(NA_character_)
  candidates <- unique(c(base, list.dirs(base, recursive = FALSE, full.names = TRUE)))
  candidates <- candidates[dir.exists(file.path(candidates, "tables"))]
  if (!is.null(MAIN_ANALYSIS_TAG)) candidates <- candidates[endsWith(basename(candidates), MAIN_ANALYSIS_TAG)]
  if (length(candidates) > 1L) stop("Several primary directories match ", code,
                                    ". Set MAIN_ANALYSIS_TAG or MAIN_ANALYSIS_DIRS explicitly: ", paste(candidates, collapse = " | "))
  if (length(candidates)) candidates[1L] else NA_character_
}

make_model_index <- function() {
  primary <- rbindlist(lapply(OUTCOMES_TO_INCLUDE, function(code) data.table(
    analysis_id = "Main", outcome_code = code, model_variant = "main",
    model_role = "Primary model", analysis_name = "Primary analysis",
    root = resolve_primary_directory(code))))
  parts <- list(primary)
  for (spec in SENSITIVITIES_TO_INCLUDE) {
    variants <- if (spec == "S1_pollution") paste0(pollution_models(), "_adjusted") else "main"
    d <- CJ(outcome_code = OUTCOMES_TO_INCLUDE, model_variant = variants, sorted = FALSE)
    d[, `:=`(analysis_id = spec, analysis_name = unname(SENSITIVITY_NAMES[[spec]]),
             model_role = ifelse(grepl("_adjusted$", model_variant), "Pollution-adjusted", "Sensitivity model"),
             root = file.path(SENSITIVITY_RESULTS_ROOT, spec, outcome_code, model_variant))]
    parts[[length(parts) + 1L]] <- d
  }
  z <- rbindlist(parts, use.names = TRUE)
  z[, analysis_tag := ifelse(analysis_id == "Main", if (is.null(MAIN_ANALYSIS_TAG))
    basename(root) else MAIN_ANALYSIS_TAG, analysis_id)]
  z[, scope := "model"]
  z[]
}

make_inference_index <- function(models) {
  main_root <- MAIN_HETEROGENEITY_DIR
  if (is.null(main_root)) {
    base <- file.path(PRIMARY_RESULTS_ROOT, "_core_heterogeneity_inference")
    if (!is.null(MAIN_ANALYSIS_TAG)) main_root <- file.path(base, MAIN_ANALYSIS_TAG) else {
      candidates <- list.dirs(base, recursive = FALSE, full.names = TRUE)
      candidates <- candidates[dir.exists(file.path(candidates, "tables"))]
      if (length(candidates) > 1L) stop("Set MAIN_HETEROGENEITY_DIR to select the primary inference directory.")
      main_root <- if (length(candidates)) candidates[1L] else NA_character_
    }
  }
  groups <- unique(models[, .(analysis_id, analysis_name, model_variant, model_role)])
  groups[, root := ifelse(analysis_id == "Main", main_root,
                          file.path(SENSITIVITY_RESULTS_ROOT, "_heterogeneity", analysis_id, model_variant))]
  groups[, `:=`(outcome_code = NA_character_, scope = "inference",
                analysis_tag = ifelse(analysis_id == "Main", if (is.null(MAIN_ANALYSIS_TAG)) basename(main_root) else MAIN_ANALYSIS_TAG, analysis_id))]
  groups[]
}

current_analysis_entry <- function(entry) {
  if (identical(entry$analysis_id, "Main")) return(TRUE)
  if (length(entry$root) != 1L || is.na(entry$root)) return(FALSE)
  name <- if (entry$scope == "model") "sensitivity_completed.rds" else "heterogeneity_completed.rds"
  file <- file.path(entry$root, "models", name)
  if (!file.exists(file)) return(FALSE)
  marker <- tryCatch(readRDS(file), error = function(e) NULL)
  if (!is.list(marker) || !identical(marker$run_version, SENSITIVITY_RUN_VERSION)) return(FALSE)
  if (entry$scope == "inference") return(isTRUE(marker$completed))
  identical(marker$outcome_code, entry$outcome_code) &&
    identical(marker$specification, entry$analysis_id) &&
    identical(marker$variant, entry$model_variant)
}

source_files <- function(root) {
  if (length(root) != 1L || is.na(root) || !dir.exists(file.path(root, "tables"))) return(character())
  list.files(file.path(root, "tables"), pattern = "[.]csv$", full.names = FALSE, ignore.case = TRUE)
}

extend_catalog <- function(catalog, models) {
  if (!INCLUDE_ADDITIONAL_MODEL_TABLES) return(catalog)
  current <- vapply(seq_len(nrow(models)), function(i) current_analysis_entry(models[i]), logical(1))
  available <- sort(unique(unlist(lapply(models[current]$root, source_files))))
  extra <- setdiff(available, catalog[scope == "model", primary_file])
  known <- c(table_04a_joint_logrr_covariance.csv = "Coefficient_covariance",
             table_05_state_share_sum_check_annual.csv = "State_share_checks_annual",
             table_06_state_share_sum_check_mean_annual.csv = "State_share_checks_mean",
             table_08_shapley_additivity_check_annual.csv = "Additivity_checks_annual",
             table_08_shapley_additivity_check_mean_annual.csv = "Additivity_checks_mean",
             table_08_shapley_share_sum_check_annual.csv = "Component_share_checks_annual",
             table_08_shapley_share_sum_check_mean_annual.csv = "Component_share_checks_mean",
             table_13_requested_mean_annual_summary.csv = "Combined_mean_summary",
             table_00_input_source_audit.csv = "Input_sample_summary")
  extra <- intersect(extra, names(known))
  used <- c("Readme", "Contents", "Analysis_key", "Availability", "Source_index", catalog$sheet)
  for (file in extra) {
    label <- if (file %in% names(known)) unname(known[[file]]) else
      gsub("[^A-Za-z0-9_]", "_", sub("[.]csv$", "", sub("^table_[0-9]+[a-z]?_", "", file)))
    label <- substr(label, 1, 31); candidate <- label; k <- 1L
    while (tolower(candidate) %in% tolower(used)) {
      candidate <- paste0(substr(label, 1, 26), "_", k); k <- k + 1L
    }
    used <- c(used, candidate)
    title <- gsub("_", " ", sub("[.]csv$", "", sub("^table_[0-9]+[a-z]?_", "", file)))
    note <- if (file == "table_04a_joint_logrr_covariance.csv")
      "Each matrix row retains its state or term label; coefficient columns retain their original names. Matrices are copied without rescaling." else
        "Additional source result table; all original columns are retained."
    catalog <- rbind(catalog, data.table(scope = "model", sheet = candidate,
                                         primary_file = file, title = title, sensitivity_file = file, core = FALSE, note = note), use.names = TRUE)
  }
  catalog[]
}

# 3. Read and combine the current source tables -----------------------------------
preserve_column <- function(x, name) {
  if (!name %in% names(x)) return(invisible(NULL))
  new <- paste0("source_", name); k <- 1L
  while (new %in% names(x)) {new <- paste0("source_", name, "_", k); k <- k + 1L}
  setnames(x, name, new)
  invisible(NULL)
}

label_source <- function(x, entry, filename) {
  x <- copy(x)
  if (entry$scope == "model") {
    if ("outcome_code" %in% names(x)) preserve_column(x, "outcome_code")
    if ("outcome" %in% names(x)) preserve_column(x, "outcome")
    x[, `:=`(outcome_code = entry$outcome_code, outcome = unname(OUTCOME_NAMES[[entry$outcome_code]]))]
  } else {
    if ("outcome_code" %in% names(x)) x <- x[outcome_code %in% OUTCOMES_TO_INCLUDE]
    if (all(c("disease_1", "disease_2") %in% names(x)))
      x <- x[disease_1 %in% OUTCOMES_TO_INCLUDE & disease_2 %in% OUTCOMES_TO_INCLUDE]
    if (!"outcome_code" %in% names(x)) x[, outcome_code := NA_character_]
    if (!"outcome" %in% names(x)) x[, outcome := unname(OUTCOME_NAMES[outcome_code])]
  }
  for (name in c("analysis_id", "analysis_name", "model_variant", "model_role", "analysis_tag", "source_table"))
    preserve_column(x, name)
  x[, `:=`(analysis_id = entry$analysis_id, analysis_name = entry$analysis_name,
           model_variant = entry$model_variant, model_role = entry$model_role,
           analysis_tag = entry$analysis_tag, source_table = filename)]
  x
}

order_result <- function(x, sheet) {
  if (!nrow(x)) return(x)
  x[, export_analysis_order := match(analysis_id, c("Main", names(SENSITIVITY_NAMES)))]
  x[, export_outcome_order := match(outcome_code, names(OUTCOME_NAMES))]
  order_cols <- c("export_outcome_order", "export_analysis_order", "model_variant")
  for (col in c("year", "comparison_order", "component_order", "mask", "state", "component", "disease_1", "disease_2"))
    if (col %in% names(x)) order_cols <- c(order_cols, col)
  setorderv(x, order_cols, na.last = TRUE)
  x[, c("export_analysis_order", "export_outcome_order") := NULL]
  keys <- c("analysis_id", "outcome_code", "model_variant", "comparison_code", "state", "component", "year", "disease_1", "disease_2")
  values <- if (sheet == "Joint_state_RR") c("beta", "se", "rr", "rr_low", "rr_high", "n_cases",
                                             "n_referents", "n_rows", "n_sets", "support_status") else c("AN", "AN_low", "AN_high", "AN_CI_low", "AN_CI_high",
                                                                                                         "AF_percent", "AF_percent_low", "AF_percent_high", "AF_percent_CI_low", "AF_percent_CI_high",
                                                                                                         "share_percent", "share_percent_low", "share_percent_high", "share_CI_low", "share_CI_high",
                                                                                                         "chi_square", "df", "p_value", "p_holm_across_outcomes", "p_holm", "status", "profile_status")
  trailing <- c("analysis_name", "outcome", "model_role", "analysis_tag", "source_table")
  leading <- intersect(c(keys, values), names(x))
  setcolorder(x, c(leading, setdiff(names(x), c(leading, trailing)), intersect(trailing, names(x))))
  x[]
}

collect_tables <- function(models, inference, catalog) {
  entries <- rbindlist(list(models, inference), use.names = TRUE, fill = TRUE)
  source_index <- list(); pieces <- setNames(vector("list", nrow(catalog)), catalog$sheet)
  availability <- list()
  for (i in seq_len(nrow(entries))) {
    entry <- entries[i]
    current <- current_analysis_entry(entry)
    available <- if (current) source_files(entry$root) else character()
    rr_available <- "table_04_joint_9_state_rr.csv" %in% available
    burden_available <- "table_12_shapley_4_component_burden_mean_annual.csv" %in% available
    availability[[i]] <- entry[, .(analysis_id, outcome_code, model_variant, scope,
                                   analysis_name, model_role, analysis_tag)]
    availability[[i]][, `:=`(directory_available = !is.na(entry$root) && dir.exists(file.path(entry$root, "tables")),
                             analysis_version_available = current, csv_files_available = length(available), risk_table_available = if (entry$scope == "model") rr_available else NA,
                             component_burden_available = if (entry$scope == "model") burden_available else NA)]
    types <- catalog[scope == entry$scope]
    for (j in seq_len(nrow(types))) {
      def <- types[j]
      filename <- if (entry$analysis_id == "Main") def$primary_file else def$sensitivity_file
      if (is.na(filename)) next
      present <- filename %in% available
      if (!present && !def$core) next
      error <- ""; n_source <- 0L; retained <- 0L; codes <- character()
      state <- if (present) "read" else "missing"
      if (present) {
        x <- tryCatch(fread(file.path(entry$root, "tables", filename), keepLeadingZeros = TRUE,
                            check.names = FALSE), error = function(e) {error <<- conditionMessage(e); NULL})
        if (is.null(x)) {
          state <- "read_error"
          if (STOP_ON_READ_ERROR) stop("Cannot read ", file.path(entry$root, "tables", filename), ": ", error)
        } else {
          if (!ncol(x) || anyDuplicated(names(x))) stop("Missing or duplicate column names: ", filename)
          n_source <- nrow(x)
          for (col in names(x)) {
            if (inherits(x[[col]], c("integer64", "Date", "POSIXt")))
              set(x, j = col, value = as.character(x[[col]]))
          }
          x <- label_source(x, entry, filename)
          retained <- nrow(x)
          codes <- sort(unique(x$outcome_code[!is.na(x$outcome_code)]))
          if (retained) pieces[[def$sheet]][[length(pieces[[def$sheet]]) + 1L]] <- x
        }
      }
      source_index[[length(source_index) + 1L]] <- data.table(analysis_id = entry$analysis_id,
                                                              outcome_code = entry$outcome_code, model_variant = entry$model_variant, scope = entry$scope,
                                                              worksheet = def$sheet, source_table = filename, status = state, source_rows = n_source,
                                                              exported_rows = retained, outcomes_in_table = paste(codes, collapse = "; "), note = error)
    }
  }
  tables <- lapply(names(pieces), function(sheet) {
    z <- rbindlist(pieces[[sheet]], use.names = TRUE, fill = TRUE)
    if (!ncol(z)) return(NULL)
    for (col in names(z)[grepl("date", names(z), ignore.case = TRUE)]) {
      value <- z[[col]]
      if (is.character(value) && any(!is.na(value)) &&
          all(is.na(value) | grepl("^[0-9]{4}-[0-9]{2}-[0-9]{2}$", value))) {
        converted <- as.Date(value)
        if (all(is.na(value) | !is.na(converted))) set(z, j = col, value = converted)
      }
    }
    order_result(z, sheet)
  })
  names(tables) <- names(pieces)
  tables <- tables[!vapply(tables, is.null, logical(1))]
  list(tables = tables, availability = rbindlist(availability, fill = TRUE),
       sources = rbindlist(source_index, fill = TRUE))
}

# 4. Workbook formatting ---------------------------------------------------------
column_width <- function(name) {
  if (name == "model_variant") return(33)
  if (name == "analysis_id") return(20)
  if (name == "outcome_code") return(21)
  if (name %in% c("state", "component", "year", "df")) return(13)
  if (name %in% c("title", "note", "definition")) return(85)
  if (grepl("file|source_table|analysis_tag", name)) return(49)
  if (grepl("outcome|label|name|hypothesis|interpretation|status|warning|assumption|family", name)) return(38)
  if (grepl("CI$|_text$", name)) return(30)
  max(17, min(26, nchar(name) * 0.65 + 5))
}

number_format <- function(name, value) {
  if (name == "year") return("0")
  if (grepl("^(joint_state)?(D0|DC|DH)_(N0|NC|NH)$", name)) return("0.000000E+00")
  if (grepl("^p_value$|^p_holm|^p_[a-z].*value$", name)) return("0.000E+00")
  if (name %in% c("year", "df", "source_rows", "exported_rows", "rows", "columns", "csv_files_available")) return("#,##0")
  if (grepl("^n_|count$|_order$|^mask$", name) && all(is.na(value) | value == floor(value))) return("#,##0")
  if (grepl("^rr$|^rr_|^beta$|^se$|^highest_RR", name)) return("0.0000")
  if (grepl("^AN$|^AN_|_AN$|n_deaths", name)) return("#,##0.000")
  if (grepl("percent|_pp$", name)) return("0.0000")
  if (grepl("error|difference|covariance", name)) return("0.000000E+00")
  "0.0000"
}

write_sheet <- function(wb, name, x, title, index) {
  if (nrow(x) + 4L > 1048576L || ncol(x) > 16384L)
    stop("Excel worksheet limits exceeded for ", name, "; narrow the outcome/specification selection.")
  openxlsx::addWorksheet(wb, name, gridLines = FALSE)
  openxlsx::writeData(wb, name, title, startRow = 1, colNames = FALSE)
  openxlsx::writeData(wb, name, "Main = primary analysis. Definitions and source coverage are listed in Contents and Source_index.",
                      startRow = 2, colNames = FALSE)
  openxlsx::addStyle(wb, name, openxlsx::createStyle(fontSize = 14, textDecoration = "bold"), rows = 1, cols = 1)
  openxlsx::addStyle(wb, name, openxlsx::createStyle(fontSize = 10, fontColour = "#525252"), rows = 2, cols = 1)
  header <- openxlsx::createStyle(fontColour = "#FFFFFF", fgFill = "#334155", textDecoration = "bold",
                                  halign = "center", valign = "center", wrapText = TRUE, border = "bottom", borderColour = "#CBD5E1")
  display <- copy(x)
  infinite_cells <- list()
  for (j in seq_len(ncol(display))) if (is.numeric(display[[j]])) {
    hit <- which(is.infinite(display[[j]]))
    if (length(hit)) {
      infinite_cells[[length(infinite_cells) + 1L]] <- list(col = j, rows = hit, values = as.character(display[[j]][hit]))
      set(display, i = hit, j = j, value = NA_real_)
    }
  }
  openxlsx::writeDataTable(wb, name, display, startRow = 4, tableName = paste0("Results_", index),
                           tableStyle = "TableStyleLight9", headerStyle = header, keepNA = FALSE, withFilter = TRUE)
  for (item in infinite_cells) for (k in seq_along(item$rows))
    openxlsx::writeData(wb, name, item$values[k], startRow = item$rows[k] + 4L,
                        startCol = item$col, colNames = FALSE)
  widths <- vapply(names(x), column_width, numeric(1))
  if ("value" %in% names(x) && is.character(x$value)) widths[names(x) == "value"] <- 55
  openxlsx::setColWidths(wb, name, seq_len(ncol(x)), widths)
  openxlsx::setRowHeights(wb, name, 1:4, c(24, 20, 8, 52))
  if (nrow(x)) {
    rows <- 5:(nrow(x) + 4L)
    openxlsx::setRowHeights(wb, name, rows, 32)
    for (j in seq_len(ncol(x))) {
      numeric <- is.numeric(x[[j]])
      style <- openxlsx::createStyle(halign = if (numeric) "right" else "left", valign = "center",
                                     wrapText = !numeric, numFmt = if (inherits(x[[j]], "Date")) "yyyy-mm-dd" else
                                       if (numeric) number_format(names(x)[j], x[[j]]) else "GENERAL")
      openxlsx::addStyle(wb, name, style, rows = rows, cols = j, gridExpand = TRUE, stack = TRUE)
    }
    text_cols <- which(!vapply(x, is.numeric, logical(1)))
    if (length(text_cols)) {
      lines <- rep(1, nrow(x))
      for (j in text_cols) {
        value <- as.character(x[[j]]); value[is.na(value)] <- ""
        capacity <- max(8, floor(widths[j] * 1.15))
        needed <- vapply(strsplit(value, "\n", fixed = TRUE), function(s)
          sum(pmax(1, ceiling(nchar(s, type = "width") / capacity))), numeric(1))
        lines <- pmax(lines, needed)
      }
      openxlsx::setRowHeights(wb, name, rows, pmin(409, pmax(32, 14 * lines + 6)))
    }
  }
  openxlsx::freezePane(wb, name, firstActiveRow = 5,
                       firstActiveCol = if (all(c("analysis_id", "outcome_code", "model_variant") %in% names(x))) 4 else 1)
  openxlsx::pageSetup(wb, name, orientation = "landscape", paperSize = 9,
                      fitToWidth = 1, fitToHeight = 0, printTitleRows = 1:4)
  invisible(NULL)
}

write_workbook <- function(collected, catalog, models) {
  tables <- collected$tables
  if (!length(tables)) stop("No readable result tables were found in the configured directories.")
  definitions <- data.table(item = c("Scope", "Primary comparator", "Pollution variants", "Percentages",
                                     "Precision", "Missing values", "Intervals", "Multiplicity", "Coverage", "Refresh", "Version", "Selection"),
                            definition = c("One worksheet per result type, pooling all available selected outcomes and analysis specifications.",
                                           "Main rows are the original primary-analysis results, included once per outcome per result type.",
                                           "Four pollutant-adjusted model variants are collected from the specified analysis version.",
                                           "AF_percent and share_percent are on the 0-100 scale. Columns ending in _pp are percentage-point differences.",
                                           "CSV numeric precision is retained within Excel numeric precision. Number formats only change the display. P values use scientific notation.",
                                           "Blank cells represent unavailable source columns or missing source values, never imputed zeros. Infinite numeric entries are displayed as Inf or -Inf.",
                                           "All confidence intervals are copied as stored. Component_profile_inference contains stability-screened profile intervals. No uncertainty is recalculated here.",
                                           "Each source's P values, multiplicity corrections and test eligibility are retained. Global tests are not recalculated when filtering outcomes.",
                                           "Availability lists expected model and inference groups. Source_index lists imported files, missing core tables and row-level outcome coverage. Missing optional files are not treated as failed analyses.",
                                           "Each invocation reads current per-model and inference CSV files. Recalculate burden and inference upstream after changing risks when those results need updating.",
                                           paste("Sensitivity analysis version:", SENSITIVITY_RUN_VERSION),
                                           paste("Outcomes:", paste(OUTCOMES_TO_INCLUDE, collapse = ", "), "; sensitivities:",
                                                 paste(SENSITIVITIES_TO_INCLUDE, collapse = ", "))))
  contents <- catalog[, .(worksheet = sheet, title, note)]
  contents[, `:=`(rows = vapply(worksheet, function(s) if (s %in% names(tables)) nrow(tables[[s]]) else 0L, integer(1)),
                  columns = vapply(worksheet, function(s) if (s %in% names(tables)) ncol(tables[[s]]) else 0L, integer(1)))]
  contents[, availability := ifelse(rows > 0, "Included", "No available rows")]
  key <- unique(models[, .(analysis_id, model_variant, analysis_name, model_role)])
  all_tables <- c(list(Readme = definitions, Contents = contents, Analysis_key = key), tables,
                  list(Availability = collected$availability, Source_index = collected$sources))
  titles <- c(Readme = "Primary and sensitivity results", Contents = "Workbook contents",
              Analysis_key = "Analysis and model identifiers", setNames(catalog$title, catalog$sheet),
              Availability = "Availability of model and inference results", Source_index = "Source result tables and outcome coverage")
  if (anyDuplicated(tolower(names(all_tables))) || any(nchar(names(all_tables)) > 31L)) stop("Invalid worksheet names.")
  wb <- openxlsx::createWorkbook(creator = "Cardiocerebrovascular mortality study")
  openxlsx::modifyBaseFont(wb, fontSize = 10, fontName = "Arial", fontColour = "#202020")
  for (i in seq_along(all_tables)) {
    sheet <- names(all_tables)[i]
    write_sheet(wb, sheet, all_tables[[i]], titles[[sheet]], i)
  }
  # Tabular worksheets contain no drawings or comments.
  for (i in seq_along(wb$worksheets_rels)) {
    rels <- wb$worksheets_rels[[i]]
    wb$worksheets_rels[[i]] <- rels[!grepl('/relationships/(drawing|vmlDrawing)"', rels)]
  }
  dir.create(dirname(OUTPUT_FILE), recursive = TRUE, showWarnings = FALSE)
  openxlsx::saveWorkbook(wb, OUTPUT_FILE, overwrite = TRUE)
  if (!identical(openxlsx::getSheetNames(OUTPUT_FILE), names(all_tables))) stop("Workbook sheet verification failed.")
  log_message("Saved:", OUTPUT_FILE)
  log_message("Result worksheets:", length(tables), "| imported source tables:", sum(collected$sources$status == "read"),
              "| missing core tables:", sum(collected$sources$status == "missing"))
  invisible(all_tables)
}

# 5. Execute collection ----------------------------------------------------------
collect_main_and_sensitivity <- function() {
  validate_selection(OUTCOMES_TO_INCLUDE, names(OUTCOME_NAMES), "OUTCOMES_TO_INCLUDE")
  validate_selection(SENSITIVITIES_TO_INCLUDE, names(SENSITIVITY_NAMES), "SENSITIVITIES_TO_INCLUDE")
  if (!is.null(MAIN_ANALYSIS_TAG) && (!is.character(MAIN_ANALYSIS_TAG) || length(MAIN_ANALYSIS_TAG) != 1L ||
                                      is.na(MAIN_ANALYSIS_TAG) || !nzchar(MAIN_ANALYSIS_TAG))) stop("MAIN_ANALYSIS_TAG must be a nonempty string or NULL.")
  if (!is.null(MAIN_ANALYSIS_DIRS) && (!is.character(MAIN_ANALYSIS_DIRS) || is.null(names(MAIN_ANALYSIS_DIRS)) ||
                                       anyNA(MAIN_ANALYSIS_DIRS) || anyDuplicated(names(MAIN_ANALYSIS_DIRS)) ||
                                       any(!names(MAIN_ANALYSIS_DIRS) %in% names(OUTCOME_NAMES)))) stop("Invalid MAIN_ANALYSIS_DIRS.")
  if (!dir.exists(PRIMARY_RESULTS_ROOT)) stop("Primary results directory not found: ", PRIMARY_RESULTS_ROOT)
  if (!dir.exists(SENSITIVITY_RESULTS_ROOT)) stop("Sensitivity results directory not found: ", SENSITIVITY_RESULTS_ROOT)
  log_message("Collecting current primary and sensitivity result tables.")
  models <- make_model_index()
  inference <- make_inference_index(models)
  catalog <- extend_catalog(result_catalog(), models)
  collected <- collect_tables(models, inference, catalog)
  output <- write_workbook(collected, catalog, models)
  invisible(list(file = OUTPUT_FILE, tables = output, model_index = models))
}

if (isTRUE(getOption("daynight.collect.run", TRUE))) collect_main_and_sensitivity()
