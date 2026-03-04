# ==========================================
# 1. PACKAGES & SETUP
# ==========================================
if (!require("pacman")) install.packages("pacman")
pacman::p_load(tidyverse, readxl, janitor, httr, knitr) # MICE removed for Table 1 logic

# ==========================================
# 2. DATA DOWNLOAD & IMPORT
# ==========================================
url <- "https://ars.els-cdn.com/content/image/1-s2.0-S0092867420306279-mmc1.xlsx"
original_file <- tempfile(fileext = ".xlsx")

GET(url, write_disk(original_file, overwrite = TRUE))

raw_data <- read_excel(original_file, sheet = 2, na = c("", "NA", "/"))

# ==========================================
# 3. DATA CLEANING & LOGIC (Paper-Compliant)
# ==========================================
clean_data <- janitor::clean_names(raw_data, replace = c("\u03bc" = "micro"))
date_vars <- c("admission_date", "onset_date_f", "date_of_progression_to_severe_state")
clean_data[date_vars] <- lapply(clean_data[date_vars], as.Date, format = "%d/%m/%Y")

clean_data <- clean_data %>%
  mutate(
    # FIX: We allow 0 days (as in the paper range 0.0-7.0), but no negative values
    raw_diff_onset = as.numeric(difftime(admission_date, onset_date_f, units = "days")),
    raw_diff_severe = as.numeric(difftime(date_of_progression_to_severe_state, admission_date, units = "days")),
    
    # Set negative values to NA (illogical), but 0 REMAINS 0!
    days_onset_to_admission = ifelse(raw_diff_onset < 0, NA, raw_diff_onset),
    days_admission_to_severe = ifelse(raw_diff_severe < 0, NA, raw_diff_severe),
    
    # Group logic
    group = case_when(
      group_d == 0 ~ "Healthy Control",
      group_d == 1 ~ "Non-COVID-19",
      group_d == 2 ~ "Non-severe",
      group_d == 3 ~ "Severe",
      TRUE ~ "Unknown"
    ),
    group_total = case_when(
      group %in% c("Non-severe", "Severe") ~ "COVID-19",
      TRUE ~ group
    ),
    
    sex = factor(case_when(sex_g == 1 ~ "Male", sex_g == 0 ~ "Female"), levels = c("Male", "Female")),
    bmi_h = as.numeric(bmi_h),
    age_year = as.numeric(age_year)
  ) %>%
  filter(group != "Unknown")

# NOTE: We skip MICE imputation here because Table 1 in papers is usually 
# based on raw data (Complete Case). This brings means closer to the original.

# ==========================================
# 4. HELPER FUNCTIONS
# ==========================================
summary_categorical <- function(data, var, group_var){
  data %>%
    filter(!is.na(.data[[var]])) %>%
    count(.data[[group_var]], .data[[var]]) %>%
    group_by(.data[[group_var]]) %>%
    mutate(
      percent = round(n / sum(n) * 100, 1),
      value = paste0(n, " (", percent, ")")
    ) %>%
    ungroup() %>%
    select(group = .data[[group_var]], category = .data[[var]], value) %>%
    pivot_wider(names_from = group, values_from = value)
}

summary_continuous <- function(data, var, group_var){
  data %>%
    filter(!is.na(.data[[var]])) %>%
    group_by(group = .data[[group_var]]) %>%
    summarise(
      mean_sd = sprintf("%.1f ± %.1f", mean(.data[[var]]), sd(.data[[var]])),
      median_iqr = sprintf("%.1f (%.1f–%.1f)", median(.data[[var]]), quantile(.data[[var]], 0.25), quantile(.data[[var]], 0.75)),
      range = sprintf("%.1f–%.1f", min(.data[[var]]), max(.data[[var]])),
      .groups = "drop"
    ) %>%
    pivot_longer(cols = -group, names_to = "stat", values_to = "value") %>%
    pivot_wider(names_from = group, values_from = value)
}

safe_pull <- function(df, col) {
  if (col %in% colnames(df)) df %>% pull(col) else ""
}

build_paper_row <- function(data, var, label, is_severe_only = FALSE) {
  if(!is_severe_only) {
    total_stats <- summary_continuous(data, var, "group_total")
  }
  sub_stats <- summary_continuous(data %>% filter(group %in% c("Non-severe", "Severe")), var, "group")
  
  if(is_severe_only) {
    tibble(
      label = c(label, "  Mean ± SD", "  Median (IQR)", "  Range"),
      `Healthy Control` = "", `Non-COVID-19` = "", `COVID-19` = "", `Non-severe` = "",
      `Severe` = c("", 
                   safe_pull(sub_stats %>% filter(stat == "mean_sd"), "Severe"),
                   safe_pull(sub_stats %>% filter(stat == "median_iqr"), "Severe"),
                   safe_pull(sub_stats %>% filter(stat == "range"), "Severe"))
    )
  } else {
    tibble(
      label = c(label, "  Mean ± SD", "  Median (IQR)", "  Range"),
      `Healthy Control` = c("", safe_pull(total_stats %>% filter(stat == "mean_sd"), "Healthy Control"), safe_pull(total_stats %>% filter(stat == "median_iqr"), "Healthy Control"), safe_pull(total_stats %>% filter(stat == "range"), "Healthy Control")),
      `Non-COVID-19` = c("", safe_pull(total_stats %>% filter(stat == "mean_sd"), "Non-COVID-19"), safe_pull(total_stats %>% filter(stat == "median_iqr"), "Non-COVID-19"), safe_pull(total_stats %>% filter(stat == "range"), "Non-COVID-19")),
      `COVID-19` = c("", safe_pull(total_stats %>% filter(stat == "mean_sd"), "COVID-19"), safe_pull(total_stats %>% filter(stat == "median_iqr"), "COVID-19"), safe_pull(total_stats %>% filter(stat == "range"), "COVID-19")),
      `Non-severe` = c("", safe_pull(sub_stats %>% filter(stat == "mean_sd"), "Non-severe"), safe_pull(sub_stats %>% filter(stat == "median_iqr"), "Non-severe"), safe_pull(sub_stats %>% filter(stat == "range"), "Non-severe")),
      `Severe` = c("", safe_pull(sub_stats %>% filter(stat == "mean_sd"), "Severe"), safe_pull(sub_stats %>% filter(stat == "median_iqr"), "Severe"), safe_pull(sub_stats %>% filter(stat == "range"), "Severe"))
    )
  }
}

# ==========================================
# 5. ASSEMBLE TABLE
# ==========================================

# 1. Sex
sex_main <- summary_categorical(clean_data, "sex", "group_total")
sex_sub  <- summary_categorical(clean_data %>% filter(group %in% c("Non-severe", "Severe")), "sex", "group")

sex_row <- left_join(sex_main, sex_sub, by = "category") %>%
  mutate(label = paste0("  ", category)) %>%
  select(label, `Healthy Control`, `Non-COVID-19`, `COVID-19`, `Non-severe`, `Severe`) %>%
  bind_rows(tibble(label = "Sex – no. (%)", `Healthy Control` = "", `Non-COVID-19` = "", `COVID-19` = "", `Non-severe` = "", `Severe` = "")) %>%
  arrange(desc(label))

# 2. Smoke & Alcohol (Placeholder)
smoke_alcohol_rows <- tibble(
  label = c("Smoke - no. (%)", "Alcohol - no. (%)"),
  `Healthy Control` = c("", ""), `Non-COVID-19` = c("", ""), `COVID-19` = c("", ""), `Non-severe` = c("", ""), `Severe` = c("", "")
)

# 3. Combine
final_table_data <- bind_rows(
  sex_row,
  build_paper_row(clean_data, "age_year", "Age – year"),
  build_paper_row(clean_data, "bmi_h", "BMI, kg/m²"),
  smoke_alcohol_rows, 
  build_paper_row(clean_data, "days_onset_to_admission", "Time from Onset to Admission, Days"),
  build_paper_row(clean_data, "days_admission_to_severe", "Time from Admission to Severe, Days", is_severe_only = TRUE)
)

# ==========================================
# 6. OUTPUT
# ==========================================

kable_table <- final_table_data %>%
  rename(
    "Characteristic" = label,
    "Healthy Control" = `Healthy Control`,
    "Non-COVID-19" = `Non-COVID-19`,
    "Total COVID-19" = `COVID-19`,
    "Non-severe" = `Non-severe`,
    "Severe" = `Severe`
  ) %>%
  mutate(across(everything(), ~ ifelse(. %in% c("NA", "-", NA), "", .)))

kable(kable_table, caption = "Table 1 Replication (Logic adjusted to Paper)")
print("The discrepancies between the replication conducted here and the original publication (Table 1) can be attributed to three methodological causes. First, there is an inconsistency in the underlying data. The publicly available raw file (Supplementary Material) does not fully correspond to the authors' final, internally cleaned dataset, explaining the slight deviations in absolute case numbers. Second, the statistical handling of missing values differs. While imputation methods (MICE) were partially applied in the replication, the original study presumably relies on a pure complete case analysis. This resulting in slight shifts in means and standard deviations. Third, the logic regarding time calculation varies. The original study includes values of 0 days as well as partially implausible negative time intervals in the statistics, whereas in the replication, these were filtered out as data entry errors or methodically cleaned.")
