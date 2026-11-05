# LMM_Analysis — ANOVA designs fitted as linear (mixed) models

`LMM_Analysis_v6.m` is an interactive MATLAB tool that analyses one-way and two-way ANOVA designs, with or without repeated measures. It fits each design as a linear model or a linear mixed model and runs a full set of post hoc comparisons. Every result is written to a single Excel workbook.

You choose the design, data file, columns and options through dialog boxes, so no programming is needed.

| # | Design | Example | Model |
|---|---|---|---|
| 1 | **One-way ANOVA**, independent groups | Licks per session in 3 groups of different animals | Linear model |
| 2 | **One-way repeated-measures ANOVA** | Licks per session across 5 days, same animals every day | Linear mixed model |
| 3 | **Two-way ANOVA**, independent groups | Licks by genotype × treatment, different animals in every combination | Linear model |
| 4 | **Two-way repeated-measures ANOVA**, both factors within-subject | Licks by day × trial type, every animal in every combination | Linear mixed model |
| 5 | **Two-way mixed ANOVA**, one within-subject and one between-subjects factor | Licks by day (within) × group (between) | Linear mixed model |

### Why use mixed models instead of classical ANOVA?

- **Same answers with complete data.** With complete, balanced data, the models give the same F-tests and post hoc results as the classical (sphericity-assumed) ANOVAs.
- **Missing values don't cost whole subjects.** Classical repeated-measures ANOVA drops any subject with a missing value. Here, that subject's other measurements are still used.
- **Incomplete designs still work.** Designs where some factor combinations have no data at all are analysed with a cell-means model (see [Incomplete and nested designs](#incomplete-and-nested-designs)).

## Contents

- [Requirements](#requirements)
- [Preparing the input file](#preparing-the-input-file)
- [Running the analysis](#running-the-analysis)
- [How the analysis works](#how-the-analysis-works)
- [Output workbook](#output-workbook)
- [Interpreting and reporting the results](#interpreting-and-reporting-the-results)
- [Version history](#version-history)

## Requirements

- **MATLAB R2020b or later**
- **Statistics and Machine Learning Toolbox** (`fitlm`, `fitlme`, `coefTest`, `mvtcdf`, …)
- **Desktop MATLAB.** The script is driven by dialog boxes, so it cannot run in `-batch` mode or without a display.

No other files are needed. All helper functions are included at the end of the script.

## Preparing the input file

**Accepted formats:** `.xlsx`, `.xls`, `.xlsm`, `.xlsb` or `.csv`. If an Excel file has several sheets, you are asked which one to use.

### Data must be in long format

Use **one row per measurement**, with one column per variable and the **column names in row 1**. The column names can be anything, because you pick which column plays which role while the script runs. Extra columns are ignored.

| Column | Needed for designs | Contents |
|---|---|---|
| **Response** | all | The measured value (numeric) |
| **Factor A** | all | The (first) factor, such as group, day or condition |
| **Factor B** | 3, 4, 5 | The second factor |
| **Subject ID** | 2, 4, 5 | Identifies which subject (for example, which animal) each row belongs to |

Each role must come from a different column.

**Example: two-way mixed ANOVA** (design 5, with day within subjects and group between subjects):

| Animal | Group | Day | Licks |
|---|---|---|---|
| M01 | Control | 1 | 42 |
| M01 | Control | 2 | 55 |
| M01 | Control | 3 | 61 |
| M02 | Control | 1 | 38 |
| M02 | Control | 2 | NA |
| … | … | … | … |
| M07 | Treated | 1 | 40 |
| M07 | Treated | 2 | 47 |

**Wide-format data** (one row per subject, one column per day) must be converted to long format first. In MATLAB, `stack` does this; in Excel, use Power Query → **Unpivot Columns**.

### Factor levels

- **Factor levels can be text or numbers.** Levels that are all numeric (such as days 1, 2, …, 10) are sorted numerically, so 10 comes after 9 rather than after 1.
- **Factors are always treated as categorical,** never as continuous slopes. Day 1, Day 2, … are separate levels.
- **Level names appear throughout the output,** so short, clear names make the results easier to read.

### Missing values

Empty cells, and cells containing `NA`, `N/A`, `NaN`, `.`, `-`, `--`, `NULL`, `missing` or `#N/A` (in upper or lower case), are treated as missing. Non-numeric text in the response column also counts as missing.

A row is removed only if it has a missing value in a column used by the analysis. Every removed row is listed, with its Excel row number and the reason, in the `Rows_Removed` sheet. Row numbers assume the column names are in row 1.

### Subject IDs (repeated-measures designs)

- **Each subject needs one consistent ID** across all of its rows.
- **Two-way mixed design (5):** each subject must belong to only one level of the between-subjects factor. If an ID appears in more than one group, the script asks whether to treat these as different subjects (ID + group) or to cancel.
- **Duplicate values:** if a subject has **more than one value in the same condition**, the script warns you, since this usually means two animals share an ID by mistake. You can cancel to fix the file, or continue and treat the values as repeated measurements of one subject.

The outcome of these checks is recorded in the `Data_Checks` sheet.

## Running the analysis

Open `LMM_Analysis_v6.m` in MATLAB and press **Run**, or type `LMM_Analysis_v6` in the Command Window with the script's folder on the path. Then answer the dialogs in this order:

1. **Analysis type:** choose one of the five designs.
2. **Data file**, and then the **sheet** if the workbook has more than one.
3. **Columns:**
   - the response column
   - factor A (for design 5, the **within-subject** factor, such as day)
   - factor B (for design 5, the **between-subjects** factor, such as group)
   - the subject ID column
4. **Names for the output** (optional). Give readable names for the response and each factor, such as `Licks`, `Day` and `Group`. These names label every table and sheet. The defaults are the column names.
5. **Data checks:** answer any duplicate-ID or subject-in-several-groups warnings.
6. *(Model fitting runs, and the model and ANOVA table are printed in the Command Window.)*
7. **Reference comparisons** (optional). Answer **Yes** to compare conditions against a reference, such as each test day vs. baseline. Then:
   - Choose one or more **reference** conditions. Choosing several averages them into a single baseline.
   - Choose the conditions to **compare with** the reference.

   For two-way designs, the conditions are the individual A × B combinations.
8. **Save location** for the results workbook. The default is `<input name>_<design>_LMM_results.xlsx` next to the input file, where `<design>` is `oneway`, `oneway_rm`, `twoway`, `twoway_rm` or `mixed`. Cancelling this dialog saves to the default location. An existing file with the same name is overwritten, so close it in Excel first.

A message box confirms when the analysis is finished.

## How the analysis works

### Models

All factors are categorical and use effects (sum-to-zero) coding, so the ANOVA F-tests are **Type III**.

| Design | Model (Wilkinson notation) | MATLAB function |
|---|---|---|
| One-way ANOVA | `Y ~ A` | `fitlm` |
| One-way RM ANOVA | `Y ~ A + (1\|Subject)` | `fitlme` |
| Two-way ANOVA | `Y ~ A*B` | `fitlm` |
| Two-way RM ANOVA | `Y ~ A*B + (1\|Subject) + (1\|Subject:A) + (1\|Subject:B)` | `fitlme` |
| Two-way mixed ANOVA | `Y ~ A*B + (1\|Subject)` (A within, B between) | `fitlme` |

- **Mixed models** are fitted by REML. If REML fails, the script falls back to ML.
- **Degrees of freedom** use the Satterthwaite approximation for repeated-measures designs, falling back to residual df if that cannot be computed. Independent-groups designs use the residual df.
- **Two-way RM ANOVA:** the three random intercepts give the same error strata as a classical two-way RM ANOVA. If this structure cannot be estimated (for example, with very little data), the model falls back to `(1|Subject)` only, and a note is added to `Model_Fit`.
- **Sphericity:** random intercepts imply compound symmetry, equivalent to sphericity-assumed ANOVA. No Greenhouse-Geisser or Huynh-Feldt correction is applied.

### Incomplete and nested designs

If some A × B combinations contain no data, the full factorial model cannot be estimated. The script then:

- fits a **cell-means model**, with one mean per combination that has data;
- builds the ANOVA F-tests from **estimable contrasts** only:
  - A and B marginal means, each averaged over the levels of the other factor that have data;
  - interaction contrasts whose four cells all have data;
  - an overall test of "any difference among all cell means".

  Terms that cannot be tested are marked as not estimable;
- reports which combinations have data in `Design_Check`.

**Nested layouts** (two-way RM ANOVA) are a special case, where each level of one factor occurs with only one level of the other. An example is when each training day belongs to only one training type. Here the random effects become `(1|Subject) + (1|Subject:outer factor)`:

- The outer factor is tested against subject-to-subject variation in its effect.
- The ANOVA table adds a pooled test of the inner factor within the outer one.
- The Subject × inner-factor term cannot be separated from residual error, so it is omitted. A note explaining this is added to `Model_Fit`.

### Estimated marginal means (EMMs)

Post hoc tests compare **estimated marginal means**: the model's estimate of each level's mean, in the units of the response.

- Marginal means of one factor average the cells of the other factor **equally** (unweighted), whatever the number of observations in each cell.
- EMMs can therefore differ from the raw means in `Descriptives` when data are unbalanced or missing. The EMMs are the values the tests are based on.

### Post hoc comparisons

Each comparison is a t-test of a linear contrast of the model's estimates. The difference is reported as **first minus second**. The families of comparisons are:

| Design | Sheet | Comparisons |
|---|---|---|
| One-way | `PH_<A>` | All pairs of levels of A |
| Two-way | `PH_<A>_main` | All pairs of levels of A, averaged over B |
| | `PH_<B>_main` | All pairs of levels of B, averaged over A |
| | `PH_<A>_within_<B>` | Pairs of levels of A **at each level of B** (simple effects) |
| | `PH_<B>_within_<A>` | Pairs of levels of B **at each level of A** (simple effects) |
| | `PH_Interaction_contrasts` | Whether the A difference changes between two levels of B: `(a1 − a2 at b1) − (a1 − a2 at b2)` |
| | `PH_All_Cells` | Every A × B combination vs. every other |
| Any (optional) | `PH_vs_Reference` | Selected conditions minus the chosen reference |

**P-value adjustments are applied separately within each `Group`** of each sheet. For example, in `PH_Day_within_Group`, the day comparisons inside each group form their own family. Every adjustment is reported side by side, so you can report whichever one you specified in advance:

| Column | Method | Controls |
|---|---|---|
| `p_unadjusted` | None | — |
| `p_Tukey` | Tukey / Tukey-Kramer: exact studentized range p-value (numerical integration), the same method as R's `ptukey`, SPSS, GraphPad and Origin. Given only for families of all pairwise comparisons; empty otherwise. | FWER |
| `p_Bonferroni` | Bonferroni | FWER |
| `p_Holm_Bonferroni` | Holm-Bonferroni step-down: a good general-purpose choice | FWER |
| `p_Sidak` | Šidák | FWER |
| `p_Hochberg` | Hochberg step-up | FWER |
| `p_FDR_BH` | Benjamini-Hochberg | False discovery rate |
| `p_FDR_BY` | Benjamini-Yekutieli (valid under any dependence) | False discovery rate |
| `p_SingleStep_maxT` | Single-step max-\|t\|, from 100,000 simulations of the contrasts' joint distribution. This is the Tukey-type adjustment for families that are not all-pairwise, such as interaction contrasts. | FWER |
| `p_Scheffe` | Scheffé: protects all possible contrasts in the family | FWER |
| `p_Dunnett` | *(`PH_vs_Reference` only)* Dunnett many-to-one, from the exact multivariate t probability. Uses the smallest df in the family, so it is slightly conservative. | FWER |

(FWER = family-wise error rate, the chance of at least one false positive in the family.)

The simulation-based p-values use a fixed random seed (`rng(1)`), so re-running the same analysis gives identical values. They carry a Monte Carlo error of about ±0.001.

## Output workbook

The workbook contains the sheets below, in this order. Sheet names use your factor names, with characters Excel does not allow (`: \ / ? * [ ]`) replaced by `_`. Names are shortened to Excel's 31-character limit, and a suffix such as `_2` is added if two names collide. A post hoc sheet with no possible comparisons contains a single `Note` row.

| Sheet | Contents |
|---|---|
| `README` | What every sheet and column means, plus the input file, sheet and columns used, and the run date |
| `Model_Fit` | Analysis type, formula (with your factor names), relation to classical ANOVA, fit method, df method, number of observations, log-likelihood, AIC, BIC, residual variance and SD, R², and any fitting warnings. For RM designs it also gives the number of subjects, the subject variance and the ICC. Notes about incomplete, nested or fallback models are added here too. |
| `ANOVA_TypeIII` | One row per term: `Term`, `FStat`, `DF1`, `DF2`, `pValue`. For incomplete designs, a `Note` column explains each test. |
| `Fixed_Effects` | Model coefficients with SE, t, df, p and 95 % CI. With effects coding these are deviations from the grand mean and are rarely reported directly; use the EMM sheets. For incomplete designs, a `Meaning` column gives the cell each coefficient represents. |
| `Random_Effects_Variance` | *(RM designs)* Variance components: `Group`, `Name1`, `Name2`, `Type`, `Estimate`, `Lower`, `Upper`, and `Variance` (the square of standard-deviation estimates). Includes the residual (error) term. |
| `Random_Effects_Subjects` | *(RM designs)* Each subject's predicted random effect: its estimated deviation from the average |
| `EMM_<A>`, `EMM_<B>`, `EMM_<A>_x_<B>` | Estimated marginal means: `Factor`, `Level`, `Mean`, `SE`, `df`, `CI95_lower`, `CI95_upper`, and a `Note` when a mean averages over only the levels with data. One-way designs have only `EMM_<A>`. |
| `PH_…` | Post hoc comparisons (see the table above and the column list below) |
| `Descriptives` | Raw summary of the data used, per level or cell: `N_obs`, `N_subjects` (RM designs), `Mean`, `SD`, `SEM`, `Median`. `SEM` is `SD / √N_obs`, computed over observations rather than subjects. |
| `Design_Check` | Number of observations in every A × B combination, and whether it has data |
| `Data_Checks` | Results of the subject-ID checks |
| `Missing_Data` | Rows in the file, rows removed for each reason, and rows analysed |
| `Rows_Removed` | `ExcelRow` and `Reason` for every removed row |
| `Data_Used` | Every analysed row: `ExcelRow`, `Value`, factor levels, `Subject`, `Fitted`, `Residual_raw`, `Residual_standardized` |

**Post hoc sheet columns:**

| Column | Meaning |
|---|---|
| `Family` | The set of comparisons, such as "Day (main effect)" |
| `Group` | The subset within which p-values are adjusted, such as "Within Group Control" |
| `Comparison` | Which means are compared, written as `first − second` |
| `Difference`, `SE`, `df`, `t` | Estimated difference, its standard error, degrees of freedom and t statistic |
| `CI95_lower`, `CI95_upper` | **Unadjusted** 95 % confidence interval for the difference |
| `p_unadjusted` … `p_Scheffe` | p-values (see the adjustment table above). `p_Dunnett` appears only in `PH_vs_Reference`. |
| `q_Tukey`, `k_means` | Studentized range statistic and number of means used for the Tukey p-value |
| `Significant_Holm_Bonferroni_0p05`, `Significant_FDR_0p05` | `TRUE` if the Holm-Bonferroni or Benjamini-Hochberg p-value is below 0.05 |

## Interpreting and reporting the results

- **Decide on one adjustment method in advance** and report that one. Holm-Bonferroni is a sound default for controlling false positives. Tukey suits complete sets of pairwise comparisons, Dunnett suits comparisons against a control or baseline, and Benjamini-Hochberg suits exploratory analyses with many comparisons.
- **Report EMMs, not coefficients.** Report means and differences from the `EMM_…` and `PH_…` sheets. The `Fixed_Effects` coefficients are not directly interpretable under effects coding.
- **Choose post hoc families to match the ANOVA result.** If the interaction is significant, the simple-effects sheets (`PH_<A>_within_<B>`, `PH_<B>_within_<A>`) are usually more informative than the main-effect sheets.
- **`PH_All_Cells` is strict.** It adjusts across every pair of cells, so its p-values are more conservative than those of the more targeted sheets.
- **Check the model assumptions.** Plot `Residual_standardized` against `Fitted` from `Data_Used` to check for roughly constant spread, and look at a histogram or Q-Q plot of the residuals to check approximate normality.
- **Reproducibility.** The `README` and `Model_Fit` sheets record the input file, columns, design, formula and methods used. Keep the workbook with any reported result.

## In this study

<!-- TODO: which datasets were analysed with which design, the response and factor names used,
     which adjustment was reported, and which figure / table each result appears in. -->

LINEAR MIXED EFFECTS MODELS WERE RUN FOR REPEATED MEASUREMENTS OF BEHAVIORAL PERFORMANCE ACROSS PARADIGMS IN FIGURES 4 AND S5, AS WELL AS REPEATED MEASURMENTS OF DRN CALCIUM ACTIVITY DURING AR-SAT IN FIGURE 6. 
POST HOC COMPARISONS WERE MADE USING THE HOLM-BONFERRONI CORRECTION FOR P-VALUES.

## Version history

- **v1:** linear mixed model (`fitlme`) with post hoc tests and Excel export.
- **v2:** handles designs where some factor combinations have no data (cell-means model, estimable-contrast ANOVA, `Design_Check` sheet).
- **v3:** choice of five ANOVA designs; user-defined factor names; subject-ID checks.
- **v4:** nested two-way RM layouts, with `(1|Subject) + (1|Subject:outer factor)` random effects.
- **v5:** exact Tukey / Tukey-Kramer p-values from the studentized range distribution; simulated max-t kept for non-pairwise families.
- **v6:** optional comparisons against a reference condition, with Dunnett-adjusted p-values (`PH_vs_Reference`).
