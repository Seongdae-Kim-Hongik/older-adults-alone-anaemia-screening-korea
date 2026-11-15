# Anaemia flagged at national health screening and older adults living alone in Korean districts, 2011-2024

Analysis code for the manuscript

> *Is more anaemia flagged at national health screening in Korean districts where older adults live alone?*

Archived version: Zenodo, https://doi.org/10.5281/zenodo.24265670 (v1.0.0).

This repository is public under the MIT licence.

## Contents

This repository contains code only. The processed district-year data (master table, analytic panel, expected counts, result tables, adjacency graph and boundary file) are not included; they are available from the author on reasonable request. All source statistics are public (see Data sources).

| Path | What it is |
|---|---|
| `01_스크립트/` | Analysis scripts (below) |
| `04_분석/ANEM/sets_ANEM_*.json`, `config_ANEM_std3.json`, `hypothesis_table_ANEM.csv` | Covariates, expected directions, forced controls and model settings (principal model `std3`; sensitivity versions `s577`, `s2`, `s8`) |
| `04_분석/ANEM/MS/` | Scripts that read the result tables and write every number, table and figure in the manuscript and Online Resource 1, including `content_anem.py`, which holds the manuscript text |

Folder and some column names are in Korean because the source tables are Korean; each script has a header comment describing its inputs and outputs.

## Data sources

All inputs are public, aggregated statistics. No individual-level data are included.

- General health screening attendees with a judgement and attendees flagged as suspected anaemia by district (KOSIS DT_35007_N098, DT_35007_N103) and nationally by sex and age group (DT_35007_N099, DT_35007_N104): National Health Insurance Service via KOSIS, 2011-2024.
- Population by sex and 5-year age group: Ministry of the Interior and Safety, resident registration statistics (Changwon 2023 corrected; `population_override`).
- District covariates: KOSIS regional statistics, the Community Health Survey, National Health Insurance Service screening statistics, Health Insurance Review and Assessment Service, vital statistics and the census. Comparison outcomes (people treated for gastric cancer and for infectious diseases) come from National Health Insurance regional statistics.

## Pipeline

Run from the repository root (set `ST_DATA_ROOT` and `ST_FBDZIP_CORRECTED` to the local copies of the public source files). `rslot.sh` limits the number of concurrent R processes.

| Step | Script | Output |
|---|---|---|
| 1 | `99A2_build_Estd3_anem.py` (uses `38_build_Estd_all.py`, `40_region_crosswalk.py`, `50_build_master_panel.py`) | Attendees and flagged counts by district; expected numbers standardised by sex and age among attendees; population-based expected numbers |
| 2 | `99c_extra_canonical_pool_anem.py` | Basic livelihood security (2021-2023 correction), tertiary education, internal medicine specialists, nutrition-label literacy and other added covariates, joined by district code |
| 3 | `80_build_disease_panel_v.py ANEM` (`ST_VER=std3`) | Analytic panel |
| 4 | `81_fit_inla.R ANEM` (`ST_VER=std3`, `ST_GRAPH=04_분석/ANEM/graph_593.graph`, `ST_THREADS=1`; see `04_분석/ANEM/run_std3_chain.sh`) | Principal model, set comparison, 200-subset multiverse, window models |
| 5 | `85_window_maps.R`, `86_yearly_maps_ehsa.py` (with `86b_ehsa_modified_mk.py`), `87_window_grid.py` | Maps and emerging hot spot analysis |
| 6 | `88g_sens_prep.py ANEM std3 INFECT,GCA`, `88g_sens_suite.R` (`GROUP=1..6`), `88h_ppc_varpart.R`, `88k_struct_kh.R` (nb, struct, review), `88k_struct_kh_anem.R` with `kh_constraints_anem.R` (kh), `89g_sens_combine.py` | Sensitivity analyses, comparison outcomes, Knorr-Held models, review outputs |
| 7 | `04_분석/ANEM/MS/build/run_all_anem.sh` (`build_refs_anem.py`, `figs_anem.py`, `build_ejn_anem.py`, `esm_anem.py`, `check_ejn_anem.py`) | Manuscript, Online Resource 1, figures and checks |

The boundary file is found at `data/boundary/final.shp` by default; set `ST_SHP` to use another path.

## Software

R 4.6.0 with R-INLA 25.10.19, spdep and sf; Python 3.9 with pandas, geopandas, numpy, scipy, matplotlib and python-docx.

## Licence

MIT (see `LICENSE`).
