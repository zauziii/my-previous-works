# This file is used to multi-morbidity variables (for all observed years)
# Zhi Zhou
# Date: 2024-05-27 - The file first created. Remove private providers.

##### Packages #####
library(tidyverse)
library(readxl)
library(parallel)

##### Functions #####
# Help functions will be used later
dir(pattern = "^MM", 
    path = "/home/zzhh/Desktop/RRP-Proj3/RRP_Project3/R/",  
    recursive = T, 
    full.names = T) %>% 
  lapply(source) %>% 
  invisible()

##### Read in data #####
# datrrp all .sas7bdat files
datrrp <- "/home/zzhh/mnt/cdata/RRP_Project3/data/Analysis_data/proj_data/RRP_rawdata_20240527.RData" |> 
  load() |> 
  get()

# Read in diagnosis code
code <- read_excel(
  "/home/zzhh/mnt/documents/HYRA_RRP/Proj3_palvelutarve_kustannukset/MM luokitus_final_ICD10 ja ICPC2_100223_210623_korjattu.xlsx",
  .name_repair = ~ vctrs::vec_as_names(..., repair = "unique", quiet = TRUE)
)
code <- map2(
  # Split data by `Sairausryhmän koodi`
  split(code, f = code$`Sairausryhmän koodi`) %>%
    map( ~ .x %>% subset(select = c("ICD-10 dg:t"))),
  # as to get the codes for each chronic disease
  split(code, f = code$`Sairausryhmän koodi`) %>%
    map( ~ .x %>% subset(select = c("ICPC-2 dg:t"))),
  bind_cols
)
# A list of 35 elements (tbl_df). 
# Each element has 2 columns: one for ICD, another for ICPC



##### STEP 1 - Find ll and ul #####

#### H00 #####
# Year range
# No missing value in KOKO_VPVM in h00 survey data
h00_yrRange <- datrrp$t00_v2023_004_2023_03_24 |>
  subset(select = c("_RANDOM_ID_", "KOKO_VPVM")) |>
  mutate(
    # set origin to 1960-01-01, so the dates are correct
    KOKO_VPVM = as.Date(KOKO_VPVM, origin = "1960-01-01") |>
      decimal_date(),
    LL = KOKO_VPVM - 2,
    UL = KOKO_VPVM
  ) |>
  rename(havtun = "_RANDOM_ID_")
# Check the data type of each column in h00_yrRange
# str(h00_yrRange)
# hvatun: chr
# KOKO_PVM, LL, UL: decimal date

##### H11 #####
# Year Range 
h11_yrRange <- datrrp$t11_v2023_004_2023_04_20 |>
  # Connect X_RANDOM_ID_ with GUMM85ID
  left_join(datrrp$t11_v2023_004_2023_02_09_id, by = "_RANDOM_ID_") |>
  # Missing value appears
  subset(select = c("_RANDOM_ID_", "KOKO_VPVM", "GUMM85ID")) |>
  mutate(
    # format dates
    KOKO_VPVM = as.Date(KOKO_VPVM, origin = "1960-01-01"),
    # Impute missing values with median
    KOKO_VPVM = case_when(is.na(KOKO_VPVM) ~ median(KOKO_VPVM, na.rm = T),
                          TRUE ~ KOKO_VPVM) |>
      decimal_date(),
    LL = KOKO_VPVM - 2,
    UL = KOKO_VPVM,
    GUMM85ID = as.character(GUMM85ID)
  ) |>
  rename(havtun = "_RANDOM_ID_")
# Check the data type of each column in h00_2yrRange
# str(h11_yrRange)
# hvatun, GUMM85OD: chr
# KOKO_PVM, LL, UL: decimal date

##### FT17 #####
# Get 1-year range for participants based on the date from data v2023_004_ft17_2023023 variable ft17_otos_kayntipvm_kaikki
ft17_yrRange <- datrrp$v2023_004_ft17_20230223 |>
  subset(select = c("HAVTUN", "ft17_otos_kayntipvm_kaikki")) |>
  mutate(
    ft17_otos_kayntipvm_kaikki = dmy(ft17_otos_kayntipvm_kaikki),
    ft17_otos_kayntipvm_kaikki = case_when(
      is.na(ft17_otos_kayntipvm_kaikki) ~ median(ft17_otos_kayntipvm_kaikki, na.rm = T),
      TRUE ~ ft17_otos_kayntipvm_kaikki
    ) |>
      decimal_date(),
    LL = ft17_otos_kayntipvm_kaikki - 2,
    UL = ft17_otos_kayntipvm_kaikki
  ) |>
  rename(havtun = HAVTUN) 

##### TS22 #####
ts22_yrRange <- datrrp$data_ts22_20240216 %>%
  subset(select = c("gumm85id", "lopa_created")) %>%
  mutate(
    lopa_created = lopa_created |> as.Date() |> decimal_date(),
    lopa_created = case_when(
      is.na(lopa_created) ~ median(lopa_created, na.rm = T),
      TRUE ~ lopa_created
    ),
    LL = lopa_created - 2,
    UL = lopa_created
  ) %>%
  rename(havtun = "gumm85id")
# Check the data type of each column in ts22_2yrRange
# str(ts22_yrRange)
# hvatun: chr, others: num



##### STEP 2 - Extract hilmo data and limit by range #####

##### H00 #####
# Let's first check the data type
# compare_df_cols_same(datrrp$t00_v2023_004_hilmo_1969_2015_2, datrrp$t00_v2023_004_erik_hilmo_98_13_2)
# By comparison, LAHTO_PVM, SAIRAALA_E_ALA, SAIRAALA_PAL_AL, TULO_PVM are different
h00_hilmo <- datrrp$t00_v2023_004_hilmo_1969_2015_2 |> 
  mutate(TULO_PVM = ymd(TULO_PVM) |> decimal_date(),
         LAHTO_PVM = ymd(LAHTO_PVM) |> decimal_date()) |> 
  bind_rows(datrrp$t00_v2023_004_erik_hilmo_98_13_2 |>  
              mutate(TULO_PVM = as.Date(TULO_PVM, origin = "1960-01-01") %>% decimal_date(),
                     LAHTO_PVM = as.Date(LAHTO_PVM, origin = "1960-01-01") |> decimal_date())) |> 
  rename(havtun = "_RANDOM_ID_",
         paltu = SAIRAALA,
         ea = SAIRAALA_E_ALA,
         pala = SAIRAALA_PAL_AL) |> 
  right_join(h00_yrRange, by = "havtun") |> 
  subset(TULO_PVM >= LL) |> 
  dplyr::select(-UL, -LL, -KOKO_VPVM)

##### H11 #####
# Let's compared the data type
# compare_df_cols(datrrp$t00_v2023_004_hilmo_1969_2015_2, datrrp$terveys2011_hilmo_1994_2013, datrrp$t00_v2023_004_erik_hilmo_98_13_2)
# We need to change SAIRAALA_PAL_AL, SAIRAALA_PAL_AL, TULO_PVM, LAHTO_PVM
h11_hilmo <- bind_rows(
  datrrp$t00_v2023_004_hilmo_1969_2015_2 |>
    right_join(datrrp$t11_v2023_004_2023_02_09_id, by = "_RANDOM_ID_") |>
    mutate(TULO_PVM = ymd(TULO_PVM) |> decimal_date()) |> 
    rename(paltu = SAIRAALA,
           ea = SAIRAALA_E_ALA,
           pala = SAIRAALA_PAL_AL), 
  datrrp$terveys2011_hilmo_1994_2013 |> 
    right_join(datrrp$t11_v2023_004_2023_02_09_id, by = "GUMM85ID") %>% 
    rename(TULO_PVM = HILMO_TUPVA,
           paltu = HILMO_PALTU,
           ea = HILMO_EA,
           pala = HILMO_PALA)  |> 
    mutate(TULO_PVM = ymd(TULO_PVM) |> decimal_date()),
  datrrp$t00_v2023_004_erik_hilmo_98_13_2 |> 
    mutate(TULO_PVM = as.Date(TULO_PVM, origin = "1960-01-01") |> decimal_date(),
           LAHTO_PVM = as.Date(LAHTO_PVM, origin = "1960-01-01")) |>  
    right_join(datrrp$t11_v2023_004_2023_02_09_id, by = "_RANDOM_ID_") |> 
    rename(paltu = SAIRAALA,
           ea = SAIRAALA_E_ALA,
           pala = SAIRAALA_PAL_AL)) |> 
  rename(havtun = "_RANDOM_ID_") |> 
  right_join(h11_yrRange, by = "havtun") |> 
  subset(TULO_PVM >= LL) |> 
  dplyr::select(-UL, -LL, -KOKO_VPVM, -GUMM85ID.x, -GUMM85ID.y)

##### FT17 #####
# Extract hilmo data from df (v2023_004_ft17_hilmo) and limit it according to 5yrRange
ft17_hilmo <- datrrp$v2023_004_ft17_hilmo_2024_03 |>                                
  rename(havtun = HAVTUN,
         paltu = PALTU,
         pala = PALA,
         ea = EA) |>                                     
  right_join(ft17_yrRange, by = "havtun") |>                      
  mutate(tulopvm = dmy(tulopvm) |>  decimal_date()) |>                                   
  subset(tulopvm >= LL) |> 
  dplyr::select(-UL, -LL, -ft17_otos_kayntipvm_kaikki) 

##### TS22 #####
ts22_hilmo <- datrrp$data_ts22_hilmo_96_22_toka |>                              
  rename(havtun = gumm85id) |>  
  mutate(havtun = havtun |> as.character()) |>                                                                                                                      
  right_join(ts22_yrRange, by = "havtun") |>                             
  mutate(tulo_pvm = dmy(tulo_pvm) |> decimal_date()) |> 
  subset(tulo_pvm >= LL) |> 
  dplyr::select(-UL, -LL, -lopa_created) 



##### STEP 3 - Tidy hilmo data #####

##### H00 #####
h00_hilmo_tidy <- h00_hilmo |>
  filter(!((substr(paltu, 1, 1) == '6' &
              !(
                paltu %in% c(
                  "60776", "60718", "60836", "60818", "60827", "60711",
                  "60828", "60792", "60821", "60839", "60834", "60838",
                  "60820", "60788", "60804", "60629", "66260", "60840",
                  "60825"
                )
              )) |
             substr(paltu, 1, 2) == "95" |
             paltu == "01901")) |> 
  mutate(
    # Here we replace blanks & ATC codes with missing value
    # Pattern of ATC code: [A-Z]{1}[0-9]+[A-Z]+
    # [A-Z]{1}: 1 upper class letter
    # [0-9]+: at least 1 number
    # [A-Z]+: at least 1 letter
    across(!c("TULO_PVM", "LAHTO_PVM"),                                                   
           ~ifelse(. == ""|str_detect(., "[A-Z]{1}[0-9]+[A-Z]+"), 
                   NA, 
                   as.character(.))),
    # Match columns with _10, they are using ICD10 codes and/or ATC code
    across(matches("_10"),                                                               
           ~str_extract_all(., "[[:upper:]]\\d{2}", simplify = T) |> 
             str_replace_all(fixed(" "), "")),
    # Match columns paltu|ea|pala
    across(matches("paltu|pala|ea"),                                                      
           ~ str_replace_all(., fixed(" "), ""))                                          
  ) %>%                                                                                  
  do.call(data.frame, .)   
     
##### H11 #####
h11_hilmo_tidy <- h11_hilmo |> 
  filter(!((substr(paltu, 1, 1) == '6' &
              !(
                paltu %in% c(
                  "60776", "60718", "60836", "60818", "60827", "60711",
                  "60828", "60792", "60821", "60839", "60834", "60838",
                  "60820", "60788", "60804", "60629", "66260", "60840",
                  "60825"
                )
              )) |
             substr(paltu, 1, 2) == "95" |
             paltu == "01901")) |> 
  mutate(
    # Change blanks, dots & ATC code to NA
    across(
      !c("HILMO_LPVM", "HILMO_TOIPVM", "TULO_PVM", "LAHTO_PVM", "HILMO_JOPVM"),
      ~ ifelse(
        . == "" | . == "." | str_detect(., "[A-Z]{1}[0-9]+[A-Z]+"),
        NA,
        as.character(.)
      )
    ),
    # Focus on columns with "_10" from H00 hilmo data, and HILMO_?DG??(O|E) columns
    across(
      matches("_10|HILMO_[A-Z]{1}DG[A-Z0-9]{0, }[O|E]"),
      ~ str_extract_all(., "[[:upper:]]\\d{2}", simplify = T)
    ),
    across(matches("paltu|pala|ea"),
           ~ str_replace_all(., fixed(" "), ""))
  ) %>%
  do.call(data.frame, .)

##### FT17 #####
# Clean ft17_hilmo data - diagnosis codes to official form
ft17_hilmo_tidy <- ft17_hilmo |> 
  filter(!((substr(paltu, 1, 1) == '6' &
              !(
                paltu %in% c(
                  "60776", "60718", "60836", "60818", "60827", "60711",
                  "60828", "60792", "60821", "60839", "60834", "60838",
                  "60820", "60788", "60804", "60629", "66260", "60840",
                  "60825"
                )
              )) |
             substr(paltu, 1, 2) == "95" |
             paltu == "01901")) |> 
  mutate(
    across(!c("tulopvm", "lahtopvm"), 
           ~ifelse(. == ""|str_detect(., "[A-Z]{1}[0-9]+[A-Z]+"),
                   NA, 
                   as.character(.))),
    # Code in hilmo data look like, M161, (or H3538)
    # It is a different of writing M16.1, however, 
    #  it is a subcategory of M16
    across(                                                              
      matches("DG"),                                                     
      ~ str_extract_all(., "[[:upper:]]\\d{2}", simplify = T)            
    ),
    across(matches("paltu|pala|ea"),
           ~ str_replace_all(., fixed(" "), ""))            
  ) %>%        
  do.call(data.frame, .) %>% 
  rename(TULO_PVM = tulopvm,
         LAHTO_PVM = lahtopvm)

##### TS22 #####
ts22_hilmo_tidy <- ts22_hilmo |> 
  filter(!((substr(paltu, 1, 1) == '6' &
              !(
                paltu %in% c(
                  "60776", "60718", "60836", "60818", "60827", "60711",
                  "60828", "60792", "60821", "60839", "60834", "60838",
                  "60820", "60788", "60804", "60629", "66260", "60840",
                  "60825"
                )
              )) |
             substr(paltu, 1, 2) == "95" |
             paltu == "01901")) |> 
  mutate(
    across(matches("icd10_"), 
           ~ifelse(. == ""|str_detect(., "[A-Z]{1}[0-9]+[A-Z]{2}"),
                   NA, 
                   as.character(.))),
    across(                                                              
      matches("icd10_"),                                                 
      ~ str_extract_all(., "[[:upper:]]\\d{2}", simplify = T)            
    ),
    across(matches("paltu|pala|ea"),
           ~ str_replace_all(., fixed(" "), ""))) %>%        
  do.call(data.frame, .) %>% 
  rename(TULO_PVM = tulo_pvm,
         LAHTO_PVM = lahto_pvm) |> 
  # To limit the ICD variables ??? Do we still need this, if there is few differences.
  dplyr::select(!"icd10_4o":"icd10_20e") |>                                     
  dplyr::distinct(havtun, ea, TULO_PVM, .keep_all = T)



##### Step 4: Derivation #####

##### H00 #####
# Split by havtun
h00_hilmo_split <- split(h00_hilmo_tidy, f = h00_hilmo_tidy$havtun)
# Derive time-to-event mm variable
h00_hilmo_Mltc <- derive_time_mm(h00_hilmo_split, cores = 8)

##### H11 #####
# Split by havtun
h11_hilmo_split <- split(h11_hilmo_tidy, f = h11_hilmo_tidy$havtun)
# Derive MM variable 
h11_hilmo_Mltc <- derive_time_mm(h11_hilmo_split, cores = 8)

##### FT17 #####
# split by havtun
ft17_hilmo_split <- split(ft17_hilmo_tidy, f = ft17_hilmo_tidy$havtun)
# Derive MM variable
ft17_hilmo_Mltc <- derive_time_mm(ft17_hilmo_split)

##### TS22 #####
# Split by havtun
ts22_hilmo_split <- split(ts22_hilmo_tidy, f = ts22_hilmo_tidy$havtun)
# Derive time-to-event mm variable
ts22_hilmo_Mltc <- derive_time_mm(ts22_hilmo_split, cores = 8)



##### Step 5: Merge and Combination #####
##### H00 #####
# Extract age and sex
h00_hilmo_bgvars <- datrrp$t00_v2023_004_2023_03_24 %>% 
  dplyr::select("_RANDOM_ID_", IKA2, SP2) %>% 
  set_names(c("havtun", "age", "sex"))
# Combine
h00_hilmo_mm <- h00_hilmo_Mltc %>% 
  right_join(h00_hilmo_bgvars, by = "havtun") %>% 
  replace_na(list(mm = 0,
                  time = 0))

##### H00 #####
# Extract
h11_hilmo_bgvars <- datrrp$t11_v2023_004_2023_04_20 %>% 
  dplyr::select("_RANDOM_ID_", IKA2, SP2) %>% 
  set_names(c("havtun", "age", "sex")) 
# Combine
h11_hilmo_mm <- h11_hilmo_Mltc %>% 
  right_join(h11_hilmo_bgvars, by = "havtun") %>% 
  replace_na(list(mm = 0,
                  time = 0))

##### FT17 #####
# Extract basic background variable (age, sex) from (v2023_004_ft17_20230329)
ft17_hilmo_bgvars <- datrrp$v2023_004_ft17_20230329 %>% 
  dplyr::select(HAVTUN, ft17_otos_sukup, ika_poiminta, fr_alue4) %>% 
  mutate(ika_poiminta = floor(ika_poiminta),
         HAVTUN = as.character(HAVTUN)) %>% 
  set_names(c("havtun", "sex", "age", "area"))
# Combine MM variable with background variables
ft17_hilmo_mm <- ft17_hilmo_Mltc %>% 
  right_join(ft17_hilmo_bgvars, by = "havtun") %>% 
  replace_na(list(mm = 0,
                  time = 0))

##### TS22 #####
# Extract age and sex
ts22_hilmo_bgvars <- datrrp$data_ts22_20240216 %>% 
  dplyr::select(gumm85id, ika2, sukupuoli) %>% 
  set_names(c("havtun", "age", "sex"))
# Combine
ts22_hilmo_mm <- ts22_hilmo_Mltc %>% 
  right_join(ts22_hilmo_bgvars, by = "havtun") %>% 
  replace_na(list(mm = 0,
                  time = 0))



##### Step 6: Save #####
save(
  h00_hilmo_mm,
  h11_hilmo_mm,
  ft17_hilmo_mm,
  ts22_hilmo_mm,
  h00_hilmo_tidy,
  h11_hilmo_tidy,
  ft17_hilmo_tidy,
  ts22_hilmo_tidy,
  datrrp,
  file = "/home/zzhh/mnt/cdata/RRP_Project3/data/Analysis_data/proj_data/MM_20240527.RData"
)

##### End of This File #####