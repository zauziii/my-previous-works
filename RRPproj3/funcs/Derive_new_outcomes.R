# This file is used to derive new outcome variables
# Zhi Zhou
# Date: 2024-03-21

##### Packages #####
library(tidyverse)
library(haven)
library(janitor)

##### Load #####
datrrp <- "/home/zzhh/mnt/cdata/RRP_Project3/data/Analysis_data/proj_data/RRP_rawdata_20240527.RData" |> 
  load() |> 
  get()

##### STEP 1 - Find ll and ul after exam date #####

#### H00 #####
# Year range
# No missing value in KOKO_VPVM in h00 survey data
h00_yrRange <- datrrp$t00_v2023_004_2023_03_24 |>
  subset(select = c("_RANDOM_ID_", "KOKO_VPVM")) |>
  mutate(
    # set origin to 1960-01-01, so the dates are correct
    KOKO_VPVM = as.Date(KOKO_VPVM, origin = "1960-01-01") |>
      decimal_date(),
    LL = KOKO_VPVM,
    UL = KOKO_VPVM + 1
  ) |>
  rename(havtun = "_RANDOM_ID_")

# Check the data type of each column in h00_yrRange
# str(h00_yrRange)
# hvatun: chr
# KOKO_PVM, LL, UL: decimal date

aux_h00_yrRange <- datrrp$t00_v2023_004_2023_03_24 |>
  subset(select = c("_RANDOM_ID_", "KOKO_VPVM")) |>
  mutate(
    # set origin to 1960-01-01, so the dates are correct
    KOKO_VPVM = as.Date(KOKO_VPVM, origin = "1960-01-01") |>
      decimal_date(),
    LL = KOKO_VPVM - 1,
    UL = KOKO_VPVM
  ) |>
  rename(havtun = "_RANDOM_ID_")


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
    LL = KOKO_VPVM,
    UL = KOKO_VPVM + 1,
    GUMM85ID = as.character(GUMM85ID)
  ) |>
  rename(havtun = "_RANDOM_ID_")

# Check the data type of each column in h00_2yrRange
# str(h11_yrRange)
# hvatun, GUMM85OD: chr
# KOKO_PVM, LL, UL: decimal date

aux_h11_yrRange <- datrrp$t11_v2023_004_2023_04_20 |>
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
    LL = KOKO_VPVM - 1,
    UL = KOKO_VPVM,
    GUMM85ID = as.character(GUMM85ID)
  ) |>
  rename(havtun = "_RANDOM_ID_")


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
    LL = ft17_otos_kayntipvm_kaikki,
    UL = ft17_otos_kayntipvm_kaikki + 1
  ) |>
  rename(havtun = HAVTUN) 

aux_ft17_yrRange <- datrrp$v2023_004_ft17_20230223 |>
  subset(select = c("HAVTUN", "ft17_otos_kayntipvm_kaikki")) |>
  mutate(
    ft17_otos_kayntipvm_kaikki = dmy(ft17_otos_kayntipvm_kaikki),
    ft17_otos_kayntipvm_kaikki = case_when(
      is.na(ft17_otos_kayntipvm_kaikki) ~ median(ft17_otos_kayntipvm_kaikki, na.rm = T),
      TRUE ~ ft17_otos_kayntipvm_kaikki
    ) |>
      decimal_date(),
    LL = ft17_otos_kayntipvm_kaikki - 1,
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
    LL = lopa_created,
    UL = lopa_created + 1
  ) %>%
  rename(havtun = "gumm85id")
# Check the data type of each column in ts22_2yrRange
# str(ts22_yrRange)
# hvatun: chr, others: num

aux_ts22_yrRange <- datrrp$data_ts22_20240216 %>%
  subset(select = c("gumm85id", "lopa_created")) %>%
  mutate(
    lopa_created = lopa_created |> as.Date() |> decimal_date(),
    lopa_created = case_when(
      is.na(lopa_created) ~ median(lopa_created, na.rm = T),
      TRUE ~ lopa_created
    ),
    LL = lopa_created - 1,
    UL = lopa_created
  ) %>%
  rename(havtun = "gumm85id")

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
  subset(between(TULO_PVM, LL, UL)) |> 
  dplyr::select(-UL, -LL, -KOKO_VPVM)

aux_h00_hilmo <- datrrp$t00_v2023_004_hilmo_1969_2015_2 |> 
  mutate(TULO_PVM = ymd(TULO_PVM) |> decimal_date(),
         LAHTO_PVM = ymd(LAHTO_PVM) |> decimal_date()) |> 
  bind_rows(datrrp$t00_v2023_004_erik_hilmo_98_13_2 |>  
              mutate(TULO_PVM = as.Date(TULO_PVM, origin = "1960-01-01") %>% decimal_date(),
                     LAHTO_PVM = as.Date(LAHTO_PVM, origin = "1960-01-01") |> decimal_date())) |> 
  rename(havtun = "_RANDOM_ID_",
         paltu = SAIRAALA,
         ea = SAIRAALA_E_ALA,
         pala = SAIRAALA_PAL_AL) |> 
  right_join(aux_h00_yrRange, by = "havtun") |> 
  subset(between(TULO_PVM, LL, UL)) |> 
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
  subset(between(TULO_PVM, LL, UL)) |> 
  dplyr::select(-UL, -LL, -KOKO_VPVM, -GUMM85ID.x, -GUMM85ID.y)

aux_h11_hilmo <- bind_rows(
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
  right_join(aux_h11_yrRange, by = "havtun") |> 
  subset(between(TULO_PVM, LL, UL)) |> 
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
  subset(between(tulopvm, LL, UL)) |> 
  dplyr::select(-UL, -LL, -ft17_otos_kayntipvm_kaikki) 

aux_ft17_hilmo <- datrrp$v2023_004_ft17_hilmo_2024_03 |>                                
  rename(havtun = HAVTUN,
         paltu = PALTU,
         pala = PALA,
         ea = EA) |>                                     
  right_join(aux_ft17_yrRange, by = "havtun") |>                      
  mutate(tulopvm = dmy(tulopvm) |>  decimal_date()) |>                                   
  subset(between(tulopvm, LL, UL)) |> 
  dplyr::select(-UL, -LL, -ft17_otos_kayntipvm_kaikki) 


##### TS22 #####
ts22_hilmo <- ts22_hilmo_raw$data_ts22_hilmo_96_22_toka |>                              
  rename(havtun = gumm85id) |>  
  mutate(havtun = havtun |> as.character()) |>                                                                                                                      
  right_join(ts22_yrRange, by = "havtun") |>                             
  mutate(tulo_pvm = dmy(tulo_pvm) |> decimal_date()) |> 
  subset(between(tulo_pvm, LL, UL)) |> 
  dplyr::select(-UL, -LL, -lopa_created) 

aux_ts22_hilmo <- ts22_hilmo_raw$data_ts22_hilmo_96_22_toka |>                              
  rename(havtun = gumm85id) |>  
  mutate(havtun = havtun |> as.character()) |>                                                                                                                      
  right_join(aux_ts22_yrRange, by = "havtun") |>                             
  mutate(tulo_pvm = dmy(tulo_pvm) |> decimal_date()) |> 
  subset(between(tulo_pvm, LL, UL)) |> 
  dplyr::select(-UL, -LL, -lopa_created) 


##### STEP 3 - Tidy hilmo data #####
##### H00 #####
h00_hilmo_tidy <- h00_hilmo %>% 
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

aux_h00_hilmo_tidy <- aux_h00_hilmo %>% 
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
h11_hilmo_tidy <- h11_hilmo %>%
  mutate(
    # Change blanks, dots & ATC code to NA
    across(
      !c("HILMO_LPVM", "HILMO_TOIPVM", "LAHTO_PVM", "HILMO_JOPVM"),
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

aux_h11_hilmo_tidy <- aux_h11_hilmo %>%
  mutate(
    # Change blanks, dots & ATC code to NA
    across(
      !c("HILMO_LPVM", "HILMO_TOIPVM", "LAHTO_PVM", "HILMO_JOPVM"),
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
ft17_hilmo_tidy <- ft17_hilmo %>% 
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

aux_ft17_hilmo_tidy <- aux_ft17_hilmo %>% 
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
ts22_hilmo_tidy <- ts22_hilmo %>% 
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

aux_ts22_hilmo_tidy <- aux_ts22_hilmo %>% 
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

##### STEP 4 - Create new variable #####
##### A function #####
hclass <- function(hilmo_tidy) {
  cols <- c(yhteystapa = NA, kiireellisyys = NA)
  hclass <- hilmo_tidy |>
    (\(.) add_column(., !!!cols[setdiff(names(cols),names(.))]))() |> 
    mutate(
      kaynti_typpi = case_when(
        (pala %in% c("1", "5", "6") | (pala == "" &
                                         yhteystapa == "R80")) ~ "vuode",
        ea == "15" |
          pala == "91" |
          (pala == "" &
             yhteystapa != "R80" & kiireellisyys %in% c("5", "6"))
        ~ "paivystys",
        TRUE ~ "muu_avo"
      ),
      del_ea = case_match(ea,
                          c("70", "74", "75", "00", "99") ~ 1,
                          .default = 0)
    ) |>
    filter(del_ea == 0 &
             !((substr(paltu, 1, 1) == '6' &
                  !(
                    paltu %in% c(
                      "60776", "60718", "60836", "60818", "60827",
                      "60711", "60828", "60792", "60821", "60839", 
                      "60834", "60838", "60820", "60788", "60804",
                      "60629", "66260", "60840", "60825"
                    )
                  )) |
                 substr(paltu, 1, 2) == "95" |
                 paltu == "01901")) |>  
    (\(.) bind_cols(., . |> sjmisc::to_dummy(kaynti_typpi, suffix = "label")))() |>
    rename(muu_avo = kaynti_typpi_muu_avo,
           paivystys = kaynti_typpi_paivystys,
           vuode = kaynti_typpi_vuode) |> 
    dplyr::select(havtun, vuode, paivystys, muu_avo) |>
    group_by(havtun) |>
    reframe(
      n_vuode = sum(vuode),
      n_paivystys = sum(paivystys),
      n_muu = sum(muu_avo)
    ) |> 
    rename(id = havtun)
}

##### H00 #####
h00_hclass <- hclass(h00_hilmo_tidy) |> 
  mutate(year = 2000)

aux_h00_hclass <- hclass(aux_h00_hilmo_tidy) |> 
  mutate(year = 2000) |> 
  set_names(c("id", "n_vuode_aux", "n_paivystys_aux", "n_muu_aux", "year"))


##### H11 #####
h11_hclass <- hclass(h11_hilmo_tidy) |> 
  mutate(year = 2011)

aux_h11_hclass <- hclass(aux_h11_hilmo_tidy) |> 
  mutate(year = 2011) |> 
  set_names(c("id", "n_vuode_aux", "n_paivystys_aux", "n_muu_aux", "year"))

##### FT17 #####
ft17_hclass <- hclass(ft17_hilmo_tidy) |> 
  mutate(year = 2017)

aux_ft17_hclass <- hclass(aux_ft17_hilmo_tidy) |> 
  mutate(year = 2017) |> 
  set_names(c("id", "n_vuode_aux", "n_paivystys_aux", "n_muu_aux", "year"))

##### TS22 #####
ts22_hclass <- hclass(ts22_hilmo_tidy) |> 
  mutate(year = 2022)

aux_ts22_hclass <- hclass(aux_ts22_hilmo_tidy) |> 
  mutate(year = 2022) |> 
  set_names(c("id", "n_vuode_aux", "n_paivystys_aux", "n_muu_aux", "year"))

##### STEP 5 - Combine the harmonized data #####
##### Read harmonized data #####
hdata <-
  read.csv2(
    "/home/zzhh/mnt/cdata/RRP_Project3/data/Analysis_data/analysis_data08042024.csv"
  ) %>%
  rename(area = my_area) %>%
  mutate(
    age = case_when(year == 2017 & ft17_survey == 1 & age >= 100 ~ 100,
                    TRUE ~ age),
    exam_date = exam_date %>% dmy() %>% decimal_date(),
    year = case_when(year == 2023 ~ 2022,
                     TRUE ~ year)
  ) 

##### H00 #####
hc.list <- list()
aux.list <- list()
hc.list$h00 <- hdata |> 
  filter(year == 2000) |> 
  left_join(h00_hclass, by = join_by(id, year))

aux.list$h00 <- hdata |> 
  filter(year == 2000) |> 
  left_join(aux_h00_hclass, by = join_by(id, year))

##### H11 #####
hc.list$h11 <- hdata |> 
  filter(year == 2011) |> 
  left_join(h11_hclass, by = join_by(id, year))

aux.list$h11 <- hdata |> 
  filter(year == 2011) |> 
  left_join(aux_h11_hclass, by = join_by(id, year))

##### FT17 #####
hc.list$ft17 <- hdata |> 
  filter(year == 2017) |> 
  left_join(ft17_hclass, by = join_by(id, year))

aux.list$ft17 <- hdata |> 
  filter(year == 2017) |> 
  left_join(aux_ft17_hclass, by = join_by(id, year))

##### TS22 #####
hc.list$ts22 <- hdata |> 
  filter(year == 2022) |> 
  left_join(ts22_hclass, by = join_by(id, year))

aux.list$ts22 <- hdata |> 
  filter(year == 2022) |> 
  left_join(aux_ts22_hclass, by = join_by(id, year))

##### Combine new outcome variables and auxiliary variables #####
new_out <- hc.list |> 
  bind_rows() |> 
  dplyr::select(1, year, sex, 48:50)

aux_out <- aux.list |> 
  bind_rows() |> 
  dplyr::select(1, year, sex, 48:50)

new_Outcomes <- new_out |>
  inner_join(aux_out, by = join_by(id, year, sex)) |>
  replace_na(
    list(
      n_vuode = 0,
      n_paivystys = 0,
      n_muu = 0,
      n_vuode_aux = 0,
      n_paivystys_aux = 0,
      n_muu_aux = 0
    )
  )

NewOCV <- new_Outcomes |>
  mutate_at(vars(n_vuode, n_paivystys, n_muu), ~ case_when(year == 2022 ~ NA_integer_, TRUE ~ .)) |> 
  mutate(year = case_when(year == 2022 ~ 2023,
                          TRUE ~ year))

save(NewOCV, file = "/home/zzhh/Desktop/RRP-Proj3/RRP_Project3/newOutcomes_20240524.Rdata")

##### END #####