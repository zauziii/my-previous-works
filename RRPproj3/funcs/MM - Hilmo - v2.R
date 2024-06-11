# File:     Helper function - Hilmo.R
# Author:   ZZ
# Date:     2023-05-22
# Summary:  A function for connecting ICPC-10 code and Hilmo data. A update version as to get the time to event mm variable.
# Update:   2023-05-22 Birthday.
#           2024-02-27 Added pattern (TS22).

func_hilmo_v2 <- function(x, y){
  
  pattern <- "[A-Z]DG[A-Z0-9]{0, }[O|E]|s_[[:lower:]]{0, }dg|DG_[[:upper:]]{1, }_[[:upper:]]{1, }_10|HILMO_[A-Z]{1}DG[A-Z0-9]{0, }[O|E]|icd10_[0-9]{1,2}[a-z]{1}"
  
  # ts22_hilmo: icd10_[0-9]{1,2}[a-z]{1}
  # icd10_: exact match
  # [0-9]: numbers (0-9)
  # {1,2}: number of matches, 1 or 2. As we have 1-20
  # [a-z]: lower case letter (o or e)
  # {1}: 1 match
  # selected varaibles: icd10_1o, icd10_1e, icd10_2o, icd10_2e, icd10_3o, icd10_3e, icd10_4o, icd10_4e, icd10_5o
  #   icd10_5e, icd10_6o, icd10_6e, icd10_7o, icd10_7e, icd10_8o, icd10_8e, icd10_9o, icd10_9e, icd10_10o, icd10_10e,
  #   icd10_11o, icd10_11e, icd10_12o, icd10_12e, icd10_13o, icd10_13e, icd10_14e, icd10_15o, icd10_15e, icd10_16o,
  #   icd10_16e, icd10_17o, icd10_17e, icd10_18o, icd10_18e, icd10_19o, icd10_19e, icd10_20o, icd10_20e

  # ft17_hilmo: [A-Z]DG[A-Z0-9]{0, }[O|E]
  # [A-Z]: a upper class letter (A to Z)
  # DG: exact match
  # [A-Z0-9]: upper class letter(A to Z) or/and numbers (0 to 9)
  # {0, }: specified the number of matches, 0 or more. 
  # [O|E]: end with O or E
  # Selected variables: PDGO, PDGE, SDG10, SDG1E, SDG2O, SDG2E
  
  # fr_hilmo: s_[[:lower:]]{0, }dg
  # [[:lower:]]: lower case letter (a to z)
  # {0, }: number of matches, 0 or more. So it allows pattern like s_dgxxx or s_paadg
  # s_, dg: exact match
  # Selected variables: s_dg2, s_dg3, s_dg4, s_paadg, s_paadge, s_dg2e, s_dg3e, s_dg4e, s_dg5, s_dg5e,
  #   s_dg6, s_dg6e, s_dg7, s_dg7e, s_dg8,  s_dg8e, s_dg9, s_dg9e, s_dg10, s_dg10e, s_dg11, 
  #   s_dg11e, s_dg12, s_dg12e, s_dg13, s_dg13e, s_dg14, s_dg14e, s_dg15, s_dg15e, s_dg16, s_dg16e, 
  #   s_dg17, s_dg17e, s_dg18, s_dg18e, s_dg19, s_dg19e, s_dg20, s_dg20e, s_dg21, s_dg21e, 
  #   s_dg22, s_dg22e, s_dg23, s_dg23e, s_dg24, s_dg24e, s_dg25, s_dg25e, s_dg26, s_dg26e" 
  
  # h00 (& h11): DG_[[:upper:]]{1, }_[[:upper:]]{1, }_10
  # DG_: exact match
  # [[:upper:]]: any upper class letter
  # {1, }: number of upper class letter, 1 or more. So it give _E_, _SIVU_, _PAA_ and such
  # _, _10: exact match
  # Selected variables: DG_PAA_SYY_10, DG_SIVU_SYY_10_1, DG_SIVU_SYY_10_2, DG_PAA_OIRE_10,
  #  DG_SIVU_OIRE_10_1 DG_SIVU_OIRE_10_2 DG_E_KOODI_10
  
  # h11: HILMO_[A-Z]{1}DG[A-Z0-9]{0, }[O|E] 
  # HILMO_: exact match
  # [A-Z]{1}: 1 upper class letter 
  # DG: exact match
  # [A-Z0-9]{0, }: 0 or more upper class letters or numbers 
  # [O|E]: end with 0 or E
  # variables: HILMO_PDGO, HILMO_PDGE, HILMO_SDG1O, HILMO_SDG1E, HILMO_SDG2O, HILMO_SDG2E
  
  y %>%
    arrange(TULO_PVM) %>% 
    mutate(evnt := case_when(
      if_any(everything(), function(z) z %in% x$`ICD-10 dg:t`) ~ 1,
      TRUE ~ 0
    )
    )
}
