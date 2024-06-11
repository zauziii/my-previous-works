# File:     Helper function - Hilmo.R
# Author:   ZZ
# Date:     2023-04-11
# Summary:  A function for connecting ICPC-10 code and Hilmo data.
# Update:   2023-04-11 First commit.

func_hilmo <- function(x, y){
  
  pattern <- "[A-Z]DG[A-Z0-9]{0, }[O|E]|s_[[:lower:]]{0, }dg|DG_[[:upper:]]{1, }_[[:upper:]]{1, }_10|HILMO_[A-Z]{1}DG[A-Z0-9]{0, }[O|E]"
  
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
    select(matches(pattern)) %>%        
    unlist() %>%                                                                                    
    unique() %>%                                                                           
    t() %>%                                                                                
    data.frame() %>% 
    mutate(" " := case_when(
      if_any(everything(), function(z) z %in% x$`ICD-10 dg:t`) ~ 1,
      TRUE ~ 0
    )) %>% 
    subset(select = " ")
}