# File:     Helper function - Avohilmo.R
# Author:   ZZ
# Date:     2023-04-11
# Summary:  A function for connecting ICPC-10 code and avohilmo data.
# Update:   2023-04-11 First commit.

# A function: used in map() to connect code and avohilmo data
# This avohilmo function may be useful later
func_avohilmo <- function(x, y){
  bind_cols(y %>% 
              select(matches("icd")) %>%                          # For each element in ft17_avohilmo_split, find columns whose name contains "icd"
              unlist() %>%                                        # unlist and get the unique code - so to get the codes a participant have and not based on row(visits)
              unique() %>%                                        # but based on all the visits a participant have 
              t() %>% 
              data.frame() %>% 
              set_names(c(paste0("icd_", {1:length(.)}))),
            y %>%                                                 # Same here, but matches by "icpc"
              select(matches("icpc")) %>%
              unlist() %>% 
              unique() %>% 
              t() %>% 
              data.frame() %>% 
              set_names(c(paste0("icpc_", {1:length(.)})))) %>%   # Combine icd and icpc 
    mutate(" " := case_when(
      if_any(matches("icd"), function(z) z %in% x$`ICD-10 dg:t`) | if_any(matches("icpc"), function(z) z %in% na.omit(x$`ICPC-2 dg:t`)) ~ 1,
      TRUE ~ 0                                                    # if any of the icd variable in `ICD-10 dg:t` or any of the icpc variable in `ICPC-2 dg:t` -> 1, 
    )) %>%                                                        # else -> 0
    subset(select = " ")
}
