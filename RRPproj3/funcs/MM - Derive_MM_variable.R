# File:     Helper function - Derive_MM_variable.R
# Author:   ZZ
# Date:     2023-04-11
# Summary:  A function for connecting ICPC-10 code and avohilmo data.
# Update:   2023-04-11 First commit.

# A function: used to derive MM variable
list_func <- function(list, func) {
  map(code, func, list) %>% 
    bind_cols(.name_repair = ~ vec_as_names(..., repair = "unique", quiet = TRUE)) %>%      # Combine columns and silence messages
    setNames(., names(code)) %>%                                                            # Rename: each column = 1 chronic disease
    mutate(SUM = rowSums(across(where(is.numeric))),                                        # Sum 
           mltc = ifelse(SUM >= 2, 1, 0))}                                                  # Define MM variable 