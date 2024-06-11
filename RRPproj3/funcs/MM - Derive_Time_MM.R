# File:     MM - Derive_Time_MM.R
# Author:   ZZ
# Date:     2023-05-22
# Summary:  A function to derive time-to-event mm variable.
# Update:   2023-05-24 Birthday.
#           2024-02-27 Add pattern (TS22).
#           2024-02-28 foreach to mclapply.
#           2024-02-29 set default cores. 

derive_time_mm <- function(split_by_id, cores = 10) {
  # Pattern used to select related columns
  pattern <-
    "[A-Z]DG[A-Z0-9]{0, }[O|E]|s_[[:lower:]]{0, }dg|DG_[[:upper:]]{1, }_[[:upper:]]{1, }_10|HILMO_[A-Z]{1}DG[A-Z0-9]{0, }[O|E]|icd10_[0-9]{1,2}[a-z]{1}"
  
  # A function to split the data by date for each participant,
  # and collect diagnosis info for each visit
  splitandcollect <- function(y) {
    y %>%
      dplyr::select(matches(pattern), TULO_PVM) %>%
      arrange(TULO_PVM) %>%
      group_split(TULO_PVM) %>%
      map(., ~ .x %>% select_if( ~ !any(is.na(.)))) %>%
      set_names(sort(unique(y$TULO_PVM))) -> split_bydate
    mclapply(split_bydate,
             \(x) {
               map(code, func_hilmo_v2, x) %>%
                 map( ~ .x %>% data.frame() %>% subset(select = c("evnt"))) %>%
                 bind_cols(x$TULO_PVM,
                           .name_repair = ~ vctrs::vec_as_names(..., repair = "unique", quiet = TRUE)) %>%
                 setNames(., c(names(code), "TULO_PVM"))
             },
             mc.cores = cores) |>
      bind_rows()
  }
  
  # Part 1: Turn each element in split to form that
  #  with only diagnosis code and TULO_PVM
  split_and_collect <-
    mclapply(split_by_id, \(.) splitandcollect(.), mc.cores = cores)
  
  # Part 2: Find the first "1" in each diagnosis group and let the others in the same column = 0
  Findfirstone <- mclapply(split_and_collect,
                           \(.) . |>
                             mutate(across(
                               !TULO_PVM,
                               ~ if_else(row_number() > match(1, .), NA_integer_, .)
                             )) |>
                             mutate(across(everything(), \(x) replace(x, is.na(
                               x
                             ), 0))),
                           mc.cores = cores)
  
  # Part 3: Calculate row sum, row sum means how many codes are put by doctor every visit.
  # also calculate the cumulative sum for row and column, so to know when n(diagnosis code) >= 2
  mmtime <- mclapply(
    Findfirstone,
    \(.) . |>
      mutate(
        rowSum = dplyr::select(., 1:35) %>% rowSums(),
        cumSum = cumsum(rowSum),
        mm = case_when(sum(rowSum) >= 2 ~ 1,
                       TRUE ~ 0)
      ) %>%
      # Find the first cumsum >= 2, let the others become 0
      mutate(across(
        cumSum,
        \(x) if_else(row_number() > match(TRUE, x >= 2), 0, x)
      )) %>%
      # Subset the cumSum >= 2
      filter(row_number() == match(TRUE, cumSum >= 2)) %>%
      # reframe it, as we only need mm and time
      reframe(
        time = ifelse(nrow(.) == 0, 0, TULO_PVM),
        mm = ifelse(nrow(.) == 0, 0, 1)
      ),
    mc.cores = cores
  ) |>
    bind_rows() |>
    mutate(havtun = names(Findfirstone))
  
  return(mmtime)
}


