# This file is used to create predictive data frame for the future
# Author: Zhi Zhou
# Date:
#     2024-03-28: update
#     2024-05-07: make it a function


PDFArea <- function(file = "/home/zzhh/mnt/groups4/HYRA_RRP/Proj3/work/Zhi/data/003_139f_2040_20240202-132934.csv",
                    future_years,
                    from_age = 18,
                    size,
                    different = TRUE) {
  '
  file: Projection forcasts from statistics Finland (With wellbeing services county).
  future_years: The future years (or predicted years). For example, c(2025:2030) or c(2025, 2030, 2035)
  from_age: The lower limit of age. By default from 18. No upper limit was defined as we tend to use people aged from 18.
  size: The sample size. When different = T, it means the size for counties are not the same, the size could be like 28000;
        if different = F, it means the size for counties are the same for each wsc, the size could be 1000.
  different: Whether to use same size for each county or not. 
  
  Notice: Final size not always equal to the input size because of rounding.
  
  Example:
  test <- PDFArea(file, future_years = c(2025:2030), from_age = 18, size = 1000, different = F)
  table(test$year)
  table(test$year, test$sex)
  table(test$year, test$wsc_area)
  '
  pop_forecasts <- read.table(
    file,
    sep = ";",
    skip = 2,
    header = T,
    check.names = F,
    fileEncoding = "ISO-8859-1"
  ) |>
    subset(Age != "Total") |>
    mutate_at("Age", # Age = "100 -" (character), mean 100 and above, we need numeric age
              # use 100 to replace "100 -". However, in our case, we don't need this "100 -".
              ~ replace(., Age == "100 -", 100) |>
                as.integer()) |>
    # Replace "Population 31 Dec (projection 2021)" by empty character
    rename_with( ~ str_replace(
      .,
      "\\s(Population 31 Dec)\\s(\\()(projection)\\s(2021)(\\))",
      ""
    )) |>
    # Replace "MA1" & "well-being services county" by empty character
    rename_with( ~ str_replace(., "( MA1|wellbeing services county )", "")) |>
    # Put all numbers in a string to the end
    rename_with( ~ paste(gsub("\\d+ ", "", .), sapply(., function(x)
      paste(str_extract_all(x, "\\d+")[[1]], collapse = " "))))
  
  # Get total amount
  total <- pop_forecasts |>
    dplyr::select(grep(".*Total|Age", names(pop_forecasts), value = F)) |>
    rename(Age = "Age ") |>
    pivot_longer(!Age) |>
    (\(.) bind_cols(., regmatches(
      x = .[["name"]],
      m = regexpr(pattern = "[[:digit:]]+", text = .[["name"]])
    ), .name_repair = "unique_quiet"))() |>
    rename(year = "...4") |>
    (\(.) split(., f = as.factor(.$year)))() |>
    map(\(.) . |> dplyr::select(-year) |> pivot_wider())
  
  # Get amount of men
  males <- pop_forecasts |>
    dplyr::select(grep(".*Males|Age", names(pop_forecasts), value = F)) |>
    rename(Age = "Age ") |>
    pivot_longer(!Age) |>
    (\(.) bind_cols(., regmatches(
      x = .[["name"]],
      m = regexpr(pattern = "[[:digit:]]+", text = .[["name"]])
    ), .name_repair = "unique_quiet"))() |>
    rename(year = "...4") |>
    (\(.) split(., f = as.factor(.$year)))() |>
    map(\(.) . |> dplyr::select(-year) |> pivot_wider())
  
  # Get amount of women
  females <- pop_forecasts |>
    dplyr::select(grep(".*Females|Age", names(pop_forecasts), value = F)) |>
    rename(Age = "Age ") |>
    pivot_longer(!Age) |>
    (\(.) bind_cols(., regmatches(
      x = .[["name"]],
      m = regexpr(pattern = "[[:digit:]]+", text = .[["name"]])
    ), .name_repair = "unique_quiet"))() |>
    rename(year = "...4") |>
    (\(.) split(., f = as.factor(.$year)))() |>
    map(\(.) . |> dplyr::select(-year) |> pivot_wider())
  
  # Whole population for the year
  vec_year <- c("Age") |> append(future_years) |> as.character()
  mainfi <- total |>
    # Select columns with contains MAINLAND in total
    lapply(\(.) . |> dplyr::select(contains("MAINLAND"))) |>
    bind_cols(total$`2021`$Age, .name_repair = "unique_quiet") |>
    dplyr::rename(Age = last_col()) |>
    dplyr::select(all_of(contains(vec_year))) |>
    filter(Age >= from_age)
  
  # Well-being service counties
  wsc <- c(
    "Lapland",
    "North Ostrobothnia",
    "Kainuu",
    "Central Ostrobothnia",
    "Central Finland",
    "North Savo",
    "North Karelia",
    "South Savo",
    "South Ostrobothnia",
    "Pirkanmaa",
    "Kanta-Häme",
    "Ostrobothnia",
    "Satakunta",
    "Southwest Finland",
    "West Uusimaa",
    "Central Uusimaa",
    "East Uusimaa",
    "Päijät-Häme",
    "Kymenlaakso",
    "South Karelia",
    "Vantaa and Kerava",
    "City of Helsinki"
  )
  
  # Total population by area, year, age
  total.wsc <- mclapply(set_names(wsc), \(x) {
    total |>
      map( ~ .x |>
             dplyr::select(starts_with(x))) |>
      bind_cols(total$`2021`$Age, .name_repair = "unique_quiet") |>
      rename(Age = last_col()) |>
      dplyr::select(all_of(contains(vec_year))) |>
      filter(Age >= from_age)
  })
  
  # No. of males by area, year, age
  male.wsc <- mclapply(set_names(wsc), \(x) {
    males |>
      map( ~ .x |>
             dplyr::select(starts_with(x))) |>
      bind_cols(total$`2021`$Age, .name_repair = "unique_quiet") |>
      rename(Age = last_col()) |>
      dplyr::select(all_of(contains(vec_year))) |>
      filter(Age >= from_age)
  })
  
  female.wsc <- mclapply(set_names(wsc), \(x) {
    females |>
      map( ~ .x |>
             dplyr::select(starts_with(x))) |>
      bind_cols(total$`2021`$Age, .name_repair = "unique_quiet") |>
      rename(Age = last_col()) |>
      dplyr::select(all_of(contains(vec_year))) |>
      filter(Age >= from_age)
  })
  
  if (different) {
    ##### The Code Below Creates Sample Equal to "size" ######
    ##### If You Want The Same Size of Each WSC, Please Use The Others #####
    # Compute proportion for men
    pwsc.m <- mclapply(
      male.wsc,
      \(x) map2(
        names(x)[-1],
        names(mainfi)[-1],
        \(xx, yy) (x |> dplyr::select(all_of(xx))) / (mainfi |> dplyr::select(all_of(yy)) |> sum())
      ) |>
        bind_cols() |>
        bind_cols(x[1]) |>
        rename_with(~ str_replace(., "( Males)", "")) |>
        dplyr::select(all_of(contains(vec_year)))
    )
    
    # Compute proportion for women
    pwsc.w <- mclapply(
      female.wsc,
      \(x) map2(names(x)[-1], names(mainfi)[-1], \(xx, yy) ((x |> dplyr::select(all_of(xx))) / (mainfi |> dplyr::select(all_of(yy)) |> sum())
      )) |>
        bind_cols() |>
        bind_cols(x[1]) |>
        rename_with( ~ str_replace(., "( Females)", "")) |>
        dplyr::select(all_of(contains(vec_year)))
    )
    
    # Get men and women separately with sample size mentioned above
    # Men
    pop.m <- mclapply(
      set_names(vec_year[-1]),
      \(xx)
      pwsc.m |>
        map( ~ .x |>
               dplyr::select(contains(
                 xx |> as.character()
               ))) |>
        bind_cols() |>
        bind_cols(pwsc.m$`East Uusimaa`$Age, .name_repair = "unique_quiet") |>
        rename(Age = last_col()) |>
        dplyr::select(all_of(contains(vec_year))) |>
        mutate(across(!Age, \(.) (. * size) |> round(0))) |>
        mutate(year = xx, sex = 1) |>
        pivot_longer(cols = !c("year", "sex", "Age")) |>
        mutate(name = name |> str_remove(" [[:digit:]]+")) |>
        apply(
          1,
          \(.)
          data.frame(
            year = rep(.[["year"]], times = .[["value"]]),
            age = rep(.[["Age"]], times = .[["value"]]),
            sex = rep(.[["sex"]], times = .[["value"]]),
            wsc_area = rep(.[["name"]], times = .[["value"]])
          )
        ) |>
        bind_rows()
    ) |>
      bind_rows()
    
    # Women
    pop.w <- mclapply(
      set_names(vec_year[-1]),
      \(xx)
      pwsc.w |>
        map( ~ .x |>
               dplyr::select(contains(
                 xx |> as.character()
               ))) |>
        bind_cols() |>
        bind_cols(pwsc.m$`East Uusimaa`$Age, .name_repair = "unique_quiet") |>
        rename(Age = last_col()) |>
        dplyr::select(all_of(contains(vec_year))) |>
        mutate(across(!Age, \(.) (. * size) |> round(0))) |>
        mutate(year = xx, sex = 2) |>
        pivot_longer(cols = !c("year", "sex", "Age")) |>
        mutate(name = name |> str_remove(" [[:digit:]]+")) |>
        apply(
          1,
          \(.)
          data.frame(
            year = rep(.[["year"]], times = .[["value"]]),
            age = rep(.[["Age"]], times = .[["value"]]),
            sex = rep(.[["sex"]], times = .[["value"]]),
            wsc_area = rep(.[["name"]], times = .[["value"]])
          )
        ) |>
        bind_rows()
    ) |>
      bind_rows()
  } else{
    ##### Code For Creating Same Size For Each WSC #####
    # Proportion Men
    prop.same.m <-
      map2(
        male.wsc,
        total.wsc,
        \(x, y) map2(
          names(x)[-1],
          names(y)[-1],
          \(xx, yy) (x |> dplyr::select(all_of(xx))) / (y |> dplyr::select(all_of(yy)) |> sum())
        ) |> bind_cols() |>
          bind_cols(x[1]) |>
          rename_with(~ str_replace(., "( Males)", "")) |>
          dplyr::select(all_of(contains(vec_year)))
      )
    # Proportion Women
    prop.same.w <-
      map2(
        female.wsc,
        total.wsc,
        \(x, y) map2(
          names(x)[-1],
          names(y)[-1],
          \(xx, yy) (x |> dplyr::select(all_of(xx))) / (y |> dplyr::select(all_of(yy)) |> sum())
        ) |> bind_cols() |>
          bind_cols(x[1]) |>
          rename_with( ~ str_replace(., "( Females)", "")) |>
          dplyr::select(all_of(contains(vec_year)))
      )
    
    # To get the sample
    # Men
    pop.m <- mclapply(
      set_names(vec_year[-1]),
      \(xx)
      prop.same.m |>
        map( ~ .x |>
               dplyr::select(contains(
                 xx |> as.character()
               ))) |>
        bind_cols() |>
        bind_cols(prop.same.m$`East Uusimaa`$Age, .name_repair = "unique_quiet") |>
        rename(Age = last_col()) |>
        dplyr::select(all_of(contains(vec_year))) |>
        mutate(across(!Age, \(.) (. * size) |> round(0))) |>
        mutate(year = xx, sex = 1) |>
        pivot_longer(cols = !c("year", "sex", "Age")) |>
        mutate(name = name |> str_remove(" [[:digit:]]+")) |>
        apply(
          1,
          \(.)
          data.frame(
            year = rep(.[["year"]], times = .[["value"]]),
            age = rep(.[["Age"]], times = .[["value"]]),
            sex = rep(.[["sex"]], times = .[["value"]]),
            wsc_area = rep(.[["name"]], times = .[["value"]])
          )
        ) |>
        bind_rows()
    ) |>
      bind_rows()
    # Women
    pop.w <- mclapply(
      set_names(vec_year[-1]),
      \(xx)
      prop.same.w |>
        map( ~ .x |>
               dplyr::select(contains(
                 xx |> as.character()
               ))) |>
        bind_cols() |>
        bind_cols(prop.same.w$`East Uusimaa`$Age, .name_repair = "unique_quiet") |>
        rename(Age = last_col()) |>
        dplyr::select(all_of(contains(vec_year))) |>
        mutate(across(!Age, \(.) (. * size) |> round(0))) |>
        mutate(year = xx, sex = 2) |>
        pivot_longer(cols = !c("year", "sex", "Age")) |>
        mutate(name = name |> str_remove(" [[:digit:]]+")) |>
        apply(
          1,
          \(.)
          data.frame(
            year = rep(.[["year"]], times = .[["value"]]),
            age = rep(.[["Age"]], times = .[["value"]]),
            sex = rep(.[["sex"]], times = .[["value"]]),
            wsc_area = rep(.[["name"]], times = .[["value"]])
          )
        ) |>
        bind_rows()
    ) |>
      bind_rows()
  }
  
  # Get the predictive data frame for the future 2025-2034
  pop.future <- bind_rows(pop.m, pop.w) |>
    mutate(
      year = year |> as.numeric(),
      age = age |> as.numeric(),
      sex = sex |> as.factor(),
      wsc_area = wsc_area |> as.factor()
    ) |>
    arrange(year, age, sex)
}
