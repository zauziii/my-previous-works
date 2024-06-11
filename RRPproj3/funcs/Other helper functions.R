# Helper function
plot_trends <- function(compl_data, groups, vars, ylab) {
  "
  A helper function for plotting the trends.

  compl_data: Completed data from mids object.
  groups: Grouping varibales. For example, year or year, sex_2
  vars: Variables need to be summerized.
  ylab: Y label.
  "
  
  # Define labels
  if (!"obese_2" %in% vars) {
    labels =
      c("n_muu" = "Other visits",
        "n_paivystys" = "Emergency visit",
        "n_vuode" = "Inpatient care")
  } else{
    labels = c("obese_2" = "Moderate", "obese_3" = "Severe")
  }
  # Define colors
  if (length(groups) == 2) {
    colors = palette_thl(3, name = "quali")[-1]
  } else {
    colors = palette_thl(2, name = "quali")[-1]
  }
  
  # Main body
  compl_data |>
    mutate(sex_2 = sex_2 |> factor(levels = 0:1, labels = c("Men", "Women"))) |>
    fastDummies::dummy_cols("obese") |>
    group_by(across(all_of(groups))) |>
    reframe(across(all_of(vars), ~ list(qwraps2::mean_ci(.x, na_rm = TRUE)))) |>
    pivot_longer(!groups) |>
    unnest_wider(value) |>
    ungroup() |>
    mutate_at(c("mean", "lcl", "ucl"), as.numeric) |>
    ggplot(aes(x = year, y = mean)) +
    ggh4x::geom_pointpath(
      size = 1,
      shape = 1,
      stroke = 1,
      linewidth = 1,
      if (length(groups) == 2) {
        aes(group = sex_2, color = sex_2)
      } else{
        aes(color = colors)
      }
    ) +
    scale_x_continuous_thl(breaks = unique(compl_data$year),
                           guide = guide_axis(check.overlap = TRUE)) +
    scale_y_continuous(limits = c(0, NA)) +
    ylab(ylab) + 
      {
        if (length(groups) == 1)
          guides(colour = "none")
      } +
    facet_wrap( ~ name,
                ncol = 1,
                scales = "free_y",
                labeller = labeller(name = labels)) +
    scale_color_manual(name = "", values = colors) +
    theme_thl() +
    theme(
      aspect.ratio = 1 / 3,
      axis.ticks.x = element_line(linewidth = thlPtsConvert(1.5) * 1.5, colour = "#606060"),
      legend.position = "top",
      axis.title.y = element_text(margin = unit(c(0, 1, 0, 0), "mm"))
    )
}

# Helper function - 2
plot_wgttrends <- function(unwgt_mean,
                           com_wgt_mean,
                           groups,
                           vars,
                           ylab) {
  # Define colors
  if (length(groups) == 2) {
    colors = palette_thl(3, name = "quali")[-1]
  } else {
    colors = palette_thl(2, name = "quali")[-1]
  }
  # Define labels
  if (!"obese_2" %in% vars) {
    labels =
      c("n_muu" = "Other visits",
        "n_paivystys" = "Emergency visit",
        "n_vuode" = "Inpatient care")
  } else{
    labels = c("obese_2" = "Moderate", "obese_3" = "Severe")
  }
  
  map2(
    unwgt_mean |> list.subset(vars),
    com_wgt_mean |> list.subset(vars),
    \(x, y) x |>
      filter(year > 2023) |>
      dplyr::select(-name) |>
      bind_rows(y |>
                  dplyr::select(groups, mean, lcl, ucl)) |>
      arrange(across(all_of(groups))) |>
      mutate(type = unique(x$name))
  ) |>
    bind_rows() |>
    ggplot(aes(x = year, y = mean)) +
    ggh4x::geom_pointpath(
      size = 1,
      shape = 1,
      stroke = 1,
      linewidth = 1,
      if (length(groups) == 2) {
        aes(group = sex_2, color = sex_2)
      } else{
        aes(color = colors)
      }
    ) +
    scale_x_continuous_thl(breaks = unique(sce1.compl$year),
                           guide = guide_axis(check.overlap = TRUE)) +
    scale_y_continuous(limits = c(0, NA)) +
    facet_wrap(
      ~ type,
      ncol = 1,
      scales = "free_y",
      labeller = labeller(type = labels)
    ) +
    scale_color_manual(name = "", values = colors) +
    ylab(ylab) +
    theme_thl() +
    theme(
      aspect.ratio = 1 / 3,
      axis.ticks.x = element_line(size = thlPtsConvert(1.5) * 1.5, colour = "#606060"),
      legend.position = "top",
      axis.title.y = element_text(margin = unit(c(0, 1, 0, 0), "mm"))
    )
}

# Helper function - 3
combined_res <- function(mids, compl_data, groups){
  
  # Pool
  unwgt.mean <- compl_data |> 
    mutate(sex_2 = sex_2 |> factor(levels = 0:1, labels = c("Men", "Women"))) |> 
    fastDummies::dummy_cols("obese") |> 
    group_by(across(all_of(groups))) |>
    reframe(across(
      c("n_muu", "n_paivystys", "n_vuode", "obese_2", "obese_3"),
      ~ list(qwraps2::mean_ci(.x, na_rm = TRUE))
    )) |> 
    pivot_longer(!groups) |> 
    unnest_wider(value) |> 
    ungroup() |> 
    mutate_at(c("mean", "lcl", "ucl"), as.numeric) |> 
    group_split(name) |> 
    set_names(c("n_muu", "n_paivystys", "n_vuode", "obese_2", "obese_3"))
  
  # MI and survey design
  sce_list <-
    imputationList(lapply(1:mids$m, function(n)
      mice::complete(mids, action = n) |> fastDummies::dummy_cols("obese")))
  sce_list$imputations <- map(
    sce_list$imputations,
    \(.) . |>
      filter(year < 2025) |>
      bind_cols(rrp.all |>
                  arrange(year, age, sex) |> 
                  dplyr::select(w_anal, stratum, cluster)) |>
      filter(!is.na(cluster))
  )
  # survey design
  des.sce <-
    svydesign(
      id =  ~ cluster,
      strata = ~ stratum,
      weights = ~ w_anal,
      data = sce_list,
      nest = TRUE
    )
  
  # get weighted mean 
  wgt.mean <- mclapply(c("n_muu", "n_paivystys", "n_vuode", "obese_2", "obese_3") |> set_names(),
                       \(.) if (length(groups) == 1) {
                         with(des.sce, svyby(~ get(.), ~ year, svymean, na.rm = T))
                       } else{
                         with(des.sce, svyby(~ get(.), ~ year + sex_2, svymean, na.rm = T))
                       })
  # combine results
  com.wgt.mean <- map(wgt.mean,
                      \(.) . |>
                        MIcombine() |>
                        summary() |>
                        rownames_to_column() |>
                        (\(.) if (length(groups) == 2) {
                          . |>
                            separate(col = "rowname",
                                     into = c("year", "sex_2"),
                                     sep = "\\.") |>
                            rename(mean = results, lcl = "(lower", ucl = "upper)") |>
                            mutate(sex_2 = sex_2 |> factor(levels = 0:1, 
                                                           labels = c("Men", "Women")))
                        } else{
                          . |>
                            rename(
                              year = rowname,
                              mean = results,
                              lcl = "(lower",
                              ucl = "upper)"
                            ) 
                        })() |> 
                        mutate(year = year |> as.numeric()))
  mylist <- list(unwgt.mean, com.wgt.mean) |> set_names(c("unwgt.mean", "com.wgt.mean"))
  return(mylist)
} 
