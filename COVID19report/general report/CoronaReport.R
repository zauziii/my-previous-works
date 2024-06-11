data_dir <- "/home/zzhh/mnt/ZZHH/_Desktop/Tasks/Koronaraportointi/data/data.xlsx"

PTH_raw_dir <- "/home/zzhh/mnt/ZZHH/_Desktop/Tasks/Koronaraportointi/PTH_updated.xls"
PTH_new_dir <- "/home/zzhh/mnt/ZZHH/_Desktop/Tasks/Koronaraportointi/covid_pth_tilannekuva_2022-09-08.xls"
PTH_updated_dir <- "/home/zzhh/mnt/ZZHH/_Desktop/Tasks/Koronaraportointi/PTH_updated1.xls"

ESH_raw_dir <- "/home/zzhh/mnt/ZZHH/_Desktop/Tasks/Koronaraportointi/ESH_updated.xls"
ESH_new_dir <- "/home/zzhh/mnt/ZZHH/_Desktop/Tasks/Koronaraportointi/covid_esh_tilannekuva_2022-09-08.xls"
ESH_updated_dir <- "/home/zzhh/mnt/ZZHH/_Desktop/Tasks/Koronaraportointi/ESH_updated1.xls"

SOS_raw_dir <- "/home/zzhh/mnt/ZZHH/_Desktop/Tasks/Koronaraportointi/SOS_updated.xls"
SOS_new_dir <- "/home/zzhh/mnt/ZZHH/_Desktop/Tasks/Koronaraportointi/covid_sos_tilannekuva_2022-09-08.xls"
SOS_updated_dir <- "/home/zzhh/mnt/ZZHH/_Desktop/Tasks/Koronaraportointi/SOS_updated1.xls"

ppt_sturcture_dir <- "/home/zzhh/mnt/ZZHH/_Desktop/Tasks/Koronaraportointi/PPT_report_structure.xlsx"
ppt_template_dir <- "/home/zzhh/mnt/ZZHH/_Desktop/Tasks/Koronaraportointi/ppt/ppt-template-v2.pptx"

date <- "14.12.2020-08.09.2022"
ppt_save_dir <- "/home/zzhh/mnt/ZZHH/_Desktop/Tasks/Koronaraportointi/ppt/THL_STM kuntaseuranta_Kuvioita 100%_20220908_vk35.pptx"



# Load packages
library("xlsx")
library("readxl")
library("magrittr")
library("tidyverse")
library("lubridate")
library("ggplot2")
library("officer")

# Load data
# Raw data
# Get sheet names
sheet_names <- excel_sheets(data_dir)
# Read all sheets to list
list_all <- lapply(sheet_names, function(x) {         
  as.data.frame(read_excel(data_dir, sheet = x)) 
})
# Rename list elements
names(list_all) <- sheet_names       
# New data
# 
PTH_new <- read.xlsx2(PTH_raw_dir,
                      sheetIndex = 1, 
                      header = T) 
PTH_new[PTH_new == "ERROR"] <- ""
PTH_add <- read.xlsx2(PTH_new_dir,
                      sheetIndex = 1, 
                      header = T) 

PTH_add[, PTH_VarName_new[!PTH_VarName_new %in% names(PTH_add)]] <- NA
PTH_new[, PTH_VarName_raw[!PTH_VarName_raw %in% names(PTH_new)]] <- NA

PTH_ori <- list_all$PTH_ori
PTH_ori[, PTH_VarName_raw[!PTH_VarName_raw %in% names(PTH_ori)]] <- NA


PTH_add <- PTH_add %>% 
  subset(select = PTH_VarName_new)
names(PTH_add) <- names(PTH_new)  <- c("Viikko","Sairaanhoitopiiri","Järjestäjäorganisaatio","lopa_subject","lopa_created","lopa_updated","lopa_draft","PTH.avosh.ps","pth_service_availability_comment_1","PTH.vos.ps","pth_service_availability_comment_2","PTH.neuv.ps","pth_service_availability_comment_3","PTH.kouluth.ps","pth_service_availability_comment_4","PTH.opiskth.ps","pth_service_availability_comment_5","PTH.kunt.ps","pth_service_availability_comment_6","PTH.suunth.ps","pth_service_availability_comment_7","PTH.mieli.ps","pth_service_availability_comment_8","PTH.päihde.ps","pth_service_availability_comment_9","PTH.roko.ps","pth_service_action_comment_1","PTH.test.ps","pth_service_action_comment_2","PTH.tjälj.ps","pth_service_action_comment_3","PTH.tneuvo.ps","pth_service_passenger_trafic_comment_1","PTH.todtark.ps","pth_service_passenger_trafic_comment_2","PTH.testmtulo.ps","pth_service_passenger_trafic_comment_3","PTH.2.test.ps","pth_service_passenger_trafic_comment_4","PTH.lapsimttperus.ps","pth_service_mental_health_child_comment_1","PTH.lapsimtttesh","pth_service_mental_health_child_comment_2","PTH.nuorimttperus.ps","pth_service_mental_health_young_comment_1","PTH.nuorimttesh.ps","pth_service_mental_health_young_comment_2","PTH.aikuismttperus","pth_service_mental_health_adult_comment_1","PTH.aikuismttesh.ps","pth_service_mental_health_adult_comment_2","PTH.kiirsupist.ps","PTH.toimuud.ps","PTH.etäplis.ps","PTH.ostop.ps","PTH.muuten.ps","PTH.tmpriittäv.ps","PTH.avosh.hr","pth_personnel_sufficiency_comment_1","PTH.vos.hr","pth_personnel_sufficiency_comment_2","PTH.neuv.hr","pth_personnel_sufficiency_comment_3","PTH.kouluth.hr","pth_personnel_sufficiency_comment_4","PTH.opiskth.hr","pth_personnel_sufficiency_comment_5","PTH.kunt.hr","pth_personnel_sufficiency_comment_6","PTH.suunth.hr","pth_personnel_sufficiency_comment_7","PTH.mieli.hr","pth_personnel_sufficiency_comment_8","PTH.päihde.hr","pth_personnel_sufficiency_comment_9","PTH.roko.hr","pth_personnel_sufficiency_action_comment_1","PTH.test.hr","pth_personnel_sufficiency_action_comment_2","PTH.tjälj.hr","pth_personnel_sufficiency_action_comment_3","PTH.tneuv.hr","pth_personnel_sufficiency_passenger_trafic_comment_1","PTH.todtark.hr","pth_personnel_sufficiency_passenger_trafic_comment_2","PTH.testmsaap.hr","pth_personnel_sufficiency_passenger_trafic_comment_3","PTH.2.test.hr","pth_personnel_sufficiency_passenger_trafic_comment_4","PTH.lainsääd.hr","PTH.henksiir.hr","PTH.henkrekry.hr","PTH.neuvot.hr","PTH.ostop.hr","PTH.muuten.hr","PTH.tmpriittäv.hr", "pth_service_availability_option_10", "pth_service_availability_comment_10", "pth_service_availability_option_11", "pth_service_availability_comment_11" , "pth_personnel_sufficiency_option_10", "pth_personnel_sufficiency_comment_10", "pth_personnel_sufficiency_option_11", "pth_personnel_sufficiency_comment_11", "pth_personnel_sufficiency_situation_manage_6", "pth_protection_work_option_1", "pth_protection_work_option_2", "pth_protection_work_option_3", "pth_protection_work_option_4", "pth_protection_work_option_5", "pth_protection_work_option_6", "pth_protection_work_option_7", "pth_protection_work_option_8", "pth_protection_work_option_9", "pth_protection_work_option_10", "pth_protection_work_option_11")
names(list_all$PTH_ori) <- c("Viikko","Sairaanhoitopiiri","Järjestäjäorganisaatio","lopa_subject","lopa_created","lopa_updated","lopa_draft","PTH.avosh.ps","pth_service_availability_comment_1","PTH.vos.ps","pth_service_availability_comment_2","PTH.neuv.ps","pth_service_availability_comment_3","PTH.kouluth.ps","pth_service_availability_comment_4","PTH.opiskth.ps","pth_service_availability_comment_5","PTH.kunt.ps","pth_service_availability_comment_6","PTH.suunth.ps","pth_service_availability_comment_7","PTH.mieli.ps","pth_service_availability_comment_8","PTH.päihde.ps","pth_service_availability_comment_9","PTH.roko.ps","pth_service_action_comment_1","PTH.test.ps","pth_service_action_comment_2","PTH.tjälj.ps","pth_service_action_comment_3","PTH.tneuvo.ps","pth_service_passenger_trafic_comment_1","PTH.todtark.ps","pth_service_passenger_trafic_comment_2","PTH.testmtulo.ps","pth_service_passenger_trafic_comment_3","PTH.2.test.ps","pth_service_passenger_trafic_comment_4","PTH.lapsimttperus.ps","pth_service_mental_health_child_comment_1","PTH.lapsimtttesh","pth_service_mental_health_child_comment_2","PTH.nuorimttperus.ps","pth_service_mental_health_young_comment_1","PTH.nuorimttesh.ps","pth_service_mental_health_young_comment_2","PTH.aikuismttperus","pth_service_mental_health_adult_comment_1","PTH.aikuismttesh.ps","pth_service_mental_health_adult_comment_2","PTH.kiirsupist.ps","PTH.toimuud.ps","PTH.etäplis.ps","PTH.ostop.ps","PTH.muuten.ps","PTH.tmpriittäv.ps","PTH.avosh.hr","pth_personnel_sufficiency_comment_1","PTH.vos.hr","pth_personnel_sufficiency_comment_2","PTH.neuv.hr","pth_personnel_sufficiency_comment_3","PTH.kouluth.hr","pth_personnel_sufficiency_comment_4","PTH.opiskth.hr","pth_personnel_sufficiency_comment_5","PTH.kunt.hr","pth_personnel_sufficiency_comment_6","PTH.suunth.hr","pth_personnel_sufficiency_comment_7","PTH.mieli.hr","pth_personnel_sufficiency_comment_8","PTH.päihde.hr","pth_personnel_sufficiency_comment_9","PTH.roko.hr","pth_personnel_sufficiency_action_comment_1","PTH.test.hr","pth_personnel_sufficiency_action_comment_2","PTH.tjälj.hr","pth_personnel_sufficiency_action_comment_3","PTH.tneuv.hr","pth_personnel_sufficiency_passenger_trafic_comment_1","PTH.todtark.hr","pth_personnel_sufficiency_passenger_trafic_comment_2","PTH.testmsaap.hr","pth_personnel_sufficiency_passenger_trafic_comment_3","PTH.2.test.hr","pth_personnel_sufficiency_passenger_trafic_comment_4","PTH.lainsääd.hr","PTH.henksiir.hr","PTH.henkrekry.hr","PTH.neuvot.hr","PTH.ostop.hr","PTH.muuten.hr","PTH.tmpriittäv.hr")
PTH_New <- rbind(PTH_new, PTH_add)

# ESH
ESH_new <- read.xlsx2(ESH_raw_dir,
                      sheetIndex = 1, 
                      header = T)
ESH_new[ESH_new == "ERROR"] <- ""
ESH_add <- read.xlsx2(ESH_new_dir,
                      sheetIndex = 1, 
                      header = T) 

ESH_add[, ESH_VarName_new[!ESH_VarName_new %in% names(ESH_add)]] <- NA
ESH_new[, ESH_VarName_new[!ESH_VarName_raw %in% names(ESH_new)]] <- NA

ESH_ori <- list_all$ESH_ori
ESH_ori[, ESH_VarName_raw[!ESH_VarName_raw %in% names(ESH_ori)]] <- NA

ESH_add <- ESH_add %>% 
  subset(select = ESH_VarName_new)

names(ESH_add) <- names(ESH_new) <- c("Viikko","Sairaanhoitopiiri","Järjestäjäorganisaatio","lopa_subject","lopa_created","lopa_updated","lopa_draft","ESH.somaattinen.ps","esh_service_availability_comment_11","ESH.psykiatria.ps","esh_service_availability_comment_12","ESH.som.pkl.hr","esh_personnel_sufficiency_comment_1","ESH.som.vos.hr","esh_personnel_sufficiency_comment_2","ESH.psyk.pkl.hr","esh_personnel_sufficiency_comment_3","ESH.psyk.vos.hr","esh_personnel_sufficiency_comment_4","ESH.teho.vos.hr","esh_personnel_sufficiency_comment_5","ESH.lainsääd.hr","ESH.henksiir.hr","ESH.henkrekryt.hr","ESH.neuv.hr","ESH.ostop.hr","ESH.muuten.hr","ESH.tmp.riittäv.hr", "esh_service_availability_option_1","esh_service_availability_comment_1", "esh_service_availability_option_2", "esh_service_availability_comment_2", "esh_service_availability_option_3", "esh_service_availability_comment_3", "esh_service_availability_option_4", "esh_service_availability_comment_4", "esh_service_availability_option_5", "esh_service_availability_comment_5", "esh_personnel_sufficiency_situation_manage_6", "esh_protection_work_option_1", "esh_protection_work_option_2","esh_protection_work_option_3","esh_protection_work_option_4","esh_protection_work_option_5","esh_protection_work_option_12")
names(list_all$ESH_ori) <- c("Viikko","Sairaanhoitopiiri","Järjestäjäorganisaatio","lopa_subject","lopa_created","lopa_updated","lopa_draft","ESH.somaattinen.ps","esh_service_availability_comment_11","ESH.psykiatria.ps","esh_service_availability_comment_12","ESH.som.pkl.hr","esh_personnel_sufficiency_comment_1","ESH.som.vos.hr","esh_personnel_sufficiency_comment_2","ESH.psyk.pkl.hr","esh_personnel_sufficiency_comment_3","ESH.psyk.vos.hr","esh_personnel_sufficiency_comment_4","ESH.teho.vos.hr","esh_personnel_sufficiency_comment_5","ESH.lainsääd.hr","ESH.henksiir.hr","ESH.henkrekryt.hr","ESH.neuv.hr","ESH.ostop.hr","ESH.muuten.hr","ESH.tmp.riittäv.hr")
ESH_New <- rbind(ESH_new, ESH_add)


# SOS
SOS_new <- read.xlsx2(SOS_raw_dir,
                      sheetIndex = 1, 
                      header = T)
SOS_new[SOS_new == "ERROR"] <- ""
SOS_add <- read.xlsx2(SOS_new_dir,
                      sheetIndex = 1, 
                      header = T) 
SOS_add[, SOS_VarName_new[!SOS_VarName_new %in% names(SOS_add)]] <- NA
SOS_new[, SOS_VarName_raw[!SOS_VarName_raw %in% names(SOS_new)]] <- NA

SOS_ori <- list_all$SOS_ori
SOS_ori[, SOS_VarName_raw[!SOS_VarName_raw %in% names(SOS_ori)]] <- NA

SOS_add <- SOS_add %>% 
  subset(select = SOS_VarName_new)
names(SOS_add) <- names(SOS_new) <- c("Viikko","Sairaanhoitopiiri","Järjestäjäorganisaatio","lopa_subject","lopa_created","lopa_updated","lopa_draft","SOS.iäkkäät.kotiin.ps","sos_service_availability_comment_1","SOS.iäkkäät.ympäriv.ps","sos_service_availability_comment_2","SOS.lapsip.ps","sos_service_availability_comment_3","SOS.lastens.ps","sos_service_availability_comment_4","SOS.vamm.ps","sos_service_availability_comment_5","SOS.työikä.ps","sos_service_availability_comment_6","SOS.perheoik.ps","sos_service_availability_comment_7","SOS.ttt.ps","sos_service_availability_comment_8","SOS.tilapasum.ps","sos_service_availability_comment_9","SOS.sospäiv.ps","sos_service_availability_comment_10","SOS.päihdet.ps","sos_service_availability_comment_11","SOS.omaishoit.ps","sos_service_availability_comment_12","SOS.kiiresup.ps","SOS.toimuud.ps","SOS.etäp.ps","SOS.ostop.ps","SOS.muuten.ps","SOS.riittäv.ps","SOS.iäkkäät.kotiin.hr","sos_personnel_sufficiency_comment_1","SOS.iäkkäät.ympäriv.hr","sos_personnel_sufficiency_comment_2","SOS.lapsip.hr","sos_personnel_sufficiency_comment_3","SOS.lastens.hr","sos_personnel_sufficiency_comment_4","SOS.vamm.hr","sos_personnel_sufficiency_comment_5","SOS.työikä.hr","sos_personnel_sufficiency_comment_6","SOS.perheoik.hr","sos_personnel_sufficiency_comment_7","SOS.ttt.hr","sos_personnel_sufficiency_comment_8","SOS.tilapasum.hr","sos_personnel_sufficiency_comment_9","SOS.sospäiv.hr","sos_personnel_sufficiency_comment_10","SOS.päihdet.hr","sos_personnel_sufficiency_comment_11","SOS.omaishoit.hr","sos_personnel_sufficiency_comment_12","SOS.lainsääd.hr","SOS.henksiir.hr","SOS.henkrekryt.hr","SOS.neuv.hr","SOS.ostop.hr","SOS.muuten.hr","SOS.tmp.riittäv.hr", "sos_personnel_sufficiency_situation_manage_6","sos_protection_work_option_1","sos_protection_work_option_2","sos_protection_work_option_3","sos_protection_work_option_4","sos_protection_work_option_5","sos_protection_work_option_6","sos_protection_work_option_7","sos_protection_work_option_8","sos_protection_work_option_9","sos_protection_work_option_10","sos_protection_work_option_11","sos_protection_work_option_12")
names(list_all$SOS_ori) <- c("Viikko","Sairaanhoitopiiri","Järjestäjäorganisaatio","lopa_subject","lopa_created","lopa_updated","lopa_draft","SOS.iäkkäät.kotiin.ps","sos_service_availability_comment_1","SOS.iäkkäät.ympäriv.ps","sos_service_availability_comment_2","SOS.lapsip.ps","sos_service_availability_comment_3","SOS.lastens.ps","sos_service_availability_comment_4","SOS.vamm.ps","sos_service_availability_comment_5","SOS.työikä.ps","sos_service_availability_comment_6","SOS.perheoik.ps","sos_service_availability_comment_7","SOS.ttt.ps","sos_service_availability_comment_8","SOS.tilapasum.ps","sos_service_availability_comment_9","SOS.sospäiv.ps","sos_service_availability_comment_10","SOS.päihdet.ps","sos_service_availability_comment_11","SOS.omaishoit.ps","sos_service_availability_comment_12","SOS.kiiresup.ps","SOS.toimuud.ps","SOS.etäp.ps","SOS.ostop.ps","SOS.muuten.ps","SOS.riittäv.ps","SOS.iäkkäät.kotiin.hr","sos_personnel_sufficiency_comment_1","SOS.iäkkäät.ympäriv.hr","sos_personnel_sufficiency_comment_2","SOS.lapsip.hr","sos_personnel_sufficiency_comment_3","SOS.lastens.hr","sos_personnel_sufficiency_comment_4","SOS.vamm.hr","sos_personnel_sufficiency_comment_5","SOS.työikä.hr","sos_personnel_sufficiency_comment_6","SOS.perheoik.hr","sos_personnel_sufficiency_comment_7","SOS.ttt.hr","sos_personnel_sufficiency_comment_8","SOS.tilapasum.hr","sos_personnel_sufficiency_comment_9","SOS.sospäiv.hr","sos_personnel_sufficiency_comment_10","SOS.päihdet.hr","sos_personnel_sufficiency_comment_11","SOS.omaishoit.hr","sos_personnel_sufficiency_comment_12","SOS.lainsääd.hr","SOS.henksiir.hr","SOS.henkrekryt.hr","SOS.neuv.hr","SOS.ostop.hr","SOS.muuten.hr","SOS.tmp.riittäv.hr")
SOS_New <- rbind(SOS_new, SOS_add)

# Define a plot function
CoronaReport <- function(Img1Var, NameImg1Var, Img2Var = NULL, NameImg2Var = NULL, source, NameList){
  if(NameList == "combined_new"){
    # Format data frame for left-hand side figure
    left <- source[[NameList]] %>%
      select(grep(paste0(".*", Img1Var, "|Viikko"), 
                  names(source[[NameList]]), value = F))  %>% 
      filter(!get(Img1Var) %in% c("", 9, NA, "ERROR")) %>% 
      dplyr::count(Viikko, get(Img1Var)) %>% 
      group_by(Viikko) %>% 
      mutate(group = NameImg1Var)
    # Format data frame for right-hand side figure
    right <- source[[NameList]] %>%
      select(grep(paste0(".*", Img2Var, "|Viikko"), 
                  names(source[[NameList]]), value = F)) %>% 
      filter(!get(Img2Var) %in% c("", 9, NA)) %>% 
      dplyr::count(Viikko, get(Img2Var)) %>% 
      group_by(Viikko) %>% 
      mutate(group = NameImg2Var)
    
    # Combine and relevel
    colnames(right) <- colnames(left)
    slide <- rbind(left, right)
    slide$group <- factor(slide$group, levels = c(NameImg1Var, NameImg2Var))
    colnames(slide) <- c("Viikko", "levels", "n", "group")
    slide$levels <- factor(slide$levels, levels = c(1:4))
    slide$Viikko <- as.factor(slide$Viikko)
    
    # Plot
    pSlide <- ggplot(slide, aes(fill = levels, y = n, x = Viikko)) + 
      geom_bar(position = position_fill(reverse = TRUE), stat="identity", color = c("black"), size = 0.3) +
      geom_hline(yintercept = 0, 
                 linetype = "solid", 
                 color = "black", 
                 size = 0.5) +
      coord_cartesian(ylim = c(0, 1), clip = "off") +
      scale_fill_manual(values = c("chartreuse4", "#99CC33", "yellow", "red")) +
      scale_y_continuous(name = " ",
                         breaks = seq(0, 1, 0.1),
                         labels = c("0%", "10%", "20%", "30%", "40%", "50%", "60%", "70%", "80%", "90%", "100%")) +
      annotate("label",
               x = "2021-33",
               y = -0.1,
               label = "2021",
               label.r = unit(0, "lines")) +
      annotate("label",
               x = "2022-02",
               y = -0.1,
               label = "2022",
               label.r = unit(0, "lines")) +
      scale_x_discrete(labels = substr(unique(sort(slide$Viikko)), 6, 7)) + 
      theme_classic() +
      facet_wrap(~ group) +
      theme(panel.grid.major.x = element_blank(),
            panel.grid.major.y = element_line(color = "grey"),
            axis.ticks.y.left = element_line(color = "grey"),
            axis.ticks.x.bottom = element_blank(),
            axis.text.x = element_text(size = 12, 
                                       vjust = 6,
                                       family = "Arial"),
            axis.text.y = element_text(size = 12, family = "Arial"),
            axis.title.x.bottom = element_blank(),
            axis.line = element_blank(),
            panel.spacing = unit(2, "lines"),
            strip.text.x = element_text(lineheight = 2, size = 18, face = "bold", family = "Source Sans Pro"),
            strip.background = element_rect(size = 0.75),
            legend.position = "none",
            plot.margin = margin(0.5, 0.5, 2, 0, "lines"))
    
    # Create label - "Kalenteriviikko"
    f_label <- data.frame(group = c(NameImg1Var, NameImg2Var), 
                          label = c("    ", "Kalenteriviikko")) 
    f_label$group <- as.factor(f_label$group)
    
    pSlide +
      geom_text(data = f_label, 
                aes(label = label), 
                x = levels(slide$Viikko)[length(unique(slide$Viikko))-2], 
                y = -0.165,
                inherit.aes = FALSE,
                size = 5.25,
                family = "Arial")
    
  } else if(is.null(Img2Var)){
    
    Main <- list_all[[Img1Var]] %>% 
      pivot_longer(cols = 2:5, names_to = "levels", values_to = "n") %>% 
      mutate(levels = as.factor(case_when(levels == "1" ~ 1,
                                          levels == "2" ~ 2,
                                          levels == "3" ~ 3,
                                          TRUE ~ 4)),
             n = as.numeric(as.character(n)),
             group = NameImg1Var) %>% 
      data.frame()
    
    
    # Format data frame for right-hand side figure
    main <- source[[NameList]] %>%
      select(grep(paste0(".*", Img1Var, "|Viikko"), 
                  names(source[[NameList]]), value = F)) %>% 
      filter(!get(Img1Var) %in% c("", 9)) %>% 
      dplyr::count(Viikko, get(Img1Var)) %>% 
      group_by(Viikko) %>% 
      mutate(group = NameImg1Var) %>% 
      data.frame()
    
    # Combine and relevel
    colnames(main) <- colnames(Main)
    slide <- rbind(Main, main)
    slide$levels <- factor(slide$levels, levels = c(1:4))
    slide$Viikko <- as.factor(slide$Viikko)
    
    # Plot
    pSlide <- ggplot(slide, aes(fill = levels, y = n, x = Viikko)) + 
      geom_bar(position = position_fill(reverse = TRUE), stat="identity", color = c("black"), size = 0.3) +
      geom_hline(yintercept = 0, 
                 linetype = "solid", 
                 color = "black", 
                 size = 0.5) +
      coord_cartesian(ylim = c(0, 1), clip = "off") +
      scale_fill_manual(values = c("chartreuse4", "#99CC33", "yellow", "red")) +
      scale_y_continuous(name = " ",
                         breaks = seq(0, 1, 0.1),
                         labels = c("0%", "10%", "20%", "30%", "40%", "50%", "60%", "70%", "80%", "90%", "100%")) +
      annotate("label",
               x = "2020-51",
               y = -0.1,
               label = "2020",
               label.r = unit(0, "lines")) +
      annotate("label",
               x = "2021-01",
               y = -0.1,
               label = "2021",
               label.r = unit(0, "lines")) +
      annotate("label",
               x = "2022-02",
               y = -0.1,
               label = "2022",
               label.r = unit(0, "lines")) +
      scale_x_discrete(labels = substr(unique(slide$Viikko), 6, 7)) + 
      theme_classic() +
      facet_wrap(~ group) +
      theme(panel.grid.major.x = element_blank(),
            panel.grid.major.y = element_line(color = "grey"),
            axis.ticks.y.left = element_line(color = "grey"),
            axis.ticks.x.bottom = element_blank(),
            axis.text.x = element_text(size = 12, 
                                       vjust = 6,
                                       family = "Arial"),
            axis.text.y = element_text(size = 12, family = "Arial"),
            axis.title.x.bottom = element_blank(),
            axis.line = element_blank(),
            panel.spacing = unit(2, "lines"),
            strip.text.x = element_text(lineheight = 2, size = 18, face = "bold", family = "Source Sans Pro"),
            strip.background = element_rect(size = 0.75),
            legend.position = "none",
            plot.margin = margin(0.5, 0.5, 2, 0, "lines"))
    
    f_label <- data.frame(group = c(NameImg1Var), 
                          label = c("Kalenteriviikko")) 
    pSlide +
      geom_text(data = f_label, 
                aes(label = label), 
                x = levels(slide$Viikko)[length(unique(slide$Viikko))-2], 
                y = -0.165,
                inherit.aes = FALSE,
                size = 5.25,
                family = "Arial")
  } else{
    LEFT <- list_all[[Img1Var]] %>% 
      pivot_longer(cols = 2:5, names_to = "levels", values_to = "n") %>% 
      mutate(levels = as.factor(case_when(levels == "1" ~ 1,
                                          levels == "2" ~ 2,
                                          levels == "3" ~ 3,
                                          TRUE ~ 4)),
             n = as.numeric(as.character(n)),
             group = NameImg1Var) %>% 
      data.frame()
    
    RIGHT <- list_all[[Img2Var]] %>% 
      pivot_longer(cols = 2:5, names_to = "levels", values_to = "n") %>% 
      mutate(levels = as.factor(case_when(levels == "1" ~ 1,
                                          levels == "2" ~ 2,
                                          levels == "3" ~ 3,
                                          TRUE ~ 4)),
             n = as.numeric(as.character(n)),
             group = NameImg2Var) %>% 
      data.frame()
    
    left <- source[[NameList]] %>%
      select(grep(paste0(".*", Img1Var, "|Viikko"), 
                  names(source[[NameList]]), value = F))  %>% 
      filter(!get(Img1Var) %in% c("", 9)) %>% 
      dplyr::count(Viikko, get(Img1Var)) %>% 
      group_by(Viikko) %>% 
      mutate(group = NameImg1Var) %>% 
      data.frame()
    
    right <- source[[NameList]] %>%
      select(grep(paste0(".*", Img2Var, "|Viikko"), 
                  names(source[[NameList]]), value = F))  %>% 
      filter(!get(Img2Var) %in% c("", 9)) %>% 
      dplyr::count(Viikko, get(Img2Var)) %>% 
      group_by(Viikko) %>% 
      mutate(group = NameImg2Var) %>% 
      data.frame()
    
    names(left) <- names(right) <- names(LEFT)
    
    slide <- rbind(LEFT, RIGHT, left, right) 
    slide$group <- factor(slide$group, levels = c(NameImg1Var, NameImg2Var))
    slide$Viikko <- as.factor(slide$Viikko)
    
    labels <- levels(as.factor(slide$Viikko))
    labels[2 * 0:(length(levels(as.factor(slide$Viikko)))*0.5 - 1 )] <- ""
    
    pSlide <- ggplot(slide, aes(fill = levels, y = n, x = Viikko)) + 
      geom_bar(position = position_fill(reverse = TRUE), stat="identity", color = c("black"), size = 0.3, width = 0.8) +
      geom_hline(yintercept = 0, 
                 linetype = "solid", 
                 color = "black", 
                 size = 0.5) +
      coord_cartesian(ylim = c(0, 1), clip = "off") +
      scale_fill_manual(values = c("chartreuse4", "#99CC33", "yellow", "red")) +
      scale_y_continuous(name = " ",
                         breaks = seq(0, 1, 0.1),
                         labels = c("0%", "10%", "20%", "30%", "40%", "50%", "60%", "70%", "80%", "90%", "100%")) +
      annotate("label",
               x = "2020-52",
               y = -0.1,
               label = "2020",
               label.r = unit(0, "lines")) +
      annotate("label",
               x = "2021-02",
               y = -0.1,
               label = "2021",
               label.r = unit(0, "lines")) +
      annotate("label",
               x = "2021-52",
               y = -0.1,
               label = "2022",
               label.r = unit(0, "lines")) +
      scale_x_discrete(labels = substr(labels, 6, 7)) +  # unique(slide$Viikko)
      theme_classic() +
      facet_wrap(~ group) +
      theme(panel.grid.major.x = element_blank(),
            panel.grid.major.y = element_line(color = "grey"),
            axis.ticks.y.left = element_line(color = "grey"),
            axis.ticks.x.bottom = element_blank(),
            axis.text.x = element_text(size = 9, 
                                       vjust = 9,
                                       family = "Arial"),
            axis.text.y = element_text(size = 12, family = "Arial"),
            axis.title.x.bottom = element_blank(),
            axis.line = element_blank(),
            panel.spacing = unit(1, "lines"),
            strip.text.x = element_text(lineheight = 2, size = 18, face = "bold", family = "Source Sans Pro"),
            strip.background = element_rect(size = 0.75),
            legend.position = "none",
            plot.margin = margin(0.5, 0.5, 2.5, 0, "lines"))
    
    # Create label - "Kalenteriviikko"
    f_label <- data.frame(group = c(NameImg1Var, NameImg2Var), 
                          label = c("    ", "Kalenteriviikko")) 
    f_label$group <- as.factor(f_label$group)
    pSlide +
      geom_text(data = f_label, 
                aes(label = label), 
                x = levels(slide$Viikko)[length(unique(slide$Viikko))-4], 
                y = -0.165,
                inherit.aes = FALSE,
                size = 5.25,
                family = "Arial")
  }
}

CoronaReport_work <- function(Img1Var, NameImg1Var, Img2Var = NULL, NameImg2Var = NULL, source, NameList){
  
  if(is.null(Img2Var)){
    # Format data frame for right-hand side figure
    main <- source[[NameList]] %>%
      select(grep(paste0(".*", Img1Var, "|Viikko"), 
                  names(source[[NameList]]), value = F)) %>% 
      filter(!get(Img1Var) %in% c("", 9 , NA, "ERROR")) %>% 
      dplyr::count(Viikko, get(Img1Var)) %>% 
      group_by(Viikko) %>% 
      mutate(group = NameImg1Var) %>% 
      data.frame()
    
    
    # Combine and relevel
    colnames(main) <- c("Viikko", "levels", "n", "group")
    main$levels <- factor(main$levels, levels = c(1:4))
    main$Viikko <- as.factor(main$Viikko)
    
    # Plot
    pSlide <- ggplot(main, aes(fill = levels, y = n, x = Viikko)) + 
      geom_bar(position = position_fill(reverse = TRUE), stat="identity", color = c("black"), size = 0.3) +
      geom_hline(yintercept = 0, 
                 linetype = "solid", 
                 color = "black", 
                 size = 0.5) +
      coord_cartesian(ylim = c(0, 1), clip = "off") +
      scale_fill_manual(
        breaks = c(1, 2, 3, 4),
        values = c("chartreuse4", "#99CC33", "yellow", "red")) +
      scale_y_continuous(name = " ",
                         breaks = seq(0, 1, 0.1),
                         labels = c("0%", "10%", "20%", "30%", "40%", "50%", "60%", "70%", "80%", "90%", "100%")) +
      annotate("label",
               x = "2022-13",
               y = -0.1,
               label = "2022",
               label.r = unit(0, "lines")) +
      scale_x_discrete(
        limits = c(as.character(unique(main$Viikko)), rep(NA, 15)),
        labels = c(substr(unique(main$Viikko), 6, 7), " ")) + 
      theme_classic() +
      facet_wrap(~ group) +
      theme(panel.grid.major.x = element_blank(),
            panel.grid.major.y = element_line(color = "grey"),
            axis.ticks.y.left = element_line(color = "grey"),
            axis.ticks.x.bottom = element_blank(),
            axis.text.x = element_text(size = 12, 
                                       vjust = 6,
                                       family = "Arial"),
            axis.text.y = element_text(size = 12, family = "Arial"),
            axis.title.x.bottom = element_blank(),
            axis.line = element_blank(),
            panel.spacing = unit(2, "lines"),
            strip.text.x = element_text(size = 18, face = "bold", family = "Source Sans Pro"),
            strip.background = element_rect(size = 0.75),
            legend.position = "none",
            plot.margin = margin(0.5, 0.5, 2, 0, "lines"))
    
    f_label <- data.frame(group = c(NameImg1Var), 
                          label = c("Kalenteriviikko")) 
    pSlide +
      geom_text(data = f_label, 
                aes(label = label, hjust = -5.25), 
                x = "2022-13", 
                y = -0.165,
                inherit.aes = FALSE,
                size = 5.25,
                family = "Arial")
  } else{
    # Format data frame for left-hand side figure
    left <- source[[NameList]] %>%
      select(grep(paste0(".*", Img1Var, "|Viikko"), 
                  names(source[[NameList]]), value = F))  %>% 
      filter(!get(Img1Var) %in% c("", 9, NA, "ERROR")) %>% 
      dplyr::count(Viikko, get(Img1Var)) %>% 
      group_by(Viikko) %>% 
      mutate(group = NameImg1Var)
    # Format data frame for right-hand side figure
    right <- source[[NameList]] %>%
      select(grep(paste0(".*", Img2Var, "|Viikko"), 
                  names(source[[NameList]]), value = F)) %>% 
      filter(!get(Img2Var) %in% c("", 9, NA, "ERROR")) %>% 
      dplyr::count(Viikko, get(Img2Var)) %>% 
      group_by(Viikko) %>% 
      mutate(group = NameImg2Var)
    
    # Combine and relevel
    colnames(right) <- colnames(left)
    slide <- rbind(left, right)
    slide$group <- factor(slide$group, levels = c(NameImg1Var, NameImg2Var))
    colnames(slide) <- c("Viikko", "levels", "n", "group")
    slide$levels <- factor(slide$levels, levels = c(1:4))
    slide$Viikko <- as.factor(slide$Viikko)
    
    
    pSlide <- ggplot(slide, aes(fill = levels, y = n, x = Viikko)) + 
      geom_bar(position = position_fill(reverse = TRUE), stat="identity", color = c("black"), size = 0.3) +
      geom_hline(yintercept = 0, 
                 linetype = "solid", 
                 color = "black", 
                 size = 0.5) +
      coord_cartesian(ylim = c(0, 1), clip = "off") +
      scale_fill_manual(values = c("chartreuse4", "#99CC33", "yellow", "red")) +
      scale_y_continuous(name = " ",
                         breaks = seq(0, 1, 0.1),
                         labels = c("0%", "10%", "20%", "30%", "40%", "50%", "60%", "70%", "80%", "90%", "100%")) +
      annotate("label",
               x = "2022-13",
               y = -0.1,
               label = "2022",
               label.r = unit(0, "lines")) +
      scale_x_discrete(
        limits = c(as.character(unique(slide$Viikko)), rep(NA, 10)),
        labels = c(substr(unique(slide$Viikko), 6, 7), " ")) + 
      theme_classic() +
      facet_wrap(~ group, labeller = label_wrap_gen()) +
      theme(panel.grid.major.x = element_blank(),
            panel.grid.major.y = element_line(color = "grey"),
            axis.ticks.y.left = element_line(color = "grey"),
            axis.ticks.x.bottom = element_blank(),
            axis.text.x = element_text(size = 12, 
                                       vjust = 6,
                                       family = "Arial"),
            axis.text.y = element_text(size = 12, family = "Arial"),
            axis.title.x.bottom = element_blank(),
            axis.line = element_blank(),
            panel.spacing = unit(2, "lines"),
            strip.text.x = element_text(size = 18, face = "bold", family = "Source Sans Pro"),
            strip.background = element_rect(size = 0.75),
            legend.position = "none",
            plot.margin = margin(0.5, 0.5, 2, 0, "lines"))
    
    # Create label - "Kalenteriviikko"
    f_label <- data.frame(group = c(NameImg1Var, NameImg2Var), 
                          label = c("    ", "Kalenteriviikko")) 
    pSlide +
      geom_text(data = f_label, 
                aes(label = label, hjust = -2.25), 
                x = "2022-13", 
                y = -0.165,
                inherit.aes = FALSE,
                size = 5.25,
                family = "Arial")
  }
}

# Define a function make week biweekly
weekly <- function(source_new, source_ori){
  weeks <- lubridate::floor_date(ymd(str_extract_all(source_new$lopa_created, "\\d{4}-\\d{2}-\\d{2}")), unit = "week")
  #isodd <- as.POSIXlt(weeks)$yday %% 2 == 1
  #weeks[!isodd] <- weeks[!isodd] - 7L
  source_new$Viikko <- format(weeks, "%G-%V")
  names(source_new) <- names(source_ori)
  new <- rbind(source_ori, source_new)
  new$Viikko <- factor(as.character(new$Viikko))
  list <- list(source_new, new)
  names(list) <- c("new", "combined_new")
  return(list)
}

# Combine new data with raw data

PTH_combined_new <- weekly(PTH_New, PTH_ori)
PTH_combined_new[["new"]]$Viikko <- as.factor(PTH_combined_new[["new"]]$Viikko)
PTH_combined_new[["combined_new"]]$Viikko <- as.factor(as.character(PTH_combined_new[["combined_new"]]$Viikko))

ESH_combined_new <- weekly(ESH_New, ESH_ori)
ESH_combined_new[["new"]]$Viikko <- as.factor(ESH_combined_new[["new"]]$Viikko)
ESH_combined_new[["combined_new"]]$Viikko <- as.factor(as.character(ESH_combined_new[["combined_new"]]$Viikko))

SOS_combined_new <- weekly(SOS_New, SOS_ori)
SOS_combined_new[["new"]]$Viikko <- as.factor(SOS_combined_new[["new"]]$Viikko)
SOS_combined_new[["combined_new"]]$Viikko <- as.factor(as.character(SOS_combined_new[["combined_new"]]$Viikko))

# Data set for p5, 6, 26 and 39
# Define 
getFull <- function(source_ori, source_new, N, group){
  org <- source_new[["new"]] %>% 
    select(grep(paste0(".*", "Järjestäjäorganisaatio", "|Viikko"), 
                names(source_new[["new"]]), value = F)) %>% 
    group_by(Viikko) %>% 
    unique() %>% 
    dplyr::count() %>% 
    mutate(N = N) %>% 
    data.frame()
  
  org_1 <- rbind(list_all[[source_ori]], org)
  org_2 <- org_1 %>% 
    mutate(n = N - n)
  
  org_new <- rbind(org_1, org_2) %>% 
    arrange(Viikko) %>% 
    mutate(levels = as.factor(rep(c(1,2), length(levels(as.factor(org_1$Viikko))))),
           pct = round(n/N * 100, 0),
           group = group)
  
  org_new$pct[2 * c(0:length(levels(as.factor(org_1$Viikko))))] <- NA
  
  return(org_new)
}

PTH_full <- getFull(source_ori = "S5_PTH", 
                    source_new =  PTH_combined_new, 
                    N = 133, 
                    group = "PTH organisaatiot (133)")
ESH_full <- getFull(source_ori = "S5_ESH", 
                    source_new =  ESH_combined_new, 
                    N = 21, 
                    group = "ESH organisaatiot (21)")
SOS_full <- getFull(source_ori = "S5_SOS", 
                    source_new =  SOS_combined_new, 
                    N = 209, 
                    group = "SOS organisaatiot (209)")

CoronaFullplot <- function(source = NULL){
  if(is.null(source)){
    ORG <- rbind(PTH_full, ESH_full, SOS_full)
    ORG$group <- factor(ORG$group, c("PTH organisaatiot (133)", "SOS organisaatiot (209)", "ESH organisaatiot (21)"))
    labels <- levels(as.factor(ORG$Viikko))
    labels[2 * 0:(length(levels(as.factor(ORG$Viikko)))*0.5)] <- ""
    p <- ggplot(ORG, aes(fill = levels, y = n, x = Viikko)) + 
      geom_bar(position = position_fill(reverse = TRUE), stat = "identity", color = c("black"), size = 0.3, width = 0.8) +
      geom_label(aes(label = pct, color = "black", size = 3), 
                 position = position_fill(reverse = TRUE),
                 label.r = unit(0, "lines"),
                 size = 2.25,
                 fill = "yellow",
                 color = "black",
                 label.size = 0,
                 inherit.aes = T,
                 vjust = 1.0,
                 label.padding = unit(0.125, "lines")) +
      geom_hline(yintercept = 0, 
                 linetype = "solid", 
                 color = "black", 
                 size = 0.5) +
      coord_cartesian(ylim = c(0, 1), clip = "off") +
      scale_fill_manual(values = c("chartreuse4", "white")) +
      scale_y_continuous(name = " ",
                         breaks = seq(0, 1, 0.1),
                         labels = c("0%", "10%", "20%", "30%", "40%", "50%", "60%", "70%", "80%", "90%", "100%")) +
      annotate("label",
               x = "2020-52",
               y = -0.12,
               label = "2020",
               label.r = unit(0, "lines"),
               size = 3) +
      annotate("label",
               x = "2021-02",
               y = -0.12,
               label = "2021",
               label.r = unit(0, "lines"),
               size = 3) +
      annotate("label",
               x = "2022-02",
               y = -0.12,
               label = "2022",
               label.r = unit(0, "lines"),
               size = 3) +
      scale_x_discrete(labels = substr(labels, 6, 7)) + 
      theme_classic() +
      facet_wrap(~ group) +
      theme(panel.grid.major.x = element_blank(),
            panel.grid.major.y = element_line(color = "grey"),
            axis.ticks.y.left = element_line(color = "grey"),
            axis.ticks.x.bottom = element_blank(),
            axis.text.x = element_text(size = 12, 
                                       vjust = 5,
                                       family = "Arial"),
            axis.text.y = element_text(size = 12, family = "Arial"),
            axis.title.x.bottom = element_blank(),
            axis.line = element_blank(),
            panel.spacing = unit(0.75, "lines"),
            strip.text = element_text(size = 18, face = "bold", lineheight = 2, family = "Source Sans Pro"),
            legend.position = "none",
            plot.margin = margin(0.5, 2, 3, 0, "lines")) 
    
    # Create label - "Kalenteriviikko"
    f_label <- data.frame(group = c("PTH organisaatiot (133)", "SOS organisaatiot (209)", "ESH organisaatiot (21)"), 
                          label = c(" ", " ", "Kalenteriviikko")) 
    f_label$group <- as.factor(f_label$group)
    
    p +
      geom_text(data = f_label, 
                aes(label = label), 
                x = levels(as.factor(ORG$Viikko))[length(unique(as.factor(ORG$Viikko)))-4], 
                y = -0.18,
                inherit.aes = FALSE,
                size = 5.25)
  } else{
    p <- ggplot(source, aes(fill = levels, y = n, x = Viikko)) + 
      geom_bar(position = position_fill(reverse = TRUE), stat = "identity", color = c("black"), size = 0.3) +
      geom_label(aes(label = pct, color = "black"), 
                 position = position_fill(reverse = TRUE),
                 label.r = unit(-0, "lines"),
                 size = 3,
                 fill = "white",
                 color = "black",
                 label.size = 0,
                 inherit.aes = T,
                 vjust = 1.1) +
      geom_hline(yintercept = 0, 
                 linetype = "solid", 
                 color = "black", 
                 size = 0.5) +
      coord_cartesian(ylim = c(0, 1), clip = "off") +
      scale_fill_manual(values = c("chartreuse4", "white")) +
      scale_y_continuous(name = " ",
                         breaks = seq(0, 1, 0.1),
                         labels = c("0%", "10%", "20%", "30%", "40%", "50%", "60%", "70%", "80%", "90%", "100%")) +
      annotate("label",
               x = "2020-51",
               y = -0.12,
               label = "2020",
               label.r = unit(0, "lines")) +
      annotate("label",
               x = "2021-01",
               y = -0.12,
               label = "2021",
               label.r = unit(0, "lines")) +
      annotate("label",
               x = "2022-02",
               y = -0.12,
               label = "2022",
               label.r = unit(0, "lines")) +
      scale_x_discrete(labels = substr(levels(as.factor(source$Viikko)), 6, 7)) + 
      theme_classic() +
      facet_wrap(~ group) +
      theme(panel.grid.major.x = element_blank(),
            panel.grid.major.y = element_line(color = "grey"),
            axis.ticks.y.left = element_line(color = "grey"),
            axis.ticks.x.bottom = element_blank(),
            axis.text.x = element_text(size = 12, 
                                       vjust = 5,
                                       family = "Arial"),
            axis.text.y = element_text(size = 12, family = "Arial"),
            axis.title.x.bottom = element_blank(),
            axis.line = element_blank(),
            panel.spacing = unit(0.75, "lines"),
            strip.text = element_text(size = 18, face = "bold", lineheight = 2, family = "Source Sans Pro"),
            legend.position = "none",
            plot.margin = margin(0.5, 2, 3, 0, "lines"),
            text = element_text(size = 16)) 
    
    # Create label - "Kalenteriviikko"
    f_label <- data.frame(group = c(unique(source$group)), 
                          label = c("Kalenteriviikko")) 
    p +
      geom_text(data = f_label, 
                aes(label = label), 
                x = levels(as.factor(source$Viikko))[length(unique(as.factor(source$Viikko)))-4], 
                y = -0.2,
                inherit.aes = FALSE,
                size = 5.25,
                family = "Arial")
    
  }
  
}

p5 <- CoronaFullplot(source = NULL)
p6 <- CoronaFullplot(source = PTH_full)
p26 <- CoronaFullplot(source = SOS_full)
p39 <- CoronaFullplot(source = ESH_full)

pics <- list()
pics[[5]] <- p5
pics[[6]] <- p6
pics[[26]] <- p26
pics[[39]] <- p39

PPT_report_structure <- read_excel(ppt_sturcture_dir)
PPT_report_structure <- as.data.frame(PPT_report_structure[-c(1:6, 26, 39),])


for (i in c(PPT_report_structure$SlideNo)) {
  if(i %in% c(16:25)){
    pics[[i]] <- CoronaReport(Img1Var = PPT_report_structure[PPT_report_structure$SlideNo == i, 6],
                              NameImg1Var = PPT_report_structure[PPT_report_structure$SlideNo == i, 5],
                              Img2Var = PPT_report_structure[PPT_report_structure$SlideNo == i, 8],
                              NameImg2Var = PPT_report_structure[PPT_report_structure$SlideNo == i, 7],
                              source = PTH_combined_new,
                              NameList = "combined_new")
  } else if(i %in% c(41:45)){
    pics[[i]] <- CoronaReport(Img1Var = PPT_report_structure[PPT_report_structure$SlideNo == i, 6],
                              NameImg1Var = PPT_report_structure[PPT_report_structure$SlideNo == i, 5],
                              source = ESH_combined_new,
                              NameList = "new")
  } else if (i %in% c(7:15)){
    pics[[i]] <- CoronaReport(Img1Var = PPT_report_structure[PPT_report_structure$SlideNo == i, 6],
                              NameImg1Var = PPT_report_structure[PPT_report_structure$SlideNo == i, 5],
                              Img2Var = PPT_report_structure[PPT_report_structure$SlideNo == i, 8],
                              NameImg2Var = PPT_report_structure[PPT_report_structure$SlideNo == i, 7],
                              source = PTH_combined_new,
                              NameList = "new")
    
  } else if(i %in% c(27:36)){
    pics[[i]] <- CoronaReport(Img1Var = PPT_report_structure[PPT_report_structure$SlideNo == i, 6],
                              NameImg1Var = PPT_report_structure[PPT_report_structure$SlideNo == i, 5],
                              Img2Var = PPT_report_structure[PPT_report_structure$SlideNo == i, 8],
                              NameImg2Var = PPT_report_structure[PPT_report_structure$SlideNo == i, 7],
                              source = SOS_combined_new,
                              NameList = "new")
  } else if(i %in% c(37:38)){
    pics[[i]] <- CoronaReport(Img1Var = PPT_report_structure[PPT_report_structure$SlideNo == i, 6],
                              NameImg1Var = PPT_report_structure[PPT_report_structure$SlideNo == i, 5],
                              Img2Var = PPT_report_structure[PPT_report_structure$SlideNo == i, 8],
                              NameImg2Var = PPT_report_structure[PPT_report_structure$SlideNo == i, 7],
                              source = SOS_combined_new,
                              NameList = "combined_new")
  } else{
    pics[[i]] <- CoronaReport(Img1Var = PPT_report_structure[PPT_report_structure$SlideNo == i, 6],
                              NameImg1Var = PPT_report_structure[PPT_report_structure$SlideNo == i, 5],
                              Img2Var = PPT_report_structure[PPT_report_structure$SlideNo == i, 8],
                              NameImg2Var = PPT_report_structure[PPT_report_structure$SlideNo == i, 7],
                              source = ESH_combined_new,
                              NameList = "combined_new")
  }
}

pics[[46]] <- CoronaReport_work(Img1Var = "pth_service_availability_option_10",
                                NameImg1Var = "Palvelujen saatavuus",
                                Img2Var = "pth_personnel_sufficiency_option_10",
                                NameImg2Var = "Henkilöstön riittävyys",
                                source = PTH_combined_new,
                                NameList = "combined_new")

pics[[47]] <- CoronaReport_work(Img1Var = "pth_service_availability_option_11",
                                NameImg1Var = "Palvelujen saatavuus",
                                Img2Var = "pth_personnel_sufficiency_option_11",
                                NameImg2Var = "Henkilöstön riittävyys",
                                source = PTH_combined_new,
                                NameList = "combined_new")

pics[[48]] <- CoronaReport_work(Img1Var = "esh_service_availability_option_1",
                                NameImg1Var = "Tehohoidon palvelujen saatavuus",
                                Img2Var = "esh_service_availability_option_2",
                                NameImg2Var = "Päivystyksen/ensihoidon palv. saatavuus",
                                source = ESH_combined_new,
                                NameList = "new")

pics[[49]] <- CoronaReport_work(Img1Var = "esh_service_availability_option_3",
                                NameImg1Var = "Leikkaus&anestesiahoidon saatavuus",
                                Img2Var = "esh_service_availability_option_4",
                                NameImg2Var = "Synnytystoiminnan saatavuus",
                                source = ESH_combined_new,
                                NameList = "new")

pics[[50]] <- CoronaReport_work(Img1Var = "esh_service_availability_option_5",
                                NameImg1Var = "Laboratorio&kuvantamispalveluiden saatavuus",
                                source = ESH_combined_new,
                                NameList = "combined_new")


myppt <- read_pptx(ppt_template_dir)
mytmp <- layout_properties(myppt, layout = "Custom Layout", master = "THL_ppt-pohja_FI_2019")


for (i in 1:50) {
  ppt <- myppt %>% 
    on_slide(index = i) %>% 
    ph_remove(type = "dt") %>% 
    ph_with(value = date, ph_location_type(type = "dt"))
  if (i %in% 5:50){
    ppt <- myppt %>% 
      on_slide(index = i) %>% 
      ph_with(value = pics[[i]], location = ph_location_label(ph_label = "Content Placeholder 7")) 
  } 
}

print(ppt, ppt_save_dir)

# Save updates when everything is alright
# PTH
write.xlsx(PTH_New, PTH_updated_dir, row.names = F)
# ESH
write.xlsx(ESH_New, ESH_updated_dir, row.names = F)
# SOS
write.xlsx(SOS_New, SOS_updated_dir, row.names = F)
