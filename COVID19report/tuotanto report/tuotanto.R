library("xlsx")

wb <- loadWorkbook("/home/zzhh/mnt/ZZHH/_Desktop/Tasks/Koronaraportointi/type3/Tuotanto_report_template_2022-04-01.xlsx")
WeekNumber <- "2022-35"
output_dir <- "/home/zzhh/mnt/ZZHH/_Desktop/Tasks/Koronaraportointi/type3/THL_STM kuntakyselyt_TUOTANTO_SHP_20220908_vk35.xlsx"

##### Color Setting #####


Fill0 <- Fill(foregroundColor = "white")    # create fill object # 0
CellStyle0 <- CellStyle(wb, fill = Fill0) + Border(color = "black", position = c("BOTTOM", "LEFT", "TOP", "RIGHT"))

Fill1 <- Fill(foregroundColor = "#009933")  # create fill object # 1
CellStyle1 <- CellStyle(wb, fill = Fill1) + Border(color = "black", position = c("BOTTOM", "LEFT", "TOP", "RIGHT"))

Fill2 <- Fill(foregroundColor = "#99CC33")  # create fill object # 2
CellStyle2 <- CellStyle(wb, fill = Fill2) + Border(color = "black", position = c("BOTTOM", "LEFT", "TOP", "RIGHT"))

Fill3 <- Fill(foregroundColor = "yellow")   # create fill object # 3
CellStyle3 <- CellStyle(wb, fill = Fill3) + Border(color = "black", position = c("BOTTOM", "LEFT", "TOP", "RIGHT"))

Fill4 <- Fill(foregroundColor = "red")      # create fill object # 4
CellStyle4 <- CellStyle(wb, fill = Fill4) + Border(color = "black", position = c("BOTTOM", "LEFT", "TOP", "RIGHT"))

FontStyle <- CellStyle(wb, fill = Fill4) + 
  Font(wb, name = "Arial", heightInPoints = 16, color = "ghostwhite", isBold = TRUE, isItalic = FALSE) +
  Alignment(horizontal = "ALIGN_CENTER") +
  Border(color = "black", position = c("BOTTOM", "LEFT", "TOP", "RIGHT"))

FontStyle1 <- CellStyle(wb, fill = Fill0) + 
  Font(wb, name = "Arial", heightInPoints = 14, color = "gray5", isBold = FALSE, isItalic = FALSE) +
  Alignment(horizontal = "ALIGN_CENTER", vertical = "VERTICAL_CENTER") +
  Border(color = "black", position = c("BOTTOM", "LEFT", "TOP", "RIGHT"))

FontStyle2 <- CellStyle(wb, fill = Fill0) + 
  Font(wb, name = "Tempus Sans ITC", heightInPoints = 14, color = "gray5", isBold = TRUE, isItalic = FALSE) +
  Alignment(horizontal = "ALIGN_CENTER") +
  Border(color = "black", position = c("BOTTOM", "LEFT", "TOP", "RIGHT"))

headerStyle <- CellStyle(wb,
                         fill = Fill(foregroundColor = "white"),
                         font = Font(wb, name = "Dyuthi", isBold = TRUE, color = "#009933", heightInPoints = 20),
                         alignment = Alignment(wrapText = TRUE,
                                               horizontal = "ALIGN_CENTER",
                                               vertical = "VERTICAL_CENTER"),
                         border = Border(color = "gray5", 
                                         position = c("BOTTOM", "LEFT", "TOP", "RIGHT"),
                                         pen = c("BORDER_THIN", "BORDER_THIN","BORDER_THIN","BORDER_THIN")))

BlockStyle0 <- CellStyle(wb,
                         fill = Fill(foregroundColor = "white"),
                         alignment = Alignment(wrapText = TRUE,
                                               horizontal = "ALIGN_CENTER",
                                               vertical = "VERTICAL_CENTER"),
                         border = Border(color = "black", 
                                         position = c("BOTTOM", "LEFT", "TOP", "RIGHT"),
                                         pen = c("BORDER_THIN", "BORDER_THIN","BORDER_THIN","BORDER_THIN")))
BlockStyle1 <- CellStyle(wb,
                         fill = Fill(foregroundColor = "#009933"),
                         alignment = Alignment(wrapText = TRUE,
                                               horizontal = "ALIGN_CENTER",
                                               vertical = "VERTICAL_CENTER"),
                         border = Border(color = "black", 
                                         position = c("BOTTOM", "LEFT", "TOP", "RIGHT"),
                                         pen = c("BORDER_THIN", "BORDER_THIN","BORDER_THIN","BORDER_THIN")))

BlockStyle2 <- CellStyle(wb,
                         fill = Fill(foregroundColor = "yellow"),
                         alignment = Alignment(wrapText = TRUE,
                                               horizontal = "ALIGN_CENTER",
                                               vertical = "VERTICAL_CENTER"),
                         border = Border(color = "black", 
                                         position = c("BOTTOM", "LEFT", "TOP", "RIGHT"), 
                                         pen = c("BORDER_THIN", "BORDER_THIN","BORDER_THIN","BORDER_THIN")))
BlockStyle3 <- CellStyle(wb,
                         fill = Fill(foregroundColor = "red"),
                         alignment = Alignment(wrapText = TRUE,
                                               horizontal = "ALIGN_CENTER",
                                               vertical = "VERTICAL_CENTER"),
                         border = Border(color = "black", 
                                         position = c("BOTTOM", "LEFT", "TOP", "RIGHT"), 
                                         pen = c("BORDER_THIN", "BORDER_THIN","BORDER_THIN","BORDER_THIN")))

BlockStyle4 <- CellStyle(wb,
                         fill = Fill(foregroundColor = "white"),
                         font = Font(wb, name = "Arial", isBold = FALSE, color = "blue", heightInPoints = 14),
                         alignment = Alignment(horizontal = "ALIGN_CENTER",
                                               vertical = "VERTICAL_CENTER"),
                         border = Border(color = "gray5", 
                                         position = c("BOTTOM", "LEFT", "TOP", "RIGHT"),
                                         pen = c("BORDER_THIN", "BORDER_THIN","BORDER_THIN","BORDER_THIN")))
BlockStyle5 <- CellStyle(wb,
                         fill = Fill(foregroundColor = "white"),
                         font = Font(wb, name = "Arial", isBold = TRUE, color = "blue", heightInPoints = 14),
                         border = Border(color = "gray5", 
                                         position = c("BOTTOM", "LEFT", "TOP", "RIGHT"),
                                         pen = c("BORDER_THIN", "BORDER_THIN","BORDER_THIN","BORDER_THIN")),
                         alignment = Alignment(horizontal = "ALIGN_CENTER",
                                               vertical = "VERTICAL_CENTER")
)


##### Area and rows #####
area <- data.frame(c("Etelä-Karjala", "HUS", "Kymenlaakso", "Päijät-Häme", "Etelä-Savo", 
                     "Itä-Savo", "Keski-Suomi", "Pohjois-Karjala", "Pohjois-Savo", "Kainuu",
                     "Keski-Pohjanmaa", "Lappi", "Länsi-Pohja", "Pohjois-Pohjanmaa", "Etelä-Pohjanmaa",
                     "Kanta-Häme", "Pirkanmaa", "Satakunta", "Vaasa", "Varsinais-Suomi", "Ahvenanmaa"))
colnames(area) <- "area"
area$RowType1 <- 7:27
area$RowType2 <- 9:29
area$RowType3 <- 6:26
area$RowType4 <- 10:30

##### Sheet and Color Function #####
sheets <- xlsx::getSheets(wb)

##### PTH palv saatavuus #####
sheet1 <- sheets[["PTH palv saatavuus"]]                  # get specific sheet
rows1   <- xlsx::getRows(sheet1)         # get rows
cells1  <- xlsx::getCells(rows1)         # get cells

colorBlock1 <- function(j, WeekNum){
  
  sheetVar <- c("PTH.avosh.ps", "PTH.vos.ps", "PTH.neuv.ps", 
                "PTH.kouluth.ps", "PTH.opiskth.ps", "PTH.kunt.ps", 
                "PTH.suunth.ps", "PTH.mieli.ps", "PTH.päihde.ps", 
                "pth_service_availability_option_10", "pth_service_availability_option_11")
  k = 0
  
  setCellValue(cells1[[paste0(area[j, 2],".", 3)]], WeekNum)
  setCellStyle(cells1[[paste0(area[j, 2],".", 3)]], FontStyle1)
  
  for (i in c(sheetVar)) {
    
    source <- PTH_combined_new  
    
    count <- source[["combined_new"]] %>% 
      group_by(Viikko, Sairaanhoitopiiri, get(i)) %>% 
      dplyr::count() %>% 
      setNames(., c("Viikko", "Sairaanhoitopiiri", i, "n")) %>% 
      filter(get(i) %in% c(1:4, "1", "2", "3", "4"), 
             Viikko == WeekNum,
             Sairaanhoitopiiri == area[j, 1])
    
    k = k + 5
    
    if(1 %in% count[[i]]){
      setCellStyle(cells1[[paste0(area[j, 2],".", k)]], CellStyle1)
    } else {
      setCellStyle(cells1[[paste0(area[j, 2],".", k)]], CellStyle0)
    }
    
    if(2 %in% count[[i]]){
      setCellStyle(cells1[[paste0(area[j, 2],".", k+1)]], CellStyle2)
    } else {
      setCellStyle(cells1[[paste0(area[j, 2],".", k+1)]], CellStyle0)
    }
    
    if(any(3 %in% count[[i]])){
      setCellStyle(cells1[[paste0(area[j, 2],".", k+2)]], CellStyle3)
    } else {
      setCellStyle(cells1[[paste0(area[j, 2],".", k+2)]], CellStyle0)
    }
    
    if(4 %in% count[[i]]){
      setCellValue(cells1[[paste0(area[j, 2],".", k+3)]], count[count[[i]] == 4, ]$n)
      setCellStyle(cells1[[paste0(area[j, 2],".", k+3)]], FontStyle)
    } else{
      setCellStyle(cells1[[paste0(area[j, 2],".", k+3)]], CellStyle0)
    }
  }
} #RowType1: 7:27

for (i in 1:21) {
  colorBlock1(i, WeekNumber)
}

##### PTH palv saatavuus 2 #####
sheet2 <- sheets[["PTH palv saatavuus (2)"]]                  # get specific sheet
rows2   <- xlsx::getRows(sheet2)         # get rows
cells2  <- xlsx::getCells(rows2)         # get cells

colorBlock2 <- function(j, WeekNum){
  
  sheetVar <- c("PTH.roko.ps", "PTH.test.ps", "PTH.tjälj.ps", "PTH.tneuvo.ps", "PTH.todtark.ps", "PTH.testmtulo.ps", "PTH.2.test.ps")
  k = 0
  
  setCellValue(cells2[[paste0(area[j, 3],".", 3)]], WeekNum)
  setCellStyle(cells2[[paste0(area[j, 3],".", 3)]], FontStyle1)
  
  for (i in c(sheetVar)) {
    
    source <- PTH_combined_new
    
    count <- source[["combined_new"]] %>% 
      group_by(Viikko, Sairaanhoitopiiri, get(i)) %>% 
      dplyr::count() %>% 
      setNames(., c("Viikko", "Sairaanhoitopiiri", i, "n")) %>% 
      filter(get(i) %in% c(1:4), 
             Viikko == WeekNum,
             Sairaanhoitopiiri == area[j, 1])
    
    k = k + 5
    
    if(1 %in% count[[i]]){
      setCellStyle(cells2[[paste0(area[j, 3],".", k)]], CellStyle1)
    } else {
      setCellStyle(cells2[[paste0(area[j, 3],".", k)]], CellStyle0)
    }
    
    if(2 %in% count[[i]]){
      setCellStyle(cells2[[paste0(area[j, 3],".", k+1)]], CellStyle2)
    } else {
      setCellStyle(cells2[[paste0(area[j, 3],".", k+1)]], CellStyle0)
    }
    
    if(3 %in% count[[i]]){
      setCellStyle(cells2[[paste0(area[j, 3],".", k+2)]], CellStyle3)
    } else{
      setCellStyle(cells2[[paste0(area[j, 3],".", k+2)]], CellStyle0) 
    }
    
    if(4 %in% count[[i]]){
      setCellValue(cells2[[paste0(area[j, 3],".", k+3)]], count[count[[i]] == 4, ]$n)
      setCellStyle(cells2[[paste0(area[j, 3],".", k+3)]], FontStyle)
    } else{
      setCellStyle(cells2[[paste0(area[j, 3],".", k+3)]], CellStyle0) 
    }
  }
} #RowType2: 9:29
for (i in 1:21) {
  colorBlock2(i, WeekNumber)
}

##### PTH palv saatavuus 3 #####
sheet3 <- sheets[["PTH palv saatavuus (3)"]]                  # get specific sheet
rows3   <- xlsx::getRows(sheet3)         # get rows
cells3  <- xlsx::getCells(rows3)         # get cells

colorBlock3 <- function(j, WeekNum){
  
  sheetVar <- c("PTH.lapsimttperus.ps", "PTH.lapsimtttesh", "PTH.nuorimttperus.ps", "PTH.nuorimttesh.ps", "PTH.aikuismttperus", "PTH.aikuismttesh.ps")
  k = 0
  
  setCellValue(cells3[[paste0(area[j, 3],".", 3)]], WeekNum)
  setCellStyle(cells3[[paste0(area[j, 3],".", 3)]], FontStyle1)
  
  for (i in c(sheetVar)) {
    
    source <- PTH_combined_new
    
    count <- source[["combined_new"]] %>% 
      group_by(Viikko, Sairaanhoitopiiri, get(i)) %>% 
      dplyr::count() %>% 
      setNames(., c("Viikko", "Sairaanhoitopiiri", i, "n")) %>% 
      filter(get(i) %in% c(1:4), 
             Viikko == WeekNum,
             Sairaanhoitopiiri == area[j, 1])
    
    k = k + 5
    
    if(1 %in% count[[i]]){
      setCellStyle(cells3[[paste0(area[j, 3],".", k)]], CellStyle1)
    } else {
      setCellStyle(cells3[[paste0(area[j, 3],".", k)]], CellStyle0)
    }
    
    if(2 %in% count[[i]]){
      setCellStyle(cells3[[paste0(area[j, 3],".", k+1)]], CellStyle2)
    } else {
      setCellStyle(cells3[[paste0(area[j, 3],".", k+1)]], CellStyle0)
    }
    
    if(3 %in% count[[i]]){
      setCellStyle(cells3[[paste0(area[j, 3],".", k+2)]], CellStyle3)
    } else{
      setCellStyle(cells3[[paste0(area[j, 3],".", k+2)]], CellStyle0) 
    }
    
    if(4 %in% count[[i]]){
      setCellValue(cells3[[paste0(area[j, 3],".", k+3)]], count[count[[i]] == 4, ]$n)
      setCellStyle(cells3[[paste0(area[j, 3],".", k+3)]], FontStyle)
    } else{
      setCellStyle(cells3[[paste0(area[j, 3],".", k+3)]], CellStyle0) 
    }
  }
} #RowType2: 9:29
for (i in 1:21) {
  colorBlock3(i, WeekNumber)
}

##### PTH henk riittävyys #####
sheet5 <- sheets[["PTH henk riittävyys"]]                  # get specific sheet
rows5   <- xlsx::getRows(sheet5)         # get rows
cells5  <- xlsx::getCells(rows5)         # get cells

colorBlock5 <- function(j, WeekNum){
  
  sheetVar <- c("PTH.avosh.hr", "PTH.vos.hr", "PTH.neuv.hr", "PTH.kouluth.hr","PTH.opiskth.hr", "PTH.kunt.hr", "PTH.suunth.hr", "PTH.mieli.hr", "PTH.päihde.hr",
                "pth_personnel_sufficiency_option_10", "pth_personnel_sufficiency_option_11")
  k = 0
  
  setCellValue(cells5[[paste0(area[j, 2],".", 3)]], WeekNum)
  setCellStyle(cells5[[paste0(area[j, 2],".", 3)]], FontStyle1)
  
  for (i in c(sheetVar)) {
    
    source <- PTH_combined_new
    
    count <- source[["combined_new"]] %>% 
      group_by(Viikko, Sairaanhoitopiiri, get(i)) %>% 
      dplyr::count() %>% 
      setNames(., c("Viikko", "Sairaanhoitopiiri", i, "n")) %>% 
      filter(get(i) %in% c(1:4), 
             Viikko == WeekNum,
             Sairaanhoitopiiri == area[j, 1])
    
    k = k + 5
    
    if(1 %in% count[[i]]){
      setCellStyle(cells5[[paste0(area[j, 2],".", k)]], CellStyle1)
    } else {
      setCellStyle(cells5[[paste0(area[j, 2],".", k)]], CellStyle0)
    }
    
    if(2 %in% count[[i]]){
      setCellStyle(cells5[[paste0(area[j, 2],".", k+1)]], CellStyle2)
    } else {
      setCellStyle(cells5[[paste0(area[j, 2],".", k+1)]], CellStyle0)
    }
    
    if(3 %in% count[[i]]){
      setCellStyle(cells5[[paste0(area[j, 2],".", k+2)]], CellStyle3)
    } else{
      setCellStyle(cells5[[paste0(area[j, 2],".", k+2)]], CellStyle0) 
    }
    
    if(4 %in% count[[i]]){
      setCellValue(cells5[[paste0(area[j, 2],".", k+3)]], count[count[[i]] == 4, ]$n)
      setCellStyle(cells5[[paste0(area[j, 2],".", k+3)]], FontStyle)
    } else{
      setCellStyle(cells5[[paste0(area[j, 2],".", k+3)]], CellStyle0) 
    }
  }
} #RowType1: 7:27
for (i in 1:21) {
  colorBlock5(i, WeekNumber)
}

##### PTH henk riittävyys (2) #####
sheet6 <- sheets[["PTH henk riittävyys (2)"]]                  # get specific sheet
rows6   <- xlsx::getRows(sheet6)         # get rows
cells6  <- xlsx::getCells(rows6)         # get cells

colorBlock6 <- function(j, WeekNum){
  
  sheetVar <- c("PTH.roko.hr", "PTH.test.hr", "PTH.tjälj.hr", "PTH.tneuv.hr", "PTH.todtark.hr", "PTH.testmsaap.hr", "PTH.2.test.hr")
  k = 0
  
  setCellValue(cells6[[paste0(area[j, 3],".", 3)]], WeekNum)
  setCellStyle(cells6[[paste0(area[j, 3],".", 3)]], FontStyle1)
  
  for (i in c(sheetVar)) {
    
    if(grepl("PTH", i, fixed = TRUE)){
      source <- PTH_combined_new  
    } else{
      source <- SOS_combined_new
    }
    
    count <- source[["combined_new"]] %>% 
      group_by(Viikko, Sairaanhoitopiiri, get(i)) %>% 
      dplyr::count() %>% 
      setNames(., c("Viikko", "Sairaanhoitopiiri", i, "n")) %>% 
      filter(get(i) %in% c(1:4), 
             Viikko == WeekNum,
             Sairaanhoitopiiri == area[j, 1])
    
    k = k + 5
    
    if(1 %in% count[[i]]){
      setCellStyle(cells6[[paste0(area[j, 3],".", k)]], CellStyle1)
    } else {
      setCellStyle(cells6[[paste0(area[j, 3],".", k)]], CellStyle0)
    }
    
    if(2 %in% count[[i]]){
      setCellStyle(cells6[[paste0(area[j, 3],".", k+1)]], CellStyle2)
    } else {
      setCellStyle(cells6[[paste0(area[j, 3],".", k+1)]], CellStyle0)
    }
    
    if(3 %in% count[[i]]){
      setCellStyle(cells6[[paste0(area[j, 3],".", k+2)]], CellStyle3)
    } else{
      setCellStyle(cells6[[paste0(area[j, 3],".", k+2)]], CellStyle0) 
    }
    
    if(4 %in% count[[i]]){
      setCellValue(cells6[[paste0(area[j, 3],".", k+3)]], count[count[[i]] == 4, ]$n)
      setCellStyle(cells6[[paste0(area[j, 3],".", k+3)]], FontStyle)
    } else{
      setCellStyle(cells6[[paste0(area[j, 3],".", k+3)]], CellStyle0) 
    }
  }
} #RowType1: 9:29
for (i in 1:21) {
  colorBlock6(i, WeekNumber)
}

##### SOS palv saatavuus #####
sheet8 <- sheets[["SOS palv saatavuus"]]                  # get specific sheet
rows8   <- xlsx::getRows(sheet8)         # get rows
cells8  <- xlsx::getCells(rows8)         # get cells

colorBlock8 <- function(j, WeekNum){
  
  sheetVar <- c("SOS.iäkkäät.kotiin.ps", "SOS.iäkkäät.ympäriv.ps", "SOS.lapsip.ps", 
                "SOS.lastens.ps", "SOS.vamm.ps", "SOS.työikä.ps", 
                "SOS.perheoik.ps", "SOS.ttt.ps", "SOS.tilapasum.ps", "SOS.sospäiv.ps")
  k = 0
  
  setCellValue(cells8[[paste0(area[j, 2],".", 3)]], WeekNum)
  setCellStyle(cells8[[paste0(area[j, 2],".", 3)]], FontStyle1)
  
  for (i in c(sheetVar)) {
    
    if(grepl("PTH", i, fixed = TRUE)){
      source <- PTH_combined_new  
    } else{
      source <- SOS_combined_new
    }
    
    count <- source[["combined_new"]] %>% 
      group_by(Viikko, Sairaanhoitopiiri, get(i)) %>% 
      dplyr::count() %>% 
      setNames(., c("Viikko", "Sairaanhoitopiiri", i, "n")) %>% 
      filter(get(i) %in% c(1:4), 
             Viikko == WeekNum,
             Sairaanhoitopiiri == area[j, 1])
    
    k = k + 5
    
    if(1 %in% count[[i]]){
      setCellStyle(cells8[[paste0(area[j, 2],".", k)]], CellStyle1)
    } else {
      setCellStyle(cells8[[paste0(area[j, 2],".", k)]], CellStyle0)
    }
    
    if(2 %in% count[[i]]){
      setCellStyle(cells8[[paste0(area[j, 2],".", k+1)]], CellStyle2)
    } else {
      setCellStyle(cells8[[paste0(area[j, 2],".", k+1)]], CellStyle0)
    }
    
    if(3 %in% count[[i]]){
      setCellStyle(cells8[[paste0(area[j, 2],".", k+2)]], CellStyle3)
    } else{
      setCellStyle(cells8[[paste0(area[j, 2],".", k+2)]], CellStyle0) 
    }
    
    if(4 %in% count[[i]]){
      setCellValue(cells8[[paste0(area[j, 2],".", k+3)]], count[count[[i]] == 4, ]$n)
      setCellStyle(cells8[[paste0(area[j, 2],".", k+3)]], FontStyle)
    } else{
      setCellStyle(cells8[[paste0(area[j, 2],".", k+3)]], CellStyle0) 
    }
  }
} #RowType1: 7:27
for (i in 1:21) {
  colorBlock8(i, WeekNumber)
}

##### SOS palv saatavuus (2) #####
sheet9 <- sheets[["SOS palv saatavuus (2)"]]                  # get specific sheet
rows9   <- xlsx::getRows(sheet9)         # get rows
cells9  <- xlsx::getCells(rows9)         # get cells

colorBlock9 <- function(j, WeekNum){
  
  sheetVar <- c("SOS.päihdet.ps", "SOS.omaishoit.ps")
  k = 0
  
  setCellValue(cells9[[paste0(area[j, 2],".", 3)]], WeekNum)
  setCellStyle(cells9[[paste0(area[j, 2],".", 3)]], FontStyle1)
  
  for (i in c(sheetVar)) {
    
    if(grepl("PTH", i, fixed = TRUE)){
      source <- PTH_combined_new  
    } else{
      source <- SOS_combined_new
    }
    
    count <- source[["combined_new"]] %>% 
      group_by(Viikko, Sairaanhoitopiiri, get(i)) %>% 
      dplyr::count() %>% 
      setNames(., c("Viikko", "Sairaanhoitopiiri", i, "n")) %>% 
      filter(get(i) %in% c(1:4), 
             Viikko == WeekNum,
             Sairaanhoitopiiri == area[j, 1])
    
    k = k + 5
    
    if(1 %in% count[[i]]){
      setCellStyle(cells9[[paste0(area[j, 2],".", k)]], CellStyle1)
    } else {
      setCellStyle(cells9[[paste0(area[j, 2],".", k)]], CellStyle0)
    }
    
    if(2 %in% count[[i]]){
      setCellStyle(cells9[[paste0(area[j, 2],".", k+1)]], CellStyle2)
    } else {
      setCellStyle(cells9[[paste0(area[j, 2],".", k+1)]], CellStyle0)
    }
    
    if(3 %in% count[[i]]){
      setCellStyle(cells9[[paste0(area[j, 2],".", k+2)]], CellStyle3)
    } else{
      setCellStyle(cells9[[paste0(area[j, 2],".", k+2)]], CellStyle0) 
    }
    
    if(4 %in% count[[i]]){
      setCellValue(cells9[[paste0(area[j, 2],".", k+3)]], count[count[[i]] == 4, ]$n)
      setCellStyle(cells9[[paste0(area[j, 2],".", k+3)]], FontStyle)
    } else{
      setCellStyle(cells9[[paste0(area[j, 2],".", k+3)]], CellStyle0) 
    }
  }
} #RowType1: 7:27
for (i in 1:21) {
  colorBlock9(i, WeekNumber)
}

##### SOS henk riittävyys #####
sheet11 <- sheets[["SOS henk riittävyys"]]                  # get specific sheet
rows11   <- xlsx::getRows(sheet11)         # get rows
cells11  <- xlsx::getCells(rows11)         # get cells

colorBlock11 <- function(j, WeekNum){
  
  sheetVar <- c("SOS.iäkkäät.kotiin.hr", "SOS.iäkkäät.ympäriv.hr", "SOS.lapsip.hr", 
                "SOS.lastens.hr", "SOS.vamm.hr", "SOS.työikä.hr", 
                "SOS.perheoik.hr", "SOS.ttt.hr", "SOS.tilapasum.hr", "SOS.sospäiv.hr")
  k = 0
  
  setCellValue(cells11[[paste0(area[j, 2],".", 3)]], WeekNum)
  setCellStyle(cells11[[paste0(area[j, 2],".", 3)]], FontStyle1)
  
  for (i in c(sheetVar)) {
    
    if(grepl("PTH", i, fixed = TRUE)){
      source <- PTH_combined_new  
    } else{
      source <- SOS_combined_new
    }
    
    count <- source[["combined_new"]] %>% 
      group_by(Viikko, Sairaanhoitopiiri, get(i)) %>% 
      dplyr::count() %>% 
      setNames(., c("Viikko", "Sairaanhoitopiiri", i, "n")) %>% 
      filter(get(i) %in% c(1:4), 
             Viikko == WeekNum,
             Sairaanhoitopiiri == area[j, 1])
    
    k = k + 5
    
    if(1 %in% count[[i]]){
      setCellStyle(cells11[[paste0(area[j, 2],".", k)]], CellStyle1)
    } else {
      setCellStyle(cells11[[paste0(area[j, 2],".", k)]], CellStyle0)
    }
    
    if(2 %in% count[[i]]){
      setCellStyle(cells11[[paste0(area[j, 2],".", k+1)]], CellStyle2)
    } else {
      setCellStyle(cells11[[paste0(area[j, 2],".", k+1)]], CellStyle0)
    }
    
    if(3 %in% count[[i]]){
      setCellStyle(cells11[[paste0(area[j, 2],".", k+2)]], CellStyle3)
    } else{
      setCellStyle(cells11[[paste0(area[j, 2],".", k+2)]], CellStyle0) 
    }
    
    if(4 %in% count[[i]]){
      setCellValue(cells11[[paste0(area[j, 2],".", k+3)]], count[count[[i]] == 4, ]$n)
      setCellStyle(cells11[[paste0(area[j, 2],".", k+3)]], FontStyle)
    } else{
      setCellStyle(cells11[[paste0(area[j, 2],".", k+3)]], CellStyle0) 
    }
  }
} #RowType1: 7:27
for (i in 1:21) {
  colorBlock11(i, WeekNumber)
}

##### SOS henk riittävyys (2) #####
sheet12 <- sheets[["SOS henk riittävyys (2)"]]                  # get specific sheet
rows12   <- xlsx::getRows(sheet12)         # get rows
cells12  <- xlsx::getCells(rows12)         # get cells

colorBlock12 <- function(j, WeekNum){
  
  sheetVar <- c("SOS.päihdet.hr", "SOS.omaishoit.hr")
  k = 0
  
  setCellValue(cells12[[paste0(area[j, 2],".", 3)]], WeekNum)
  setCellStyle(cells12[[paste0(area[j, 2],".", 3)]], FontStyle1)
  
  for (i in c(sheetVar)) {
    
    if(grepl("PTH", i, fixed = TRUE)){
      source <- PTH_combined_new  
    } else{
      source <- SOS_combined_new
    }
    
    count <- source[["combined_new"]] %>% 
      group_by(Viikko, Sairaanhoitopiiri, get(i)) %>% 
      dplyr::count() %>% 
      setNames(., c("Viikko", "Sairaanhoitopiiri", i, "n")) %>% 
      filter(get(i) %in% c(1:4), 
             Viikko == WeekNum,
             Sairaanhoitopiiri == area[j, 1])
    
    k = k + 5
    
    if(1 %in% count[[i]]){
      setCellStyle(cells12[[paste0(area[j, 2],".", k)]], CellStyle1)
    } else {
      setCellStyle(cells12[[paste0(area[j, 2],".", k)]], CellStyle0)
    }
    
    if(2 %in% count[[i]]){
      setCellStyle(cells12[[paste0(area[j, 2],".", k+1)]], CellStyle2)
    } else {
      setCellStyle(cells12[[paste0(area[j, 2],".", k+1)]], CellStyle0)
    }
    
    if(3 %in% count[[i]]){
      setCellStyle(cells12[[paste0(area[j, 2],".", k+2)]], CellStyle3)
    } else{
      setCellStyle(cells12[[paste0(area[j, 2],".", k+2)]], CellStyle0) 
    }
    
    if(4 %in% count[[i]]){
      setCellValue(cells12[[paste0(area[j, 2],".", k+3)]], count[count[[i]] == 4, ]$n)
      setCellStyle(cells12[[paste0(area[j, 2],".", k+3)]], FontStyle)
    } else{
      setCellStyle(cells12[[paste0(area[j, 2],".", k+3)]], CellStyle0) 
    }
  }
} #RowType1: 7:27
for (i in 1:21) {
  colorBlock12(i, WeekNumber)
}

##### ESH palv saatavuus #####
sheet14 <- sheets[["ESH palv saatavuus"]]                  # get specific sheet
rows14   <- xlsx::getRows(sheet14)         # get rows
cells14  <- xlsx::getCells(rows14)         # get cells

colorBlock14 <- function(j, WeekNum){
  
  sheetVar <- c("esh_service_availability_option_1", "esh_service_availability_option_2",
                "esh_service_availability_option_3", "esh_service_availability_option_4",
                "esh_service_availability_option_5", "ESH.psykiatria.ps")
  k = 0
  
  setCellValue(cells14[[paste0(area[j, 2],".", 3)]], WeekNum)
  setCellStyle(cells14[[paste0(area[j, 2],".", 3)]], FontStyle1)
  
  for (i in c(sheetVar)) {
    
    source <- ESH_combined_new
    
    count <- source[["combined_new"]] %>% 
      group_by(Viikko, Sairaanhoitopiiri, get(i)) %>% 
      dplyr::count() %>% 
      setNames(., c("Viikko", "Sairaanhoitopiiri", i, "n")) %>% 
      filter(get(i) %in% c(1:4), 
             Viikko == WeekNum,
             Sairaanhoitopiiri == area[j, 1])
    
    k = k + 5
    
    if(1 %in% count[[i]]){
      setCellStyle(cells14[[paste0(area[j, 2],".", k)]], CellStyle1)
    } else {
      setCellStyle(cells14[[paste0(area[j, 2],".", k)]], CellStyle0)
    }
    
    if(2 %in% count[[i]]){
      setCellStyle(cells14[[paste0(area[j, 2],".", k+1)]], CellStyle2)
    } else {
      setCellStyle(cells14[[paste0(area[j, 2],".", k+1)]], CellStyle0)
    }
    
    if(3 %in% count[[i]]){
      setCellStyle(cells14[[paste0(area[j, 2],".", k+2)]], CellStyle3)
    } else{
      setCellStyle(cells14[[paste0(area[j, 2],".", k+2)]], CellStyle0) 
    }
    
    if(4 %in% count[[i]]){
      setCellValue(cells14[[paste0(area[j, 2],".", k+3)]], count[count[[i]] == 4, ]$n)
      setCellStyle(cells14[[paste0(area[j, 2],".", k+3)]], FontStyle)
    } else{
      setCellStyle(cells14[[paste0(area[j, 2],".", k+3)]], CellStyle0) 
    }
  }
} #RowType3: 7:27
for (i in 1:21) {
  colorBlock14(i, WeekNumber)
}

##### ESH henk riittävyys #####
sheet15 <- sheets[["ESH henk riittävyys"]]                  # get specific sheet
rows15   <- xlsx::getRows(sheet15)         # get rows
cells15  <- xlsx::getCells(rows15)         # get cells

colorBlock15 <- function(j, WeekNum){
  
  sheetVar <- c("ESH.som.pkl.hr", "ESH.som.vos.hr", "ESH.psyk.pkl.hr", "ESH.psyk.vos.hr", "ESH.teho.vos.hr")
  k = 0
  
  setCellValue(cells15[[paste0(area[j, 2],".", 3)]], WeekNum)
  setCellStyle(cells15[[paste0(area[j, 2],".", 3)]], FontStyle1)
  
  for (i in c(sheetVar)) {
    
    if(grepl("PTH", i, fixed = TRUE)){
      source <- PTH_combined_new  
    } else if (grepl("SOS", i, fixed = TRUE)){
      source <- SOS_combined_new
    } else{
      source <- ESH_combined_new
    }
    
    count <- source[["combined_new"]] %>% 
      group_by(Viikko, Sairaanhoitopiiri, get(i)) %>% 
      dplyr::count() %>% 
      setNames(., c("Viikko", "Sairaanhoitopiiri", i, "n")) %>% 
      filter(get(i) %in% c(1:4), 
             Viikko == WeekNum,
             Sairaanhoitopiiri == area[j, 1])
    
    k = k + 5
    
    if(1 %in% count[[i]]){
      setCellStyle(cells15[[paste0(area[j, 2],".", k)]], CellStyle1)
    } else {
      setCellStyle(cells15[[paste0(area[j, 2],".", k)]], CellStyle0)
    }
    
    if(2 %in% count[[i]]){
      setCellStyle(cells15[[paste0(area[j, 2],".", k+1)]], CellStyle2)
    } else {
      setCellStyle(cells15[[paste0(area[j, 2],".", k+1)]], CellStyle0)
    }
    
    if(3 %in% count[[i]]){
      setCellStyle(cells15[[paste0(area[j, 2],".", k+2)]], CellStyle3)
    } else{
      setCellStyle(cells15[[paste0(area[j, 2],".", k+2)]], CellStyle0) 
    }
    
    if(4 %in% count[[i]]){
      setCellValue(cells15[[paste0(area[j, 2],".", k+3)]], count[count[[i]] == 4, ]$n)
      setCellStyle(cells15[[paste0(area[j, 2],".", k+3)]], FontStyle)
    } else{
      setCellStyle(cells15[[paste0(area[j, 2],".", k+3)]], CellStyle0) 
    }
  }
} #RowType3: 7:27
for (i in 1:21) {
  colorBlock15(i, WeekNumber)
}

##### PTH palv saatav varm #####
sheet4 <- sheets[["PTH palv saatav varm"]]                  # get specific sheet
rows4   <- xlsx::getRows(sheet4)         # get rows
cells4  <- xlsx::getCells(rows4)         # get cells

colorBlock4 <- function(j, WeekNum){
  
  sheetVar <- c("PTH.kiirsupist.ps","PTH.toimuud.ps","PTH.etäplis.ps","PTH.ostop.ps","PTH.muuten.ps","PTH.tmpriittäv.ps")
  k = 0
  
  setCellValue(cells4[[paste0(area[j, 5],".", 3)]], WeekNum)
  setCellStyle(cells4[[paste0(area[j, 5],".", 3)]], FontStyle1)
  
  for (i in c(sheetVar)) {
    
    if(grepl("PTH", i, fixed = TRUE)){
      source <- PTH_combined_new  
    } else if (grepl("SOS", i, fixed = TRUE)){
      source <- SOS_combined_new
    } else{
      source <- ESH_combined_new
    }
    
    count <- source[["combined_new"]] %>% 
      group_by(Viikko, Sairaanhoitopiiri, get(i)) %>% 
      dplyr::count() %>% 
      setNames(., c("Viikko", "Sairaanhoitopiiri", i, "n")) %>% 
      filter(get(i) %in% c(1:4), 
             Viikko == WeekNum,
             Sairaanhoitopiiri == area[j, 1])
    
    k = k + 5
    
    if(i != "PTH.tmpriittäv.ps"){
      
      if(1 %in% count[[i]]){
        CellBlock <- CellBlock(sheet4, area[j, 5], k, 1, 3, create = F)
        CB.setRowData(
          CellBlock,
          "???",
          1,
          colOffset = 0,
          showNA = F,
          rowStyle = headerStyle
        )
      } else {
        CellBlock <- CellBlock(sheet4, area[j, 5], k, 1, 3, create = F)
        CB.setRowData(
          CellBlock,
          " ",
          1,
          colOffset = 0,
          showNA = F,
          rowStyle = headerStyle
        )
      }
      
    } else{
      
      if(1 %in% count[[i]]){
        CellBlock <- CellBlock(sheet4, area[j, 5], k, 1, 3, create = F)
        CB.setRowData(
          CellBlock,
          "",
          1,
          colOffset = 0,
          showNA = F,
          rowStyle = BlockStyle1
        )
      } else {
        CellBlock <- CellBlock(sheet4, area[j, 5], k, 1, 3, create = F)
        CB.setRowData(
          CellBlock,
          " ",
          1,
          colOffset = 0,
          showNA = F,
          rowStyle = BlockStyle0
        )
      }
      
      if(2 %in% count[[i]]){
        CellBlock <- CellBlock(sheet4, area[j, 5], k+5, 1, 3, create = F)
        CB.setRowData(
          CellBlock,
          "",
          1,
          colOffset = 0,
          showNA = F,
          rowStyle = BlockStyle2
        )
      } else {
        CellBlock <- CellBlock(sheet4, area[j, 5], k+5, 1, 3, create = F)
        CB.setRowData(
          CellBlock,
          " ",
          1,
          colOffset = 0,
          showNA = F,
          rowStyle = BlockStyle0
        )
      }
      
      if(3 %in% count[[i]]){
        CellBlock <- CellBlock(sheet4, area[j, 5], k+10, 1, 3, create = F)
        CB.setRowData(
          CellBlock,
          "",
          1,
          colOffset = 0,
          showNA = F,
          rowStyle = BlockStyle3
        )
      } else {
        CellBlock <- CellBlock(sheet4, area[j, 5], k+10, 1, 3, create = F)
        CB.setRowData(
          CellBlock,
          " ",
          1,
          colOffset = 0,
          showNA = F,
          rowStyle = BlockStyle0
        )
      }
      
    }
  }
}
for (i in 1:21) {
  colorBlock4(i, WeekNumber)
}

##### PTH henk riittäv varm #####
sheet7 <- sheets[["PTH henk riittäv varm"]]                  # get specific sheet
rows7   <- xlsx::getRows(sheet7)         # get rows
cells7  <- xlsx::getCells(rows7)         # get cells

colorBlock7 <- function(j, WeekNum){
  
  sheetVar <- c("PTH.lainsääd.hr","PTH.henksiir.hr","PTH.henkrekry.hr",
                "PTH.neuvot.hr","PTH.ostop.hr", "pth_personnel_sufficiency_situation_manage_6", "PTH.muuten.hr","PTH.tmpriittäv.hr")
  k = 0
  
  setCellValue(cells7[[paste0(area[j, 5],".", 3)]], WeekNum)
  setCellStyle(cells7[[paste0(area[j, 5],".", 3)]], FontStyle1)
  
  for (i in c(sheetVar)) {
    
    source <- PTH_combined_new
    
    count <- source[["combined_new"]] %>% 
      group_by(Viikko, Sairaanhoitopiiri, get(i)) %>% 
      dplyr::count() %>% 
      setNames(., c("Viikko", "Sairaanhoitopiiri", i, "n")) %>% 
      filter(get(i) %in% c(1:4), 
             Viikko == WeekNum,
             Sairaanhoitopiiri == area[j, 1])
    
    k = k + 5
    
    if(i != "PTH.tmpriittäv.hr"){
      
      if(1 %in% count[[i]]){
        CellBlock <- CellBlock(sheet7, area[j, 5], k, 1, 3, create = F)
        CB.setRowData(
          CellBlock,
          "???",
          1,
          colOffset = 0,
          showNA = F,
          rowStyle = headerStyle
        )
      } else {
        CellBlock <- CellBlock(sheet7, area[j, 5], k, 1, 3, create = F)
        CB.setRowData(
          CellBlock,
          " ",
          1,
          colOffset = 0,
          showNA = F,
          rowStyle = headerStyle
        )
      }
      
    } else{
      
      if(1 %in% count[[i]]){
        CellBlock <- CellBlock(sheet7, area[j, 5], k, 1, 3, create = F)
        CB.setRowData(
          CellBlock,
          "",
          1,
          colOffset = 0,
          showNA = F,
          rowStyle = BlockStyle1
        )
      } else {
        CellBlock <- CellBlock(sheet7, area[j, 5], k, 1, 3, create = F)
        CB.setRowData(
          CellBlock,
          " ",
          1,
          colOffset = 0,
          showNA = F,
          rowStyle = BlockStyle0
        )
      }
      
      if(2 %in% count[[i]]){
        CellBlock <- CellBlock(sheet7, area[j, 5], k+5, 1, 3, create = F)
        CB.setRowData(
          CellBlock,
          "",
          1,
          colOffset = 0,
          showNA = F,
          rowStyle = BlockStyle2 
        )
      } else {
        CellBlock <- CellBlock(sheet7, area[j, 5], k+5, 1, 3, create = F)
        CB.setRowData(
          CellBlock,
          " ",
          1,
          colOffset = 0,
          showNA = F,
          rowStyle = BlockStyle0 
        )
      }
      
      if(3 %in% count[[i]]){
        CellBlock <- CellBlock(sheet7, area[j, 5], k+10, 1, 3, create = F)
        CB.setRowData(
          CellBlock,
          "",
          1,
          colOffset = 0,
          showNA = F,
          rowStyle = BlockStyle3
        )
      } else {
        CellBlock <- CellBlock(sheet7, area[j, 5], k+10, 1, 3, create = F)
        CB.setRowData(
          CellBlock,
          " ",
          1,
          colOffset = 0,
          showNA = F,
          rowStyle = BlockStyle0 
        )
      }
      
    }
  }
}
for (i in 1:21) {
  colorBlock7(i, WeekNumber)
}

##### SOS palv saatav varm #####
sheet10 <- sheets[["SOS palv saatav varm"]]                  # get specific sheet
rows10   <- xlsx::getRows(sheet10)         # get rows
cells10  <- xlsx::getCells(rows10)         # get cells

colorBlock10 <- function(j, WeekNum){
  
  sheetVar <- c("SOS.kiiresup.ps","SOS.toimuud.ps","SOS.etäp.ps","SOS.ostop.ps","SOS.muuten.ps","SOS.riittäv.ps")
  k = 0
  
  setCellValue(cells10[[paste0(area[j, 5],".", 3)]], WeekNum)
  setCellStyle(cells10[[paste0(area[j, 5],".", 3)]], FontStyle1)
  
  for (i in c(sheetVar)) {
    
    if(grepl("PTH", i, fixed = TRUE)){
      source <- PTH_combined_new  
    } else if (grepl("SOS", i, fixed = TRUE)){
      source <- SOS_combined_new
    } else{
      source <- ESH_combined_new
    }
    
    count <- source[["combined_new"]] %>% 
      group_by(Viikko, Sairaanhoitopiiri, get(i)) %>% 
      dplyr::count() %>% 
      setNames(., c("Viikko", "Sairaanhoitopiiri", i, "n")) %>% 
      filter(get(i) %in% c(1:4), 
             Viikko == WeekNum,
             Sairaanhoitopiiri == area[j, 1])
    
    k = k + 5
    
    if(i != "SOS.riittäv.ps"){
      
      if(1 %in% count[[i]]){
        CellBlock <- CellBlock(sheet10, area[j, 5], k, 1, 3, create = F)
        CB.setRowData(
          CellBlock,
          "???",
          1,
          colOffset = 0,
          showNA = F,
          rowStyle = headerStyle
        )
      } else {
        CellBlock <- CellBlock(sheet10, area[j, 5], k, 1, 3, create = F)
        CB.setRowData(
          CellBlock,
          " ",
          1,
          colOffset = 0,
          showNA = F,
          rowStyle = headerStyle
        )
      }
      
    } else{
      
      if(1 %in% count[[i]]){
        CellBlock <- CellBlock(sheet10, area[j, 5], k, 1, 3, create = F)
        CB.setRowData(
          CellBlock,
          "",
          1,
          colOffset = 0,
          showNA = F,
          rowStyle = BlockStyle1
        )
      } else {
        CellBlock <- CellBlock(sheet10, area[j, 5], k, 1, 3, create = F)
        CB.setRowData(
          CellBlock,
          " ",
          1,
          colOffset = 0,
          showNA = F,
          rowStyle = BlockStyle0
        )
      }
      
      if(2 %in% count[[i]]){
        CellBlock <- CellBlock(sheet10, area[j, 5], k+5, 1, 3, create = F)
        CB.setRowData(
          CellBlock,
          "",
          1,
          colOffset = 0,
          showNA = F,
          rowStyle = BlockStyle2
        )
      } else {
        CellBlock <- CellBlock(sheet10, area[j, 5], k+5, 1, 3, create = F)
        CB.setRowData(
          CellBlock,
          " ",
          1,
          colOffset = 0,
          showNA = F,
          rowStyle = BlockStyle0 
        )
      }
      
      if(3 %in% count[[i]]){
        CellBlock <- CellBlock(sheet10, area[j, 5], k+10, 1, 3, create = F)
        CB.setRowData(
          CellBlock,
          "",
          1,
          colOffset = 0,
          showNA = F,
          rowStyle = BlockStyle3
        )
      } else {
        CellBlock <- CellBlock(sheet10, area[j, 5], k+10, 1, 3, create = F)
        CB.setRowData(
          CellBlock,
          " ",
          1,
          colOffset = 0,
          showNA = F,
          rowStyle = BlockStyle0
        )
      }
      
    }
  }
}
for (i in 1:21) {
  colorBlock10(i, WeekNumber)
}

##### SOS henk riittäv varm #####
sheet13 <- sheets[["SOS henk riittäv varm"]]                  # get specific sheet
rows13   <- xlsx::getRows(sheet13)         # get rows
cells13  <- xlsx::getCells(rows13)         # get cells

colorBlock13 <- function(j, WeekNum){
  
  sheetVar <- c("SOS.lainsääd.hr","SOS.henksiir.hr","SOS.henkrekryt.hr","SOS.neuv.hr","SOS.ostop.hr",
                "sos_personnel_sufficiency_situation_manage_6", "SOS.muuten.hr","SOS.tmp.riittäv.hr")
  k = 0
  
  setCellValue(cells13[[paste0(area[j, 5],".", 3)]], WeekNum)
  setCellStyle(cells13[[paste0(area[j, 5],".", 3)]], FontStyle1)
  
  for (i in c(sheetVar)) {
    
    source <- SOS_combined_new
    
    count <- source[["combined_new"]] %>% 
      group_by(Viikko, Sairaanhoitopiiri, get(i)) %>% 
      dplyr::count() %>% 
      setNames(., c("Viikko", "Sairaanhoitopiiri", i, "n")) %>% 
      filter(get(i) %in% c(1:4), 
             Viikko == WeekNum,
             Sairaanhoitopiiri == area[j, 1])
    
    k = k + 5
    
    if(i != "SOS.tmp.riittäv.hr"){
      
      if(1 %in% count[[i]]){
        CellBlock <- CellBlock(sheet13, area[j, 5], k, 1, 3, create = F)
        CB.setRowData(
          CellBlock,
          "???",
          1,
          colOffset = 0,
          showNA = F,
          rowStyle = headerStyle 
        )
      } else {
        CellBlock <- CellBlock(sheet13, area[j, 5], k, 1, 3, create = F)
        CB.setRowData(
          CellBlock,
          " ",
          1,
          colOffset = 0,
          showNA = F,
          rowStyle = headerStyle 
        )
      }
      
    } else{
      
      if(1 %in% count[[i]]){
        CellBlock <- CellBlock(sheet13, area[j, 5], k, 1, 3, create = F)
        CB.setRowData(
          CellBlock,
          "",
          1,
          colOffset = 0,
          showNA = F,
          rowStyle = BlockStyle1
        )
      } else {
        CellBlock <- CellBlock(sheet13, area[j, 5], k, 1, 3, create = F)
        CB.setRowData(
          CellBlock,
          "",
          1,
          colOffset = 0,
          showNA = F,
          rowStyle = BlockStyle0 
        )
      }
      
      if(2 %in% count[[i]]){
        CellBlock <- CellBlock(sheet13, area[j, 5], k+5, 1, 3, create = F)
        CB.setRowData(
          CellBlock,
          "",
          1,
          colOffset = 0,
          showNA = F,
          rowStyle = BlockStyle2 
        )
      } else {
        CellBlock <- CellBlock(sheet13, area[j, 5], k+5, 1, 3, create = F)
        CB.setRowData(
          CellBlock,
          "",
          1,
          colOffset = 0,
          showNA = F,
          rowStyle = BlockStyle0
        )
      }
      
      if(3 %in% count[[i]]){
        CellBlock <- CellBlock(sheet13, area[j, 5], k+10, 1, 3, create = F)
        CB.setRowData(
          CellBlock,
          "",
          1,
          colOffset = 0,
          showNA = F,
          rowStyle = BlockStyle3
        )
      } else {
        CellBlock <- CellBlock(sheet13, area[j, 5], k+10, 1, 3, create = F)
        CB.setRowData(
          CellBlock,
          " ",
          1,
          colOffset = 0,
          showNA = F,
          rowStyle = BlockStyle0
        )
      }
      
    }
  }
}
for (i in 1:21) {
  colorBlock13(i, WeekNumber)
}

##### ESH henk riittäv varm #####
sheet16 <- sheets[["ESH henk riittäv varm"]]                  # get specific sheet
rows16   <- xlsx::getRows(sheet16)         # get rows
cells16  <- xlsx::getCells(rows16)         # get cells

colorBlock16 <- function(j, WeekNum){
  
  sheetVar <- c("ESH.lainsääd.hr","ESH.henksiir.hr","ESH.henkrekryt.hr","ESH.neuv.hr","ESH.ostop.hr",
                "esh_personnel_sufficiency_situation_manage_6", "ESH.muuten.hr","ESH.tmp.riittäv.hr")
  k = 0
  
  setCellValue(cells16[[paste0(area[j, 5],".", 3)]], WeekNum)
  setCellStyle(cells16[[paste0(area[j, 5],".", 3)]], FontStyle1)
  
  for (i in c(sheetVar)) {
    
    source <- ESH_combined_new
    
    count <- source[["combined_new"]] %>% 
      group_by(Viikko, Sairaanhoitopiiri, get(i)) %>% 
      dplyr::count() %>% 
      setNames(., c("Viikko", "Sairaanhoitopiiri", i, "n")) %>% 
      filter(get(i) %in% c(1:4), 
             Viikko == WeekNum,
             Sairaanhoitopiiri == area[j, 1])
    
    k = k + 5
    
    if(i != "ESH.tmp.riittäv.hr"){
      
      if(1 %in% count[[i]]){
        CellBlock <- CellBlock(sheet16, area[j, 5], k, 1, 3, create = F)
        CB.setRowData(
          CellBlock,
          "???",
          1,
          colOffset = 0,
          showNA = F,
          rowStyle = headerStyle
        )
      } else {
        CellBlock <- CellBlock(sheet16, area[j, 5], k, 1, 3, create = F)
        CB.setRowData(
          CellBlock,
          " ",
          1,
          colOffset = 0,
          showNA = F,
          rowStyle = headerStyle
        )
      }
      
    } else{
      
      if(1 %in% count[[i]]){
        CellBlock <- CellBlock(sheet16, area[j, 5], k, 1, 3, create = F)
        CB.setRowData(
          CellBlock,
          "",
          1,
          colOffset = 0,
          showNA = F,
          rowStyle = BlockStyle1
        )
      } else {
        CellBlock <- CellBlock(sheet16, area[j, 5], k, 1, 3, create = F)
        CB.setRowData(
          CellBlock,
          " ",
          1,
          colOffset = 0,
          showNA = F,
          rowStyle = BlockStyle0
        )
      }
      
      if(2 %in% count[[i]]){
        CellBlock <- CellBlock(sheet16, area[j, 5], k+5, 1, 3, create = F)
        CB.setRowData(
          CellBlock,
          "",
          1,
          colOffset = 0,
          showNA = F,
          rowStyle = BlockStyle2
        )
      } else {
        CellBlock <- CellBlock(sheet16, area[j, 5], k+5, 1, 3, create = F)
        CB.setRowData(
          CellBlock,
          " ",
          1,
          colOffset = 0,
          showNA = F,
          rowStyle = BlockStyle0
        )
      }
      
      if(3 %in% count[[i]]){
        CellBlock <- CellBlock(sheet16, area[j, 5], k+10, 1, 3, create = F)
        CB.setRowData(
          CellBlock,
          "",
          1,
          colOffset = 0,
          showNA = F,
          rowStyle = BlockStyle3
        )
      } else {
        CellBlock <- CellBlock(sheet16, area[j, 5], k+10, 1, 3, create = F)
        CB.setRowData(
          CellBlock,
          " ",
          1,
          colOffset = 0,
          showNA = F,
          rowStyle = BlockStyle0
        )
      }
      
    }
  }
}
for (i in 1:21) {
  colorBlock16(i, WeekNumber)
}

##### Vastaajat #####
count_PTH <- PTH_combined_new[["combined_new"]] %>% 
  group_by(Viikko, Sairaanhoitopiiri, Järjestäjäorganisaatio) %>% 
  dplyr::count() %>% 
  filter(Viikko == WeekNumber) %>% 
  group_by(Sairaanhoitopiiri) %>% 
  dplyr::count() 
count_PTH$Sairaanhoitopiiri <- as.factor(count_PTH$Sairaanhoitopiiri)
count_PTH$Sairaanhoitopiiri <- factor(count_PTH$Sairaanhoitopiiri, levels = area$area)
count_PTH <- count_PTH %>% 
  arrange(Sairaanhoitopiiri)
N_PTH <- sum(count_PTH$n)

count_SOS <- SOS_combined_new[["combined_new"]] %>% 
  group_by(Viikko, Sairaanhoitopiiri, Järjestäjäorganisaatio) %>% 
  dplyr::count() %>% 
  filter(Viikko == WeekNumber) %>% 
  group_by(Sairaanhoitopiiri) %>% 
  dplyr::count()
count_SOS$Sairaanhoitopiiri <- as.factor(count_SOS$Sairaanhoitopiiri)
count_SOS$Sairaanhoitopiiri <- factor(count_SOS$Sairaanhoitopiiri, levels = area$area)
count_SOS <- count_SOS %>% 
  arrange(Sairaanhoitopiiri)
N_SOS <- sum(count_SOS$n)

count_ESH <- ESH_combined_new[["combined_new"]] %>% 
  group_by(Viikko, Sairaanhoitopiiri, Järjestäjäorganisaatio) %>% 
  dplyr::count() %>% 
  filter(Viikko == WeekNumber) %>% 
  group_by(Sairaanhoitopiiri) %>% 
  dplyr::count() 
count_ESH$Sairaanhoitopiiri <- as.factor(count_ESH$Sairaanhoitopiiri)
count_ESH$Sairaanhoitopiiri <- factor(count_ESH$Sairaanhoitopiiri, levels = area$area)
count_ESH <- count_ESH %>% 
  arrange(Sairaanhoitopiiri)
N_ESH <- sum(count_ESH$n)

sheet17 <- sheets[["Vastaajat"]]                  # get specific sheet
rows17   <- xlsx::getRows(sheet17)         # get rows
cells17  <- xlsx::getCells(rows17)         # get cells

colorBlock17 <- function(j, WeekNum){
  
  setCellValue(cells17[[paste0(area[j, 3],".", 3)]], WeekNum)
  setCellStyle(cells17[[paste0(area[j, 3],".", 3)]], FontStyle1)
  
  # PTH
  CellBlock <- CellBlock(sheet17, 8, 5, 1, 4, create = F)
  CB.setRowData(
    CellBlock,
    N_PTH,
    1,
    colOffset = 0,
    showNA = F,
    rowStyle = BlockStyle5)
  
  if(area[j, 1] %in% count_PTH$Sairaanhoitopiiri){
    CellBlock <- CellBlock(sheet17, area[j, 3], 5, 1, 4, create = F)
    CB.setRowData(
      CellBlock,
      as.numeric(count_PTH[count_PTH$Sairaanhoitopiiri == area[j, 1], "n"]),
      1,
      colOffset = 0,
      showNA = F,
      rowStyle = BlockStyle4
    )
  } else {
    CellBlock <- CellBlock(sheet17, area[j, 3], 5, 1, 4, create = F)
    CB.setRowData(
      CellBlock,
      "0",
      1,
      colOffset = 0,
      showNA = F,
      rowStyle = BlockStyle4
    )
  }
  # SOS
  CellBlock <- CellBlock(sheet17, 8, 15, 1, 4, create = F)
  CB.setRowData(
    CellBlock,
    N_SOS,
    1,
    colOffset = 0,
    showNA = F,
    rowStyle = BlockStyle5)
  
  if(area[j, 1] %in% count_SOS$Sairaanhoitopiiri){
    CellBlock <- CellBlock(sheet17, area[j, 3], 15, 1, 4, create = F)
    CB.setRowData(
      CellBlock,
      as.numeric(count_SOS[count_SOS$Sairaanhoitopiiri == area[j, 1], "n"]),
      1,
      colOffset = 0,
      showNA = F,
      rowStyle = BlockStyle4
    )
  } else {
    CellBlock <- CellBlock(sheet17, area[j, 3], 15, 1, 4, create = F)
    CB.setRowData(
      CellBlock,
      0,
      1,
      colOffset = 0,
      showNA = F,
      rowStyle = BlockStyle4
    )
  }
  
  #ESH
  CellBlock <- CellBlock(sheet17, 8, 25, 1, 4, create = F)
  CB.setRowData(
    CellBlock,
    N_ESH,
    1,
    colOffset = 0,
    showNA = F,
    rowStyle = BlockStyle5)
  
  if(area[j, 1] %in% count_ESH$Sairaanhoitopiiri){
    CellBlock <- CellBlock(sheet17, area[j, 3], 25, 1, 4, create = F)
    CB.setRowData(
      CellBlock,
      as.numeric(count_ESH[count_ESH$Sairaanhoitopiiri == area[j, 1], "n"]),
      1,
      colOffset = 0,
      showNA = F,
      rowStyle = BlockStyle4
    )
  } else {
    CellBlock <- CellBlock(sheet17, area[j, 3], 25, 1, 4, create = F)
    CB.setRowData(
      CellBlock,
      0,
      1,
      colOffset = 0,
      showNA = F,
      rowStyle = BlockStyle4
    )
  }
}

for (i in 1:21) {
  colorBlock17(i, WeekNumber)
}

##### PTH henk riittävyys tote #####
sheet18 <- sheets[["PTH henk riittävyys tote"]]                  # get specific sheet
rows18   <- xlsx::getRows(sheet18)         # get rows
cells18  <- xlsx::getCells(rows18)         # get cells

colorBlock18 <- function(j, WeekNum){
  
  sheetVar <- c("pth_protection_work_option_1", "pth_protection_work_option_2", "pth_protection_work_option_3", 
                "pth_protection_work_option_4", "pth_protection_work_option_5", "pth_protection_work_option_6", 
                "pth_protection_work_option_7", "pth_protection_work_option_8", "pth_protection_work_option_9", 
                "pth_protection_work_option_10", "pth_protection_work_option_11")
  k = 0
  
  setCellValue(cells18[[paste0(area[j, 2],".", 3)]], WeekNum)
  setCellStyle(cells18[[paste0(area[j, 2],".", 3)]], FontStyle1)
  
  for (i in c(sheetVar)) {
    
    source <- PTH_combined_new  
    
    count <- source[["combined_new"]] %>% 
      group_by(Viikko, Sairaanhoitopiiri, get(i)) %>% 
      dplyr::count() %>% 
      setNames(., c("Viikko", "Sairaanhoitopiiri", i, "n")) %>% 
      filter(get(i) %in% c(1:3, 9, "1", "2", "3", "9"), 
             Viikko == WeekNum,
             Sairaanhoitopiiri == area[j, 1])
    
    k = k + 5
    
    if(1 %in% count[[i]]){
      setCellValue(cells18[[paste0(area[j, 2],".", k)]], "X")
      setCellStyle(cells18[[paste0(area[j, 2],".", k)]], FontStyle2)
    } else {
      setCellStyle(cells18[[paste0(area[j, 2],".", k)]], CellStyle0)
    }
    
    if(2 %in% count[[i]]){
      setCellValue(cells18[[paste0(area[j, 2],".", k+1)]], "X")
      setCellStyle(cells18[[paste0(area[j, 2],".", k+1)]], FontStyle2)
    } else {
      setCellStyle(cells18[[paste0(area[j, 2],".", k+1)]], CellStyle0)
    }
    
    if(3 %in% count[[i]]){
      setCellValue(cells18[[paste0(area[j, 2],".", k+2)]], "X")
      setCellStyle(cells18[[paste0(area[j, 2],".", k+2)]], FontStyle2)
    } else{
      setCellStyle(cells18[[paste0(area[j, 2],".", k+2)]], CellStyle0) 
    }
    
    if(9 %in% count[[i]]){
      setCellValue(cells18[[paste0(area[j, 2],".", k+3)]], "X")
      setCellStyle(cells18[[paste0(area[j, 2],".", k+3)]], FontStyle2)
    } else{
      setCellStyle(cells18[[paste0(area[j, 2],".", k+3)]], CellStyle0) 
    }
  }
} #RowType1: 7:27
for (i in 1:21) {
  colorBlock18(i, WeekNumber)
}
##### SOS henk riittävyys tote #####
sheet19 <- sheets[["SOS henk riittävyys tote"]]                  # get specific sheet
rows19   <- xlsx::getRows(sheet19)         # get rows
cells19  <- xlsx::getCells(rows19)         # get cells

colorBlock19 <- function(j, WeekNum){
  
  sheetVar <- c("sos_protection_work_option_1", "sos_protection_work_option_2", "sos_protection_work_option_3", 
                "sos_protection_work_option_4", "sos_protection_work_option_5", "sos_protection_work_option_6", 
                "sos_protection_work_option_7", "sos_protection_work_option_8", "sos_protection_work_option_9", 
                "sos_protection_work_option_10")
  k = 0
  
  setCellValue(cells19[[paste0(area[j, 2],".", 3)]], WeekNum)
  setCellStyle(cells19[[paste0(area[j, 2],".", 3)]], FontStyle1)
  
  for (i in c(sheetVar)) {
    
    source <- SOS_combined_new
    
    count <- source[["combined_new"]] %>% 
      group_by(Viikko, Sairaanhoitopiiri, get(i)) %>% 
      dplyr::count() %>% 
      setNames(., c("Viikko", "Sairaanhoitopiiri", i, "n")) %>% 
      filter(get(i) %in% c(1:3, 9, "1", "2", "3", "9"), 
             Viikko == WeekNum,
             Sairaanhoitopiiri == area[j, 1])
    
    k = k + 5
    
    if(1 %in% count[[i]]){
      setCellValue(cells19[[paste0(area[j, 2],".", k)]], "X")
      setCellStyle(cells19[[paste0(area[j, 2],".", k)]], FontStyle2)
    } else {
      setCellStyle(cells19[[paste0(area[j, 2],".", k)]], CellStyle0)
    }
    
    if(2 %in% count[[i]]){
      setCellValue(cells19[[paste0(area[j, 2],".", k+1)]], "X")
      setCellStyle(cells19[[paste0(area[j, 2],".", k+1)]], FontStyle2)
    } else {
      setCellStyle(cells19[[paste0(area[j, 2],".", k+1)]], CellStyle0)
    }
    
    if(3 %in% count[[i]]){
      setCellValue(cells19[[paste0(area[j, 2],".", k+2)]], "X")
      setCellStyle(cells19[[paste0(area[j, 2],".", k+2)]], FontStyle2)
    } else{
      setCellStyle(cells19[[paste0(area[j, 2],".", k+2)]], CellStyle0) 
    }
    
    if(9 %in% count[[i]]){
      setCellValue(cells19[[paste0(area[j, 2],".", k+3)]], "X")
      setCellStyle(cells19[[paste0(area[j, 2],".", k+3)]], FontStyle2)
    } else{
      setCellStyle(cells19[[paste0(area[j, 2],".", k+3)]], CellStyle0) 
    }
  }
} #RowType1: 7:27
for (i in 1:21) {
  colorBlock19(i, WeekNumber)
}

##### SOS henk riittävyys tote (2) #####
sheet20 <- sheets[["SOS henk riittävyys tote (2)"]]                  # get specific sheet
rows20   <- xlsx::getRows(sheet20)         # get rows
cells20  <- xlsx::getCells(rows20)         # get cells

colorBlock20 <- function(j, WeekNum){
  
  sheetVar <- c("sos_protection_work_option_11", "sos_protection_work_option_12")
  k = 0
  
  setCellValue(cells20[[paste0(area[j, 2],".", 3)]], WeekNum)
  setCellStyle(cells20[[paste0(area[j, 2],".", 3)]], FontStyle1)
  
  for (i in c(sheetVar)) {
    
    source <- SOS_combined_new
    
    count <- source[["combined_new"]] %>% 
      group_by(Viikko, Sairaanhoitopiiri, get(i)) %>% 
      dplyr::count() %>% 
      setNames(., c("Viikko", "Sairaanhoitopiiri", i, "n")) %>% 
      filter(get(i) %in% c(1:3, 9, "1", "2", "3", "9"), 
             Viikko == WeekNum,
             Sairaanhoitopiiri == area[j, 1])
    
    k = k + 5
    
    if(1 %in% count[[i]]){
      setCellValue(cells20[[paste0(area[j, 2],".", k)]], "X")
      setCellStyle(cells20[[paste0(area[j, 2],".", k)]], FontStyle2)
    } else {
      setCellStyle(cells20[[paste0(area[j, 2],".", k)]], CellStyle0)
    }
    
    if(2 %in% count[[i]]){
      setCellValue(cells20[[paste0(area[j, 2],".", k+1)]], "X")
      setCellStyle(cells20[[paste0(area[j, 2],".", k+1)]], FontStyle2)
    } else {
      setCellStyle(cells20[[paste0(area[j, 2],".", k+1)]], CellStyle0)
    }
    
    if(3 %in% count[[i]]){
      setCellValue(cells20[[paste0(area[j, 2],".", k+2)]], "X")
      setCellStyle(cells20[[paste0(area[j, 2],".", k+2)]], FontStyle2)
    } else{
      setCellStyle(cells20[[paste0(area[j, 2],".", k+2)]], CellStyle0) 
    }
    
    if(9 %in% count[[i]]){
      setCellValue(cells20[[paste0(area[j, 2],".", k+3)]], "X")
      setCellStyle(cells20[[paste0(area[j, 2],".", k+3)]], FontStyle2)
    } else{
      setCellStyle(cells20[[paste0(area[j, 2],".", k+3)]], CellStyle0) 
    }
  }
} #RowType1: 7:27
for (i in 1:21) {
  colorBlock20(i, WeekNumber)
}

##### ESH henk riittävyys tote #####
sheet21 <- sheets[["ESH henk riittävyys tote"]]                  # get specific sheet
rows21   <- xlsx::getRows(sheet21)         # get rows
cells21  <- xlsx::getCells(rows21)         # get cells

colorBlock21 <- function(j, WeekNum){
  
  sheetVar <- c("esh_protection_work_option_1", "esh_protection_work_option_2",
                "esh_protection_work_option_3", "esh_protection_work_option_4",
                "esh_protection_work_option_5", "esh_protection_work_option_12")
  k = 0
  
  setCellValue(cells21[[paste0(area[j, 2],".", 3)]], WeekNum)
  setCellStyle(cells21[[paste0(area[j, 2],".", 3)]], FontStyle1)
  
  for (i in c(sheetVar)) {
    
    source <- ESH_combined_new
    
    count <- source[["combined_new"]] %>% 
      group_by(Viikko, Sairaanhoitopiiri, get(i)) %>% 
      dplyr::count() %>% 
      setNames(., c("Viikko", "Sairaanhoitopiiri", i, "n")) %>% 
      filter(get(i) %in% c(1:3, 9, "1", "2", "3", "9"), 
             Viikko == WeekNum,
             Sairaanhoitopiiri == area[j, 1])
    
    k = k + 5
    
    if(1 %in% count[[i]]){
      setCellValue(cells21[[paste0(area[j, 2],".", k)]], "X")
      setCellStyle(cells21[[paste0(area[j, 2],".", k)]], FontStyle2)
    } else {
      setCellStyle(cells21[[paste0(area[j, 2],".", k)]], CellStyle0)
    }
    
    if(2 %in% count[[i]]){
      setCellValue(cells21[[paste0(area[j, 2],".", k+1)]], "X")
      setCellStyle(cells21[[paste0(area[j, 2],".", k+1)]], FontStyle2)
    } else {
      setCellStyle(cells21[[paste0(area[j, 2],".", k+1)]], CellStyle0)
    }
    
    if(3 %in% count[[i]]){
      setCellValue(cells21[[paste0(area[j, 2],".", k+2)]], "X")
      setCellStyle(cells21[[paste0(area[j, 2],".", k+2)]], FontStyle2)
    } else{
      setCellStyle(cells21[[paste0(area[j, 2],".", k+2)]], CellStyle0) 
    }
    
    if(9 %in% count[[i]]){
      setCellValue(cells21[[paste0(area[j, 2],".", k+3)]], "X")
      setCellStyle(cells21[[paste0(area[j, 2],".", k+3)]], FontStyle2)
    } else{
      setCellStyle(cells21[[paste0(area[j, 2],".", k+3)]], CellStyle0) 
    }
  }
} #RowType3: 7:27
for (i in 1:21) {
  colorBlock21(i, WeekNumber)
}

##### write workbook #####
saveWorkbook(wb, output_dir)

