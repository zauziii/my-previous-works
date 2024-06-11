library(rmarkdown)

SHP$SHParea <- c("Keski-Pohjanmaa", "HUS", "Kymenlaakso", "Pohjois-Pohjanmaa", "Pohjois-Karjala", 
                 "Etelä-Karjala", "Kainuu", "Päijät-Häme", "Varsinais-Suomi", "Ahvenanmaa",
                 "Itä-Savo", "Lappi", "Keski-Suomi", "Kanta-Häme", "Pirkanmaa",
                 "Pohjois-Savo", "Etelä-Pohjanmaa", "Länsi-Pohja", "Satakunta", "Vaasa", "Etelä-Savo"
)
SHP$code <- c("17", "25", "08", "18", "12", 
              "09", "19", "07", "03", "00", 
              "11", "21", "14", "05", "06", 
              "13", "15", "20", "04", "16", "10")


render_report <- function(code, SHParea, date) {
  rmarkdown::render(
    "templ.Rmd", params = list(
      Sairaanhoitopiiri = SHParea,
      date = date),
    output_format = pdf_document(),
    output_dir = "/home/zzhh/mnt/ZZHH/_Desktop/Tasks/Koronaraportointi/PDF/0201",
    output_file = paste0(code, "_", SHParea, "_", date, ".pdf")
    )
}

for (i in 1:21) {
  render_report(code = SHP[i, 2], 
                SHParea = as.character(SHP[i, 1]), 
                date = "01.02.2022")
}
