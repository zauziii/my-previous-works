render_kooste <- function(){
  rmarkdown::render(
    "kooste3n4.Rmd",
    output_dir = "/home/zzhh/mnt/ZZHH/_Desktop/Tasks/Koronaraportointi/Word/0322",
    output_file = "kooste.docx"
  )
}

render_kooste()
