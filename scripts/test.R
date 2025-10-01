library(haven)
epen2022 <- as.data.frame(read_sav(
  "/home/renato/Dropbox/PUBLIC/CURSO_CHILE/DIA02/datos/intro/EPEN 2022 BD_Publicación Dpto.SAV"
))
saveRDS(epen2022, "datos/intro/epen2022_bd_dpto.rds")
write.csv(
  epen2024,
  "/home/renato/Dropbox/PUBLIC/CURSO_CHILE/DIA02/datos/intro/EPEN 2022 BD_Publicación Dpto.csv"
)
library(arrow)
write_parquet(
  epen2024,
  "/home/renato/Dropbox/PUBLIC/CURSO_CHILE/DIA02/datos/intro/EPEN 2022 BD_Publicación Dpto.parquet"
)


for (i in 1:length(years)) {
  assign(
    paste("fdres", years[i], sep = ""),
    as.matrix(read.table(
      paste(getwd(), "/", years[i], "/fdresa-", years[i], ".csv", sep = ""),
      header = TRUE,
      dec = ".",
      sep = ",",
      row.names = 1
    ))
  )
  assign(
    paste("prodres", years[i], sep = ""),
    as.matrix(read.table(
      paste(getwd(), "/", years[i], "/prodresa-", years[i], ".csv", sep = ""),
      header = TRUE,
      dec = ".",
      sep = ",",
      row.names = 1
    ))
  )
}
