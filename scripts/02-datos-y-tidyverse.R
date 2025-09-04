# Limpiar el área de trabajo
rm(list = ls())

# Importar datos

# Valores separados por comas (CSV)

# El archivo fue guardado desde Excel con
# locale chileno sin separadores de miles
empleo_csv <- read.csv(
  "datos/intro/empleo_chl.csv",
  sep = ";",
  dec = ",",
  encoding = "UTF-8"
)

# Excel
# llamamos librería
library(readxl)

equivalencias_empleo_chl <- read_excel(
  "datos/intro/Excel_empleo_chl.xlsx",
  sheet = "Equivalencia"
)

# SPSS o STATA
# library(haven)

# # Datos de Perú
# epen2024 <- read_sav("datos/intro/EPEN 2022 BD_Publicación Dpto.SAV")

# Datos Geográficos (no llamamos las librerías)

# # Leer el dato geográfico
# adm1 <- sf::read_sf("datos/intro/gis/chl_adm1.shp")

# ggplot2::ggplot(adm1) +
#   geom_sf()

## Tidyverse (dplyr, tidyr, stringr, etc.)
library(tidyverse)

# Ejemplo de un verbo (Seleccionar)
equivalencias_empleo_chl <- equivalencias_empleo_chl |>
  select(!`Ocupados (miles)`)


# El "tubo" |>
# También se puede %>%

empleo_xlsx <- read_excel(
  "datos/intro/Excel_empleo_chl.xlsx",
  sheet = "Empleo Regional"
) |>
  mutate(`Código Región` = str_extract(Región, "(?<=\\()[A-Z]{2}(?=\\))")) |>
  select(!`Población ocupada (total)`) |>
  pivot_longer(
    cols = c(3:24),
    names_to = "Actividad",
    values_to = "Valor"
  ) |>
  left_join(
    equivalencias_empleo_chl,
    join_by(Actividad)
  ) |>
  filter(`Año` == 2022) |>
  group_by(
    `Código Región`,
    `Región`,
    `Correlativo MIP12`,
    `Nombre MIP`,
    `Código MIP`
  ) |>
  summarize(
    Total = sum(Valor * 1000, na.rm = T) # Miles de personas
  ) |>
  ungroup() |>
  pivot_wider(
    id_cols = c(`Código Región`, `Región`),
    names_from = c(`Código MIP`, `Nombre MIP`),
    values_from = Total,
    names_sep = "_",
    names_sort = T,
    values_fill = 0
  ) |>
  rename(
    `Región de Chile` = `Región`
  )

# # Otro ejemplo con Perú
# equivalencia <- read_excel(
#   "datos/intro/PER_epem_equivalencia.xlsx",
#   col_types = c("numeric", "text", "text", "text")
# )

# empleo_per <- epen2024 |>
#   left_join(
#     equivalencia,
#     join_by(C309_COD)
#   ) |>
#   filter(
#     C310 %in% c(1:7),
#     !is.na(C310),
#     !C309_COD %in% c("9700", "9900")
#   ) |>
#   group_by(CCDD, as_factor(CCDD), io_code, io_name) |>
#   summarize(
#     total = sum(FAC300_ANUAL)
#   ) |>
#   ungroup() |>
#   pivot_wider(
#     id_cols = c(CCDD, `as_factor(CCDD)`),
#     names_from = c(io_code, io_name),
#     values_from = total,
#     names_sep = "_",
#     names_sort = T,
#     values_fill = 0
#   ) |>
#   rename(
#     region_code = CCDD,
#     region_name = `as_factor(CCDD)`
#   )

# Guardar a Excel
library(openxlsx)

# Guardar rápidamente
write.xlsx(
  x = empleo_xlsx,
  file = "salidas/empleo_Excel.xlsx",
  asTable = FALSE
)

# Guardar más complejo
wb <- loadWorkbook("datos/ejemplos/Conversion-a-formato-plano.xlsx")

writeData(
  wb,
  "Datos Originales",
  empleo_xlsx,
  startCol = 1,
  startRow = 40
)
saveWorkbook(wb, "salidas/empleo_MIP.xlsx", overwrite = TRUE)
