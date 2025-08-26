# Script de procesamiento COU
# Librerías
library(readxl)    # Extraer datos del Excel
library(tidyverse) # Manipular datos
library(openxlsx)  # Exportar a Excel

# Limpiar el área de trabajo
rm(list = ls())

# Datos
archivo <- "datos/cou/COU_2022_PRECIOSCORRIENTES_111x181.xlsx"

# Hojas del Excel
hojas <- excel_sheets(archivo)

# Hojas de interés:
# Matriz de producción (of)
# Oferta total (of)
# Utilización intermedia (ut)
# Utilización total (ut)
# Valor Agregado (va)

pes <- c("1", "2", "5", "6")

# Oferta total
info <- read_excel(
  archivo,
  range = paste0("'", pes[1], "'!B4:B6"),
  col_names = FALSE,
  col_types = "text"
)


