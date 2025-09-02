# Script de procesamiento COU
# Renato Vargas

# Librerías
library(readxl) # Importar datos de Excel
library(dplyr) # Manipulación de datos
library(tidyr) # Limpieza de datos
library(stringr) # Manipulación de textos
library(openxlsx) # Exportación a Excel

# Limpiar el área de trabajo
rm(list = ls())

# Datos
archivo <- "datos/cou/COU_2022_PRECIOSCORRIENTES_111x181.xlsx"

# Hojas del Excel
hoja_todas <- excel_sheets(archivo)

# hoja de interés:
# Matriz de producción (mp)

cuadrante <- "mp"
cuadro <- "01 Oferta"
hoja <- "1"
rango <- "C14:DJ194"

excluir_columnas <- c(112)
excluir_filas <- integer(0)

# Se refiere al número de elemento de
# cualquiera de las listas arriba
i <- 1 # Para nuestro bucle "a mano"

anio <- 2022
unidad <- "miles de millones de pesos"

# Importar el cuerpo de datos
datos1 <- read_excel(
  archivo,
  range = "'1'!C14:DJ194",
  col_names = FALSE,
  col_types = "numeric"
)

# Crear códigos de columna y fila estables
n_filas <- nrow(datos1)
n_columnas <- ncol(datos1)
cod_filas <- sprintf(
  "%s_f%03d",
  "mp",
  seq_len(n_filas)
)
cod_columnas <- sprintf(
  "%s_c%03d",
  "mp",
  seq_len(n_columnas)
)

# Reemplazar las celdas vacías con ceros
datos2 <- datos1 |>
  mutate(
    across(
      everything(),
      ~ replace_na(.x, 0)
    )
  ) |>
  setNames(cod_columnas) |>
  mutate(
    Filas = cod_filas,
    .before = 1
  )

# Deshacernos de filas y columnas redundantes o vacías
datos3 <- datos2 |>
  filter(
    !Filas %in% excluir_filas
  ) |>
  select(
    -cod_columnas[excluir_columnas]
  )

# Alargar Y agregar columnas informativas
mp <- datos3 |> # matriz de producción
  pivot_longer(
    cols = -Filas,
    names_to = "Columnas",
    values_to = "Valor"
  ) |>
  mutate(
    Filas,
    Columnas,
    `Año` = anio,
    Cuadro = cuadro,
    Cuadrante = cuadrante,
    Unidades = unidad,
    Precios = "corrientes",
    Valor,
    .keep = "none" # Lo mismo que si usamos transmute.
  )
