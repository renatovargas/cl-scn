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
# Matriz de producción (of)
# Oferta total (ot)
# Utilización intermedia (ui)
# Utilización total (ut)
# Valor Agregado (va)

cuadrante <- c(
  "mp",
  "ot",
  "ui",
  "ut",
  "va"
)
cuadro <- c(
  "01 Oferta",
  "01 Oferta",
  "02 Utilización",
  "02 Utilización",
  "03 Valor Agregado"
)
hoja <- c("1", "2", "5", "6", "23")
rango <- c(
  "C14:DJ194",
  "C15:K195",
  "C14:DJ194",
  "D16:L196",
  "D20:DK29"
)
excluir_columnas <- list(
  c(112), # mp
  c(1, 3, 6, 9), # ot
  c(112), # ui
  c(1, 8, 9), # ut
  c(112) # va
)
excluir_filas <- list(
  integer(0), # mp
  integer(0), # ot
  integer(0), # ui
  integer(0), # ut
  c(2, 3, 5, 6, 8, 9) # va
)

# Se refiere al número de elemento de
# cualquiera de las listas arriba
i <- 1 # Para nuestro bucle "a mano"


# Cuadrante Matriz de Producción (mp)
# Información general de la hoja
info <- read_excel(
  archivo,
  range = sprintf("'%s'!B4:B6", hoja[i]),
  col_names = FALSE,
  col_types = "text"
)

anio <- as.integer((str_extract(info[3, ], "\\d{4}")))
unidad <- info[3, ]

# Importar el cuerpo de datos
datos1 <- read_excel(
  archivo,
  range = sprintf("'%s'!%s", hoja[i], rango[i]),
  col_names = FALSE,
  col_types = "numeric"
)

# Crear códigos de columna y fila estables
n_filas <- nrow(datos1)
n_columnas <- ncol(datos1)
cod_filas <- sprintf("%s_f%03d", cuadrante[i], seq_len(n_filas))
cod_columnas <- sprintf("%s_c%03d", cuadrante[i], seq_len(n_columnas))

# Reemplazar las celdas vacías con ceros
datos2 <- datos1 |>
  mutate(across(everything(), ~ replace_na(.x, 0))) |>
  setNames(cod_columnas) |>
  mutate(Filas = cod_filas, .before = 1)

# Deshacernos de filas y columnas redundantes o vacías
datos3 <- datos2 |>
  filter(
    !Filas %in% cod_filas[excluir_filas[[i]]]
  ) |>
  select(
    -cod_columnas[excluir_columnas[[i]]]
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
    Cuadro = cuadro[i],
    Cuadrante = cuadrante[i],
    Unidades = str_extract(unidad, "(?<=\\().*?(?= de \\d{4}\\))"),
    Precios = "corrientes",
    Valor,
    .keep = "none" # Lo mismo que si usamos transmute.
  )

# Repetir para el segundo cuadrante (ot)
# ...

i <- 2 # Para el segundo elemento de las listas

# Cuadrante oferta total (ot)
# Información general de la hoja
info <- read_excel(
  archivo,
  range = sprintf("'%s'!B4:B6", hoja[i]),
  col_names = FALSE,
  col_types = "text"
)

anio <- as.integer((str_extract(info[3, ], "\\d{4}")))
unidad <- info[3, ]

# Importar el cuerpo de datos
datos1 <- read_excel(
  archivo,
  range = sprintf("'%s'!%s", hoja[i], rango[i]),
  col_names = FALSE,
  col_types = "numeric"
)

# Crear códigos de columna y fila estables
n_filas <- nrow(datos1)
n_columnas <- ncol(datos1)
cod_filas <- sprintf("%s_f%03d", cuadrante[i], seq_len(n_filas))
cod_columnas <- sprintf("%s_c%03d", cuadrante[i], seq_len(n_columnas))

# Reemplazar las celdas vacías con ceros
datos2 <- datos1 |>
  mutate(across(everything(), ~ replace_na(.x, 0))) |>
  setNames(cod_columnas) |>
  mutate(Filas = cod_filas, .before = 1)

# Deshacernos de filas y columnas redundantes o vacías
datos3 <- datos2 |>
  filter(
    !Filas %in% cod_filas[excluir_filas[[i]]]
  ) |>
  select(
    -cod_columnas[excluir_columnas[[i]]]
  )

# Alargar Y agregar columnas informativas
ot <- datos3 |> # matriz de producción
  pivot_longer(
    cols = -Filas,
    names_to = "Columnas",
    values_to = "Valor"
  ) |>
  mutate(
    Filas,
    Columnas,
    `Año` = anio,
    Cuadro = cuadro[i],
    Cuadrante = cuadrante[i],
    Unidades = str_extract(unidad, "(?<=\\().*?(?= de \\d{4}\\))"),
    Precios = "corrientes",
    Valor,
    .keep = "none" # Lo mismo que si usamos transmute.
  )

i <- 3 # Para el tercer elemento de las listas

# Cuadrante oferta total (ot)
# Información general de la hoja
info <- read_excel(
  archivo,
  range = sprintf("'%s'!B4:B6", hoja[i]),
  col_names = FALSE,
  col_types = "text"
)

anio <- as.integer((str_extract(info[3, ], "\\d{4}")))
unidad <- info[3, ]

# Importar el cuerpo de datos
datos1 <- read_excel(
  archivo,
  range = sprintf("'%s'!%s", hoja[i], rango[i]),
  col_names = FALSE,
  col_types = "numeric"
)

# Crear códigos de columna y fila estables
n_filas <- nrow(datos1)
n_columnas <- ncol(datos1)
cod_filas <- sprintf("%s_f%03d", cuadrante[i], seq_len(n_filas))
cod_columnas <- sprintf("%s_c%03d", cuadrante[i], seq_len(n_columnas))

# Reemplazar las celdas vacías con ceros
datos2 <- datos1 |>
  mutate(across(everything(), ~ replace_na(.x, 0))) |>
  setNames(cod_columnas) |>
  mutate(Filas = cod_filas, .before = 1)

# Deshacernos de filas y columnas redundantes o vacías
datos3 <- datos2 |>
  filter(
    !Filas %in% cod_filas[excluir_filas[[i]]]
  ) |>
  select(
    -cod_columnas[excluir_columnas[[i]]]
  )

# Alargar Y agregar columnas informativas
ui <- datos3 |> # matriz de producción
  pivot_longer(
    cols = -Filas,
    names_to = "Columnas",
    values_to = "Valor"
  ) |>
  mutate(
    Filas,
    Columnas,
    `Año` = anio,
    Cuadro = cuadro[i],
    Cuadrante = cuadrante[i],
    Unidades = str_extract(unidad, "(?<=\\().*?(?= de \\d{4}\\))"),
    Precios = "corrientes",
    Valor,
    .keep = "none" # Lo mismo que si usamos transmute.
  )

i <- 4 # Para el cuarto elemento de las listas

# Cuadrante oferta total (ot)
# Información general de la hoja
info <- read_excel(
  archivo,
  range = sprintf("'%s'!B4:B6", hoja[i]),
  col_names = FALSE,
  col_types = "text"
)

anio <- as.integer((str_extract(info[3, ], "\\d{4}")))
unidad <- info[3, ]

# Importar el cuerpo de datos
datos1 <- read_excel(
  archivo,
  range = sprintf("'%s'!%s", hoja[i], rango[i]),
  col_names = FALSE,
  col_types = "numeric"
)

# Crear códigos de columna y fila estables
n_filas <- nrow(datos1)
n_columnas <- ncol(datos1)
cod_filas <- sprintf("%s_f%03d", cuadrante[i], seq_len(n_filas))
cod_columnas <- sprintf("%s_c%03d", cuadrante[i], seq_len(n_columnas))

# Reemplazar las celdas vacías con ceros
datos2 <- datos1 |>
  mutate(across(everything(), ~ replace_na(.x, 0))) |>
  setNames(cod_columnas) |>
  mutate(Filas = cod_filas, .before = 1)

# Deshacernos de filas y columnas redundantes o vacías
datos3 <- datos2 |>
  filter(
    !Filas %in% cod_filas[excluir_filas[[i]]]
  ) |>
  select(
    -cod_columnas[excluir_columnas[[i]]]
  )

# Alargar Y agregar columnas informativas
ut <- datos3 |> # matriz de producción
  pivot_longer(
    cols = -Filas,
    names_to = "Columnas",
    values_to = "Valor"
  ) |>
  mutate(
    Filas,
    Columnas,
    `Año` = anio,
    Cuadro = cuadro[i],
    Cuadrante = cuadrante[i],
    Unidades = str_extract(unidad, "(?<=\\().*?(?= de \\d{4}\\))"),
    Precios = "corrientes",
    Valor,
    .keep = "none" # Lo mismo que si usamos transmute.
  )

i <- 5 # Para el quinto elemento de las listas

# Cuadrante oferta total (ot)
# Información general de la hoja
info <- read_excel(
  archivo,
  range = sprintf("'%s'!B4:B6", hoja[i]),
  col_names = FALSE,
  col_types = "text"
)

anio <- as.integer((str_extract(info[3, ], "\\d{4}")))
unidad <- info[3, ]

# Importar el cuerpo de datos
datos1 <- read_excel(
  archivo,
  range = sprintf("'%s'!%s", hoja[i], rango[i]),
  col_names = FALSE,
  col_types = "numeric"
)

# Crear códigos de columna y fila estables
n_filas <- nrow(datos1)
n_columnas <- ncol(datos1)
cod_filas <- sprintf("%s_f%03d", cuadrante[i], seq_len(n_filas))
cod_columnas <- sprintf("%s_c%03d", cuadrante[i], seq_len(n_columnas))

# Reemplazar las celdas vacías con ceros
datos2 <- datos1 |>
  mutate(across(everything(), ~ replace_na(.x, 0))) |>
  setNames(cod_columnas) |>
  mutate(Filas = cod_filas, .before = 1)

# Deshacernos de filas y columnas redundantes o vacías
datos3 <- datos2 |>
  filter(
    !Filas %in% cod_filas[excluir_filas[[i]]]
  ) |>
  select(
    -cod_columnas[excluir_columnas[[i]]]
  )

# Alargar Y agregar columnas informativas
va <- datos3 |> # matriz de producción
  pivot_longer(
    cols = -Filas,
    names_to = "Columnas",
    values_to = "Valor"
  ) |>
  mutate(
    Filas,
    Columnas,
    `Año` = anio,
    Cuadro = cuadro[i],
    Cuadrante = cuadrante[i],
    Unidades = str_extract(unidad, "(?<=\\().*?(?= de \\d{4}\\))"),
    Precios = "corrientes",
    Valor,
    .keep = "none" # Lo mismo que si usamos transmute.
  )

scn_chl_2022 <- rbind(mp, ot, ui, ut, va)

#Le damos significado a las filas y columnas
clasificacionColumnas <- read_xlsx(
  "datos/equivalencias/CHL_Equivalencias.xlsx",
  sheet = "Columnas",
  col_names = TRUE,
)
clasificacionFilas <- read_xlsx(
  "datos/equivalencias/CHL_Equivalencias.xlsx",
  sheet = "Filas",
  col_names = TRUE,
)

# Hacemos una unión
scn_chl_2022 <- left_join(scn_chl_2022, clasificacionColumnas, by = "Columnas")
scn_chl_2022 <- left_join(scn_chl_2022, clasificacionFilas, by = "Filas")
