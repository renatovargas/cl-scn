# Propuesta estimación de cuentas de cultura

# Cargamos las librerías necesarias
library(tidyverse)
library(readxl)

# Limpiamos el área de trabajo
rm(list = ls())

# Cargamos los datos

# Sistema de Cuentas Nacionales de Chile
scn <- readRDS("salidas/chl_scn_bd.rds")

# Cargamos los coeficientes
coef_oferta1 <- read_xlsx(
  "datos/equivalencias/coeficientes_cultura.xlsx",
  sheet = "coef_oferta",
  col_names = T
) |> 
  mutate(
    `Año` = as.integer(`Año`)
  )

# Convertimos los coeficientes en dataset
# Nos deshacemos de las celdas vacías (values_drop_na)
# Hacemos que los nombres de columna coincidan con la scn_bd
# Si son numericos hay que asegurarse que el tipo de números 
# es el mismo (double con double, integer con integer)
coef_oferta <- coef_oferta1 |> 
  pivot_longer(
    cols = c(4:ncol(coef_oferta1)),
    names_to = "Código Actividades (CAE)",
    values_to = "Coeficiente Cultural",
    values_drop_na = T
  ) |> 
  rename(
    `Código Productos (CUP)` = CUP
  ) |> 
    mutate(
    `Código Actividades (CAE)` = as.numeric(`Código Actividades (CAE)`)
  )  

# Asignamos los coeficientes en donde corresponda
# Nótese que se le asigna a toda la base a pesar
# que nuestra base de coeficientes es solo para mp y
# el año 2022. En este ejercicio filtramos, pero 
# la idea es que la preparación de coeficientes 
# se haga para todo.

scn_chl_mp2022_1 <- scn_chl |> 
  filter(
    Año == 2022
  )



