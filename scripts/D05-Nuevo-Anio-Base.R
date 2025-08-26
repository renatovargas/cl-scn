# Preámbulo
library (readxl)
library (openxlsx)
library (reshape2)
library (stringr)
library (plyr)

rm(list = ls())
iso <- "ECU"

# Datos
 
archivo <- "data/mcs_tou_2018_2022.xlsx"
hojas <- excel_sheets(archivo)

# Para no tener que manipular Excel
# Extraemos las hojas de oferta y utilización por separado.
hojas_of <- hojas[grep("^Of \\d{4} C$", hojas)]
hojas_ut <- hojas[grep("^Dm \\d{4} C$", hojas)]

# Inicializamos una lista con un elemento a borrar después.
lista <- c("inicio")

# Para pruebas de un solo año
i <- 1

# Itero por todas las hojas del archivo.
for (i in 1:length(hojas_of)) {
  
  # Año de la pestaña
  anio <- as.numeric((str_extract(hojas_of[i], "\\d{4}")))
  unidad <- "Miles de dólares"
  precios <- "Corrientes"
  
  # Cuadro de Oferta
  oferta <- as.matrix(read_excel(
    archivo,
    range = paste("'" , hojas_of[i], "'!B9:CN704", sep = ""),
    col_names = FALSE,
    col_types = "numeric"
  ))
  
  oferta[is.na(oferta)] <- 0
  
  # Correlativos de fila
  rownames(oferta) <- sprintf(
    paste(iso, "of%03d", sep = ""), 
    seq(1, dim(oferta)[1]))
  
  # Correlativos de columna
  colnames(oferta) <- sprintf(
    paste(iso, "oc%03d", sep = ""), 
    seq(1, dim(oferta)[2]))
  
  # Eliminamos filas y columnas de más
  # Nos quedamos con la distinción de componente
  # Nacional e Importado y eliminamos el Total
  # con una secuencia de tres en tres.
  oferta1 <- oferta[
    -c(seq(from = 1, to = dim(oferta)[1], by = 3)), 
    -c(4:7, 84, 85, 91)]
  
  # Desdoblamos
  oferta2 <- melt(oferta1)
  oferta <- cbind(anio, precios, 1, "Oferta", oferta2, unidad)
  
  # Y asignamos nombres
  colnames(oferta) <- c(
    "Año", 
    "Precios", 
    "No. Cuadro", 
    "Cuadro", 
    "Filas", 
    "Columnas", 
    "Valor", 
    "Unidades")
  
  # Borramos cuadros intermedios
  rm(oferta1, oferta2)
  
  # Cuadro de Utilización
  
  utilizacion <- as.matrix(read_excel(
    archivo,
    range = paste("'" , hojas_ut[i], "'!B9:CW704", sep = ""),
    col_names = FALSE,
    col_types = "numeric"
  ))
  
  utilizacion[is.na(utilizacion)] <- 0.0
  
  rownames(utilizacion) <- sprintf(
    paste(iso, "uf%03d", sep = ""), 
    seq(1, dim(utilizacion)[1]))
  
  colnames(utilizacion) <- sprintf(
    paste(iso, "uc%03d", sep = ""), 
    seq(1, dim(utilizacion)[2]))
  
  utilizacion1 <- utilizacion[
    -c(seq(from = 1, to = dim(oferta)[1], by = 3)), 
    -c(1:8, 85, 86, 89, 92, 96, 99, 100)]
  
  # as.matrix(rowSums(utilizacion1))
  # # Verificamos que vayamos haciendo las cosas bien
  # as.matrix(rowSums(oferta1))-as.matrix(rowSums(utilizacion1))
  
  # Desdoblamos
  utilizacion <- cbind(
    anio, 
    precios, 
    2, 
    "Utilización", 
    melt(utilizacion1), 
    unidad)
  
  colnames(utilizacion) <-
    c("Año",
      "Precios",
      "No. Cuadro",
      "Cuadro",
      "Filas",
      "Columnas",
      "Valor",
      "Unidades")


# >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>># 
# >>>>>>>>>>>>>>>>>>>> PASO 8 >>>>>>>>>>>>>>>>>>>#
#                 Valor Agregado                 #
# >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>># 
      valorAgregado <- as.data.frame(read_excel(
      archivo,
      range = paste("'" ,hojas[i],"'!c571:bu571",sep = ""),
      col_names = FALSE,
      col_types = "numeric"
    ))
      
  rownames(valorAgregado) <-
    c(sprintf(paste(iso, "vf%03d", sep = ""), seq(1, dim(valorAgregado)[1])))
  colnames(valorAgregado) <-
    c(sprintf(paste(iso, "vc%03d", sep = ""), seq(1, dim(valorAgregado)[2])))
  valorAgregado[is.na(valorAgregado)]<- 0.0

    # Desdoblamos
    valorAgregado <-
      cbind(anio,
            precios,
            3,
            "Valor Agregado",
            "ECUvf001",
            melt(valorAgregado),
            unidad)
    
    colnames(valorAgregado) <-
      c("Año",
        "Precios",
        "No. Cuadro",
        "Cuadro",
        "Filas",
        "Columnas",
        "Valor",
        "Unidades")
# >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>># 
# >>>>>>>>>>>>>>>>>>>> PASO 9 >>>>>>>>>>>>>>>>>>>#
#               Cuadro de Empleo                 #
# >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>># 
     empleo <- as.data.frame(read_excel(
       archivo,
       range = paste("'" , hojas[i], "'!c572:bu572", sep = ""),
       col_names = FALSE,
       col_types = "numeric"
     ))
    rownames(empleo) <- 
      c(sprintf(paste(iso, "ef%03d", sep = ""), seq(1, dim(empleo)[1])))
    colnames(empleo) <- 
      c(sprintf(paste(iso, "ec%03d", sep = ""), seq(1, dim(empleo)[2])))
  
#Desdoblamos
    empleo <- cbind(anio,
                    precios,
                    4,
                    "Empleo",
                    "ECUef001",
                    melt(empleo),
                    "Puestos de trabajo")
    
    colnames(empleo) <- c("Año",
                          "Precios",
                          "No. Cuadro",
                          "Cuadro",
                          "Filas",
                          "Columnas",
                          "Valor",
                          "Unidades")
}