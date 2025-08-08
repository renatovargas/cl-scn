# Llamar librerías
library(readxl)
library(openxlsx)
library(reshape2)
library(stringr)
library(plyr)

# Limpiar el área de trabajo
rm(list = ls())
# Ruta de trabajo 
wd <- "D:/github/SIBIO/datos/CHL/datos"
# Cambio de ruta de trabajo
setwd(wd)
# País
# Chile = CHL
iso <- "CHL"
# Lógica recursiva
# ================
# info es lo que resulta de leer el archivo

archivo <- "Base.xlsx"
hojas <- excel_sheets(archivo)
#crea objetos con nombre de las hoajs de excel.
#hojas <- hojas[-c(11)]
lista <- c("inicio")
i<- 1
for (i in 1:length(hojas)) {
  # Extracción de año y unidad de medida
  info <- read_excel(
    archivo,
    range = paste("'", hojas[i], "'!A1:A2", sep = "") ,
    col_names = FALSE,
    col_types = "text",
  )
  # Extracción del texto de la cadena de caracteres
  anio <- as.numeric((str_extract(info[1, ], "\\d{4}")))
  unidad <- toString(info[2, ]) 
  # unidad de medida
    precios <- "Corrientes"
    
  # Cuadro de Oferta
  # ================
  
  oferta <- as.matrix(read_excel(
    archivo,
    range = paste("'" , hojas[i], "'!B6:dr186", sep = ""),
    # Nótese que no se incluyó la fila de totales
    col_names = FALSE,
    col_types = "numeric"
  ))
  oferta[is.na(oferta)]<- 0.0
  rownames(oferta) <- c(sprintf(paste(iso, "of%03d", sep = ""), seq(1, dim(oferta)[1])))
  colnames(oferta) <- c(sprintf(paste(iso, "oc%03d", sep = ""), seq(1, dim(oferta)[2])))
  
  # Columnas a eliminar con subtotales y totales
  #Total(oc112)
  #Producto	(oc113)
  #Producción bruta precio básico	(oc114)
  #Oferta total precio básico	(oc116)
  #Oferta total precio productor	(oc119)
  #Nótese que no se eliminan filas.
  oferta1 <- oferta[, -c(112,113,114,116,119)]

  # Desdoblamos
  oferta <- cbind(anio, precios,1, "Oferta", melt(oferta1), unidad)
  
  colnames(oferta) <-
    c("Año",
      "Precios",
      "No. Cuadro",
      "Cuadro",
      "Filas",
      "Columnas",
      "Valor",
      "Unidades")
  
  # Cuadro de utilización
  # =====================
  
  utilizacion <- as.matrix(read_excel(
    archivo,
    range = paste("'" , hojas[i], "'!b191:dq371", sep = ""),
    # Nótese que no se incluyó la fila de totales
    col_names = FALSE,
    col_types = "numeric"
  ))
  utilizacion[is.na(utilizacion)]<- 0.0
    rownames(utilizacion) <-
      c(sprintf(paste(iso, "uf%03d", sep = ""), seq(1, dim(utilizacion)[1])))
    colnames(utilizacion) <-
      c(sprintf(paste(iso, "uc%03d", sep = ""), seq(1, dim(utilizacion)[2])))
  
    # Columnas a eliminar con subtotales y totales
    #Total(oc112)
    #Producto	(oc113)
    #Consumo intermedio	(oc114)
  utilizacion1 <- utilizacion[, -c(112,113,114)]
  as.matrix(rowSums(utilizacion1))
  as.matrix(rowSums(oferta1))-as.matrix(rowSums(utilizacion1))

   # Desdoblamos
  utilizacion <-
    cbind(anio, 
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

  # Cuadros de Valor Agregado y empleo solo para precios Corrientes
  # ========================
      # Cuadro de Valor Agregado
    # ========================
      valorAgregado <- as.data.frame(read_excel(
      archivo,
      range = paste("'" ,hojas[i],"'!b374:dh374",sep = ""),
      col_names = FALSE,
      col_types = "numeric"
    ))
  rownames(valorAgregado) <-
    c(sprintf(paste(iso, "vf%03d", sep = ""), seq(1, dim(valorAgregado)[1])))
  colnames(valorAgregado) <-
    c(sprintf(paste(iso, "vc%03d", sep = ""), seq(1, dim(valorAgregado)[2])))
  valorAgregado[is.na(valorAgregado)]<- 0.0

    #   Columnas a eliminar con subtotales y totales

    # Desdoblamos
    valorAgregado <-
      cbind(anio,
            precios,
            3,
            "Valor Agregado",
            "CHLvf001",
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

# Unimos todas las partes
if (precios == "Corrientes") {
  union <- rbind(oferta, 
                 utilizacion, 
                 valorAgregado)
  
  union <- rbind(oferta,utilizacion,valorAgregado)
  assign(paste("COU_", anio, "_", precios, sep = ""), 
         union)
  lista <- c(lista, paste("COU_", anio, "_", precios, sep = ""))
 
 }
}

# Actualizamos nuestra lista de objetos creados
lista <- lapply(lista[-1], as.name)

# Unimos los objetos de todos los años y precios
SCN <- do.call(rbind.data.frame, lista)

# Y borramos los objetos individuales
do.call(rm,lista)



#Le damos significado a las filas y columnas

clasificacionColumnas <- read_xlsx(
"CHL_Clasificaciones.xlsx",
   sheet = "columnas",
  col_names = TRUE,
)
 clasificacionFilas <- read_xlsx(
  "CHL_Clasificaciones.xlsx",
  sheet = "filas",
   col_names = TRUE,
 )

SCN <- join(SCN,clasificacionColumnas,by = "Columnas")
SCN <- join(SCN,clasificacionFilas, by = "Filas")
gc()

# Y lo exportamos a Excel
write.xlsx(
  SCN,
  "CHL_SCN_BD.xlsx",
  sheetName= "CHL_SCN_BD",
  rowNames=FALSE,
  colnames=FALSE,
  overwrite = TRUE,
  asTable = FALSE
)

# El formato CSV se exporta muy grande, pero se comprime muy bien a 3mb
# write.csv(
#   SCN,
#   "scn.csv",
#   col.names = TRUE,
#   row.names = FALSE,
#   fileEncoding = "latin1"
# )