# Matriz de insumo producto

# Primeramente llamamos las librerías que necesitamos,
# incluyendo funciones a la medida que hemos creado `aggregate_matrix.R`.
library(openxlsx)
library(tidyverse)
source("funciones/aggregate_matrix.R")

# Importamos los datos de Excel. Primero nuestra matriz `Z` con sus nombres.

# Consumo intermedio
Z <- as.matrix(read.xlsx(
  "datos/mip/ECU_MIP_BIO_2019.xlsx",
  sheet = "MIP_BIO_2019",
  rows = c(10:80),
  cols = c(5:75),
  skipEmptyRows = FALSE,
  colNames = FALSE,
  rowNames = FALSE
))

nombres <- read.xlsx(
  "datos/mip/ECU_MIP_BIO_2019.xlsx",
  sheet = "actividades",
  # rows= c(10:80),
  # cols = c(2:4),
  skipEmptyRows = FALSE,
  colNames = T,
  rowNames = F
)
colnames(Z) <- as.factor(nombres$Descripción.Actividades)
rownames(Z) <- as.factor(nombres$Descripción.Actividades)


# Ahora importamos nuestros datos de la demanda final.

DF <- as.matrix(read.xlsx(
  "datos/mip/ECU_MIP_BIO_2019.xlsx",
  sheet = "MIP_BIO_2019",
  rows = c(10:80),
  cols = c(77:82),
  skipEmptyRows = FALSE,
  colNames = FALSE,
  rowNames = FALSE
))


# Modelo de insumo producto

# Producción
x <- as.vector(rowSums(Z) + rowSums(DF))

# Demanda final
f <- as.vector(rowSums(DF))

# x sombrero
xhat <- diag(x)
xhat_inv <- solve(xhat)

# Matriz de coeficientes técnicos
A <- Z %*% solve(xhat)

# Matriz identidad
I <- diag(dim(A)[1])

# Matriz de Leontief
L <- solve(I - A)


# Coeficientes de uso de insumos bioeconómicos por unidad de producto

E_cruda <- as.matrix(read.xlsx(
  "datos/mip/ECU_MIP_BIO_2019.xlsx",
  sheet = "MIP_BIO_2019",
  rows = c(99:103),
  cols = c(4:75),
  skipEmptyRows = FALSE,
  colNames = F,
  rowNames = T
))
colnames(E_cruda) <- as.factor(nombres$Descripción.Actividades)


# Multiplicadores de insumos bioeconómicos

c_bio <- E_cruda[1, ] %*% xhat_inv
c_bio_hat <- diag(as.vector(c_bio))
H_bio <- c_bio_hat %*% L
multip_bio <- as.matrix(colSums(H_bio))


# Multiplicadores de insumos bioeconómicos extendidos.

c_bio_e <- E_cruda[2, ] %*% xhat_inv
c_bio_e_hat <- diag(as.vector(c_bio_e))
H_bio_e <- c_bio_e %*% L
multip_bio_e <- as.matrix(colSums(H_bio_e))


# Multiplicadores de insumos bioeconómicos totales.

c_bio_t <- E_cruda[3, ] %*% xhat_inv
c_bio_t_hat <- diag(as.vector(c_bio_t))
H_bio_t <- c_bio_t %*% L
multip_bio_t <- as.matrix(colSums(H_bio_t))


# Unimos las tres.

multiplicadores_bio <- cbind(multip_bio, multip_bio_e, multip_bio_t)
colnames(multiplicadores_bio) <- c(
  "Bioeconomía",
  "Bioeconomía extendida",
  "Bioeconomía total"
)
multiplicadores_bio <- cbind(nombres, multiplicadores_bio)
# Quitar ceros de la base de datos
multiplicadores_bio <- multiplicadores_bio %>%
  filter(`Bioeconomía total` > 0)


# Mediana de la bioeconomia total
bio_median <- median(multiplicadores_bio$`Bioeconomía total`)

# Classificar a las actividades unidireccionalmente
multiplicadores_bio <- multiplicadores_bio %>%
  mutate(
    Deviation_Category = case_when(
      `Bioeconomía total` > bio_median + 2 * mad(`Bioeconomía total`) ~
        "1. Muy alta",
      `Bioeconomía total` > bio_median + mad(`Bioeconomía total`) ~ "2. Alta",
      `Bioeconomía total` > bio_median ~ "3. Sobre la mediana",
      TRUE ~ "4. Mediana o abajo"
    )
  )

# Ver el dataset
print(head(multiplicadores_bio))

# Hacer el resumen
deviation_summary <- multiplicadores_bio %>%
  group_by(Deviation_Category) %>%
  summarize(
    Count = n(),
    Average_Bioeconomía = mean(`Bioeconomía total`, na.rm = TRUE)
  )

print(deviation_summary)


# Cargamos paquete de gráficos
library(ggplot2)

# Create a boxplot
ggplot(
  multiplicadores_bio,
  aes(
    x = Deviation_Category,
    y = `Bioeconomía total`,
    fill = Deviation_Category
  )
) +
  geom_boxplot() +
  labs(
    title = "Dispersión de multiplicadores de insumos bioeconómicos",
    x = "Clasificación de Desviación Absoluta Mediana",
    y = "Bioeconomía + extendida"
  ) +
  theme_minimal() +
  theme(
    legend.position = "none", # Esconder leyenda
    axis.text.x = element_text(angle = 45, hjust = 1) # Ajustar ejes
  )


# Exportamos a Excel

wb <- loadWorkbook("salidas/resultados.xlsx")
names(wb)

writeData(
  wb,
  "multiplicadores bio",
  multiplicadores_bio,
  rowNames = FALSE
)

# saveWorkbook(
#   wb,
#   "salidas/resultados-mip-ECU.xlsx",
#   overwrite = T
# )
