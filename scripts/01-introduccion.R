# Las funciones de R
## funcion( objeto, parametro1 = "", parametro2 = 2, parametro3 = F)

# Las funciones embebidas
## funcion1( funcion2(objeto, funcion2parametro1 = ...), funcion1parametro1f1 = ...)

# Limpiar el área de trabajo
rm(list = ls())

## El directorio de trabajo
getwd()

## Asignar

# Un valor

uno <- 1
uno

# Un rango de números

n10 <- c(1:10)
n10

# Una cadena de caracteres

c <- "Andrés"
c


# Un vector de cadenas de caracteres

a <- c("Uno", "Dos", "Tres")
a


# Usando la función de asignación

assign("b", a)
b

# Usando un muestreo

tresletras <- LETTERS[1:3]
caracteres <- sample(tresletras, 10, replace = TRUE)
tresletras
caracteres


## Matrices y cuadros de datos

# Una matriz

d <- matrix(c(3, 4, 5, 7, 8, 9, 12, 34, 28), nrow = 3, ncol = 3, byrow = TRUE)
d


# Un cuadro de datos

e <- data.frame(x = 1, y = 1:10, caracteres = caracteres)
e


# Extraer los valores de una columna como vector

extraccion <- e$y
extraccion

# Multiplicar ese vector por un tipo de cambio

extraccion * 2


## Concatenación de cadenas de caracteres
paste(c, "Bello")

## Índices

# ..de Vectores
a
a[c(1, 3)]

# ...de matrices o marcos de datos
d
d[c(1, 3), 3]

# Índices negativos
d
d[-1, -3]
