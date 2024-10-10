# Función para instalar un paquete solo si no está instalado
instalar_si_falta <- function(paquete) {
  if (!require(paquete, character.only = TRUE)) {
    install.packages(paquete, dependencies = TRUE)
    library(paquete, character.only = TRUE)
  }
}

# Instalar y cargar los paquetes necesarios 
instalar_si_falta("readxl")
instalar_si_falta("writexl")
instalar_si_falta("dplyr")
instalar_si_falta("openxlsx")
instalar_si_falta("fuzzyjoin")

# Cargar paquetes necesarios 
library(readxl)
library(writexl)
library(dplyr)
library(openxlsx)
library(fuzzyjoin)


##### DEFINIR RUTAS ####
# Ruta del directorio con los archivos actualizados
directorio_entrada <- choose.dir(caption = "Selecciona el directorio del archivo de entrada")
ruta_de_salida <- choose.dir(caption = "Selecciona el directorio donde se guardará el archivo de salida")

# Selecciona el archivo de entrada
rutaDirectorio <- file.choose()
rutaRPT <- file.choose()

#### Leer los xlsx  ####

# Lee el archivo Excel del directorio y selecciona automáticamente la primera hoja
hojas <- excel_sheets(rutaDirectorio)  # Obtener los nombres de las hojas
primera_hoja <- hojas[1]  # Seleccionar el nombre de la primera hoja

directorio_df <- read_excel(
  rutaDirectorio,
  sheet = primera_hoja,
  col_types = "text",
  col_names = TRUE
)


# Lee el archivo Excel de la RPT normalizada y selecciona automáticamente la primera hoja
hojas <- excel_sheets(tablaRPT)  # Obtener los nombres de las hojas
primera_hoja <- hojas[1]  # Seleccionar el nombre de la primera hoja

rpt_df <- read_excel(
  rutaRPT,
  sheet = primera_hoja,
  col_types = "text",
  col_names = TRUE
)


#### Extraemos las columnas a comparar  ####


nombres_directorio <- directorio_df %>% select(nombre, apellido1, apellido2)
nombres_rpt <- rpt_df %>% select(nombre, apellido1, apellido2)

#Pasamos todo a mayúsculas
directorio_df <- directorio_df%>%
  mutate(across(c(nombre, apellido1, apellido2), toupper))

resultado <- stringdist_full_join(rpt_df, directorio_df,
                             by = c("nombre", "apellido1", "apellido2"),
                             method = "jw",   # Jaro-Winkler
                             max_dist = 0.3,
                             distance_col = "distancia")  # Ajusta la tolerancia a errores

#### REORDENAR COLUMNAS ####
resultado <- resultado %>%
  select(
    unidadAdscripcionMadre.x,
    unidad.x,
    unidad.y,
    nombre.x, 
    apellido1.x,
    nombre.y, 
    apellido1.y,
    everything()  # Mantiene el resto de las columnas en su orden original
  )





#### ELIMINAR #N/D Y VACIAR CELDAS ####

# Reemplazamos "#N/D" por NA y NA por vacío
resultado <- resultado %>%
  mutate(across(everything(), ~ ifelse(. == "#N/D", "", .)))


#### OUTPUT ####

# Obtener la fecha actual en formato YYYYMMDD
fecha_actual <- format(Sys.Date(), "%Y%m%d")

# Base del nombre del archivo
nombre_base <- paste0(fecha_actual, "_SEDIA")

# Generar el nombre completo del archivo
nombre_archivo <- paste0(nombre_base, ".xlsx")
ruta_salida <- file.path(ruta_de_salida, nombre_archivo)

# Si el archivo ya existe, agregar un contador al final
contador <- 1
while (file.exists(ruta_salida)) {
  nombre_archivo <- paste0(nombre_base, "_", contador, ".xlsx")
  ruta_salida <- file.path(ruta_de_salida, nombre_archivo)
  contador <- contador + 1
}

# Crear un nuevo workbook y agregar una hoja con el nombre "rpt"
wb <- createWorkbook()
addWorksheet(wb, "sedia")

# Escribir los datos en la hoja
writeData(wb, "sedia", resultado)

#### CREAR TABLA Y APLICAR FORMATO ####

# Crear un rango de la tabla, desde la primera celda hasta la última celda con datos
writeDataTable(wb, sheet = "sedia", x = resultado, startCol = 1, startRow = 1, tableStyle = "TableStyleMedium2", withFilter = TRUE)

# Autoajustar el ancho de las columnas para que sea legible la información
setColWidths(wb, sheet = "sedia", cols = 1:ncol(resultado), widths = "auto")

# Crear un estilo para la cabecera (negrita)
header_style <- createStyle(textDecoration = "bold", fontSize = 12, halign = "center", valign = "center")

# Aplicar el estilo a la cabecera (primera fila)
addStyle(wb, sheet = "sedia", style = header_style, rows = 1, cols = 1:ncol(resultado), gridExpand = TRUE)

# Guardar el archivo
saveWorkbook(wb, ruta_salida, overwrite = TRUE)
