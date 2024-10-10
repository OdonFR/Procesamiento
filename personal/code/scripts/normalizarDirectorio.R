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

# Cargar paquetes necesarios 
library(readxl)
library(writexl)
library(dplyr)
library(openxlsx)


##### DEFINIR RUTAS ####
# Ruta del directorio con los archivos actualizados
directorio_entrada <- choose.dir(caption = "Selecciona el directorio del archivo de entrada")
ruta_de_salida <- choose.dir(caption = "Selecciona el directorio donde se guardará el archivo de salida")

# Selecciona el archivo de entrada
ruta_entrada <- file.choose()

#### LEE LA PRIMERA HOJA DEL DOCUMENTO  ####

# Lee el archivo Excel y selecciona automáticamente la primera hoja
hojas <- excel_sheets(ruta_entrada)  # Obtener los nombres de las hojas
primera_hoja <- hojas[1]  # Seleccionar el nombre de la primera hoja

datos <- read_excel(
  ruta_entrada,
  sheet = primera_hoja,
  col_types = "text",
  col_names = TRUE
)





#### AGREGAR COLUMNAS ADICIONALES ####


datos <- datos %>%
  mutate(
    # Crear la columna 'unidadAdscripcionMadre' basada en la columna 'unidad'
    unidadAdscripcionMadre = case_when(
      unidad %in% c("GSEDIA", "SSEDIA", "UTPRTR", "DIVPEP", "SGA") ~ "SEDIA",
      unidad %in% c("UADGDATO", "DIVDIE", "DIVDEDIE", "SGPGOP") ~ "DGDATO",
      unidad %in% c("DIVPETDANEL", "UADGPETDANEL") ~ "DGPETDANEL",
      unidad %in% c("UADGDIA", "DIVED", "SGCIBER", "SGIATHAD", "SGSOD", "SGTED") ~ "DGDIA",
      unidad %in% c("SMTDFP", "GMTDFP") ~ "MTDFP",
      unidad %in% c("GSUBSE") ~ "SUBSE",
      TRUE ~ unidad # Mantener el resto de los valores
    )
  )


datos <- datos %>%
  select(
    unidadAdscripcionMadre,
    unidad,

    everything()  # Mantiene el resto de las columnas en su orden original
  )



#### ELIMINAR #N/D Y VACIAR CELDAS ####

# Reemplazamos "#N/D" por NA y NA por vacío
datos <- datos %>%
  mutate(across(everything(), ~ ifelse(. == "#N/D", "", .)))


#### OUTPUT ####

# Obtener la fecha actual en formato YYYYMMDD
fecha_actual <- format(Sys.Date(), "%Y%m%d")

# Base del nombre del archivo
nombre_base <- paste0(fecha_actual, "_Directorio_SEDIA")

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
addWorksheet(wb, "rpt")

# Escribir los datos en la hoja
writeData(wb, "rpt", datos)

#### CREAR TABLA Y APLICAR FORMATO ####

# Crear un rango de la tabla, desde la primera celda hasta la última celda con datos
writeDataTable(wb, sheet = "rpt", x = datos, startCol = 1, startRow = 1, tableStyle = "TableStyleMedium2", withFilter = TRUE)

# Autoajustar el ancho de las columnas para que sea legible la información
setColWidths(wb, sheet = "rpt", cols = 1:ncol(datos), widths = "auto")

# Crear un estilo para la cabecera (negrita)
header_style <- createStyle(textDecoration = "bold", fontSize = 12, halign = "center", valign = "center")

# Aplicar el estilo a la cabecera (primera fila)
addStyle(wb, sheet = "rpt", style = header_style, rows = 1, cols = 1:ncol(datos), gridExpand = TRUE)

# Guardar el archivo
saveWorkbook(wb, ruta_salida, overwrite = TRUE)
