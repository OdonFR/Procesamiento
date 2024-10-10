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

#### DISTINGUIR UNIDADES DE APOYO ####

# Distinguimos entre UADGDATO, UADGDIA y UADGPETDANEL
codigo_uadgdato <- "52553"
codigo_uadgdia <- "51798"
codigo_uadgpetdanel <- "52554"

datos <- datos %>%
  mutate(
    `Denominación Unidad` = case_when(
      Unidad == codigo_uadgdato ~ "UNIDAD DE APOYO DGDATO",
      Unidad == codigo_uadgdia ~ "UNIDAD DE APOYO DGDIA",
      Unidad == codigo_uadgpetdanel ~ "UNIDAD DE APOYO DGPETDANEL",
      TRUE ~ `Denominación Unidad`
    )
  )

#### AGREGAR COLUMNAS ADICIONALES ####

datos <- datos %>%
  mutate(
    # Crear la columna 'unidad' con siglas de las unidades
    unidad = case_when(
      `Denominación Unidad` == "GABINETE SECR. DE ESTADO" ~ "GSEDIA",
      `Denominación Unidad` == "SECRETARIA SECR. DE ESTADO" ~ "SSEDIA",
      `Denominación Unidad` == "UNIDAD DE APOYO" ~ "UA",
      `Denominación Unidad` == "S.G. DE INTELIGENCIA ARTIFICIAL Y TECNOLOGIAS HABILITADORAS DIGITALES" ~ "SGIATHAD",
      `Denominación Unidad` == "DIVISION DE ECONOMIA DIGITAL" ~ "DIVED",
      `Denominación Unidad` == "S.G. PARA LA SOCIEDAD DIGITAL" ~ "SGSOD",
      `Denominación Unidad` == "S.G. DE TALENTO Y EMPRENDIMIENTO DIGITAL" ~ "SGTED",
      `Denominación Unidad` == "S.G. DE CIBERSEGURIDAD" ~ "SGCIBER",
      `Denominación Unidad` == "DIVISION DE PLANIFIC. ESTRATEGICA EN TECNOLOGIAS DIGITALES AVANZADAS Y NUEVA ECONOMIA DE LA LENGUA" ~ "DIVPETDANEL",
      `Denominación Unidad` == "UNIDAD DE APOYO DGDATO" ~ "UADGDATO",
      `Denominación Unidad` == "UNIDAD DE APOYO DGPETDANEL" ~ "UADGPETDANEL",
      `Denominación Unidad` == "UNIDAD DE APOYO DGDIA" ~ "UADGDIA",
      `Denominación Unidad` == "S.G. PROGRAMAS, GOBERNANZA Y PROMOCION" ~ "SGPGOP",
      `Denominación Unidad` == "DIVISION DE DISEÑO, INNOVACION Y EXPLOTACION" ~ "DIVDEDIE",
      `Denominación Unidad` == "DIVISION DISEÑO,INNOVACION Y EXPLOTACION" ~ "DIVDIE",
      `Denominación Unidad` == "S.G. DE AYUDAS" ~ "SGA",
      `Denominación Unidad` == "DIVISION DE PLANIFICACION Y EJECUCION DE PROGRAMAS" ~ "DIVPEP",
      `Denominación Unidad` == "UNIDAD TEMPORAL DEL PLAN DE RECUPERACION, TRANSFORMACION Y RESILIENCIA" ~ "UTPRTR",
      TRUE ~ `Denominación Unidad` # Mantener el resto de los valores
    ),
    # Crear la columna 'unidadAdscripcionMadre' basada en la columna 'unidad'
    unidadAdscripcionMadre = case_when(
      unidad %in% c("GSEDIA", "SSEDIA", "UTPRTR", "DIVPEP", "SGA") ~ "SEDIA",
      unidad %in% c("UADGDATO", "DIVDIE", "DIVDEDIE", "SGPGOP") ~ "DGDATO",
      unidad %in% c("DIVPETDANEL", "UADGPETDANEL") ~ "DGPETDANEL",
      unidad %in% c("UADGDIA", "DIVED", "SGCIBER", "SGIATHAD", "SGSOD", "SGTED") ~ "DGDIA",
      TRUE ~ unidad # Mantener el resto de los valores
    ),
    # Crear la columna 'situacionPuesto' basada en la columna 'unidad'
    situacionPuesto = case_when(
      # Ocupado
      !is.na(DNI) & Vinculo %in% c(1, 2, 3) & is.na(Obs.) ~ "Ocupado",
      # Reservado OCG - Funcionario en otro puesto
      !is.na(DNI) & Vinculo %in% c(4, 6) & Obs. == "OCG" ~ "Reservado OCG - Funcionario en otro puesto",
      # Libre no ocupable - Pendiente de finalizar proceso de selección
      is.na(DNI) & is.na(Vinculo) & Obs. %in% c("OCG", "OEP") ~ "Libre no ocupable - Pendiente de finalizar proceso de selección",
      # Reservado OEP - Funcionario en otro puesto
      !is.na(DNI) & Vinculo %in% c(4, 6) & Obs. == "OEP" ~ "Reservado OEP - Funcionario en otro puesto",
      # Reservado sin Observaciones - Funcionario en otro puesto
      !is.na(DNI) & Vinculo %in% c(4, 6) & is.na(Obs.) ~ "Reservado sin observaciones - Funcionario en otro puesto",
      # Reservado con Observaciones - Funcionario en otro puesto
      !is.na(DNI) & Vinculo %in% c(4, 6) & !is.na(Obs.) ~ "Reservado con observaciones - Funcionario en otro puesto",
      # Libre y ocupable sin restricciones
      is.na(DNI) & is.na(Vinculo) & is.na(Obs.) ~ "Libre y ocupable sin observaciones", 
      # Libre y ocupable - Para funcionario con experiencia en proyectos
      is.na(DNI) & is.na(Vinculo) & Obs. == "EJ4" ~ "Libre y ocupable - Sólo por funcionario con experiencia en proyectos", 
      # Ocupados con especificidades
      !is.na(DNI) & Vinculo %in% c(1, 2, 3) & !is.na(Obs.) ~ "Ocupado con especificidades",
      # Libre y ocupable pero con especificidades
      is.na(DNI) & is.na(Vinculo) & !(Obs. %in% c("OCG", "OEP")) ~ "Libre y ocupable, pero con especificidades",
      # Errores
      TRUE ~ "Error"
    ),
    unidadPrestacionServicios = "",  # Creamos la columna vacía para añadir datos manualmente
    observacionesAdicionales = ""    # Creamos la columna vacía para añadir datos manualmente
  )

#### ACTUALIZAR COLUMNA NOMBRE Y ELIMINAR PLAZAS NO NECESARIAS ####

#Actualizar la columna 'Nombre' según la condición en 'situacionPuesto'
datos <- datos %>%
  mutate(
    Nombre = case_when(
      situacionPuesto == "Libre y ocupable sin observaciones" ~ "VACANTE",
      situacionPuesto == "Libre y ocupable - Sólo por funcionario con experiencia en proyectos" ~ "VACANTE Funcionario con Experiencia en proyectos",
      situacionPuesto == "Libre y ocupable, pero con especificidades" ~ "VACANTE con OBSERVACIONES",
      TRUE ~ Nombre  # Mantener el valor original de 'Nombre' para las demás filas
    )
  )

# Eliminar las filas con los valores específicos en 'situacionPuesto'
datos <- datos %>%
  filter(
    !situacionPuesto %in% c(
      "Reservado con observaciones - Funcionario en otro puesto",
      "Reservado sin observaciones - Funcionario en otro puesto",
      "Reservado OEP - Funcionario en otro puesto",
      "Reservado OCG - Funcionario en otro puesto",
      "Libre no ocupable - Pendiente de finalizar proceso de selección"
    )
  )

#### ELIMINAR #N/D Y VACIAR CELDAS ####

# Reemplazamos "#N/D" por NA y NA por vacío
datos <- datos %>%
  mutate(across(everything(), ~ ifelse(. == "#N/D", "", .)))

#### REORDENAR COLUMNAS ####
datos <- datos %>%
  select(
    unidadAdscripcionMadre,
    `Denominación Unidad`,
    unidad,
    unidadPrestacionServicios,
    Puesto,
    Denominación,
    Nivel,
    `C.Esp.`,
    Nombre,
    Apellido1,
    Apellido2,
    DNI,
    observacionesAdicionales,
    everything()  # Mantiene el resto de las columnas en su orden original
  )

#### RENOMBRAR COLUMNAS ####

datos <- datos %>%
  rename(
    códigoMinisterio = `Minis.`,
    nombreMinisterio = `Denominación Ministerio`,
    codigoCentroDirectivo = `C.Dir`,
    nombreCentroDirectivo = `Denominación C.Dir`,
    codigoUnidad = `Unidad`,
    nombreUnidad = `Denominación Unidad`,
    destUni = `Dest.Uni.`,
    codigoPuesto = `Puesto`,
    nombrePuesto = `Denominación`,
    nivelFuncionario = `Nivel`,
    complementeEspecifico = `C.Esp.`,
    tipoPuesto = `T.Pto`,
    provision = `Provis.`,
    adscripcionCuerpo = `Ad.Pu`,
    grupoPuesto = `Gr/Sb.`,
    reservaPuesto = `Res.Pue`,
    agregacionCuerpo = `Agr.cuer/cuer`,
    formacion = `For.`,
    titulacionPuesto = `Tit.`,
    observaciones = `Obs.`,
    dni = `DNI`,
    nombre = `Nombre`,
    apellido1 = `Apellido1`,
    apellido2 = `Apellido2`,
    grupoPersona = `Gr/Sb.Vin`,
    vinculo = `Vinculo`,
    tipoRelacionServicio = `Tipo RS`,
    cuerpoDelFuncionario = `Cuerpo del Efectivo`,
    nombreCuerpoFuncionario = `Descripción del Cuerpo`,
    fechaUltimoCesePuesto = `Fecha último cese`,
    codigoSitAdm = `Cód. Sit. Adm.`,
    modSitAdm = `Mod. Sit. Adm.`,
    fechaNacimiento = `Fecha Nacimiento`,
    sexo = `Sexo`,
    fechaNombramiento = `Fecha Nombramiento`,
    fechaUltimaPosesion = `F.Última Posesión`,
  )

# Actualizar la columna 'situacionPuesto' según la condición en 'dni'
datos <- datos %>%
  mutate(
    situacionPuesto = case_when(
      is.na(dni) | dni == "" ~ "Vacante",  # Si 'dni' está vacío o es NA, se asigna "Vacante"
      !is.na(dni) & dni != "" ~ "Ocupado"  # Si 'dni' tiene algún valor, se asigna "Ocupado"
    )
  )

#### OUTPUT ####

# Obtener la fecha actual en formato YYYYMMDD
fecha_actual <- format(Sys.Date(), "%Y%m%d")

# Base del nombre del archivo
nombre_base <- paste0(fecha_actual, "_Personal_SEDIA")

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
