Sys.setlocale("LC_TIME", "es_EC.utf8")

library(readxl)
library(tidyverse)
library(stringi)
library(stringr)
library(openxlsx)
library(dplyr)
library(excel.link)

base_victimas <- read_excel("C:/Users/juan.minchalo/Desktop/MINEDUC/0. Información Boletín/3. Reportes Jp/2024/07. Jul/Boletín/DATA VICTIMAS.xlsx",
                            sheet = "DATA VICTIMAS")

base_victimas <- base_victimas %>%  filter(ESTADO_CASO == 1)

###########INFORMACIÓN ANUAL A NIVEL NACIONAL#################################

bv_nac_ing_sis <- base_victimas %>%
  filter(ESTADO_CASO == 1) %>% group_by(year(FECHA_INGRESO_SISTEMA)) %>%
  count(year(FECHA_INGRESO_SISTEMA))

bv_nac_ing_dist <- base_victimas %>%
  filter(ESTADO_CASO == 1) %>% group_by(year(FEC_INGRESO_DENUNCIA_DIST)) %>%
  count(year(FEC_INGRESO_DENUNCIA_DIST))

###############ELABORACION DE BASES TIENE O NO TIENE PLAN DE ACOMP#########

bv_no_nac_1 <- base_victimas %>%
  filter(ESTADO_CASO == 1) %>%
  filter(is.na(EXISTE_PLAN_ACOMPANAMIENTO))

bv_si_nac_1 <- base_victimas %>%
  filter(ESTADO_CASO == 1) %>%
  filter(!is.na(EXISTE_PLAN_ACOMPANAMIENTO))

##########INDICADORES#################################

indi_si_plan <- (nrow(bv_si_nac_1)/nrow(base_victimas))
indi_no_plan <- (nrow(bv_no_nac_1)/nrow(base_victimas))


#####FECHA INGRESO EN SISTEMA EN 2024 EN EL SISTEMA PLAN ACOMPAÑAMIENTO######

bv_total_2024 <- base_victimas %>%
  filter(ESTADO_CASO == 1, year(FECHA_INGRESO_SISTEMA) == 2024)

bv_no_2024 <- base_victimas %>%
  filter(ESTADO_CASO == 1, year(FECHA_INGRESO_SISTEMA) == 2024) %>%
  filter(is.na(EXISTE_PLAN_ACOMPANAMIENTO))

bv_si_2024 <- base_victimas %>%
  filter(ESTADO_CASO == 1, year(FECHA_INGRESO_SISTEMA) == 2024) %>%
  filter(!is.na(EXISTE_PLAN_ACOMPANAMIENTO))

indi_si_plan_2024 <- (nrow(bv_si_2024)/nrow(bv_total_2024))*100
indi_no_plan_2024 <- (nrow(bv_no_2024)/nrow(bv_total_2024))*100

bv_no_2024_1 <- bv_no_2024 %>% select(FECHA_INGRESO_SISTEMA,CODIGO_CASO, CODIGO_VICTIMA, TXT_NOM_ZONA, COD_AD_DISTRITO,
                            DISTRITO, NUM_CEDULA, TXT_NOM_VICTIMA, FEC_INGRESO_DENUNCIA_DIST)

Zona_1 <- bv_no_2024_1 %>% filter(TXT_NOM_ZONA == 1)
res_zona_1 <- Zona_1 %>% group_by(COD_AD_DISTRITO, DISTRITO) %>%
  count(DISTRITO)


# Definir la ruta base de la carpeta de destino
base_output_folder <- "C:/Users/juan.minchalo/Desktop/MINEDUC/REDEVI/3. Planes de acompañamiento/7. Julio/"

# Crear y exportar dataframes por zona y distrito
for (zona in 1:9) {
  # Filtrar por zona
  zona_data <- bv_no_2024_1 %>% filter(TXT_NOM_ZONA == zona)
  res_zona <- zona_data %>% group_by(COD_AD_DISTRITO, DISTRITO) %>% count(DISTRITO)
  
  # Crear dataframes por distrito
  zona_por_distrito <- split(zona_data, zona_data$COD_AD_DISTRITO)
  
  # Definir la ruta de la carpeta de destino para la zona
  output_folder <- paste0(base_output_folder, "Zona ", zona, "/")
  
  # Comprobar si la carpeta de destino existe y crearla si no
  if (!dir.exists(output_folder)) {
    dir.create(output_folder, recursive = TRUE)
  }
  
  # Crear un workbook para la zona
  wb_zona <- createWorkbook()
  
  # Añadir hoja con la tabla resumen como la primera hoja
  addWorksheet(wb_zona, "Resumen por Distrito")
  writeData(wb_zona, "Resumen por Distrito", res_zona)
  
  # Añadir hoja con todos los casos de la zona como la segunda hoja
  addWorksheet(wb_zona, "Casos de la Zona")
  writeData(wb_zona, "Casos de la Zona", zona_data)
  
  # Guardar el workbook con todas las hojas en la carpeta de la zona
  saveWorkbook(wb_zona, paste0(output_folder, "Resumen_no_tiene_plan_Zona_", zona, ".xlsx"), overwrite = TRUE)
  
  # Confirmación de creación del archivo resumen
  if (file.exists(paste0(output_folder, "Resumen_Zona_", zona, ".xlsx"))) {
    message(paste("Archivo resumen creado:", paste0(output_folder, "Resumen_Zona_", zona, ".xlsx")))
  } else {
    message(paste("Error al crear el archivo resumen:", paste0(output_folder, "Resumen_Zona_", zona, ".xlsx")))
  }
  
  # Exportar cada distrito a un archivo Excel individual
  lapply(names(zona_por_distrito), function(name) {
    # Crear nombre del archivo con la ruta de destino
    file_name <- paste0(output_folder, "df_", name, ".xlsx")
    
    # Obtener el dataframe correspondiente
    df_zona <- zona_por_distrito[[name]]
    
    # Exportar a archivo Excel individual
    write.xlsx(df_zona, file_name)
    
    # Confirmación de creación del archivo
    if (file.exists(file_name)) {
      message(paste("Archivo creado:", file_name))
    } else {
      message(paste("Error al crear el archivo:", file_name))
    }
  })
}



###############ANALISIS DE LOS CASOS GENERAL###########################

RESUM_GENERAL <- bv_total_2024 %>%
  filter(is.na(EXISTE_PLAN_ACOMPANAMIENTO)) %>%
  group_by(TXT_NOM_ZONA, COD_AD_DISTRITO, DISTRITO) %>%
  count(DISTRITO)

write.xlsx(RESUM_GENERAL, "C:/Users/juan.minchalo/Desktop/MINEDUC/REDEVI/3. Planes de acompañamiento/7. Julio/Resumen_casos_sin_plan_2024.xlsx")

############ANALISIS PARA CADA DISTRITO#####################################

# RESUM_GENERAL <- bv_total_2024 %>% group_by(TXT_NOM_ZONA, COD_AD_DISTRITO, DISTRITO) %>%
#   count(DISTRITO)
# 
# Zona_1 <- bv_no_2024_1 %>% filter(TXT_NOM_ZONA == 1)
# res_zona_1 <- Zona_1 %>% group_by(COD_AD_DISTRITO, DISTRITO) %>%
#   count(DISTRITO)
# 
# 
# Zona_2 <- bv_no_2024_1 %>% filter(TXT_NOM_ZONA == 2)
# res_zona_2 <- Zona_2 %>% group_by(COD_AD_DISTRITO, DISTRITO) %>%
#   count(DISTRITO)
# 
# Zona_3 <- bv_no_2024_1 %>% filter(TXT_NOM_ZONA == 3)
# res_zona_3 <- Zona_3 %>% group_by(COD_AD_DISTRITO, DISTRITO) %>%
#   count(DISTRITO)
# 
# Zona_4 <- bv_no_2024_1 %>% filter(TXT_NOM_ZONA == 4)
# res_zona_4 <- Zona_4 %>% group_by(COD_AD_DISTRITO, DISTRITO) %>%
#   count(DISTRITO)
# 
# Zona_5 <- bv_no_2024_1 %>% filter(TXT_NOM_ZONA == 5)
# res_zona_5 <- Zona_5 %>% group_by(COD_AD_DISTRITO, DISTRITO) %>%
#   count(DISTRITO)
# 
# Zona_6 <- bv_no_2024_1 %>% filter(TXT_NOM_ZONA == 6)
# res_zona_6 <- Zona_6 %>% group_by(COD_AD_DISTRITO, DISTRITO) %>%
#   count(DISTRITO)
# 
# Zona_7 <- bv_no_2024_1 %>% filter(TXT_NOM_ZONA == 7)
# res_zona_7 <- Zona_7 %>% group_by(COD_AD_DISTRITO, DISTRITO) %>%
#   count(DISTRITO)
# 
# Zona_8 <- bv_no_2024_1 %>% filter(TXT_NOM_ZONA == 8)
# res_zona_8 <- Zona_8 %>% group_by(COD_AD_DISTRITO, DISTRITO) %>%
#   count(DISTRITO)
# 
# Zona_9 <- bv_no_2024_1 %>% filter(TXT_NOM_ZONA == 9)
# res_zona_9 <- Zona_9 %>% group_by(COD_AD_DISTRITO, DISTRITO) %>%
#   count(DISTRITO)
# 
# 

# Obtener la última fecha de ingreso en el sistema para cada distrito
ultima_fecha_por_distrito <- base_victimas %>% 
  group_by(TXT_NOM_ZONA, COD_AD_DISTRITO, DISTRITO) %>% 
  summarise(ultima_fecha = max(FECHA_INGRESO_SISTEMA, na.rm = TRUE), .groups = 'drop')

# Calcular el tiempo que ha pasado desde la última fecha de ingreso hasta la fecha actual
ultima_fecha_por_distrito <- ultima_fecha_por_distrito %>%
  mutate(dias_desde_ultimo_reporte = as.numeric(difftime(Sys.Date(), ultima_fecha, units = "days")))

# Lista de distritos
distritos <- c(
  "01D01", "01D02", "01D03", "01D04", "01D05", "01D06", "01D07", "01D08",
  "02D01", "02D02", "02D03", "02D04", "03D01", "03D02", "03D03", "04D01", "04D02", "04D03",
  "05D01", "05D02", "05D03", "05D04", "05D05", "05D06", "06D01", "06D02", "06D03", "06D04",
  "06D05", "07D01", "07D02", "07D03", "07D04", "07D05", "07D06", "08D01", "08D02", "08D03",
  "08D04", "08D05", "08D06", "09D01", "09D02", "09D03", "09D04", "09D05", "09D06", "09D07",
  "09D08", "09D09", "09D10", "09D11", "09D12", "09D13", "09D14", "09D15", "09D16", "09D17",
  "09D18", "09D19", "09D20", "09D21", "09D22", "09D23", "09D24", "10D01", "10D02", "10D03",
  "11D01", "11D02", "11D03", "11D04", "11D05", "11D06", "11D07", "11D08", "11D09", "12D01",
  "12D02", "12D03", "12D04", "12D05", "12D06", "13D01", "13D02", "13D03", "13D04", "13D05",
  "13D06", "13D07", "13D08", "13D09", "13D10", "13D11", "13D12", "14D01", "14D02", "14D03",
  "14D04", "14D05", "14D06", "15D01", "15D02", "16D01", "16D02", "17D01", "17D02", "17D03",
  "17D04", "17D05", "17D06", "17D07", "17D08", "17D09", "17D10", "17D11", "17D12", "18D01",
  "18D02", "18D03", "18D04", "18D05", "18D06", "19D01", "19D02", "19D03", "19D04", "20D01",
  "21D01", "21D02", "21D03", "21D04", "22D01", "22D02", "22D03", "23D01", "23D02", "23D03",
  "24D01", "24D02"
)

# Crear un data frame de distritos
distritos_df <- data.frame(COD_AD_DISTRITO = distritos, stringsAsFactors = FALSE)

# Unir los datos de la última fecha con el data frame de distritos
distritos_resumen <- left_join(distritos_df, ultima_fecha_por_distrito, by = "COD_AD_DISTRITO")


# Función para convertir días en años, meses y días
convertir_dias_en_anos_meses_dias <- function(total_dias) {
  anos <- total_dias %/% 365
  dias_restantes <- total_dias %% 365
  meses <- dias_restantes %/% 30
  dias <- dias_restantes %% 30
  return(list(anos = anos, meses = meses, dias = dias))
}

# Crear un nuevo dataframe que transforma 'dias_desde_ultimo_reporte' en años, meses y días
distritos_resumen_transformado <- distritos_resumen %>%
  mutate(
    conversion = purrr::map(dias_desde_ultimo_reporte, convertir_dias_en_anos_meses_dias),
    anos = purrr::map_int(conversion, "anos"),
    meses = purrr::map_int(conversion, "meses"),
    dias = purrr::map_int(conversion, "dias")
  ) %>%
  select(COD_AD_DISTRITO,DISTRITO, anos, meses, dias)

write.xlsx(distritos_resumen_transformado, "C:/Users/juan.minchalo/Desktop/MINEDUC/REDEVI/3. Planes de acompañamiento/Tiemp_sin_casos/Resumen_tiempo_sin_casos_2024.xlsx")

# Rellenar los valores NA con un valor representativo (por ejemplo, Inf o NA) en la columna de días desde el último reporte
distritos_resumen$dias_desde_ultimo_reporte[is.na(distritos_resumen$dias_desde_ultimo_reporte)] <- Inf

# Obtener la lista única de zonas
zonas_unicas <- unique(distritos_resumen$TXT_NOM_ZONA)

# Directorio base donde se guardarán los archivos por zona
directorio_base <- "C:/Users/juan.minchalo/Desktop/MINEDUC/REDEVI/3. Planes de acompañamiento/Tiemp_sin_casos"

# Crear un archivo Excel por zona y guardarlo en la carpeta correspondiente
for (zona in zonas_unicas) {
  # Filtrar los datos por zona
  datos_zona <- distritos_resumen %>% filter(TXT_NOM_ZONA == zona)
  
  # Crear la ruta de la carpeta de la zona
  carpeta_zona <- file.path(directorio_base, paste0("Zona ", zona))
  
  # Crear la carpeta de la zona si no existe
  if (!dir.exists(carpeta_zona)) {
    dir.create(carpeta_zona, recursive = TRUE)
  }
  
  # Crear el nombre del archivo Excel
  nombre_archivo <- file.path(carpeta_zona, paste0("Resumen_Tiempo_Sin_Casos_2024_Zona_", zona, ".xlsx"))
  
  # Crear el workbook y la hoja
  wb <- createWorkbook()
  addWorksheet(wb, "Resumen por Distrito")
  
  # Escribir los datos en la hoja
  writeData(wb, "Resumen por Distrito", datos_zona)
  
  # Guardar el archivo Excel
  saveWorkbook(wb, nombre_archivo, overwrite = TRUE)
}



