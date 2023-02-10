options(scipen = 999)

# Cargar librerías a usar ####

packages <- c('dplyr', 'readxl', 'reshape2', 'tidyr', 'stringr', 'writexl', 'openxlsx')

new.packages <- packages[!(packages %in% installed.packages()[,"Package"])]
if(length(new.packages)) install.packages(new.packages)
lapply(packages, library, character.only = T)

rm(list = ls())

# Crear data.frame a usar posteriormente para guardar resultados ####

consolidado_cuentas_niif <- NULL
consolidado_siniestralidad <- NULL
consolidado_siniestralidad_eps <- NULL
consolidado_eps_cuentas <- NULL
consolidado_siniestralidad_niif <- NULL

`%notin%` <- Negate(`%in%`)

# Definir periodos de estudio ####

periodo <- c('2020', '2021')

for (i in periodo) {
  
  #-------------------------------------------------------------------------------
  #                     Seccion 1. Cargar bases de datos y ajustar
  #-------------------------------------------------------------------------------
  
  # Cargar bases de datos ####
  
  g1 <- read_excel(paste0("./Datos/Originales/",i,".xlsx"), sheet = "FT001-01", skip = 17)
  g2 <- read_excel(paste0("./Datos/Originales/",i,".xlsx"), sheet = "FT001-02", skip = 17)
  g6 <- read_excel(paste0("./Datos/Originales/",i,".xlsx"), sheet = "FT001-06", skip = 17)
  g7 <- read_excel(paste0("./Datos/Originales/",i,".xlsx"), sheet = "FT001-07", skip = 17)
  g8 <- read_excel(paste0("./Datos/Originales/",i,".xlsx"), sheet = "FT001-08", skip = 17)
  EPS <- read_excel(paste0("./Datos/Originales/",i,".xlsx"),  sheet = "Hoja1")
  
  if(i == "2019"){
    g7_1 <- read_excel(paste0("./Datos/Originales/",i,".xlsx"), sheet = "FT001-07_1", skip = 17)
  }else if(i == "2020"){
    g2_1 <- read_excel(paste0("./Datos/Originales/",i,".xlsx"), sheet = "FT001-02_1", skip = 17)
  }else if(i == "2021"){
    g1_1 <- read_excel(paste0("./Datos/Originales/",i,".xlsx"), sheet = "FT001-01_1", skip = 17)
    g2_1 <- read_excel(paste0("./Datos/Originales/",i,".xlsx"), sheet = "FT001-02_1", skip = 17)
  }
  
  # Organizar bases de datos ####
  
  g1 <- g1 %>% pivot_longer(cols=starts_with(c("8","9")),names_to="EPS",values_to="cost")  %>% dplyr::filter(!is.na(LENGTH))
  g2 <- g2 %>% pivot_longer(cols=starts_with(c("8","9")),names_to="EPS",values_to="cost")  %>% dplyr::filter(!is.na(LENGTH))
  g6 <- g6 %>% pivot_longer(cols=starts_with(c("8","9")),names_to="EPS",values_to="cost")  %>% dplyr::filter(!is.na(LENGTH))
  g7 <- g7 %>% pivot_longer(cols=starts_with(c("8","9")),names_to="EPS",values_to="cost")  %>% dplyr::filter(!is.na(LENGTH))
  g8 <- g8 %>% pivot_longer(cols=starts_with(c("8","9")),names_to="EPS",values_to="cost")  %>% dplyr::filter(!is.na(LENGTH))
  
  if(i == "2019"){
    g7_1 <- g7_1 %>% pivot_longer(cols=starts_with(c("8","9")),names_to="EPS",values_to="cost")  %>% dplyr::filter(!is.na(LENGTH))
  }else if(i == "2020"){
    g2_1 <- g2_1 %>% pivot_longer(cols=starts_with(c("8","9")),names_to="EPS",values_to="cost")  %>% dplyr::filter(!is.na(LENGTH))
  }else if(i == "2021"){
    g1_1 <- g1_1 %>% pivot_longer(cols=starts_with(c("8","9")),names_to="EPS",values_to="cost")  %>% dplyr::filter(!is.na(LENGTH))
    g2_1 <- g2_1 %>% pivot_longer(cols=starts_with(c("8","9")),names_to="EPS",values_to="cost")  %>% dplyr::filter(!is.na(LENGTH))
  }
  
  # Unir en la  misma tabla #### 
  
  if(i == "2019"){
    base <- rbind(g1,g2,g6,g7,g8,g7_1)
    rm(g1,g2,g6,g7,g8,g7_1)
  }else if(i == "2020"){
    base <- rbind(g1,g2,g6,g7,g8,g2_1)
    rm(g1,g2,g6,g7,g8,g2_1)
  }else if(i == "2021"){
    base <- rbind(g1,g2,g6,g7,g8,g1_1,g2_1)
    rm(g1,g2,g6,g7,g8,g1_1,g2_1)
  }else{
    base <- rbind(g1,g2,g6,g7,g8)
    rm(g1,g2,g6,g7,g8)
  }
  
  colnames(base)[6] <- colnames(EPS)[2]
  base <- left_join(base,EPS,by="NIT")
  
  # Modificar grupo según Resolución ####
  
  base <- base %>% rename(GRUPO = `GRUPO:`)
  
  base <- base %>% mutate(GRUPO = ifelse(str_starts(GRUPO,"Res. 743")==TRUE,6,GRUPO))
  base <- base %>% mutate(GRUPO = ifelse(str_starts(GRUPO,"Res. 414")==TRUE,7,GRUPO))
  base <- base %>% mutate(GRUPO = ifelse(str_starts(GRUPO,"Res. 533")==TRUE,8,GRUPO))
  
  ################################################################################
  #############################  Excluir EPS por inconsistencia   #########################
  ################################################################################
  
  if(i == "2020"){
    base <- base %>% filter(NIT %notin% c('891280008', '818000140', '860045904', '892115006', '891600091', '839000495', '824001398', '817001773',
                                          '900298372', '891856000', '899999107', '837000084', '809008362', '900604350', '900156264'))
  }else if(i == "2021"){
    base <- base %>% filter(NIT %notin% c('891280008', '818000140', '890480110', '860045904', '892115006', '839000495', '837000084', 
                                          '900604350', '830113831', '891856000', '899999107', '900298372'))
  }else{
    base <- base
  }

  
  #-------------------------------------------------------------------------------
  #                     Seccion 2. Cálculos siniestralidad
  #-------------------------------------------------------------------------------
    
  # Ingresos ####
  
  ingresos_rc_niif_1_2 <- base %>% dplyr::filter(GRUPO %in% c(1,2))
  ingresos_rc_niif_1_2 <- ingresos_rc_niif_1_2 %>% filter(COD %in% c("410227", "410232", "410233"))
  ingresos_rc_niif_1_2$niif <- ingresos_rc_niif_1_2$GRUPO
  
  ingresos_rc_rcp_6_7_8 <- base %>% dplyr::filter(GRUPO %in% c(6,7,8))
  ingresos_rc_rcp_6_7_8 <- ingresos_rc_rcp_6_7_8 %>% filter(COD %in% c("431125")) 
  ingresos_rc_rcp_6_7_8$niif <- ingresos_rc_rcp_6_7_8$GRUPO
  
  ingresos_rc <- rbind(ingresos_rc_niif_1_2,ingresos_rc_rcp_6_7_8)
  rm(ingresos_rc_niif_1_2,ingresos_rc_rcp_6_7_8)
  
  # Analisis individual por cuenta y EPS
  
  ingresos_rc_ag_eps <- ingresos_rc %>% group_by(COD, niif, NIT) %>% summarise(INGRESOS = sum(cost, na.rm=T),.groups = 'drop')
  ingresos_rc_ag_eps$Periodo <- i
  
  Ingresos_Codigos <- read_excel("Datos/Cuentas_Pres_Max.xlsx", sheet = "Ingresos")
  
  Ingresos_Codigos$COD <- as.character(Ingresos_Codigos$COD)
  
  ingresos_rc_ag_eps <- left_join(ingresos_rc_ag_eps, Ingresos_Codigos)
  
  # Analisis individual por cuenta  
  
  ingresos_rc_ag <- ingresos_rc %>% group_by(COD, niif) %>% summarise(INGRESOS = sum(cost, na.rm=T),.groups = 'drop')
  ingresos_rc_ag$Periodo <- i
  
  ingresos_rc_ag <- left_join(ingresos_rc_ag, Ingresos_Codigos)
  
  # Analisis individual por EPS
  
  ingresos_rc_eps <- ingresos_rc %>% group_by(NIT, REGIMEN, niif) %>% summarise(INGRESOS = sum(cost, na.rm=T),.groups = 'drop')
  ingresos_rc_eps$Periodo <- i
  
  # Analisis agregado
  
  ingresos_rc <- ingresos_rc %>% group_by(niif) %>% summarise(INGRESOS = sum(cost, na.rm=T),.groups = 'drop')
  ingresos_rc$Periodo <- i
  
  # Costos ####
  
  costos_rc_niif_1_2 <- base %>% dplyr::filter(GRUPO %in% c(1,2))
  costos_rc_niif_1_2 <- costos_rc_niif_1_2 %>% filter(COD %in% c("610306", "610307", "610308"))
  costos_rc_niif_1_2$niif <- costos_rc_niif_1_2$GRUPO
  
  costos_rc_rcp_6_7_8 <- base %>% dplyr::filter(GRUPO %in% c(6,7,8))
  costos_rc_rcp_6_7_8 <- costos_rc_rcp_6_7_8 %>% filter(COD %in% c("537206", "537207"))
  costos_rc_rcp_6_7_8$niif <- costos_rc_rcp_6_7_8$GRUPO
  
  costos_rc <- rbind(costos_rc_niif_1_2,costos_rc_rcp_6_7_8)
  rm(costos_rc_niif_1_2,costos_rc_rcp_6_7_8)
  
  # Analisis individual por cuenta y EPS  
  
  costos_rc_ag_eps <- costos_rc %>% group_by(COD, niif, NIT) %>% summarise(COSTOS = sum(cost, na.rm=T),.groups = 'drop')
  costos_rc_ag_eps$Periodo <- i
  
  Costos_Codigos <- read_excel("Datos/Cuentas_Pres_Max.xlsx", sheet = "Costos")
  
  Costos_Codigos$COD <- as.character(Costos_Codigos$COD)
  
  costos_rc_ag_eps <- left_join(costos_rc_ag_eps, Costos_Codigos)
  
  # Analisis individual por cuenta  
  
  costos_rc_ag <- costos_rc %>% group_by(COD, niif) %>% summarise(COSTOS = sum(cost, na.rm=T),.groups = 'drop')
  costos_rc_ag$Periodo <- i
  
  costos_rc_ag <- left_join(costos_rc_ag, Costos_Codigos)
  
  # Analisis individual por EPS
  
  costos_rc_eps <- costos_rc %>% group_by(NIT, REGIMEN, niif) %>% summarise(COSTOS = sum(cost, na.rm=T),.groups = 'drop')
  costos_rc_eps$Periodo <- i
  
  # Analisis agregado
  
  costos_rc <- costos_rc %>% group_by(niif) %>% summarise(COSTOS = sum(cost, na.rm=T),.groups = 'drop')
  costos_rc$Periodo <- i
  
  # Liberacion reservas ####
  
  liberacion_rc_niif_1_2 <- base %>% dplyr::filter(GRUPO %in% c(1,2))
  liberacion_rc_niif_1_2 <- liberacion_rc_niif_1_2 %>% filter(COD %in% c("410228", "410229"))
  liberacion_rc_niif_1_2$niif <- liberacion_rc_niif_1_2$GRUPO
  
  liberacion_rc <- liberacion_rc_niif_1_2
  rm(liberacion_rc_niif_1_2)
  
  # Analisis individual por cuenta y EPS
  
  liberacion_rc_ag_eps <- liberacion_rc %>% group_by(COD, niif, NIT) %>% summarise(LIBERACION = sum(cost, na.rm=T),.groups = 'drop')
  liberacion_rc_ag_eps$Periodo <- i
  
  Liberacion_Codigos <- read_excel("Datos/Cuentas_Pres_Max.xlsx", sheet = "Liberacion")
  
  Liberacion_Codigos$COD <- as.character(Liberacion_Codigos$COD)
  
  liberacion_rc_ag_eps <- left_join(liberacion_rc_ag_eps, Liberacion_Codigos)
  
  # Analisis individual por cuenta  
  
  liberacion_rc_ag <- liberacion_rc %>% group_by(COD, niif) %>% summarise(LIBERACION = sum(cost, na.rm=T),.groups = 'drop')
  liberacion_rc_ag$Periodo <- i
  
  liberacion_rc_ag <- left_join(liberacion_rc_ag, Liberacion_Codigos)
  
  # Analisis individual por EPS
  
  liberacion_rc_eps <- liberacion_rc %>% group_by(NIT, REGIMEN, niif) %>% summarise(LIBERACION = sum(cost, na.rm=T),.groups = 'drop')
  liberacion_rc_eps$Periodo <- i
  
  # Analisis agregado
  
  liberacion_rc <- liberacion_rc %>% group_by(niif) %>% summarise(LIBERACION = sum(cost, na.rm = T),.groups = 'drop')
  liberacion_rc$Periodo <- i
  
  # Cuentas
  
  costos_rc_ag <- costos_rc_ag %>% group_by(COD, Periodo, Cuenta) %>% summarise(Valor = sum(COSTOS, na.rm = T))
  costos_rc_ag$Tipo <- "COSTOS" 
  
  ingresos_rc_ag <- ingresos_rc_ag %>% group_by(COD, Periodo, Cuenta) %>% summarise(Valor = sum(INGRESOS, na.rm = T))
  ingresos_rc_ag$Tipo <- "INGRESOS" 
  
  liberacion_rc_ag <- liberacion_rc_ag %>% group_by(COD, Periodo, Cuenta) %>% summarise(Valor = sum(LIBERACION, na.rm = T))
  liberacion_rc_ag$Tipo <- "LIBERACION" 
  
  cuentas_rc_niif <- rbind(costos_rc_ag, ingresos_rc_ag)
  cuentas_rc_niif <- rbind(cuentas_rc_niif, liberacion_rc_ag)
  
  rm(costos_rc_ag, ingresos_rc_ag, liberacion_rc_ag)
  
  # Siniestralidad ####
  
  ingresos_rc_ag_eps <- ingresos_rc_ag_eps %>% rename(Valor = INGRESOS) 
  costos_rc_ag_eps <- costos_rc_ag_eps %>% rename(Valor = COSTOS) 
  liberacion_rc_ag_eps <- liberacion_rc_ag_eps %>% rename(Valor = LIBERACION) 
  
  ingresos_rc_ag_eps$TIPO <- "INGRESOS"
  costos_rc_ag_eps$TIPO <- "COSTOS"
  liberacion_rc_ag_eps$TIPO <- "LIBERACION"
  
  eps_cuentas_rc <- rbind(ingresos_rc_ag_eps, costos_rc_ag_eps)
  eps_cuentas_rc <- rbind(eps_cuentas_rc, liberacion_rc_ag_eps)
  eps_cuentas_rc[is.na(eps_cuentas_rc)] <- 0
  
  siniestralidad_rc_eps <- left_join(ingresos_rc_eps, costos_rc_eps)
  siniestralidad_rc_eps <- left_join(siniestralidad_rc_eps, liberacion_rc_eps)
  siniestralidad_rc_eps[is.na(siniestralidad_rc_eps)] <- 0
  siniestralidad_rc_eps <- siniestralidad_rc_eps %>% mutate(SINIESTRALIDAD = (COSTOS - LIBERACION)/INGRESOS)
  
  siniestralidad_rc_niif <- left_join(ingresos_rc, costos_rc)
  siniestralidad_rc_niif <- left_join(siniestralidad_rc_niif, liberacion_rc)
  siniestralidad_rc_niif[is.na(siniestralidad_rc_niif)] <- 0
  siniestralidad_rc_niif <- siniestralidad_rc_niif %>% mutate(SINIESTRALIDAD = (COSTOS - LIBERACION)/INGRESOS)
  
  siniestralidad_rc <- left_join(ingresos_rc, costos_rc)
  siniestralidad_rc <- left_join(siniestralidad_rc, liberacion_rc)
  siniestralidad_rc[is.na(siniestralidad_rc)] <- 0
  siniestralidad_rc <- siniestralidad_rc %>% group_by(Periodo) %>% summarise(INGRESOS = sum(INGRESOS, na.rm = T), COSTOS = sum(COSTOS, na.rm = T), LIBERACION = sum(LIBERACION, na.rm = T))  %>% mutate(SINIESTRALIDAD = (COSTOS - LIBERACION)/INGRESOS)
  
  rm(ingresos_rc_eps, costos_rc_eps, liberacion_rc_eps, ingresos_rc, costos_rc, liberacion_rc)
  
  consolidado_siniestralidad <- rbind(consolidado_siniestralidad, siniestralidad_rc)
  consolidado_siniestralidad_eps <- rbind(consolidado_siniestralidad_eps, siniestralidad_rc_eps)
  consolidado_eps_cuentas <- rbind(consolidado_eps_cuentas, eps_cuentas_rc)
  consolidado_siniestralidad_niif <- rbind(consolidado_siniestralidad_niif, siniestralidad_rc_niif)
  
}

rm(base, Costos_Codigos, costos_rc_ag_eps, cuentas_rc_niif, EPS, eps_cuentas_rc, Ingresos_Codigos, ingresos_rc_ag_eps,
   liberacion_rc_ag_eps, Liberacion_Codigos, siniestralidad_rc, siniestralidad_rc_eps, siniestralidad_rc_niif)

#-------------------------------------------------------------------------------
#                     Seccion 3. Exportar resultados
#-------------------------------------------------------------------------------

consolidado_siniestralidad <- as.data.frame(consolidado_siniestralidad)
consolidado_siniestralidad_eps <- as.data.frame(consolidado_siniestralidad_eps)
consolidado_eps_cuentas <- as.data.frame(consolidado_eps_cuentas)
consolidado_siniestralidad_niif <- as.data.frame(consolidado_siniestralidad_niif)

# Create a blank workbook
OUT <- createWorkbook()

# Add some sheets to the workbook
addWorksheet(OUT, "Cuentas EPS")
addWorksheet(OUT, "Siniestralidad")
addWorksheet(OUT, "Siniestralidad EPS")
addWorksheet(OUT, "Siniestralidad NIIF")

# Write the data to the sheets
writeData(OUT, sheet = "Siniestralidad", x = consolidado_siniestralidad)
writeData(OUT, sheet = "Siniestralidad EPS", x = consolidado_siniestralidad_eps)
writeData(OUT, sheet = "Siniestralidad NIIF", x = consolidado_siniestralidad_niif)
writeData(OUT, sheet = "Cuentas EPS", x = consolidado_eps_cuentas)

# Export the file
saveWorkbook(OUT, "./Datos/Salidas/ConsolidadoPM_FDF.xlsx")

