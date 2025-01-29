setwd("/Users/fsanta/Documents/PE_2025")
library(shiny)
library(readxl)
library(plyr)
library(dplyr)
library(openxlsx)
dir()

# Function to identify the date format
identify_format <- function(date_str) {
  # Extract the potential month part and check if it's valid (01-12)
  month_part <- substr(date_str, 5, 6)
  if (month_part %in% sprintf("%02d", 1:12)) {
    return("yyyymmdd")
  } else {
    return("yyyyddmm")
  }
}



db <- data.frame(readxl::read_excel("db_CANAPRO.xlsx", sheet = "base_datos")) %>% 
      dplyr::filter(!is.na(fecha_vinculacion)) %>%
      dplyr::mutate(
        date_format = sapply(fecha_vinculacion, identify_format),
        fecha_vinculacion = dplyr::case_when(
          date_format == "yyyymmdd" ~ lubridate::ymd(fecha_vinculacion),
          date_format == "yyyyddmm" ~ lubridate::ydm(fecha_vinculacion)),
        fecha_desembolso = lubridate::ymd(base::as.character(fecha_desembolso)),
        saldo_cuenta_ahorros = dplyr::if_else(is.na(saldo_cuenta_ahorros),0,saldo_cuenta_ahorros),
        saldo_ahorros_permanentes  = dplyr::if_else(is.na(saldo_ahorros_permanentes),0,saldo_ahorros_permanentes),
        saldo_cdat = dplyr::if_else(is.na(saldo_cdat),0,saldo_cdat)
      ) %>%
      dplyr::select(-date_format)

summary(db)

# names(db) <- base::tolower(base::names(db))
# names(db)[4] <- "fecha_desembolso"
# names(db)[12] <- "garantia"
# names(db)[16] <- "saldo_aportes_sociales"
# names(db)[17] <- "saldo_ahorros_permanentes"

aux <- db %>% dplyr::group_by(numero_identificacion) %>% 
                    dplyr::summarise(total = sum(saldo_deuda)) %>% base::data.frame()

db <- dplyr::left_join(db, aux, by = "numero_identificacion")

base::rm(aux)


db <- db %>% dplyr::mutate(
                    saldo_aportes_sociales1 = dplyr::if_else(total == 0, 0, base::round(saldo_aportes_sociales*(saldo_deuda/total),0)),
                    saldo_ahorros_permanentes1 = dplyr::if_else(total == 0, 0, base::round(saldo_ahorros_permanentes*(saldo_deuda/total),0)),
                    VEA0 = saldo_deuda,
                    VEA = base::round(dplyr::if_else(VEA0 > 0,VEA0,0),0),
                    total = NULL,
                    VEA0 = NULL,
                    incumplimiento = dplyr::if_else(linea == 3 & mora1 > 120,1,
                                           dplyr::if_else(linea  < 3 & mora1 >  90,1,0)),
                    ddi = dplyr::if_else(linea == 3 & mora1 > 120, mora1 - 120,
                                dplyr::if_else(linea < 3  & mora1 > 90, mora1 - 90, 0)),
                    PDI = dplyr::if_else((garantia==1|garantia==2|garantia==10|garantia==12) & ddi<270,0.5,
                                dplyr::if_else((garantia==1|garantia==2|garantia==10|garantia==12) & (ddi>=270 & ddi<540),0.7,
                                      dplyr::if_else((garantia==1|garantia==2|garantia==10|garantia==12) & ddi>=540,1,
                                            dplyr::if_else(garantia==6|garantia==8,0.12,
                                                  dplyr::if_else(garantia==9 & ddi<360,0.45,
                                                        dplyr::if_else(garantia==9 & (ddi>=360 & ddi<720),0.8,
                                                              dplyr::if_else(garantia==9 & ddi>=720,1,
                                                                    dplyr::if_else(garantia==11 & ddi<360,0.4,
                                                                          dplyr::if_else(garantia==11 & (ddi>=360 & ddi<720),0.7,
                                                                                dplyr::if_else(garantia==11 & ddi>=720,1,
                                                                                      dplyr::if_else(garantia==13 & ddi<210,0.6,
                                                                                            dplyr::if_else(garantia==13 & (ddi>=210 & ddi<420),0.7,
                                                                                                  dplyr::if_else(garantia==13 & ddi>=420,1,
                                                                                                        dplyr::if_else(garantia==14 & ddi<30,0.75,
                                                                                                              dplyr::if_else(garantia==14 & (ddi>=30 & ddi<90),0.85,
                                                                                                                    dplyr::if_else(garantia==14 & ddi>=90,1,0)))))))))))))))),
                    PDI = dplyr::if_else(linea == 1 & (garantia == 13 | garantia == 14) & mora1 <= 90, 0.45, PDI),
                    PDI = dplyr::if_else((linea == 2 | linea == 3) & garantia == 13 & mora1 <= 30, 0.45, PDI),
                    PDI = dplyr::if_else((linea == 2 | linea == 3) & garantia == 14 & mora1 <= 30, 0.50, PDI)
)

summary(db)

db_c_libranza=subset(db,linea==1)
db_s_libranza=subset(db,linea==2)
db_comercial=subset(db,linea==3)

rm(db)

SMMLV <- 1300000
#SMMLV <- 1423500


##### CONSUMO CON LIBRANZA
c_libranza <- db_c_libranza %>% dplyr::select(c(1:8,12,16,17,20,56:61)) %>%
                                dplyr::mutate(
                                          constante=1,
                                          EA = dplyr::if_else(db_c_libranza$estado_asociado == 1, 1, 0),
                                          AP = dplyr::if_else(db_c_libranza$saldo_aportes_sociales > 0, 1, 0),
                                          TC = dplyr::if_else(db_c_libranza$tipo_cuota == 2, 1, 0),
                                          FE = dplyr::if_else(db_c_libranza$organizacion == 5, 1, 0),
                                          ESIN = dplyr::if_else(db_c_libranza$organizacion == 4|
                                                                db_c_libranza$organizacion == 6|
                                                                db_c_libranza$organizacion == 8, 1, 0), 
                                          FAMOR = dplyr::if_else(db_c_libranza$organizacion == 5 & db_c_libranza$amortizacion > 90, 1, 0),
                                          VALCUOTA = dplyr::if_else(db_c_libranza$monto_cuota < 0.1*SMMLV & db_c_libranza$organizacion == 5, 1, 0),
                                          VALPRES = dplyr::if_else(db_c_libranza$monto_desembolsado < SMMLV & db_c_libranza$organizacion == 5, 1, 0),
                                          OCOOP = dplyr::if_else(db_c_libranza$monto_desembolsado > 7*SMMLV & db_c_libranza$organizacion != 5, 1, 0),
                                          FONAHO = dplyr::if_else((db_c_libranza$saldo_cuenta_ahorros + db_c_libranza$saldo_ahorros_permanentes) > 0 & db_c_libranza$organizacion == 5, 1, 0),
                                          COOCDAT = dplyr::if_else((db_c_libranza$organizacion == 3|db_c_libranza$organizacion == 7) & db_c_libranza$saldo_cdat > 0, 1, 0),
                                          FONDPLAZO = dplyr::if_else(db_c_libranza$plazo_deuda <= 6 & db_c_libranza$organizacion == 5,1,0),
                                          ANTIPRE1 = dplyr::if_else((db_c_libranza$fecha_desembolso - db_c_libranza$fecha_vinculacion) <= 31,1,0),
                                          MORA15 = dplyr::if_else(apply(db_c_libranza[,20:31],1,max,na.rm=TRUE)>=16 & apply(db_c_libranza[,20:31],1,max,na.rm=TRUE)<=30,1,0), 
                                          MORA1230 = dplyr::if_else(apply(db_c_libranza[,20:31],1,max,na.rm=TRUE)>=31 & apply(db_c_libranza[,20:31],1,max,na.rm=TRUE)<=60,1,0), 
                                          MORA1260 = dplyr::if_else(apply(db_c_libranza[,20:31],1,max,na.rm=TRUE)>60,1,0),
                                          MORA2430 = dplyr::if_else(apply(db_c_libranza[,20:43],1,max,na.rm=TRUE)>=31 & apply(db_c_libranza[,20:43],1,max,na.rm=TRUE)<=60,1,0), 
                                          MORA2460 = dplyr::if_else(apply(db_c_libranza[,20:43],1,max,na.rm=TRUE)>60,1,0),
                                          SINMORA = dplyr::if_else(apply(db_c_libranza[,20:55],1,max,na.rm=TRUE) == 0,1,0)
)

cp <- db_c_libranza[,20:22]

for(i in 1:3) cp[,i] <- ifelse(cp[,i] >= 31 & cp[,i] <= 60, 1, 0)
rm(i)

c_libranza$MORTRIM <- base::ifelse(apply(cp,1,sum,na.rm=TRUE)>=1,1,0)
rm(cp)


coef_c_libranza <- matrix(c(-2.2504,-0.8444,-1.0573,1.0715,-0.0139,0.4187,0.5313,
                             -0.5536,-0.3662,0.0586,-0.5981,
                             -1.3854,-0.5893,0.7833,0.8526,
                             1.4445,1.3892,0.2823,0.7515,
                             -0.6632,1.2362),ncol=1)

c_libranza <- c_libranza %>% dplyr::mutate(
                                   Puntaje = 1/(1+exp((-1)*(as.matrix(c_libranza[,19:39])%*%coef_c_libranza)))[,1]
                                  )

aux <- c_libranza %>% dplyr::group_by(numero_identificacion) %>% 
                            dplyr::summarise(Puntaje1 = max(Puntaje)) %>% base::data.frame()

c_libranza <- dplyr::left_join(c_libranza, aux, by = "numero_identificacion")

rm(aux)
                                   
c_libranza <- c_libranza %>% dplyr::mutate(
                                   Calificacion = base::factor(
                                     dplyr::if_else(Puntaje1 <= 0.0361,"A",
                                     dplyr::if_else(Puntaje1 >  0.0361 & Puntaje1 <= 0.0815, "B",
                                     dplyr::if_else(Puntaje1 >  0.0815 & Puntaje1 <= 0.2029, "C",
                                     dplyr::if_else(Puntaje1 >  0.2029 & Puntaje1 <= 0.3121, "D","E"))))
                                     ),
                                   Puntaje1 = NULL
                                )

pi_c_libranza <- readxl::read_xlsx("pi.xlsx",sheet = "c_libranza")

c_libranza <- c_libranza %>% 
                 dplyr::left_join(pi_c_libranza, by = c("organizacion","Calificacion")) %>%
                 dplyr::mutate(
                   PI = dplyr::if_else(incumplimiento == 1, 1, PI),
                   PE = base::round(PI * VEA * PDI)
                 )

aux <- c_libranza %>% dplyr::group_by(numero_identificacion) %>% 
                      dplyr::summarise(mora_actual = max(mora1)) %>% base::data.frame()

c_libranza <- dplyr::left_join(c_libranza, aux, by = "numero_identificacion") %>%
              dplyr::mutate(
                Calificacion_Homologada = dplyr::if_else(Calificacion == "A", "A",
                                                         dplyr::if_else(Calificacion == "B" & mora_actual <= 30, "A",
                                                               dplyr::if_else(Calificacion == "B" & mora_actual > 30, "B",
                                                                     dplyr::if_else(Calificacion == "C" & mora_actual <= 30, "B",
                                                                           dplyr::if_else(Calificacion == "C" & mora_actual > 30, "C",
                                                                                 dplyr::if_else(Calificacion == "D" | Calificacion == "E", "C",
                                                                                       dplyr::if_else(incumplimiento == 1 & (mora_actual >= 90 & mora_actual <= 180), "D",
                                                                                             dplyr::if_else(incumplimiento == 1 & mora_actual > 180, "E",Calificacion)))))))),
                mora1 = NULL
              )


rm(pi_c_libranza,db_c_libranza,aux,coef_c_libranza)

######## Sin libranza

s_libranza <- db_s_libranza %>% dplyr::select(c(1:8,12,16,17,20,56:61)) %>%
                                dplyr::mutate(
                                    constante = 1,
                                    EA = dplyr::if_else(db_s_libranza$estado_asociado==1,1,0),
                                    AP = dplyr::if_else(db_s_libranza$saldo_aportes_sociales>0,1,0),
                                    REEST = dplyr::if_else(db_s_libranza$reestructurado==1,1,0),
                                    CUENAHO = dplyr::if_else((db_s_libranza$saldo_cuenta_ahorros>0 & db_s_libranza$estado_asociado==1),1,0),
                                    CDAT = dplyr::if_else(db_s_libranza$saldo_cdat > 0,1,0),
                                    PER = dplyr::if_else(db_s_libranza$saldo_ahorros_permanentes> 0,1,0),
                                    ENTIDAD1 = dplyr::if_else(db_s_libranza$organizacion==2|db_s_libranza$organizacion==4,1,0),
                                    SALPRES = dplyr::if_else(db_s_libranza$saldo_deuda/db_s_libranza$monto_desembolsado<0.2,1,0),
                                    ANTIPRE1 = dplyr::if_else(db_s_libranza$fecha_desembolso-db_s_libranza$fecha_vinculacion<=30,1,0),
                                    ANTIPRE2 = dplyr::if_else(db_s_libranza$fecha_desembolso-db_s_libranza$fecha_vinculacion<=1080,1,0),
                                    VIN2 = dplyr::if_else(as.Date(Sys.time())-db_s_libranza$fecha_vinculacion<=3600,1,0),
                                    MORA1230 = dplyr::if_else(apply(db_s_libranza[,20:31],1,max,na.rm=TRUE)>=31 & apply(db_s_libranza[,20:31],1,max,na.rm=TRUE)<=60,1,0), 
                                    MORA1260 = dplyr::if_else(apply(db_s_libranza[,20:31],1,max,na.rm=TRUE)>60,1,0),
                                    MORA2430 = dplyr::if_else(apply(db_s_libranza[,20:43],1,max,na.rm=TRUE)>=31 & apply(db_s_libranza[,20:43],1,max,na.rm=TRUE)<=60,1,0), 
                                    MORA2460 = dplyr::if_else(apply(db_s_libranza[,20:43],1,max,na.rm=TRUE)>60,1,0),
                                    MORA3615 = dplyr::if_else(apply(db_s_libranza[,20:55],1,max,na.rm=TRUE)>=1 & apply(db_s_libranza[,20:55],1,max,na.rm=TRUE)<=15,1,0)
)

coef_s_libranza=matrix(c(-1.8017,-0.3758,-1.1475,0.4934,-0.387,
                         -1.0786,-0.0167,0.3204,-0.8419,
                          0.1271,-0.3912,-0.4892,0.7877,
                          2.5651,0.696,2.908,0.8114),ncol=1)

s_libranza <- s_libranza %>% dplyr::mutate(
                                   Puntaje = 1/(1+exp((-1)*(as.matrix(s_libranza[,19:35])%*%coef_s_libranza)))[,1]
)

aux <- s_libranza %>% dplyr::group_by(numero_identificacion) %>% 
  dplyr::summarise(Puntaje1 = max(Puntaje)) %>% base::data.frame()

s_libranza <- dplyr::left_join(s_libranza, aux, by = "numero_identificacion")

rm(aux)

s_libranza <- s_libranza %>% dplyr::mutate(
  Calificacion = base::factor(
    dplyr::if_else(Puntaje1 <= 0.1140,"A",
                   dplyr::if_else(Puntaje1 >  0.1140 & Puntaje1 <= 0.3931, "B",
                                  dplyr::if_else(Puntaje1 >  0.3931 & Puntaje1 <= 0.8510, "C",
                                                 dplyr::if_else(Puntaje1 >  0.8510 & Puntaje1 <= 0.9558, "D","E"))))
  ),
  Puntaje1 = NULL
)

pi_s_libranza <- readxl::read_xlsx("pi.xlsx",sheet = "s_libranza")

s_libranza <- s_libranza %>% 
  dplyr::left_join(pi_s_libranza, by = c("organizacion","Calificacion")) %>%
  dplyr::mutate(
    PI = dplyr::if_else(incumplimiento == 1, 1, PI),
    PE = base::round(PI * VEA * PDI)
  )

aux <- s_libranza %>% dplyr::group_by(numero_identificacion) %>% 
  dplyr::summarise(mora_actual = max(mora1)) %>% base::data.frame()

s_libranza <- dplyr::left_join(s_libranza, aux, by = "numero_identificacion") %>%
  dplyr::mutate(
    Calificacion_Homologada = dplyr::if_else(Calificacion == "A", "A",
                                             dplyr::if_else(Calificacion == "B" & mora_actual <= 30, "A",
                                                            dplyr::if_else(Calificacion == "B" & mora_actual > 30, "B",
                                                                           dplyr::if_else(Calificacion == "C" & mora_actual <= 30, "B",
                                                                                          dplyr::if_else(Calificacion == "C" & mora_actual > 30, "C",
                                                                                                         dplyr::if_else(Calificacion == "D" | Calificacion == "E", "C",
                                                                                                                        dplyr::if_else(incumplimiento == 1 & (mora_actual >= 90 & mora_actual <= 180), "D",
                                                                                                                                       dplyr::if_else(incumplimiento == 1 & mora_actual > 180, "E",Calificacion)))))))),
    mora1 = NULL
  )


rm(pi_s_libranza,db_s_libranza,aux,coef_s_libranza)

######## ComercialPN

comercial <- db_comercial %>% dplyr::select(c(1:8,12,16,17,20,56:61)) %>%
                              dplyr::mutate(
                                    constante = 1,
                                    CDAT = dplyr::if_else(db_saldo_cdat > 0,1,0),
                                    REEST = dplyr::if_else(db_reestructurado == 1,1,0),
                                    TC = dplyr::if_else(db_tipo_cuota==2,1,0),
                                    SALPRES = dplyr::if_else(db_saldo_deuda/db_monto_desembolsado < 0.2,1,0),
                                    ANTIPRE1 = dplyr::if_else(db_fecha_desembolso-db_fecha_vinculacion <= 30,1,0)
                                            )

cp <- db_comercial[,20:22]

for(i in 1:3) cp[,i] <- ifelse(cp[,i] >= 31, 1, 0)
rm(i)

comercial$MORTRIM <- base::ifelse(apply(cp,1,sum,na.rm=TRUE)>=1,1,0)
rm(cp)


cp <- db_comercial[,20:31]
for(i in 1:12) cp[,i] = dplyr::if_else(cp[,i]>=31 & cp[,i]<=60,1,0)
rm(i)

comercial <- comercial %>% dplyr::mutate(
                                  `1MORA30` = dplyr::if_else(apply(cp,1,sum,na.rm=TRUE)>=1,1,0),
                                  `2MORA30` = dplyr::if_else(apply(cp,1,sum,na.rm=TRUE)>=2,1,0),
                                  `1MORA30M3` = dplyr::if_else(apply(cp[,1:3],1,sum,na.rm=TRUE)==1,1,0)
                                        )

rm(cp)

cp <- db_comercial[,20:22]

for(i in 1:3) cp[,i] <- ifelse(cp[,i] > 60, 1, 0)
rm(i)

comercial$`1MORA60M3` <- base::ifelse(apply(cp,1,sum,na.rm=TRUE)>=1,1,0)
rm(cp)

comercial <- comercial %>% dplyr::mutate(
MORA1230 = dplyr::if_else(apply(db_comercial[,20:31],1,max,na.rm=TRUE)>=31 & apply(db_comercial[,20:31],1,max,na.rm=TRUE)<=60,1,0),
MORA1260 = dplyr::if_else(apply(db_comercial[,20:31],1,max,na.rm=TRUE)>60,1,0),
MORA2430 = dplyr::if_else(apply(db_comercial[,20:43],1,max,na.rm=TRUE)>=31 & apply(db_comercial[,20:43],1,max,na.rm=TRUE)<=60,1,0), 
MORA2460 = dplyr::if_else(apply(db_comercial[,20:43],1,max,na.rm=TRUE)>60,1,0),
MORA3630 = dplyr::if_else(apply(db_comercial[,20:55],1,max,na.rm=TRUE)>=31 & apply(db_comercial[,20:55],1,max,na.rm=TRUE)<=60,1,0),
MORA3660 = dplyr::if_else(apply(db_comercial[,20:55],1,max,na.rm=TRUE)>60,1,0),
antiguedad = lubridate::interval(fecha_desembolso, base::Sys.Date()) %/% lubridate::months(1),
NODO1 = dplyr::if_else(MORA1260==1 & MORTRIM == 0,1,0),
NODO2 = dplyr::if_else(antiguedad >= 13 & (MORA2430 == 1 | MORA2460 == 1) & (MORA1230 == 0 & MORA1260 == 0),1,0),
NODO3 = dplyr::if_else(antiguedad >= 25 & (MORA2430 == 0 | MORA2460 == 0) & (MORA3630 == 1 | MORA3660 == 1),1,0),
NODO4 = dplyr::if_else(antiguedad < 13  & plazo_deuda > 60,1,0)
)

coef_comercial=matrix(c(-0.5973,-1.827,1.562,-0.024,1.699,0.017,0.024,0.713,0.213,
                        -4.017,-2.463,-0.715,-1.809,-1.020,0.0),ncol=1)

comercial$Puntaje=c(1/(1+exp((-1)*(as.matrix(comercial[,c(16:24,33:38)])%*%coef_comercial))))

comercial$Calificacion=factor(ifelse(comercial$Puntaje<=0.04552,"A",
                              ifelse(comercial$Puntaje>0.04552 & comercial$Puntaje<=0.2194,"B",
                              ifelse(comercial$Puntaje>0.2194 & comercial$Puntaje<=0.4904,"C",
                              ifelse(comercial$Puntaje>0.4904 & comercial$Puntaje<=0.7323,"D","E")))))

comercial$PI=ifelse(db_comercial$organizacion==1 & comercial$Calificacion=="A",0,
             ifelse(db_comercial$organizacion==1 & comercial$Calificacion=="B",0,
             ifelse(db_comercial$organizacion==1 & comercial$Calificacion=="C",0,
             ifelse(db_comercial$organizacion==1 & comercial$Calificacion=="D",0,
             ifelse(db_comercial$organizacion==1 & comercial$Calificacion=="E",0,
                                                 
             ifelse(db_comercial$organizacion==2 & comercial$Calificacion=="A",0,
             ifelse(db_comercial$organizacion==2 & comercial$Calificacion=="B",0,
             ifelse(db_comercial$organizacion==2 & comercial$Calificacion=="C",0,
             ifelse(db_comercial$organizacion==2 & comercial$Calificacion=="D",0,
             ifelse(db_comercial$organizacion==2 & comercial$Calificacion=="E",0,
                                                                                    
             ifelse(db_comercial$organizacion==3 & comercial$Calificacion=="A",0.0116,
             ifelse(db_comercial$organizacion==3 & comercial$Calificacion=="B",0.0742,
             ifelse(db_comercial$organizacion==3 & comercial$Calificacion=="C",0.3589,
             ifelse(db_comercial$organizacion==3 & comercial$Calificacion=="D",0.4144,
             ifelse(db_comercial$organizacion==3 & comercial$Calificacion=="E",0.8947,
                                                                                                                       
             ifelse(db_comercial$organizacion==4 & comercial$Calificacion=="A",0.0120,
             ifelse(db_comercial$organizacion==4 & comercial$Calificacion=="B",0.0757,
             ifelse(db_comercial$organizacion==4 & comercial$Calificacion=="C",0.3640,
             ifelse(db_comercial$organizacion==4 & comercial$Calificacion=="D",0.4144,
             ifelse(db_comercial$organizacion==4 & comercial$Calificacion=="E",0.8947,
                                                                                                                                                          
             ifelse(db_comercial$organizacion==5 & comercial$Calificacion=="A",0.0439,
             ifelse(db_comercial$organizacion==5 & comercial$Calificacion=="B",0.0757,
             ifelse(db_comercial$organizacion==5 & comercial$Calificacion=="C",0.5250,
             ifelse(db_comercial$organizacion==5 & comercial$Calificacion=="D",0.7099,
             ifelse(db_comercial$organizacion==5 & comercial$Calificacion=="E",0.8947,
                                                                                                                                                                                             
             ifelse(db_comercial$organizacion==6 & comercial$Calificacion=="A",0.0120,
             ifelse(db_comercial$organizacion==6 & comercial$Calificacion=="B",0.0757,
             ifelse(db_comercial$organizacion==6 & comercial$Calificacion=="C",0.3640,
             ifelse(db_comercial$organizacion==6 & comercial$Calificacion=="D",0.4144,
             ifelse(db_comercial$organizacion==6 & comercial$Calificacion=="E",0.8947,
                                                                                                                                                                                                                                
             ifelse(db_comercial$organizacion==7 & comercial$Calificacion=="A",0.0174,
             ifelse(db_comercial$organizacion==7 & comercial$Calificacion=="B",0.0859,
             ifelse(db_comercial$organizacion==7 & comercial$Calificacion=="C",0.2152,
             ifelse(db_comercial$organizacion==7 & comercial$Calificacion=="D",0.5550,
             ifelse(db_comercial$organizacion==7 & comercial$Calificacion=="E",0.8947,
             
             ifelse(db_comercial$organizacion==8 & comercial$Calificacion=="A",0.0069,
             ifelse(db_comercial$organizacion==8 & comercial$Calificacion=="B",0.0777,
             ifelse(db_comercial$organizacion==8 & comercial$Calificacion=="C",0.3387,
             ifelse(db_comercial$organizacion==8 & comercial$Calificacion=="D",0.6105,0.8824)))))))))))))))))))))))))))))))))))))))

comercial$PI=ifelse(comercial$incumplimiento==1,1,comercial$PI)
comercial$PE=round(comercial$PI*comercial$VEA*comercial$PDI,0)

if(dim(c_libranza)[1]==0) c_libranza=data.frame(x=c(0,0,0),y=c(0,0,0))
if(dim(s_libranza)[1]==0) s_libranza=data.frame(x=c(0,0,0),y=c(0,0,0))
if(dim(comercial)[1]==0) comercial=data.frame(x=c(0,0,0),y=c(0,0,0))
df_list <- list(ConLibranza=c_libranza, SinLibranza=s_libranza)
write.xlsx(x = df_list , file = "PE_CANAPRO_28_01_2025.xlsx", rowNames = FALSE)
