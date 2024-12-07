rsconnect::setAccountInfo(name='asoriesgo1', 
                          token='E56ECE41BFFA65F87BEE1CF669015B72', 
                          secret='eEvNr2rBsjzvgVibOj5vjbb7NoLFl7D0AvQTDg5Z')


base::source("functions.R")

library(openxlsx)
library(shiny)
library(readxl)
library(plyr)
library(dplyr)

    ui = shinyUI(fluidPage(
      titlePanel("Calculo Perdida Esperada"),
      sidebarLayout(
        sidebarPanel(
          fileInput('file1', 'Seleccione archivo con extension .xlsx',
                    accept = c(".xlsx"),buttonLabel = "Subir...",multiple = FALSE
          )
        ),
        mainPanel(
          tableOutput('contents'),
          downloadButton("dl","Descargar resultados...")
      )
    )
    )
    )
    server = shinyServer(function(input, output,session){
      dataInput <- reactive({
        req(input$file1)
        
        inFile <- input$file1
        
        data.frame(read_excel(inFile$datapath, 1))
      })
      
      output$contents <- renderTable({
        db=dataInput()
        table(factor(db$linea))
      },colnames=FALSE)
      
      
      output$dl <- downloadHandler(
        
        filename = function() {
          "Calculo_PE.xlsx"
        },
        content = function(filename){
          db=dataInput()
          aux=data.frame(db %>% group_by(numero_identificacion) %>% summarise(total=sum(saldo_deuda)))
          db=merge(db,aux,by=c('numero_identificacion','numero_identificacion'))
          rm(aux)
          db$saldo_aportes_sociales1=round(db$saldo_aportes_sociales*(db$saldo_deuda/db$total),0)
          db$VEA0=db$saldo_deuda-db$saldo_aportes_sociales1
          db$VEA=round(ifelse(db$VEA0>0,db$VEA0,0),0)
          db$total=NULL
          db$VEA0=NULL
          db$incumplimiento=ifelse(db$linea==3 & db$mora1 > 120,1,
                                   ifelse(db$linea<3  & db$mora1>90,1,0))
          
          db$ddi=ifelse(db$linea==3 & db$mora1>120,db$mora1-120,
                        ifelse(db$linea<3  & db$mora1>90,db$mora1-90,0))
          
          db <- db %>% dplyr::mutate(PDI = calculate_pdi(db))
          
          db_c_libranza=subset(db,linea==1)
          db_s_libranza=subset(db,linea==2)
          db_comercial=subset(db,linea==3)
          
          SMMLV=908526
          
          ##### CONSUMO CON LIBRANZA
          c_libranza=data.frame(db_c_libranza[,c(1:8,12,16,56:60)])
          if(dim(c_libranza)[1]!=0){
          c_libranza$constante=1
          c_libranza$EA=ifelse(db_c_libranza$estado_asociado==1,1,0)
          c_libranza$AP=ifelse(db_c_libranza$saldo_aportes_sociales>0,1,0)
          c_libranza$TC=ifelse(db_c_libranza$tipo_cuota==2,1,0)
          c_libranza$FE=ifelse(db_c_libranza$organizacion==5,1,0)
          c_libranza$ESIN=ifelse(db_c_libranza$organizacion==4|
                                   db_c_libranza$organizacion==6|
                                   db_c_libranza$organizacion==8,1,0) 
          c_libranza$FAMOR=ifelse(db_c_libranza$organizacion==5 & db_c_libranza$amortizacion>90,1,0)
          c_libranza$VALCUOTA=ifelse(db_c_libranza$monto_cuota<0.1*SMMLV & db_c_libranza$organizacion==5,1,0)
          c_libranza$VALPRES=ifelse(db_c_libranza$monto_desembolsado < SMMLV & db_c_libranza$organizacion==5,1,0)
          c_libranza$OCOOP=ifelse(db_c_libranza$monto_desembolsado > 7*SMMLV & db_c_libranza$organizacion!=5,1,0)
          c_libranza$FONAHO=ifelse((db_c_libranza$saldo_cuenta_ahorros+ db_c_libranza$saldo_ahorros_permanentes)> 0 & db_c_libranza$organizacion==5,1,0)
          c_libranza$COOCDAT=ifelse((db_c_libranza$organizacion==3|db_c_libranza$organizacion==7) & db_c_libranza$saldo_cdat > 0,1,0)
          c_libranza$FONDPLAZO=ifelse(db_c_libranza$plazo_deuda<=6 & db_c_libranza$organizacion==5,1,0)
          c_libranza$ANTIPRE1=ifelse(db_c_libranza$fecha_desembolso-db_c_libranza$fecha_vinculacion<=31,1,0)
          c_libranza$MORA15=ifelse(apply(db_c_libranza[,20:31],1,max,na.rm=TRUE)>=16 & apply(db_c_libranza[,20:31],1,max,na.rm=TRUE)<=30,1,0) 
          c_libranza$MORA1230=ifelse(apply(db_c_libranza[,20:31],1,max,na.rm=TRUE)>=31 & apply(db_c_libranza[,20:31],1,max,na.rm=TRUE)<=60,1,0) 
          c_libranza$MORA1260=ifelse(apply(db_c_libranza[,20:31],1,max,na.rm=TRUE)>60,1,0)
          c_libranza$MORA2430=ifelse(apply(db_c_libranza[,20:43],1,max,na.rm=TRUE)>=31 & apply(db_c_libranza[,20:43],1,max,na.rm=TRUE)<=60,1,0) 
          c_libranza$MORA2460=ifelse(apply(db_c_libranza[,20:43],1,max,na.rm=TRUE)>60,1,0)
          c_libranza$SINMORA=ifelse(apply(db_c_libranza[,20:55],1,max,na.rm=TRUE)==0,1,0)
          cp=db_c_libranza[,20:22]
          for(i in 1:3) cp[,i]=ifelse(cp[,i]>=31 & cp[,i]<=60,1,0)
          rm(i)
          c_libranza$MORTRIM=ifelse(apply(cp,1,sum,na.rm=TRUE)>=1,1,0)
          rm(cp)
          
          coef_c_libranza=matrix(c(-2.2504,-0.8444,-1.0573,1.0715,-0.0139,0.4187,0.5313,-0.5536,-0.3662,0.0586,-0.5981,
                                   -1.3854,-0.5893,0.7833,0.8526,1.4445,1.3892,0.2823,0.7515,-0.6632,1.2362),ncol=1)
          
          c_libranza$Puntaje=c(1/(1+exp((-1)*(as.matrix(c_libranza[,16:36])%*%coef_c_libranza))))
          
          c_libranza$Calificacion=factor(ifelse(c_libranza$Puntaje<=0.0174,"A",
                                                ifelse(c_libranza$Puntaje>0.0174 & c_libranza$Puntaje<=0.0337,"B",
                                                       ifelse(c_libranza$Puntaje>0.0337 & c_libranza$Puntaje<=0.0479,"C",
                                                              ifelse(c_libranza$Puntaje>0.0479 & c_libranza$Puntaje<=0.0812,"D","E")))))
          
          c_libranza$PI=ifelse(db_c_libranza$organizacion==1 & c_libranza$Calificacion=="A",0.0067,
                               ifelse(db_c_libranza$organizacion==1 & c_libranza$Calificacion=="B",0.0209,
                                      ifelse(db_c_libranza$organizacion==1 & c_libranza$Calificacion=="C",0.0462,
                                             ifelse(db_c_libranza$organizacion==1 & c_libranza$Calificacion=="D",0.0624,
                                                    ifelse(db_c_libranza$organizacion==1 & c_libranza$Calificacion=="E",0.2566,
                                                           
                                                           ifelse(db_c_libranza$organizacion==2 & c_libranza$Calificacion=="A",0.0095,
                                                                  ifelse(db_c_libranza$organizacion==2 & c_libranza$Calificacion=="B",0.0216,
                                                                         ifelse(db_c_libranza$organizacion==2 & c_libranza$Calificacion=="C",0.1061,
                                                                                ifelse(db_c_libranza$organizacion==2 & c_libranza$Calificacion=="D",0.2526,
                                                                                       ifelse(db_c_libranza$organizacion==2 & c_libranza$Calificacion=="E",0.3811,
                                                                                              
                                                                                              ifelse(db_c_libranza$organizacion==3 & c_libranza$Calificacion=="A",0.0050,
                                                                                                     ifelse(db_c_libranza$organizacion==3 & c_libranza$Calificacion=="B",0.0060,
                                                                                                            ifelse(db_c_libranza$organizacion==3 & c_libranza$Calificacion=="C",0.0441,
                                                                                                                   ifelse(db_c_libranza$organizacion==3 & c_libranza$Calificacion=="D",0.0448,
                                                                                                                          ifelse(db_c_libranza$organizacion==3 & c_libranza$Calificacion=="E",0.2273,
                                                                                                                                 
                                                                                                                                 ifelse(db_c_libranza$organizacion==4 & c_libranza$Calificacion=="A",0.0229,
                                                                                                                                        ifelse(db_c_libranza$organizacion==4 & c_libranza$Calificacion=="B",0.0254,
                                                                                                                                               ifelse(db_c_libranza$organizacion==4 & c_libranza$Calificacion=="C",0.0337,
                                                                                                                                                      ifelse(db_c_libranza$organizacion==4 & c_libranza$Calificacion=="D",0.0412,
                                                                                                                                                             ifelse(db_c_libranza$organizacion==4 & c_libranza$Calificacion=="E",0.3281,
                                                                                                                                                                    
                                                                                                                                                                    ifelse(db_c_libranza$organizacion==5 & c_libranza$Calificacion=="A",0.0058,
                                                                                                                                                                           ifelse(db_c_libranza$organizacion==5 & c_libranza$Calificacion=="B",0.0274,
                                                                                                                                                                                  ifelse(db_c_libranza$organizacion==5 & c_libranza$Calificacion=="C",0.0678,
                                                                                                                                                                                         ifelse(db_c_libranza$organizacion==5 & c_libranza$Calificacion=="D",0.1205,
                                                                                                                                                                                                ifelse(db_c_libranza$organizacion==5 & c_libranza$Calificacion=="E",0.2709,
                                                                                                                                                                                                       
                                                                                                                                                                                                       ifelse(db_c_libranza$organizacion==6 & c_libranza$Calificacion=="A",0.0067,
                                                                                                                                                                                                              ifelse(db_c_libranza$organizacion==6 & c_libranza$Calificacion=="B",0.0209,
                                                                                                                                                                                                                     ifelse(db_c_libranza$organizacion==6 & c_libranza$Calificacion=="C",0.0462,
                                                                                                                                                                                                                            ifelse(db_c_libranza$organizacion==6 & c_libranza$Calificacion=="D",0.0624,
                                                                                                                                                                                                                                   ifelse(db_c_libranza$organizacion==6 & c_libranza$Calificacion=="E",0.2566,
                                                                                                                                                                                                                                          
                                                                                                                                                                                                                                          ifelse(db_c_libranza$organizacion==7 & c_libranza$Calificacion=="A",0.0040,
                                                                                                                                                                                                                                                 ifelse(db_c_libranza$organizacion==7 & c_libranza$Calificacion=="B",0.0217,
                                                                                                                                                                                                                                                        ifelse(db_c_libranza$organizacion==7 & c_libranza$Calificacion=="C",0.0406,
                                                                                                                                                                                                                                                               ifelse(db_c_libranza$organizacion==7 & c_libranza$Calificacion=="D",0.1027,
                                                                                                                                                                                                                                                                      ifelse(db_c_libranza$organizacion==7 & c_libranza$Calificacion=="E",0.2263,
                                                                                                                                                                                                                                                                             
                                                                                                                                                                                                                                                                             ifelse(db_c_libranza$organizacion==8 & c_libranza$Calificacion=="A",0.0116,
                                                                                                                                                                                                                                                                                    ifelse(db_c_libranza$organizacion==8 & c_libranza$Calificacion=="B",0.0209,
                                                                                                                                                                                                                                                                                           ifelse(db_c_libranza$organizacion==8 & c_libranza$Calificacion=="C",0.0407,
                                                                                                                                                                                                                                                                                                  ifelse(db_c_libranza$organizacion==8 & c_libranza$Calificacion=="D",0.0737,0.2013)))))))))))))))))))))))))))))))))))))))
          
          c_libranza$PI=ifelse(c_libranza$incumplimiento==1,1,c_libranza$PI)
          c_libranza$PE=round(c_libranza$PI*c_libranza$VEA*c_libranza$PDI,0)
          }
          ######## Sin libranza
          
          s_libranza=data.frame(db_s_libranza[,c(1:8,12,16,56:60)])
          if(dim(s_libranza)[1]!=0){
          s_libranza$constante=1
          s_libranza$EA=ifelse(db_s_libranza$estado_asociado==1,1,0)
          s_libranza$AP=ifelse(db_s_libranza$saldo_aportes_sociales>0,1,0)
          s_libranza$REEST=ifelse(db_s_libranza$reestructurado==1,1,0)
          s_libranza$CUENAHO=ifelse((db_s_libranza$saldo_cuenta_ahorros>0 & db_s_libranza$estado_asociado==1),1,0)
          s_libranza$CDAT=ifelse(db_s_libranza$saldo_cdat > 0,1,0)
          s_libranza$PER=ifelse(db_s_libranza$saldo_ahorros_permanentes> 0,1,0)
          s_libranza$ENTIDAD1=ifelse(db_s_libranza$organizacion==2|db_s_libranza$organizacion==4,1,0)
          s_libranza$SALPRES=ifelse(db_s_libranza$saldo_deuda/db_s_libranza$monto_desembolsado<0.2,1,0)
          s_libranza$ANTIPRE1=ifelse(db_s_libranza$fecha_desembolso-db_s_libranza$fecha_vinculacion<=30,1,0)
          s_libranza$ANTIPRE2=ifelse(db_s_libranza$fecha_desembolso-db_s_libranza$fecha_vinculacion<=1080,1,0)
          s_libranza$VIN2=ifelse(as.POSIXlt(Sys.time())-db_s_libranza$fecha_vinculacion<=3600,1,0)
          s_libranza$MORA1230=ifelse(apply(db_s_libranza[,20:31],1,max,na.rm=TRUE)>=31 & apply(db_s_libranza[,20:31],1,max,na.rm=TRUE)<=60,1,0) 
          s_libranza$MORA1260=ifelse(apply(db_s_libranza[,20:31],1,max,na.rm=TRUE)>60,1,0)
          s_libranza$MORA2430=ifelse(apply(db_s_libranza[,20:43],1,max,na.rm=TRUE)>=31 & apply(db_s_libranza[,20:43],1,max,na.rm=TRUE)<=60,1,0) 
          s_libranza$MORA2460=ifelse(apply(db_s_libranza[,20:43],1,max,na.rm=TRUE)>60,1,0)
          s_libranza$MORA3615=ifelse(apply(db_s_libranza[,20:55],1,max,na.rm=TRUE)>=1 & apply(db_s_libranza[,20:55],1,max,na.rm=TRUE)<=15,1,0)
          
          coef_s_libranza=matrix(c(-1.8017,-0.3758,-1.1475,0.4934,-0.387,-1.0786,-0.0167,0.3204,
                                   -0.8430,0.1271,-0.3912,-0.4892,0.7877,2.5651,0.696,2.9008,0.8114),ncol=1)
          
          s_libranza$Puntaje=c(1/(1+exp((-1)*(as.matrix(s_libranza[,16:32])%*%coef_s_libranza))))
          
          s_libranza$Calificacion=factor(ifelse(s_libranza$Puntaje<=0.0559,"A",
                                                ifelse(s_libranza$Puntaje>0.0559 & s_libranza$Puntaje<=0.1066,"B",
                                                       ifelse(s_libranza$Puntaje>0.1066 & s_libranza$Puntaje<=0.2209,"C",
                                                              ifelse(s_libranza$Puntaje>0.2209 & s_libranza$Puntaje<=0.3690,"D","E")))))
          
          s_libranza$PI=ifelse(db_s_libranza$organizacion==1 & s_libranza$Calificacion=="A",0.0032,
                               ifelse(db_s_libranza$organizacion==1 & s_libranza$Calificacion=="B",0.0156,
                                      ifelse(db_s_libranza$organizacion==1 & s_libranza$Calificacion=="C",0.0294,
                                             ifelse(db_s_libranza$organizacion==1 & s_libranza$Calificacion=="D",0.0981,
                                                    ifelse(db_s_libranza$organizacion==1 & s_libranza$Calificacion=="E",0.4302,
                                                           
                                                           ifelse(db_s_libranza$organizacion==2 & s_libranza$Calificacion=="A",0.0172,
                                                                  ifelse(db_s_libranza$organizacion==2 & s_libranza$Calificacion=="B",0.1600,
                                                                         ifelse(db_s_libranza$organizacion==2 & s_libranza$Calificacion=="C",0.2657,
                                                                                ifelse(db_s_libranza$organizacion==2 & s_libranza$Calificacion=="D",0.3582,
                                                                                       ifelse(db_s_libranza$organizacion==2 & s_libranza$Calificacion=="E",0.4646,
                                                                                              
                                                                                              ifelse(db_s_libranza$organizacion==3 & s_libranza$Calificacion=="A",0.0150,
                                                                                                     ifelse(db_s_libranza$organizacion==3 & s_libranza$Calificacion=="B",0.0595,
                                                                                                            ifelse(db_s_libranza$organizacion==3 & s_libranza$Calificacion=="C",0.1382,
                                                                                                                   ifelse(db_s_libranza$organizacion==3 & s_libranza$Calificacion=="D",0.3277,
                                                                                                                          ifelse(db_s_libranza$organizacion==3 & s_libranza$Calificacion=="E",0.4171,
                                                                                                                                 
                                                                                                                                 ifelse(db_s_libranza$organizacion==4 & s_libranza$Calificacion=="A",0.0403,
                                                                                                                                        ifelse(db_s_libranza$organizacion==4 & s_libranza$Calificacion=="B",0.0843,
                                                                                                                                               ifelse(db_s_libranza$organizacion==4 & s_libranza$Calificacion=="C",0.0959,
                                                                                                                                                      ifelse(db_s_libranza$organizacion==4 & s_libranza$Calificacion=="D",0.2812,
                                                                                                                                                             ifelse(db_s_libranza$organizacion==4 & s_libranza$Calificacion=="E",0.3986,
                                                                                                                                                                    
                                                                                                                                                                    ifelse(db_s_libranza$organizacion==5 & s_libranza$Calificacion=="A",0.0205,
                                                                                                                                                                           ifelse(db_s_libranza$organizacion==5 & s_libranza$Calificacion=="B",0.1088,
                                                                                                                                                                                  ifelse(db_s_libranza$organizacion==5 & s_libranza$Calificacion=="C",0.2313,
                                                                                                                                                                                         ifelse(db_s_libranza$organizacion==5 & s_libranza$Calificacion=="D",0.3589,
                                                                                                                                                                                                ifelse(db_s_libranza$organizacion==5 & s_libranza$Calificacion=="E",0.5014,
                                                                                                                                                                                                       
                                                                                                                                                                                                       ifelse(db_s_libranza$organizacion==6 & s_libranza$Calificacion=="A",0,
                                                                                                                                                                                                              ifelse(db_s_libranza$organizacion==6 & s_libranza$Calificacion=="B",0,
                                                                                                                                                                                                                     ifelse(db_s_libranza$organizacion==6 & s_libranza$Calificacion=="C",0,
                                                                                                                                                                                                                            ifelse(db_s_libranza$organizacion==6 & s_libranza$Calificacion=="D",0,
                                                                                                                                                                                                                                   ifelse(db_s_libranza$organizacion==6 & s_libranza$Calificacion=="E",0,
                                                                                                                                                                                                                                          
                                                                                                                                                                                                                                          ifelse(db_s_libranza$organizacion==7 & s_libranza$Calificacion=="A",0.0186,
                                                                                                                                                                                                                                                 ifelse(db_s_libranza$organizacion==7 & s_libranza$Calificacion=="B",0.0789,
                                                                                                                                                                                                                                                        ifelse(db_s_libranza$organizacion==7 & s_libranza$Calificacion=="C",0.2029,
                                                                                                                                                                                                                                                               ifelse(db_s_libranza$organizacion==7 & s_libranza$Calificacion=="D",0.4052,
                                                                                                                                                                                                                                                                      ifelse(db_s_libranza$organizacion==7 & s_libranza$Calificacion=="E",0.4451,
                                                                                                                                                                                                                                                                             
                                                                                                                                                                                                                                                                             ifelse(db_s_libranza$organizacion==8 & s_libranza$Calificacion=="A",0.0354,
                                                                                                                                                                                                                                                                                    ifelse(db_s_libranza$organizacion==8 & s_libranza$Calificacion=="B",0.0820,
                                                                                                                                                                                                                                                                                           ifelse(db_s_libranza$organizacion==8 & s_libranza$Calificacion=="C",0.1650,
                                                                                                                                                                                                                                                                                                  ifelse(db_s_libranza$organizacion==8 & s_libranza$Calificacion=="D",0.3630,0.4327)))))))))))))))))))))))))))))))))))))))
          
          s_libranza$PI=ifelse(s_libranza$incumplimiento==1,1,s_libranza$PI)
          s_libranza$PE=round(s_libranza$PI*s_libranza$VEA*s_libranza$PDI,0)
          }
          ######## ComercialPN
          
          comercial=data.frame(db_comercial[,c(1:8,12,16,56:60)])
          if(dim(comercial)[1]!=0){
          comercial$constante=1
          comercial$CDAT=ifelse(db_comercial$saldo_cdat > 0,1,0)
          comercial$REEST=ifelse(db_comercial$reestructurado==1,1,0)
          comercial$PER=ifelse(db_comercial$saldo_ahorros_permanentes> 0,1,0)
          comercial$TC=ifelse(db_comercial$tipo_cuota==2,1,0)
          comercial$PLAZOL=ifelse(db_comercial$plazo_deuda>36,1,0)
          comercial$AMOR=ifelse(db_comercial$amortizacion>=90,1,0)
          comercial$SALPRES=ifelse(db_comercial$saldo_deuda/db_comercial$monto_desembolsado<0.2,1,0)
          comercial$ANTIPRE1=ifelse(db_comercial$fecha_desembolso-db_comercial$fecha_vinculacion<=30,1,0)
          cp=db_comercial[,20:31]
          for(i in 1:12) cp[,i]=ifelse(cp[,i]>=31 & cp[,i]<=60,1,0)
          rm(i)
          comercial$`1MORA30`=ifelse(apply(cp,1,sum,na.rm=TRUE)==1,1,0)
          comercial$`2MORA30`=ifelse(apply(cp,1,sum,na.rm=TRUE)>=2,1,0)
          comercial$`1MORA30M3`=ifelse(apply(cp[,1:3],1,sum,na.rm=TRUE)==1,1,0)
          rm(cp)
          comercial$MORA1230=ifelse(apply(db_comercial[,20:31],1,max,na.rm=TRUE)>=31 & apply(db_comercial[,20:31],1,max,na.rm=TRUE)<=60,1,0) 
          comercial$MORA1260=ifelse(apply(db_comercial[,20:30],1,max,na.rm=TRUE)>60,1,0)
          comercial$MORA2430=ifelse(apply(db_comercial[,20:43],1,max,na.rm=TRUE)>=31 & apply(db_comercial[,20:43],1,max,na.rm=TRUE)<=60,1,0) 
          comercial$MORA2460=ifelse(apply(db_comercial[,20:43],1,max,na.rm=TRUE)>60,1,0)
          comercial$MORA3660=ifelse(apply(db_comercial[,20:55],1,max,na.rm=TRUE)>=31 & apply(db_comercial[,20:55],1,max,na.rm=TRUE)<=60,1,0)
          comercial$NODO4=ifelse(comercial$MORA1260==0 & comercial$MORA3660==0 & comercial$`1MORA30`==0,1,0)
          comercial$NODO5=ifelse(comercial$MORA1260==0 & comercial$MORA3660==0 & comercial$`1MORA30`==1,1,0)
          comercial$NODO7=ifelse(comercial$MORA1260==0 & comercial$MORA3660==1 & comercial$`1MORA30M3`==1,1,0)
          comercial$NODO8=ifelse(comercial$MORA1260==0 & comercial$MORA3660==1 & comercial$`1MORA30M3`==0 & comercial$`2MORA30`==0 ,1,0)
          comercial$NODO9=ifelse(comercial$MORA1260==0 & comercial$MORA3660==1 & comercial$`1MORA30M3`==0 & comercial$`2MORA30`==1 ,1,0)
          comercial$NODO1=ifelse(comercial$MORA1260==1 ,1,0)
          
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
          }
          if(dim(c_libranza)[1]==0) c_libranza=data.frame(x=c(0,0,0),y=c(0,0,0))
          if(dim(s_libranza)[1]==0) s_libranza=data.frame(x=c(0,0,0),y=c(0,0,0))
          if(dim(comercial)[1]==0) comercial=data.frame(x=c(0,0,0),y=c(0,0,0))
          df_list <- list(ConLibranza=c_libranza, SinLibranza=s_libranza, ComercialPN=comercial)
          write.xlsx(x = df_list , file = filename, rowNames = FALSE)
        }
      ) 
     })

    shinyApp(ui = ui, server = server)
