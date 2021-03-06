#Cargo Paquete
library(readxl)

#Leo el balance
#quito notacion cientfica
options(scipen=999)
#hubo un problema al leer el txt con el BSR 1917124
#pues  la descripcion tenia una " sin cerrar
balance <- read.delim2(paste(getwd(),"balance.txt",sep = "/"))
#convierto en fecha la columna FECHA.ALPHA
balance$FECHA_ALHA <- as.Date(as.character(balance$FECHA_ALHA),format = "%d/%m/%Y")
#convierto en factor columnas BSR, BCRA, NIIF y Rubro.Activo
balance$BSR <- as.factor(balance$BSR)
balance$BCRA <- as.factor(balance$BCRA)
balance$NIIF <- as.factor(balance$NIIF)
balance$Rubro.Activo <- as.factor(balance$Rubro.Activo)


#LEO PESTANA APRC-SALDOS
APRC <- read_excel(path = paste(getwd(),"excel.xlsx",sep = "/"),sheet = 13,range = "A2:R62",col_names = TRUE)

#selecciono columnas para la tabla
APRC1 <- as.data.frame(APRC[2:28,1:16])
#nuevos nombres
names(APRC1)[4:14] <- c("0%","2%","4%","20%","35%","50%","75%","100%","125%","150%","1250%")
#reemplazo NA's
APRC1[which(is.na(APRC1[,1])),1] <- ""
APRC1[which(is.na(APRC1[,2])),2] <- ""
APRC1[which(APRC1[,2]!=""),2] <- c("0%","20%","40%","50%","90%","100%")
APRC1[which(is.na(APRC1[,3])),3] <- APRC1[17,3]
APRC1[,4:16] <- 0

#convierto el codigo en factor
APRC1[,1] <- as.factor(APRC1[,1])

#leo y automatizo pestaña BG
BG <- read_excel(path = paste(getwd(),"excel.xlsx",sep = "/"),sheet = 22,range = "B1:S1590",col_names = TRUE)

#relleno columna "Concatenar Partida/Ponderador"
for(i in 1:nrow(BG)){
  BG[i,7] <-  paste(as.character(BG[i,5]),as.character(BG[i,6]),sep = "")
}

#relleno columna "Saldo NOV 2018"
#creo variable a para guardar valores de la columna "Saldo NOV 2018"
a <- c()

for(i in 1:nrow(BG)){
  a[i] <- sum(balance[which(as.character(BG[i,3])==balance[,6]),7])/1000
}

#asigno valores a la data BG
BG[,4] <- a

#detecto celdas diferentes, en este caso
#las celdas son aquellas que no se leen del balance
#sino que su formula es una suma de celdas anteriores
cu <- which(!is.na(BG[,1]))

#por otra parte existen otro tipo de
#celdas diferentes, las mimas son 
#40, 50, 665, 880, 898, 1139, 1370-1452 (suma una celda), 1512, 1525 
#1579 (celda suma de dos columnas, dos sumar si), 1587, 1589 
#13 celdas se deben hacer manual
#estas celdas no son sumas de celdas anteriores, sino que son en algunos 
#casos suma de celdas de diferentes posiciones

cd <- c(40,50,665,880,898,1139,1370,1452,1512,1525,1579,1587,1589)

#celdas suma conjunto inmediatamente anterior
#el vector a se ira recortando con el fin de sumar lo inmediatamente anterior

for(i in 1:length(cu)){
  BG[cu[i],4] <- sum(a[1:(cu[i]-1)],na.rm = TRUE)
  a[c(1:cu[i])] <- 0
}

#realizo tratamiento de celdas diferentes
#formulo "formulas" con celdas especificas
BG[cd[1],4] <- sum(BG[c(23,38),4])
BG[cd[2],4] <- sum(BG[c(16,40,48),4])
BG[cd[3],4] <- sum(BG[c(286,298,370,663),4])
BG[cd[4],4] <- sum(BG[c(50,72,103,132,194,665,706,748,754,771,788,817,828,832,872,878),4])
BG[cd[5],4] <- sum(BG[c(880,887,896),4])
BG[cd[6],4] <- sum(BG[c(1137,994,964),4])
BG[cd[7],4] <- sum(BG[c(1368),4])
BG[cd[8],4] <- sum(BG[c(1450),4])
BG[cd[9],4] <- sum(BG[c(1139,1152,1183,1213,1293,1337,1366,1370,1436,1448,1452,1510),4])
BG[cd[10],4] <- sum(BG[c(1512,1521),4])
#OJO
BG[cd[11],4] <- (sum(balance[which(5==balance[,1]),7])+sum(balance[which(6==balance[,1]),7]))/1000
BG[cd[12],4] <- sum(BG[c(1563,1556,1546,1539,1532,1579,1577),4])
BG[cd[13],4] <- sum(BG[c(1587,1525),4])

#CONTROL
#ver excel celda E1592
#notar que coincide
sum(BG[c(898,1589),4])

#sumas
#1
z1 <- sum(balance[which(1==balance[,1]),7])/1000

#2
z2 <- sum(balance[which(2==balance[,1]),7])/1000

#3
z3 <- sum(balance[which(3==balance[,1]),7])/1000

#4
z4 <- sum(balance[which(4==balance[,1]),7])/1000

#5
z5 <- sum(balance[which(5==balance[,1]),7])/1000

#6
z6 <- sum(balance[which(6==balance[,1]),7])/1000

z1+z2+z3+z4+z5+z6

#Control Activo
BG[898,4]-z1-z2
#Control Pasivo
BG[1525,4]-z3
#Control PN
BG[1587,4]-z4-z5-z6

#RELLENO COLUMNAS PONDERADOR
#ARREGLO NOMBRE DE COLUMNAS Y ASIGNO 0
names(BG)[8:18] <- c("0%","2%","4%","20%","35%","50%","75%","100%","125%","150%","1250%")
BG[,8:18] <- 0

BG[which(BG$Ponderador==0),8] <- BG[which(BG$Ponderador==0),4]
BG[which(BG$Ponderador==0.02),9] <- BG[which(BG$Ponderador==0.02),4]
BG[which(BG$Ponderador==0.04),10] <- BG[which(BG$Ponderador==0.04),4]
BG[which(BG$Ponderador==0.2),11] <- BG[which(BG$Ponderador==0.2),4]
BG[which(BG$Ponderador==0.35),12] <- BG[which(BG$Ponderador==0.35),4]
BG[which(BG$Ponderador==0.5),13] <- BG[which(BG$Ponderador==0.5),4]
BG[which(BG$Ponderador==0.75),14] <- BG[which(BG$Ponderador==0.75),4]
BG[which(BG$Ponderador==1),15] <- BG[which(BG$Ponderador==1),4]
BG[which(BG$Ponderador==1.25),16] <- BG[which(BG$Ponderador==1.25),4]
BG[which(BG$Ponderador==1.5),17] <- BG[which(BG$Ponderador==1.5),4]
BG[which(BG$Ponderador==12.5),18] <- BG[which(BG$Ponderador==12.5),4]


#AUTOMATIZACION
for(j in 8:18){
  
  bi <- as.data.frame(BG[,j])
  
  for(i in 1:length(cu)){
    BG[cu[i],j] <- sum(bi[1:(cu[i]-1),1],na.rm = TRUE)
    bi[c(1:cu[i]),1] <- 0
  }
  
}

#NUEVAS CELDAS DIFERENTES (UNA PARTE DE LAS ANTERIORES)
cd1 <- cd[1:5]

#TRATO CELDAS DIFERENTES
for(i in 8:18){
  BG[cd1[1],i] <- sum(BG[c(23,38),i])
  BG[cd1[2],i] <- sum(BG[c(16,40,48),i])
  BG[cd1[3],i] <- sum(BG[c(286,298,370,663),i])
  BG[cd1[4],i] <- sum(BG[c(50,72,103,132,194,665,706,748,754,771,788,817,828,832,872,878),i])
  BG[cd1[5],i] <- sum(BG[c(880,887,896),i])
  
}

#LEO PESTANA APRC-SALDOS
#APRC <- read_excel(path = "C:/Users/Nancy/Desktop/Excel/excel.xlsx",sheet = 13,range = "A2:R62",col_names = TRUE)
APRC <- read_excel(path = paste(getwd(),"excel.xlsx",sep = "/"),sheet = 13,range = "A2:R62",col_names = TRUE)


APRC1 <- as.data.frame(APRC[2:28,1:16])
#nuevos nombres
names(APRC1)[4:14] <- c("0%","2%","4%","20%","35%","50%","75%","100%","125%","150%","1250%")
#reemplazo NA's
APRC1[which(is.na(APRC1[,1])),1] <- ""
APRC1[which(is.na(APRC1[,2])),2] <- ""
APRC1[which(APRC1[,2]!=""),2] <- c("0%","20%","40%","50%","90%","100%")
APRC1[which(is.na(APRC1[,3])),3] <- APRC1[17,3]
APRC1[,4:16] <- 0

#convierto el codigo en factor
APRC1[,1] <- as.factor(APRC1[,1])

str(APRC1)

#Relleno celdas de ponderadores
BG <- as.data.frame(BG)

pond <- c("0.0","0.002","0.004","0.2","0.35","0.50","0.75","1","1.25","1.50","12.5")

for(j in 1:length(pond)){
  for(i in 1:nrow(APRC1)){
    APRC1[i,3+j] <- sum(BG[which(as.numeric(paste(as.character(APRC1[i,1]),pond[j],sep = ""))==BG[,7]),7+j])
  }
}


