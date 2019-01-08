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


