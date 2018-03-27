

library(xlsx)
library(gdata)
library(lubridate)
setwd("X:/Research/Analytics/Bases TV Semanal/_Temporal")

 
#################################################### Ranking TVA LD 7 targets ####################################################

db <- read.xlsx("Ranking TVA LD 7 Targets.xlsx",1)

Hojas <- db[11,seq(10,22,2)]

db <- db[-c(1:12),]

x <- 1
x1 <- NULL
while (!(is.null(db[x,1]) || is.na(db[x,1]))) {
  x1 <- c(x,x1)
  x <- x+1
}
y <- max(x1)

db <- db[c(1:y),1:23]

colnames(db) <- c("Id","Grupo","Canal","Programa","Tema","Período","Días","Inicio","Final",
                  paste(as.character(Hojas[1,1]),"Ranking",sep = " "), 
                  paste(as.character(Hojas[1,1]),"Afinidad",sep = " "),
                  paste(as.character(Hojas[1,2]),"Ranking",sep = " "),
                  paste(as.character(Hojas[1,2]),"Afinidad",sep = " "),
                  paste(as.character(Hojas[1,3]),"Ranking",sep = " "),
                  paste(as.character(Hojas[1,3]),"Afinidad",sep = " "),
                  paste(as.character(Hojas[1,4]),"Ranking",sep = " "),
                  paste(as.character(Hojas[1,4]),"Afinidad",sep = " "),
                  paste(as.character(Hojas[1,5]),"Ranking",sep = " "),
                  paste(as.character(Hojas[1,5]),"Afinidad",sep = " "),
                  paste(as.character(Hojas[1,6]),"Ranking",sep = " "),
                  paste(as.character(Hojas[1,6]),"Afinidad",sep = " "),
                  paste(as.character(Hojas[1,7]),"Ranking",sep = " "),
                  paste(as.character(Hojas[1,7]),"Afinidad",sep = " "))

db$Período <- as.Date(as.numeric(as.character(db$Período)), origin = "1899-12-30")
db$Inicio <- chron::times(as.numeric(as.character(db$Inicio)))
db$Final <- chron::times(as.numeric(as.character(db$Final)))

h1 <- db[,c(1:9,10,11)]
h2 <- db[,c(1:9,12,13)]
h3 <- db[,c(1:9,14,15)]
h4 <- db[,c(1:9,16,17)]
h5 <- db[,c(1:9,18,19)]
h6 <- db[,c(1:9,20,21)]
h7 <- db[,c(1:9,22,23)]

setwd("X:/Research/Analytics/Bases TV Semanal/_Temporal/Bases")
write.xlsx(h1, "Ranking TVA LD 7 prueba.xlsx", sheetName = as.character(Hojas[1,1]), row.names=FALSE)
write.xlsx(h2, "Ranking TVA LD 7 prueba.xlsx", sheetName = as.character(Hojas[1,2]), row.names=FALSE, append=TRUE)
write.xlsx(h3, "Ranking TVA LD 7 prueba.xlsx", sheetName = as.character(Hojas[1,3]), row.names=FALSE, append=TRUE)
write.xlsx(h4, "Ranking TVA LD 7 prueba.xlsx", sheetName = as.character(Hojas[1,4]), row.names=FALSE, append=TRUE)
write.xlsx(h5, "Ranking TVA LD 7 prueba.xlsx", sheetName = as.character(Hojas[1,5]), row.names=FALSE, append=TRUE)
write.xlsx(h6, "Ranking TVA LD 7 prueba.xlsx", sheetName = as.character(Hojas[1,6]), row.names=FALSE, append=TRUE)
write.xlsx(h7, "Ranking TVA LD 7 prueba.xlsx", sheetName = as.character(Hojas[1,7]), row.names=FALSE, append=TRUE)

remove(h1,h2,h3,h4,h5,h6,h7,db,y,Hojas,x,x1)
#################################################### Ranking TVP LV 7 Targets ####################################################
setwd("X:/Research/Analytics/Bases TV Semanal/_Temporal")
db2 <- read.xlsx("Ranking TVP LV 7 Targets.xlsx",1)

Campos <- db2[11,seq(8,20,2)]

db2 <- db2[-c(1:12),]

x <- 1
x1 <- NULL
while (!(is.null(db2[x,1]) || is.na(db2[x,1]))) {
  x1 <- c(x,x1)
  x <- x+1
}
y <- max(x1)

db2 <- db2[c(1:y),1:21]

colnames(db2) <- c("Id","Grupo","Canal","Período","Días","Inicio","Final",
                  paste(as.character(Campos[1,1]),"Ranking",sep = " "), 
                  paste(as.character(Campos[1,1]),"Afinidad",sep = " "),
                  paste(as.character(Campos[1,2]),"Ranking",sep = " "),
                  paste(as.character(Campos[1,2]),"Afinidad",sep = " "),
                  paste(as.character(Campos[1,3]),"Ranking",sep = " "),
                  paste(as.character(Campos[1,3]),"Afinidad",sep = " "),
                  paste(as.character(Campos[1,4]),"Ranking",sep = " "),
                  paste(as.character(Campos[1,4]),"Afinidad",sep = " "),
                  paste(as.character(Campos[1,5]),"Ranking",sep = " "),
                  paste(as.character(Campos[1,5]),"Afinidad",sep = " "),
                  paste(as.character(Campos[1,6]),"Ranking",sep = " "),
                  paste(as.character(Campos[1,6]),"Afinidad",sep = " "),
                  paste(as.character(Campos[1,7]),"Ranking",sep = " "),
                  paste(as.character(Campos[1,7]),"Afinidad",sep = " "))


db2$Inicio <- chron::times(as.numeric(as.character(db2$Inicio)))
db2$Final <- chron::times(as.numeric(as.character(db2$Final)))
setwd("X:/Research/Analytics/Bases TV Semanal/_Temporal/Bases")
write.xlsx(db2, "Ranking TVP LV 7 prueba.xlsx", sheetName = "Resultados", row.names=FALSE)

remove(db2,y,Campos)
#################################################### Ranking TVP SD 7 Targets ####################################################
setwd("X:/Research/Analytics/Bases TV Semanal/_Temporal")
db3 <- read.xlsx("Ranking TVP SD 7 Targets.xlsx",1)

Campos <- db3[11,seq(8,20,2)]

db3 <- db3[-c(1:12),]

x <- 1
x1 <- NULL
while (!(is.null(db3[x,1]) || is.na(db3[x,1]))) {
  x1 <- c(x,x1)
  x <- x+1
}
y <- max(x1)

db3 <- db3[c(1:y),1:21]

colnames(db3) <- c("Id","Grupo","Canal","Período","Días","Inicio","Final",
                   paste(as.character(Campos[1,1]),"Ranking",sep = " "), 
                   paste(as.character(Campos[1,1]),"Afinidad",sep = " "),
                   paste(as.character(Campos[1,2]),"Ranking",sep = " "),
                   paste(as.character(Campos[1,2]),"Afinidad",sep = " "),
                   paste(as.character(Campos[1,3]),"Ranking",sep = " "),
                   paste(as.character(Campos[1,3]),"Afinidad",sep = " "),
                   paste(as.character(Campos[1,4]),"Ranking",sep = " "),
                   paste(as.character(Campos[1,4]),"Afinidad",sep = " "),
                   paste(as.character(Campos[1,5]),"Ranking",sep = " "),
                   paste(as.character(Campos[1,5]),"Afinidad",sep = " "),
                   paste(as.character(Campos[1,6]),"Ranking",sep = " "),
                   paste(as.character(Campos[1,6]),"Afinidad",sep = " "),
                   paste(as.character(Campos[1,7]),"Ranking",sep = " "),
                   paste(as.character(Campos[1,7]),"Afinidad",sep = " "))


db3$Inicio <- chron::times(as.numeric(as.character(db3$Inicio)))
db3$Final <- chron::times(as.numeric(as.character(db3$Final)))
setwd("X:/Research/Analytics/Bases TV Semanal/_Temporal/Bases")
write.xlsx(db3, "Ranking TVP SD 7 prueba.xlsx", sheetName = "Resultados", row.names=FALSE)

remove(db3,y,Campos)
#################################################### Ranking TVA LD DDC ####################################################
setwd("X:/Research/Analytics/Bases TV Semanal/_Temporal")

db4 <- read.xlsx("Ranking TVA LD DDC.xlsx",1)

db4 <- db4[-c(1:12),]

x <- 1
x1 <- NULL
while (!(is.null(db4[x,1]) || is.na(db4[x,1]))) {
  x1 <- c(x,x1)
  x <- x+1
}
y <- max(x1)

db4 <- db4[c(1:y),1:11]

colnames(db4) <- c("Id","Grupo","Canal","Programa","Tema","Período","Días","Inicio","Final","Ranking","Afinidad")
          
db4$Período <- as.Date(as.numeric(as.character(db4$Período)), origin = "1899-12-30")
db4$Inicio <- chron::times(as.numeric(as.character(db4$Inicio)))
db4$Final <- chron::times(as.numeric(as.character(db4$Final)))
setwd("X:/Research/Analytics/Bases TV Semanal/_Temporal/Bases")
write.xlsx(db4, "Ranking TVA LD DDC prueba.xlsx", sheetName = "Resultados", row.names=FALSE)

remove(db4,y)
#################################################### Ranking TVP LV DDC ####################################################
setwd("X:/Research/Analytics/Bases TV Semanal/_Temporal")

db5 <- read.xlsx("Ranking TVP LV DDC.xlsx",1)

db5 <- db5[-c(1:12),]

x <- 1
x1 <- NULL
while (!(is.null(db5[x,1]) || is.na(db5[x,1]))) {
  x1 <- c(x,x1)
  x <- x+1
}
y <- max(x1)

db5 <- db5[c(1:y),1:9]

colnames(db5) <- c("Id","Grupo","Canal","Período","Días","Inicio","Final","Ranking","Afinidad")

db5$Inicio <- chron::times(as.numeric(as.character(db5$Inicio)))
db5$Final <- chron::times(as.numeric(as.character(db5$Final)))
setwd("X:/Research/Analytics/Bases TV Semanal/_Temporal/Bases")
write.xlsx(db5, "Ranking TVP LV DDC prueba.xlsx", sheetName = "Resultados", row.names=FALSE)

remove(db5,y)
#################################################### Ranking TVP SD DDC ####################################################
setwd("X:/Research/Analytics/Bases TV Semanal/_Temporal")

db6 <- read.xlsx("Ranking TVP SD DDC.xlsx",1)

db6 <- db6[-c(1:12),]

x <- 1
x1 <- NULL
while (!(is.null(db6[x,1]) || is.na(db6[x,1]))) {
  x1 <- c(x,x1)
  x <- x+1
}
y <- max(x1)

db6 <- db6[c(1:y),1:9]

colnames(db6) <- c("Id","Grupo","Canal","Período","Días","Inicio","Final","Ranking","Afinidad")

db6$Inicio <- chron::times(as.numeric(as.character(db6$Inicio)))
db6$Final <- chron::times(as.numeric(as.character(db6$Final)))
setwd("X:/Research/Analytics/Bases TV Semanal/_Temporal/Bases")
write.xlsx(db6, "Ranking TVP SD DDC prueba.xlsx", sheetName = "Resultados", row.names=FALSE)

remove(db6,y)

