setwd("../Proyecto_TresChanchitos")

# 1.Importación y agregación de las columnas del archivo Calendario 

library(openxlsx)
library(readxl)
library(dplyr)
library(lubridate)

calendario <- read_excel("../Proyecto_TresChanchitos/Calendario.xlsx")
calendario

calendario$Fecha <- as.Date(calendario$Fecha)
calendario <- calendario %>% 
  mutate(Nro.Mes = month(Fecha),
         Nombre_Mes = month(Fecha, 
                            label = TRUE,
                            abbr = FALSE),
         Año = year(Fecha))
calendario

# 2.Importación y unificación del archivo Clientes

clientes1 <- read_excel("./Cliente.xlsx", sheet = "Clientes1")
clientes2 <- read_excel("./Cliente.xlsx", sheet = "Clientes2")
clientes3 <- read_excel("./Cliente.xlsx", sheet = "Clientes3")

clientes1
clientes2
clientes3

Clientes <- clientes1 %>%
  full_join(clientes2, by = "IdCliente") %>%
  full_join(clientes3, by = "IdCliente")

Clientes

# 3. Importación del archivo Territorio

Territorio <- read.csv2("../Proyecto_TresChanchitos/Territorio.csv", 
                        sep = ";", 
                        stringsAsFactors = FALSE)
Territorio


Territorio$Place <- iconv(Territorio$Place, from = "latin1", to = "UTF-8")


# 4. Importación del archivo Vendedor

Vendedor <- read_xlsx("./Vendedor.xlsx")

Vendedor
Vendedor$Apellido <- tools::toTitleCase(tolower(Vendedor$Apellido))

# 5. Importación del archivo Productos

Productos <- read_xlsx("./Productos.xlsx")
Productos

# 6. Importación de la carpeta Data en un solo Dataframe

Data2017 <- read_excel("./Data/2017.xlsx")
Data2018 <- read_excel("./Data/2018.xlsx")
Data2019 <- read_excel("./Data/2019.xlsx")

Data <- bind_rows(Data2017, Data2018, Data2019)
Data$Fecha <- as.Date(Data$Fecha)
Data

# 7. Guardado de los dataframe en Excel


dir.create("Dataframe")

write_xlsx(calendario, "Dataframe/Calendario.xlsx")
write_xlsx(Clientes, "Dataframe/Clientes.xlsx")
write_xlsx(Territorio, "Dataframe/Territorio.xlsx")
write_xlsx(Vendedor, "Dataframe/Vendedor.xlsx")
write_xlsx(Productos, "Dataframe/Productos.xlsx")
write_xlsx(Data, "Dataframe/Data.xlsx")

