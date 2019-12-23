library(ggplot2)

setwd("C:/temp")

O365Data <- read.csv("O365UserLocationData.csv")



O365CountryDatawoUS <- O365Data[!grepl("United States|^$", O365Data$country),]

jpeg("Successful Logins per Country.jpg")

O365Country <- ggplot(O365CountryDatawoUS, aes(country))
O365Country + geom_bar(stat = "count", fill = "steelblue") + theme_gray() + labs(x="Country", y="Successful Logins", title="Successful Logins per Country Ouside of US")

dev.off()