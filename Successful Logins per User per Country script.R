library(ggplot2)

setwd("C:/temp")

O365Data <- read.csv("O365UserLocationData.csv")

O365CountryDatawoUS <- O365Data[!grepl("United States|^$", O365Data$country),]

jpeg("Successful Logins per User per Country.jpg")

O365CountryPerUser <- ggplot(O365CountryDatawoUS, aes(fill=O365CountryDatawoUS$country, x=O365CountryDatawoUS$UserId))
                             
O365CountryPerUser + geom_bar(stat = "count") + theme_dark() + labs(x="User Email Address", y="Successful Logins", title="Successful Logins per User per Country Ouside of US") + theme(axis.text.x = element_text(angle = -90)) + scale_fill_brewer(name = "Countries")

dev.off()