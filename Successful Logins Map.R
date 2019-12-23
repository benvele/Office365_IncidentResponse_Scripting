library(ggplot2)

setwd("C:/temp")

O365Data <- read.csv("O365UserLocationData.csv")

O365Coordinate <- paste0(O365Data$city, ", ", O365Data$country)
PSGeo <- geocode(location = PSCoordinate, output = "latlon", source = "google")
                 
plot(newmap, xlim = c(-75, 75), ylim = c(-75, 75))
points(PSGeo$lon, PSGeo$lat, col = "red")

ggmapmap <- ggmap(get_googlemap(center = c(lon = 0, lat = 0), zoom = 1, size = c(425, 425), maptype = 'terrain', color = 'color'))

ggmapmap + geom_point(aes(x = lon, y = lat, color = "red"), data = PSGeo, size = 1.75) + theme(legend.position = "none")

