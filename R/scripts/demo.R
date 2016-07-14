#install.packages('ReporteRs')
library('ReporteRs') # Load ReporteRs Package
pres.builder  <- pptx(template = '.../Master.pptx')
pres.filename <- ".../R-Meetup_Output.pptx"

# Show slide names
pres.builder

# Build Title slide
pres.builder  <- addSlide( pres.builder, "Title slide" ,bookmark = 1) # Slide name='Title slide', bookmark=1 <- overwrites first slide
pres.builder  <- addTitle( pres.builder, "R Meetup Example" ) #Add Title
pres.builder  <- addSubtitle( pres.builder , paste("Date:",Sys.Date())) # Add Subtitle

writeDoc( pres.builder, pres.filename )

# also possible to add page numbers (addPageNumber), Footer(addFooter) if added in MasterFile

pres.builder <- addSlide( pres.builder, slide.layout = 'Text 2') # add another slide, doesn't overwrite static content
pres.builder <- addTitle( pres.builder, "Example JobAd Performance")


# Add external image (or directly via RSelenium)
img.file <- file.path("C:/Users/david.geistert/Desktop/R-Meetup/R_LOGO.png" )
pres.builder <- addImage(pres.builder, img.file)      # width = 3.4, height = 3
writeDoc( pres.builder, pres.filename )

# Add table (various options to format table)
pres.builder <- addFlexTable( doc = pres.builder, vanilla.table(head(iris))) #par.properties = parCenter()
pres.builder <- addFooter(pres.builder , "David Geistert - XING AG")
writeDoc( pres.builder, pres.filename )

# add editable vector graphic
pres.builder <- addSlide( pres.builder, "Index (white)")
pres.builder <- addTitle( pres.builder, "Editable Vector Graphic")
pres.builder <- addPlot( pres.builder, function( ) hist(iris$Sepal.Width, col="lightblue"), vector.graphic = TRUE )
writeDoc( pres.builder, pres.filename )
pres.builder <- addSlide( pres.builder, "Index (white)")
pres.builder <- addTitle( pres.builder, "Print R-Code with (kind of) highlighted Syntax")
pres.builder <- addRScript(pres.builder, text = "x = rnorm(100)
  y <- c(seq(1:3))
  z <- sample(1:49,6,replace=F)
  " )


writeDoc( pres.builder, pres.filename )
