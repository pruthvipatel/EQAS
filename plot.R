setwd("C:/Users/Admin/Desktop/EQAS")
wb<-loadWorkbook ("nqa12.xlsx", create=TRUE)
QC<-readWorksheet(wb, sheet = "Sheet4", startRow = 2, endRow = 202, startCol#----here
                  = 1, endCol = 65, header = FALSE)
attach (QC)


plot<-function(X,Y){
  
  filepath<-file.path ("C:","Users","Admin","Desktop","EQAS","plot", Y)
  jpeg(file=filepath)
  hist (X, freq = T, xlab = Y, main=paste(Y,"performance", sep=" "))
  dev.off()
}

plot_CD3_29 <- plot(Col17,"QC29_CD3")
plot_CD3_29_per <- plot(Col22,"QC29_CD3percent")
plot_CD4_29 <- plot(Col27,"QC29_CD4")
plot_CD4_29_per <- plot(Col32,"QC29_CD4percent")
plot_CD3_30 <- plot(Col37,"QC30_CD3")
plot_CD3_30_per <- plot(Col42,"QC30_CD3percent")
plot_CD4_30 <- plot(Col47,"QC30_CD4")
plot_CD4_30_per <- plot(Col52,"QC30_CD4percent")