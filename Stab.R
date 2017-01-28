#x is General average (Homogeneity)
#Y is sigma pt.(AlgoA)
#Z is sd ofsample avg.(Homogeneity)
#W is no. of pt items (Homogeneity)

#-------------------------------------------------------------------------
#Stab_CD3_29<-Stability(as.numeric(Hom_CD3_29[2]),as.numeric(RobustSD_CD3_29[1]),as.numeric(Hom_CD3_29[3]),as.numeric(Hom_CD3_29[4]),8,4,3,2)
#Stab_CD3_29_per<-Stability(as.numeric(Hom_CD3_29_per[2]),as.numeric(RobustSD_CD3_29_per[1]),as.numeric(Hom_CD3_29_per[3]),as.numeric(Hom_CD3_29_per[4]),8,5,3,2)
#Stab_CD4_29<-Stability(as.numeric(Hom_CD4_29[2]),as.numeric(RobustSD_CD4_29[1]),as.numeric(Hom_CD4_29[3]),as.numeric(Hom_CD4_29[4]),8,6,3,2)
#Stab_CD4_29_per<-Stability(as.numeric(Hom_CD4_29_per[2]),as.numeric(RobustSD_CD4_29_per[1]),as.numeric(Hom_CD4_29_per[3]),as.numeric(Hom_CD4_29_per[4]),8,7,3,2)
#Stab_CD3_30<-Stability(as.numeric(Hom_CD3_30[2]),as.numeric(RobustSD_CD3_30[1]),as.numeric(Hom_CD3_30[3]),as.numeric(Hom_CD3_30[4]),8,12,3,2)
#Stab_CD3_30_per<-Stability(as.numeric(Hom_CD3_30_per[2]),as.numeric(RobustSD_CD3_30_per[1]),as.numeric(Hom_CD3_30_per[3]),as.numeric(Hom_CD3_30_per[4]),8,13,3,2)
#Stab_CD4_30<-Stability(as.numeric(Hom_CD4_30[2]),as.numeric(RobustSD_CD4_30[1]),as.numeric(Hom_CD4_30[3]),as.numeric(Hom_CD4_30[4]),8,14,3,2)
#Stab_CD4_30_per<-Stability(as.numeric(Hom_CD4_30_per[2]),as.numeric(RobustSD_CD4_30_per[1]),as.numeric(Hom_CD4_30_per[3]),as.numeric(Hom_CD4_30_per[4]),8,15,3,2)
#-------------------------------------------------------------------------

setwd("C:/Users/Admin/Desktop/EQAS")


wb<-loadWorkbook ("stability.xlsx")

Stability<-function(X,Y,Z,W,A,C,T,K){
  #t<-as.numeric(readline("Enter the number of proficiency test items: "))
  #k<-as.numeric(readline("Enter the number of test portions: "))
  t<-T
  k<-K
  Read<-readWorksheet(wb, sheet = "Sheet2", startRow = A, endRow =(A+t-1) , startCol
                      = C, endCol = (C+k-1), header = FALSE)
  
  Readings<-data.matrix(Read)
  
  y2bar<-mean(Readings)
  
  B<-abs(X-y2bar)
  
  writeWorksheet(wb,y2bar,sheet="Sheet2",startRow = (A+t+k+1),startCol = C, header= FALSE)
  writeWorksheet(wb,X,sheet="Sheet2",startRow = (A+t+k+2),startCol = C, header= FALSE)
  writeWorksheet(wb,(0.3*Y),sheet="Sheet2",startRow = (A+t+k+3),startCol = C, header= FALSE)
  writeWorksheet(wb,B,sheet="Sheet2",startRow = (A+t+k+4),startCol = C, header= FALSE)
  
  if ((0.3*Y-B)>0){
    print ("Stable")
    writeWorksheet(wb,(0.3*Y-B),sheet="Sheet2",startRow = (A+t+k+4),startCol = C, header= FALSE)
  }
  else{
    CF<-2*sqrt((1.25*sd(Readings)/sqrt(t+k))^2+(1.25*Z/sqrt(W))^2)
    writeWorksheet(wb,round((0.3*Y-B),2),sheet="Sheet2",startRow = (A+t+k+5),startCol = C, header= FALSE)
    writeWorksheet(wb,round(CF,2),sheet="Sheet2",startRow = (A+t+k+6),startCol = C, header= FALSE)
    writeWorksheet(wb,round((1.25*sd(Readings)/sqrt(t+k)),2),sheet="Sheet2",startRow = (A+t+k+7),startCol = C, header= FALSE)
    writeWorksheet(wb,round((1.25*Z/sqrt(W)),2),sheet="Sheet2",startRow = (A+t+k+8),startCol = C, header= FALSE)
    writeWorksheet(wb,round((0.3*Y+CF-B),2),sheet="Sheet2",startRow = (A+t+k+9),startCol = C, header= FALSE)
    
    saveWorkbook(wb)
    if ((0.3*Y+CF-B)>0){
      st<-"Stable"
      
    }
    else
    {
      st<-"unstable"
    }
  }
  
  return(st)
  
}

Stab_CD3_29<-Stability(as.numeric(Hom_CD3_29[2]),as.numeric(RobustSD_CD3_29[1]),as.numeric(Hom_CD3_29[3]),as.numeric(Hom_CD3_29[4]),8,4,3,2)
Stab_CD3_29_per<-Stability(as.numeric(Hom_CD3_29_per[2]),as.numeric(RobustSD_CD3_29_per[1]),as.numeric(Hom_CD3_29_per[3]),as.numeric(Hom_CD3_29_per[4]),8,5,3,2)
Stab_CD4_29<-Stability(as.numeric(Hom_CD4_29[2]),as.numeric(RobustSD_CD4_29[1]),as.numeric(Hom_CD4_29[3]),as.numeric(Hom_CD4_29[4]),8,6,3,2)
Stab_CD4_29_per<-Stability(as.numeric(Hom_CD4_29_per[2]),as.numeric(RobustSD_CD4_29_per[1]),as.numeric(Hom_CD4_29_per[3]),as.numeric(Hom_CD4_29_per[4]),8,7,3,2)
Stab_CD3_30<-Stability(as.numeric(Hom_CD3_30[2]),as.numeric(RobustSD_CD3_30[1]),as.numeric(Hom_CD3_30[3]),as.numeric(Hom_CD3_30[4]),8,12,3,2)
Stab_CD3_30_per<-Stability(as.numeric(Hom_CD3_30_per[2]),as.numeric(RobustSD_CD3_30_per[1]),as.numeric(Hom_CD3_30_per[3]),as.numeric(Hom_CD3_30_per[4]),8,13,3,2)
Stab_CD4_30<-Stability(as.numeric(Hom_CD4_30[2]),as.numeric(RobustSD_CD4_30[1]),as.numeric(Hom_CD4_30[3]),as.numeric(Hom_CD4_30[4]),8,14,3,2)
Stab_CD4_30_per<-Stability(as.numeric(Hom_CD4_30_per[2]),as.numeric(RobustSD_CD4_30_per[1]),as.numeric(Hom_CD4_30_per[3]),as.numeric(Hom_CD4_30_per[4]),8,15,3,2)
