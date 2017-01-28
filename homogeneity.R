#Hom_CD3_29<-Homogeneity(8,4,as.numeric(RobustSD_CD3_29[1]),10,2)
#Hom_CD3_29_per<-Homogeneity(8,11,as.numeric(RobustSD_CD3_29_per[1]),10,2)
#Hom_CD4_29<-Homogeneity(8,18,as.numeric(RobustSD_CD3_29_per[1]),10,2)
#Hom_CD4_29_per<-Homogeneity(8,25,as.numeric(RobustSD_CD3_29_per[1]),10,2)
#Hom_CD3_30<-Homogeneity(56,4,as.numeric(RobustSD_CD3_30[1]),10,2)
#Hom_CD3_30_per<-Homogeneity(56,11,as.numeric(RobustSD_CD3_30_per[1]),10,2)
#Hom_CD4_30<-Homogeneity(56,18,as.numeric(RobustSD_CD4_30[1]),10,2)
#Hom_CD4_30_per<-Homogeneity(56,25,as.numeric(RobustSD_CD4_30_per[1]),10,2)

setwd("C:/Users/Admin/Desktop/EQAS")


wb<-loadWorkbook ("stability.xlsx")


Homogeneity<-function(X,Y,Z,A,B){
  #t<-as.numeric(readline("Enter the number of proficiency test items: "))
  #k<-as.numeric(readline("Enter the number of test portions: "))
  t<-A
  k<-B
  Read<-readWorksheet(wb, sheet = "Sheet1", startRow = X, endRow =(X+t-1) , startCol
                      = Y, endCol = (Y+k-1), header = FALSE)
  
  Readings<-data.matrix(Read)
  
  if (k==2){
    
    Sample.avg<-array(dim=t)
    Sample.diff<-array(dim=t)
    Sum.sq<-0
    for (i in 1:t){
      Sample.avg[i]<-(Readings[i,1]+Readings[i,2])/2
      Sample.diff[i]<-abs(Readings[i,1]-Readings[i,2])
      Sum.sq<-Sum.sq+Sample.diff[i]^2
    }
    General.avg<-mean(Sample.avg)
    s.x<-sd(Sample.avg)
    s.w<-sqrt(Sum.sq/(t*2))
    s.s.sq<-max(0,(s.x^2-s.w^2/2))
    s.s<-sqrt(s.s.sq)
    
  }
  
  else if (k>2){
    Sample.var<- array(dim=c(t))
    Sample.diff.var<- array(dim=c(t))
    for (i in 1:t){
      Sample.avg[i]<-mean(Readings[i])
      
      
      for (j in 1:k){
        Sample.var[i]<-Sample.var[i]+(Readings[i,j]-Sample.avg[i])^2
      }
      Sample.diff.var[i]<-Sample.var[i]/(k-1)
      Sample.var[i]<-Sample.var[i]/k
    }
    General.avg=Sample.sum/t
    s.x<-sd(Sample.avg)
    s.w<-sqrt(mean(Sample.avg))
    s.s.sq<-max(0,(s.x^2-s.w^2/k))
    s.s<-sqrt(s.s.sq)
    
  }
  if ((0.3*Z-s.s)>0){
    st<-"ok"
  }
  else{
    st<-"not ok"
  }
  
  writeWorksheet(wb,round(as.numeric(Sample.avg),digits=2),sheet="Sheet1",startRow = X,startCol = (Y+2), header= FALSE)
  writeWorksheet(wb,round(as.numeric(Sample.diff),digits=2),sheet="Sheet1",startRow = X,startCol = (Y+3), header= FALSE)
  writeWorksheet(wb,round(General.avg,digits=2),sheet="Sheet1",startRow = (X+t),startCol =(Y+2), header=FALSE)
  writeWorksheet(wb,round(s.x,digits=2),sheet="Sheet1",startRow = (X+t+1),startCol =(Y+2), header=FALSE)
  writeWorksheet(wb,round(s.w,digits=2),sheet="Sheet1",startRow = (X+t+2),startCol =(Y+2), header=FALSE)
  writeWorksheet(wb,round(s.s,digits=2),sheet="Sheet1",startRow = (X+t+3),startCol =(Y+2),header=FALSE)
  writeWorksheet(wb,round((s.x^2),digits=2),sheet="Sheet1",startRow =(X+t+4),startCol =(Y+2), header=FALSE)
  writeWorksheet(wb,round((s.w^2),digits=2),sheet="Sheet1",startRow = (X+t+5),startCol =(Y+2), header=FALSE)
  writeWorksheet(wb,round((s.w^2/2),digits=2),sheet="Sheet1",startRow =(X+t+6),startCol =(Y+2), header=FALSE)
  writeWorksheet(wb,round(((s.x^2)-(s.w^2/2)),digits=2),sheet="Sheet1",startRow = (X+t+7),startCol =(Y+2), header=FALSE)
  writeWorksheet(wb,round(Z,digits=2),sheet="Sheet1",startRow = (X+t+8),startCol =(Y+2), header=FALSE)
  writeWorksheet(wb,round((0.3*Z),digits=2),sheet="Sheet1",startRow =(X+t+9),startCol =(Y+2),header=FALSE)
  writeWorksheet(wb,round(((0.3*Z)-(s.s)),digits=2),sheet="Sheet1",startRow = (X+t+10),startCol =(Y+2),header=FALSE)
  writeWorksheet(wb,st,sheet="Sheet1",startRow = (X+t+10),startCol =(Y+3),header=FALSE)
  saveWorkbook(wb)
  return (list(st,General.avg,s.x,t))
}

Hom_CD3_29<-Homogeneity(8,4,as.numeric(RobustSD_CD3_29[1]),10,2)
Hom_CD3_29_per<-Homogeneity(8,11,as.numeric(RobustSD_CD3_29_per[1]),10,2)
Hom_CD4_29<-Homogeneity(8,18,as.numeric(RobustSD_CD3_29_per[1]),10,2)
Hom_CD4_29_per<-Homogeneity(8,25,as.numeric(RobustSD_CD3_29_per[1]),10,2)
Hom_CD3_30<-Homogeneity(56,4,as.numeric(RobustSD_CD3_30[1]),10,2)
Hom_CD3_30_per<-Homogeneity(56,11,as.numeric(RobustSD_CD3_30_per[1]),10,2)
Hom_CD4_30<-Homogeneity(56,18,as.numeric(RobustSD_CD4_30[1]),10,2)
Hom_CD4_30_per<-Homogeneity(56,25,as.numeric(RobustSD_CD4_30_per[1]),10,2)