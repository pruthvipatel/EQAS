

#load the XLConnect package

setwd("C:/Users/Admin/Desktop/EQAS")


wb<-loadWorkbook ("nqa12.xlsx", create=TRUE)
QC<-readWorksheet(wb, sheet = "Sheet4", startRow = 2, endRow = 203, startCol
                  = 1, endCol = 65, header = FALSE)
attach (QC)



AlgorithmA<-function(X,Y,Z){
  count<-0
  temp<-array(dim = nrow(QC))
  for (i in 1:nrow(QC)){
    if (Z[i]!=1){
      count=count+1
      temp[count]<-X[i]
    }
  }

  xstar<-median(temp,na.rm=TRUE)
  sstar<-1.483*median(abs(temp-xstar), na.rm=TRUE)
  delta<-1.5*sstar
  l<-nrow(QC) 
  
  xistar<-array(dim = l)
  for (i in 1:l){
    if (is.na(temp[i])==TRUE){xistar[i]<-0}
    else if (temp[i]<(xstar-delta) && is.na(temp[i])==FALSE){xistar[i]<-(xstar-delta)}
    else if (temp[i]>(xstar+delta) && is.na(temp[i])==FALSE){xistar[i]<-(xstar+delta)}
    else xistar[i]<-temp[i]
  }
  
  xistar[xistar==0] <- NA
  
  #A = array (xistar)
  #B=previous xstar
  #C=previous sstar
  #D=original array
  
  iteration.n<-function (A,B,C,D){
    xstar<-mean(A,na.rm=TRUE)
    sstar<-1.134*sd(A,na.rm=TRUE)
    delta<-1.5*sstar
    #l<-nrow(QC)
    
    xistar<-array(dim = l)
    for (i in 1:l){
      if (is.na(D[i])==TRUE){xistar[i]<-0}
      else if (D[i]<(xstar-delta)&& is.na(D[i])==FALSE){xistar[i]<-(xstar-delta)}
      else if (D[i]>(xstar+delta)&& is.na(D[i])==FALSE){xistar[i]<-(xstar+delta)}
      else xistar[i]<-D[i]
    }
    
    xistar[xistar==0] <- NA
    
    
    chgSD<-abs(signif(B,digits = 3)-signif(xstar,digits=3))
    chgmean<-abs(signif(C,digits = 3)-signif(sstar,digits=3))
    
    if (chgSD>0 || chgmean>0){
      iteration.n (xistar,xstar,sstar,D)
    }
    else{ return (c(xstar,sstar))}
    
  }
  robust<-iteration.n(xistar,xstar,sstar,temp)
  zscore<-array(dim=l)
  result<-array(dim=l)
  zscore<-round(((X-robust[1])/robust[2]),digits=2)
  
  for (i in 1:l){
    
    if (abs(zscore[i])<2 && is.na(zscore[i])==FALSE){
      result[i]<- "satisfactory"
    }
    else if (abs(zscore[i])>=2 && abs(zscore[i])<3  && is.na(zscore[i])==FALSE){
      result[i]<-"warning"
    }
    else if (abs(zscore[i])>=3  && is.na(zscore[i])==FALSE){
      result[i]<-"unsatisfactory"
    }
  }
  df<-data.frame (zscore, result)
  r1<-round(robust[1],digits=2)
  r2<-round(robust[2],digits=2)
  for (i in 1:l){
    writeWorksheet(wb,r1,sheet="Sheet4",startRow = i+1,startCol = Y+2, header= FALSE)
    writeWorksheet(wb,r2,sheet="Sheet4",startRow = i+1,startCol = Y+3, header= FALSE)
  }
  writeWorksheet(wb,df,sheet="Sheet4",startRow = 2,startCol = Y, header= FALSE)#-------------here
  saveWorkbook(wb)
  return (list(robust[2],zscore))
}

RobustSD_CD3_29 <- AlgorithmA(Col16,17,Col58)
RobustSD_CD3_29_per <- AlgorithmA(Col21,22,Col59)
RobustSD_CD4_29 <- AlgorithmA(Col26,27,Col60)
RobustSD_CD4_29_per <- AlgorithmA(Col31,32,Col61)

RobustSD_CD3_30 <- AlgorithmA(Col36,37,Col62)
RobustSD_CD3_30_per <- AlgorithmA(Col41,42,Col63)
RobustSD_CD4_30 <- AlgorithmA(Col46,47,Col64)
RobustSD_CD4_30_per <- AlgorithmA(Col51,52,Col65)

RobustSD_CD3_29[1]
RobustSD_CD4_29[1]
RobustSD_CD3_29_per[1]
RobustSD_CD4_29_per[1]
RobustSD_CD3_30[1] 
RobustSD_CD4_30[1] 
RobustSD_CD3_30_per[1] 
RobustSD_CD4_30_per[1] 