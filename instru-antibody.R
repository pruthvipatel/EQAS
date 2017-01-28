#load the XLConnect package

setwd("C:/Users/Admin/Desktop/EQAS")


wb<-loadWorkbook ("nqa12.xlsx", create=TRUE)
QC<-readWorksheet(wb, sheet = "Sheet4", startRow = 2, endRow = 202, startCol#----here
                  = 1, endCol = 51, header = FALSE)
attach (QC)

#X:Column for the parameter (CD3, cD4 etc.) to be evaluated
#Y:Start row of the cells from where results are appended
#Z:Start column of the cells from where results are appended
#W:1=CD3, 2=CD4
instruantibody<-function(X,Y,Z,W){
  FACScaliburcount<-0
  FACScountcount<-0
  Parteccount<-0
  P<-array(dim=nrow(QC))
  Q<-array(dim=nrow(QC))
  R<-array(dim=nrow(QC))
  A<-array(dim=nrow(QC))
  B<-array(dim=nrow(QC))
  C<-array(dim=nrow(QC))
  
  for (i in 1:nrow(QC)){
    if ((Col8[i]=="FACS Calibur " && is.na(X[i])==FALSE) || (Col8[i]=="Facscalibur" && is.na(X[i])==FALSE)){
      FACScaliburcount=FACScaliburcount+1
      P[FACScaliburcount]<-X[i]
    }
    else if ((Col8[i]=="FACS count " && is.na(X[i])==FALSE) || (Col8[i]=="FACS Count" && is.na(X[i])==FALSE) || (Col8[i]=="FACS COUNT" && is.na(X[i])==FALSE) || (Col8[i]=="FACS count" && is.na(X[i])==FALSE) || (Col8[i]=="FACS Count " && is.na(X[i])==FALSE) || (Col8[i]=="FacsCount" && is.na(X[i])==FALSE)){
      FACScountcount=FACScountcount+1
      Q[FACScountcount]<-X[i]
    }
    else if (Col8[i]=="Partec " && is.na(X[i])==FALSE){
      Parteccount=Parteccount+1
      R[Parteccount]<-X[i]
    }
  }
  P<-P[!is.na(P)]
  Q<-Q[!is.na(Q)]
  R<-R[!is.na(R)]
  
  PEcy5count<-0
  PEcount<-0
  APCcount<-0
  FITCcount<-0
  if (W==1){
    for (i in 1:nrow(QC)){
      if (Col12[i]=="FITC" && is.na(X[i])==FALSE){
        FITCcount=FITCcount+1
        A[FITCcount]<-X[i]
      }
      else if (Col12[i]=="PE-cy5" && is.na(X[i])==FALSE){
        PEcy5count=PEcy5count+1
        B[PEcy5count]<-X[i]
      }
    }
    A<-A[!is.na(A)]
    B<-B[!is.na(B)]
  }
  
  else if (W==2){
    for (i in 1:nrow(QC)){
      if (Col13[i]=="FITC" && is.na(X[i])==FALSE){
        FITCcount=FITCcount+1
        A[FITCcount]<-X[i]
      }
      else if (Col13[i]=="PE" && is.na(X[i])==FALSE){
        PEcount=PEcount+1
        B[PEcount]<-X[i]
      }
      else if (Col13[i]=="APC" && is.na(X[i])==FALSE){
        APCcount=APCcount+1
        C[APCcount]<-X[i]
      }
    }
    A<-A[!is.na(A)]
    B<-B[!is.na(B)]
    C<-C[!is.na(C)]
  }
  
  Algorithm<-function(X,Y){
    xstar<-median(X,na.rm=TRUE)
    sstar<-1.483*median(abs(X-xstar), na.rm=TRUE)
    delta<-1.5*sstar
    l<-Y 
    
    xistar<-array(dim = l)
    for (i in 1:l){
      if (is.na(X[i])==TRUE){xistar[i]<-0}
      else if (X[i]<(xstar-delta) && is.na(X[i])==FALSE){xistar[i]<-(xstar-delta)}
      else if (X[i]>(xstar+delta) && is.na(X[i])==FALSE){xistar[i]<-(xstar+delta)}
      else xistar[i]<-X[i]
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
      l<-Y
      
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
    robust<-iteration.n(xistar,xstar,sstar,X)
    return (robust)
  }
  
  append<-0
  if (FACScaliburcount>=20){
    append<-Algorithm(P,FACScaliburcount)
    for (i in 1:nrow(QC)){
    writeWorksheet(wb,FACScaliburcount,sheet="Sheet5",startRow = Y+i-1,startCol = Z, header= FALSE)
    writeWorksheet(wb,round(append[1],2),sheet="Sheet5",startRow = Y+i-1,startCol = Z+1, header= FALSE)
    writeWorksheet(wb,round(append[2],2),sheet="Sheet5",startRow = Y+i-1,startCol = Z+2, header= FALSE)
  }
  }
  
  if (FACScountcount>=20){
    append<-Algorithm(Q,FACScountcount)
    for (i in 1:nrow(QC)){
    writeWorksheet(wb,FACScountcount,sheet="Sheet5",startRow = Y+i-1,startCol =Z+3, header= FALSE)
    writeWorksheet(wb,round(append[1],2),sheet="Sheet5",startRow =Y+i-1,startCol = Z+4, header= FALSE)
    writeWorksheet(wb,round(append[2],2),sheet="Sheet5",startRow =Y+i-1,startCol = Z+5, header= FALSE)
    }
  }
  
  if (Parteccount>=20){
    append<-Algorithm(R,Parteccount)
    for (i in 1:nrow(QC)){
    writeWorksheet(wb,Parteccount,sheet="Sheet5",startRow = Y+i-1,startCol = Z+6, header= FALSE)
    writeWorksheet(wb,round(append[1],2),sheet="Sheet5",startRow =Y+i-1,startCol = Z+7, header= FALSE)
    writeWorksheet(wb,round(append[2],2),sheet="Sheet5",startRow =Y+i-1,startCol = Z+8, header= FALSE)
    }
  }
  
  
  
  
  
  
  
  
  
  
  if (FITCcount>=20 && W==1){
    append<-Algorithm(A,FITCcount)
    for (i in 1:nrow(QC)){
    writeWorksheet(wb,FITCcount,sheet="Sheet5",startRow = Y+i-1,startCol = Z+27, header= FALSE)
    writeWorksheet(wb,round(append[1],2),sheet="Sheet5",startRow = Y+i-1,startCol = Z+28, header= FALSE)
    writeWorksheet(wb,round(append[2],2),sheet="Sheet5",startRow = Y+i-1,startCol = Z+29, header= FALSE)
    }
  }
  
  if (PEcy5count>=20 && W==1){
    append<-Algorithm(B,PEcy5count)
    for (i in 1:nrow(QC)){
    writeWorksheet(wb,PEcy5count,sheet="Sheet5",startRow = Y+i-1,startCol = Z+30, header= FALSE)
    writeWorksheet(wb,round(append[1],2),sheet="Sheet5",startRow = Y+i-1,startCol = Z+31, header= FALSE)
    writeWorksheet(wb,round(append[2],2),sheet="Sheet5",startRow = Y+i-1,startCol = Z+32, header= FALSE)
    }
  }
  
  if (FITCcount>=20 && W==2){
    append<-Algorithm(A,FITCcount)
    for (i in 1:nrow(QC)){
    writeWorksheet(wb,FITCcount,sheet="Sheet5",startRow = Y+i-1,startCol = Z+27, header= FALSE)
    writeWorksheet(wb,round(append[1],2),sheet="Sheet5",startRow = Y+i-1,startCol = Z+28, header= FALSE)
    writeWorksheet(wb,round(append[2],2),sheet="Sheet5",startRow = Y+i-1,startCol = Z+29, header= FALSE)
    }
  }
  
  if (PEcount>=20 && W==2){
    append<-Algorithm(B,PEcount)
    for (i in 1:nrow(QC)){
    writeWorksheet(wb,PEcount,sheet="Sheet5",startRow = Y+i-1,startCol = Z+33, header= FALSE)
    writeWorksheet(wb,round(append[1],2),sheet="Sheet5",startRow = Y+i-1,startCol = Z+34, header= FALSE)
    writeWorksheet(wb,round(append[2],2),sheet="Sheet5",startRow = Y+i-1,startCol = Z+35, header= FALSE)
    }
  }
  
  if (APCcount>=20 && W==2){
    append<-Algorithm(C,APCcount)
    for (i in 1:nrow(QC)){
    writeWorksheet(wb,APCcount,sheet="Sheet5",startRow = Y+i-1,startCol = Z+36, header= FALSE)
    writeWorksheet(wb,round(append[1],2),sheet="Sheet5",startRow = Y+i-1,startCol = Z+37, header= FALSE)
    writeWorksheet(wb,round(append[2],2),sheet="Sheet5",startRow = Y+i-1,startCol = Z+38, header= FALSE)
    }
  }
  
  saveWorkbook(wb)
  return (Parteccount)
}


CD3QC29<-instruantibody (Col16,5,2,1)
CD3perQC29 <-instruantibody (Col21,5,11,1)
CD4QC29<-instruantibody (Col26,5,20,2)
CD4perQC29 <-instruantibody (Col31,5,29,2)
CD3QC30<-instruantibody (Col36,5,86,1)
CD3perQC30<-instruantibody (Col41,5,95,1)
CD4QC30<-instruantibody (Col46,5,104,2)
CD4perQC30 <-instruantibody (Col51,5,113,2)


