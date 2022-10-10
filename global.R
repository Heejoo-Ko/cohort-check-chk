library(data.table)
library(magrittr)
library(readxl)
library(DT)
library(lubridate)
library(tableone)
library(openxlsx)
library(jstable)

setwd("/home/heejooko/ShinyApps/cohort-check-chk")

rd <- as.data.table(read_excel("ED_20220615.xlsx",sheet="Sheet1"))

# rd <- as.data.table(read_excel("LD_20220823.xlsx",sheet="DATA"))
# rd <- as.data.table(read_excel("PD_20220823.xlsx",sheet="DATA"))

# cbind(names(rd),names(rd2),names(rd3))

rd<-rd[-1,]

setnames(rd,"코호트 연구대상자번호","CohortNo")
#exclude error variables & caregiver variables for now
excludevar<-c("HXPSYEPI","HXPSYDEP","HXPSYMN","HXPSYSCH","HXPSYABUSE",
              "HXPSYALC","HXPSYHDINJ","HXPSYOTH","HXPSYOTHNAME",
              "HXDMCG","HXDMCGTR","HXHTCG","HXHTCGTR","HXHPLCG","HXHPLCGTR", 
              "HXHEARTDSCG","HXHEARTDSCGTR","HXSTROKECG","HXSTROKECGTR")
excludevar<-excludevar[excludevar %in% names(rd)]
rd<-rd[,.SD,.SDcols=!excludevar]

#Class---------------------------------------------------
date_vars<-c(grep("DAT",names(rd),value=T),"BOD","BODPRV","S2_SNSB2Date","PM_SNSB2Date")
date_vars<-date_vars[date_vars %in% names(rd)]
rd[,(date_vars):=lapply(.SD,function(x){
                              if(is.character(x)){
                                as.Date(as.numeric(x),origin="1899-12-30")
                              }else{
                                as.Date(x)
                              }
  
  }),.SDcols=date_vars]


factor_vars<-c("GENDER","EDUCAT","FMDM","HXHT","HXHTTR","HXSTR","HXSTRTR","HXHEART","HXHEARTTR","HXDM","HXDMTR","HXHL","HXHLTR","HXPSY","HXPSYEPI","HXPSYDEP","HXPSYMN","HXPSYSCH","HXPSYABUSE","HXPSYALC","HXPSYHDINJ","HXPSYOTH","HXPSYOTHNAME","DXMAIN","EDDX1","EDDX2","EDDX3","FTDDX1","FTDDX2","FTDDX3","OTDX1","OTDX2","OTDX3","CDR01","CDR02","CDR03","CDR04","CDR05","CDR06","BRAINMR","BRAINMRRS","PDGPET","PDGPETRS","AMYLOIDPET","AMYLOIDPETRS","SMOKETOTAL","SMOKECUR","HXDMCG","HXDMCGTR","HXHTCG","HXHTCGTR","HXHPLCG","HXHPLCGTR","HXHEARTDSCG","HXHEARTDSCGTR","HXSTROKECG","HXSTROKECGTR","LBAPOE","PTAMPTYP","PTAMPPOS","ISCHELVD","ISCHELVP","ISCHELVSVD","MTAGRT","MTAGLT","SNSBCERAD","AS2APYN","AS2APTYP")
factor_vars<-c("Center","Group","Level","Visit","CDRGLOBAL",factor_vars)
factor_vars<-factor_vars[factor_vars %in% names(rd)]

numeric_vars<-c("AGE","EDUYR","AGEPRV",
                "CGANPIFREQSUM","CGANPISVRSUM","CGANPIFREQSVRSUM","CGANPIBRDNSUM","BARTHELSCORE",
                "KMRSCORE","CDRSB","BMI","FTDCDRSCORE","FTDCDRSOB","KWABAQ","KWABLQ","UPDRSMSCORE","EQSCORE",
                "KIADLSCORE","KMMSE2SCORE","CDRSSCORE","CDRSSB",
                "S2_K_MMSE_total_score","S2_CDR","S2_B_ADL","S2_S1_GDS","S2_Sum_of_boxes","S2_SNSB_II_SGDS",
                "S2_K_IADL_Total_score","S2_K_IADL_NA_itemCount","S2_K_IADL_Score","S2_Global_DS",
                "PM_K_MMSE_total_score","PM_CDR","PM_B_ADL","PM_S1_GDS","PM_Sum_of_boxes","PM_GDepS",
                "PM_SGDS","PM_K_IADL_Total_score","PM_K_IADL_NA","PM_K_IADL_Score","PM_Global_DS")
numeric_vars<-numeric_vars[numeric_vars %in% names(rd)]

rd[,(factor_vars):=lapply(.SD,as.factor),.SDcols=factor_vars]
rd[,(numeric_vars):=lapply(.SD,as.numeric),.SDcols=numeric_vars]

#BRMRIDAT, PETDAT to Yes/No
rd[,BRMRIDAT_yesno:=as.factor(ifelse(is.na(BRMRIDAT),0,1)),]
rd[,PETDAT_yesno:=as.factor(ifelse(is.na(PETDAT),0,1)),]
factor_vars<-c(factor_vars,"BRMRIDAT_yesno","PETDAT_yesno")

#APOEe4
rd$LBAPOE %>% summary
rd[,APOEe4:=as.factor(ifelse(LBAPOE %in% c("E22","E23","E33"),0,
                             ifelse(LBAPOE %in% c("E24","E34"),1,
                                    ifelse(LBAPOE %in% c("E44"),2,NA))))]

#Other data error---------------------------------------------------
#date
rd[,(date_vars):=lapply(.SD,function(x){as.Date(ifelse(x>"2022-07-01",NA,x),origin="1970-01-01")}),
   .SDcols=date_vars]
#age
rd$AGEPRV<-ifelse(rd$AGEPRV>120 | rd$AGEPRV<0,NA,rd$AGEPRV)


output<-createWorkbook()
#Multivariate---------------------------------------------------

#Diagnosis
#checked no overlap between EDDX1, FTDDX1, OTDX1
rd$DXEXTRA<-as.factor(ifelse(!is.na(rd$EDDX1),as.character(rd$EDDX1),
                          ifelse(!is.na(rd$FTDDX1),as.character(rd$FTDDX1),
                                 ifelse(!is.na(rd$OTDX1),as.character(rd$OTDX1),NA))))
factor_vars<-c(factor_vars,"DXEXTRA")

addWorksheet(wb = output, sheetName = "2X2Table")

countrow<-1
writeData(wb = output, sheet="2X2Table", startRow=countrow, x="DXEXTRA X DXMAIN"); countrow<-countrow+1
mtb1<-as.matrix(table(rd$DXEXTRA,rd$DXMAIN,useNA = "always"))
rownames(mtb1)<-c(rd$DXEXTRA %>% levels,"NA"); colnames(mtb1)<-c(rd$DXMAIN %>% levels,"NA")
writeData(wb = output, sheet="2X2Table", startRow=countrow, x=mtb1, rowNames=TRUE, colNames = TRUE)
countrow<-countrow+nrow(mtb1)+3

writeData(wb = output, sheet="2X2Table", startRow=countrow, x="DXEXTRA X Group"); countrow<-countrow+1
mtb2<-as.matrix(table(rd$DXEXTRA,rd$Group,useNA = "always"))
rownames(mtb2)<-c(rd$DXEXTRA %>% levels,"NA"); colnames(mtb2)<-c(rd$Group %>% levels,"NA")
writeData(wb = output, sheet="2X2Table", startRow=countrow, x=mtb2, rowNames=TRUE, colNames=TRUE)
countrow<-countrow+nrow(mtb2)+3


#Smoking
writeData(wb = output, sheet="2X2Table", startRow=countrow, x="SMOKETOTAL X SMOKECUR"); countrow<-countrow+1
mtb3<-as.matrix(table(rd$SMOKETOTAL,rd$SMOKECUR,useNA="always"))
rownames(mtb3)<-c(rd$SMOKETOTAL %>% levels,"NA"); colnames(mtb3)<-c(rd$SMOKECUR %>% levels,"NA")
writeData(wb = output, sheet="2X2Table", startRow=countrow, x=mtb3, rowNames=TRUE, colNames=TRUE)
countrow<-countrow+nrow(mtb3)+3

#0 Never / 1 Past / 2 Current
rd$SMOKE3<-as.factor(ifelse(is.na(rd$SMOKETOTAL) & is.na(rd$SMOKECUR),NA,
                            ifelse(as.character(rd$SMOKETOTAL)=="0",0,
                                   ifelse(as.character(rd$SMOKECUR)=="3",1,
                                          ifelse((as.character(rd$SMOKECUR) %in% c("1","2")) | (as.character(rd$SMOKETOTAL) %in% c("1","2")),2,NA)))))
factor_vars<-c(factor_vars,"SMOKE3")


#HTN
writeData(wb = output, sheet="2X2Table", startRow=countrow, x="HXHT X HXHTTR"); countrow<-countrow+1
mtb4<-as.matrix(table(rd$HXHT,rd$HXHTTR,useNA="always"))
rownames(mtb4)<-c(rd$HXHT %>% levels,"NA"); colnames(mtb4)<-c(rd$HXHTTR %>% levels,"NA")
writeData(wb = output, sheet="2X2Table", startRow=countrow, x=mtb4, rowNames=TRUE, colNames=TRUE)
countrow<-countrow+nrow(mtb4)+3

rd$HXHTDX<-as.factor(ifelse(is.na(rd$HXHT) & is.na(rd$HXHTTR),NA,
                            ifelse(is.na(rd$HXHT),1,0)))
factor_vars<-c(factor_vars,"HXHTDX")


#DM
writeData(wb = output, sheet="2X2Table", startRow=countrow, x="HXDM X HXDMTR"); countrow<-countrow+1
mtb5<-as.matrix(table(rd$HXDM,rd$HXDMTR,useNA="always"))
rownames(mtb5)<-c(rd$HXDM %>% levels,"NA"); colnames(mtb5)<-c(rd$HXDMTR %>% levels,"NA")
writeData(wb = output, sheet="2X2Table", startRow=countrow, x=mtb5, rowNames=TRUE, colNames=TRUE)
countrow<-countrow+nrow(mtb5)+3

rd$HXDMDX<-as.factor(ifelse(is.na(rd$HXDM) & is.na(rd$HXDMTR),NA,
                            ifelse(is.na(rd$HXDM),1,0)))
factor_vars<-c(factor_vars,"HXDMDX")


#HL
writeData(wb = output, sheet="2X2Table", startRow=countrow, x="HXHL X HXHLTR"); countrow<-countrow+1
mtb6<-as.matrix(table(rd$HXHL,rd$HXHLTR,useNA="always"))
rownames(mtb6)<-c(rd$HXHL %>% levels,"NA"); colnames(mtb6)<-c(rd$HXHLTR %>% levels,"NA")
writeData(wb = output, sheet="2X2Table", startRow=countrow, x=mtb6, rowNames=TRUE, colNames=TRUE)
countrow<-countrow+nrow(mtb6)+3

rd$HXHLDX<-as.factor(ifelse(is.na(rd$HXHL) & is.na(rd$HXHLTR),NA,
                            ifelse(is.na(rd$HXHL),1,0)))
factor_vars<-c(factor_vars,"HXHLDX")


#Heart
writeData(wb = output, sheet="2X2Table", startRow=countrow, x="HXHEART X HXHEARTTR"); countrow<-countrow+1
mtb7<-as.matrix(table(rd$HXHEART,rd$HXHEARTTR,useNA="always"))
rownames(mtb7)<-c(rd$HXHEART %>% levels,"NA"); colnames(mtb7)<-c(rd$HXHEARTTR %>% levels,"NA")
writeData(wb = output, sheet="2X2Table", startRow=countrow, x=mtb7, rowNames=TRUE, colNames=TRUE)
countrow<-countrow+nrow(mtb7)+3

rd$HXHEARTDX<-as.factor(ifelse(is.na(rd$HXHEART) & is.na(rd$HXHEARTTR),NA,
                            ifelse(is.na(rd$HXHEART),1,0)))
factor_vars<-c(factor_vars,"HXHEARTDX")


#Stroke
writeData(wb = output, sheet="2X2Table", startRow=countrow, x="HXSTR X HXSTRTR"); countrow<-countrow+1
mtb8<-as.matrix(table(rd$HXSTR,rd$HXSTRTR,useNA="always"))
rownames(mtb8)<-c(rd$HXSTR %>% levels,"NA"); colnames(mtb8)<-c(rd$HXSTRTR %>% levels,"NA")
writeData(wb = output, sheet="2X2Table", startRow=countrow, x=mtb8, rowNames=TRUE, colNames=TRUE)
countrow<-countrow+nrow(mtb8)+3

rd$HXSTRDX<-as.factor(ifelse(is.na(rd$HXSTR) & is.na(rd$HXSTRTR),NA,
                               ifelse(is.na(rd$HXSTR),1,0)))
factor_vars<-c(factor_vars,"HXSTRDX")

#APOE, Family
writeData(wb = output, sheet="2X2Table", startRow=countrow, x="LBAPOE X FMDM"); countrow<-countrow+1
mtb9<-as.matrix(table(rd$LBAPOE,rd$FMDM,useNA="always"))
rownames(mtb9)<-c(rd$LBAPOE %>% levels,"NA"); colnames(mtb9)<-c(rd$FMDM %>% levels,"NA")
writeData(wb = output, sheet="2X2Table", startRow=countrow, x=mtb9, rowNames=TRUE, colNames=TRUE)
countrow<-countrow+nrow(mtb9)+3



#Factor levels, labels----------------------------------------------


# rd[,.SD,.SDcols=tb1_factor_vars] %>% sapply(.,levels)

rd$DXMAIN <- factor(rd$DXMAIN, level=c("1","2","3","4"))
rd$DXEXTRA <- factor(rd$DXEXTRA, levels=c("NC","SMI","AMCI","AD","PCA","fvAD","lvPPA","fAD","ADPD","CAA",
                                          "bvFTD","svPPA","nfvPPA","FTD_MND","CBS","PSPS",
                                          "LD","OND","AE","ACI","TCI","OT"))
rd$ISCHELVSVD <-factor(rd$ISCHELVSVD, levels=c("1","2","3"))

label.rd<-jstable::mk.lev(rd)

label.rd[variable=="GENDER",`:=`(val_label=c("1 남","2 여"))]
label.rd[variable=="EDUCAT",`:=`(val_label=c("1 무학","2 초등학교 졸업","3 중학교 졸업","4 고등학교 졸업","5 대학교 졸업","6 대학원 졸업 이상"))]
label.rd[variable=="SMOKE3",`:=`(val_label=c("1 Never","2 Past","3 Current"))]
label.rd[variable=="HXHTDX",`:=`(val_label=c("0 없음","1 있음"))]
label.rd[variable=="HXDMDX",`:=`(val_label=c("0 없음","1 있음"))]
label.rd[variable=="HXHLDX",`:=`(val_label=c("0 없음","1 있음"))]
label.rd[variable=="HXHEARTDX",`:=`(val_label=c("0 없음","1 있음"))]
label.rd[variable=="HXSTRDX",`:=`(val_label=c("0 없음","1 있음"))]
label.rd[variable=="FMDM",`:=`(val_label=c("0 아니오","1 예"))]
label.rd[variable=="BRMRIDAT_yesno",`:=`(val_label=c("0 시행하지 않음","1 시행함"))]
label.rd[variable=="PETDAT_yesno",`:=`(val_label=c("0 시행하지 않음","1 시행함"))]
label.rd[variable=="ISCHELVSVD",`:=`(val_label=c("1 Minimal","2 Moderate","3 Severe"))]
label.rd[variable=="PTAMPPOS",`:=`(val_label=c("0 Negative","1 Positive"))]
label.rd[variable=="DXMAIN",`:=`(val_label=c("Normal","MCI","Dementia","Parkinson"))]
label.rd[variable=="DXEXTRA" & val_label=="FTD_MND",`:=`(val_label="FTD-MND")]
label.rd[variable=="DXEXTRA" & val_label=="CBS",`:=`(val_label="Corticobasal syndrome")]
label.rd[variable=="DXEXTRA" & val_label=="PSPS",`:=`(val_label="Progressive supranuclear palsy syndrome")]

#Table 1---------------------------------------------------

tb1_vars<-c("AGE","GENDER","EDUYR","EDUCAT","SMOKE3","HXHTDX","HXDMDX","HXHLDX","HXHEARTDX","HXSTRDX",
            "BMI","LBAPOE","FMDM","BRMRIDAT_yesno","ISCHELVSVD","MTAGRT","MTAGLT","PETDAT_yesno","PTAMPPOS","KMRSCORE",
            "CDRGLOBAL","CDRSB","KIADLSCORE","BARTHELSCORE","EQSCORE","CGANPIFREQSUM","CGANPISVRSUM","CGANPIFREQSVRSUM","CGANPIBRDNSUM")
tb1_factor_vars<-tb1_vars[tb1_vars %in% factor_vars]
rdtb1<-rd[,.SD,.SDcols=c(tb1_vars,"DXMAIN","DXEXTRA")]
label.rd.tb1<-label.rd[variable %in% c(tb1_vars,"DXMAIN","DXEXTRA")]


tb1<-CreateTableOne(data = rdtb1,
                    vars = tb1_vars,
                    # strata = "DXMAIN",
                    factorVars = tb1_factor_vars,
                    # Labels = T,
                    # labeldata=label.rd.tb1,
                    includeNA = F)
tb1<-print(tb1,noSpaces=T,
           showAllLevels = T,
           varLabels = T)


tb1_1<-print(CreateTableOne2(data = rdtb1,
                             vars = tb1_vars,
                             strata = "DXMAIN",
                             factorVars = tb1_factor_vars,
                             Labels = T,
                             labeldata=label.rd.tb1,
                             includeNA = F),
             noSpaces=T,
             showAllLevels = T)

tb1_2<-print(CreateTableOne2(data = rdtb1,
                            vars = tb1_vars,
                            strata = "DXEXTRA",
                            factorVars = tb1_factor_vars,
                            Labels = T,
                            labeldata=label.rd.tb1,
                            includeNA = F),
             noSpaces=T,
             showAllLevels = T)

tb1_dt<-data.frame(VARAIBLE=rownames(tb1),Level=tb1_1[,1],Overall=tb1[,2])

tb1_dt<-cbind(tb1_dt,
          as.data.frame(tb1_1[,2:(ncol(tb1_1)-2)]),
          as.data.frame(tb1_2[,2:(ncol(tb1_2)-2)]))


addWorksheet(wb = output, sheetName = "Table1")
firstlines<-rbind(t(matrix(c(rep("",3),"DXMAIN",rep("",4),"DXEXTRA",rep("",9),"",rep("",5),"",rep("",6)))),
                  t(matrix(c(rep("",3),"",rep("",4),"EDDX",rep("",9),"FTDDX",rep("",5),"OTDX",rep("",6)))))
writeData(wb = output, colNames = F, startRow=1, sheet = "Table1", x = firstlines)
writeData(wb = output, colNames = T, startRow=3, sheet = "Table1", x = tb1_dt)
# write.csv(tb1_dt, file = "table1.csv")


#Univariate Summary---------------------------------------------------
numeric_univariate_summary<-lapply(numeric_vars,
                                   function(x){
                                     data.frame(
                                       VARIABLE=x,
                                       MEAN=mean(rd[[x]],na.rm=T),
                                       SD=sd(rd[[x]],na.rm=T),
                                       MIN=min(rd[[x]],na.rm=T),
                                       MEDIAN=median(rd[[x]],na.rm=T),
                                       MAX=max(rd[[x]],na.rm=T),
                                       NAs=sum(is.na(rd[[x]])),
                                       TOT=length(rd[[x]])
                                     )
                                   }) %>% do.call(rbind,.)
addWorksheet(wb = output, sheetName = "NumericSummary")
writeData(wb = output, sheet = "NumericSummary", x = numeric_univariate_summary)

date_univariate_summary<-lapply(date_vars,
                                function(x){
                                  data.frame(
                                    VARIABLE=x,
                                    MEAN=mean(rd[[x]],na.rm=T),
                                    SD=sd(rd[[x]],na.rm=T),
                                    MIN=min(rd[[x]],na.rm=T),
                                    MEDIAN=median(rd[[x]],na.rm=T),
                                    MAX=max(rd[[x]],na.rm=T),
                                    NAs=sum(is.na(rd[[x]])),
                                    TOT=length(rd[[x]])
                                  )
                                }) %>% do.call(rbind,.)
addWorksheet(wb = output, sheetName = "DateSummary")
writeData(wb = output, sheet = "DateSummary", x = date_univariate_summary)

factor_univariate_summary<-lapply(factor_vars,
                                  function(x){
                                    tbl<-table(rd[[x]],useNA = "always")
                                    frtb<-as.data.frame(tbl)
                                    pertb<-as.data.frame(prop.table(tbl)*100)
                                    cbind(data.frame(VARIABLE=rep(x,nrow(frtb))),frtb,pertb)
                                  }) %>% do.call(rbind,.)
factor_univariate_summary<-factor_univariate_summary[,c(1,2,3,5)]
colnames(factor_univariate_summary)<-c("VARIABLE","LEVEL","FREQUENCY","PERCENTAGE")
factor_univariate_summary$LEVEL<-ifelse(is.na(factor_univariate_summary$LEVEL),"NA",as.character(factor_univariate_summary$LEVEL))
for(i in nrow(factor_univariate_summary):2){
  if(factor_univariate_summary[i,1]==factor_univariate_summary[i-1,1]){
    factor_univariate_summary[i,1]<-""
  }
}
addWorksheet(wb = output, sheetName = "FactorSummary")
writeData(wb = output, sheet = "FactorSummary", x = factor_univariate_summary)

#Table 2---------------------------------------------------


rd[,DXMAIN_dementia:=ifelse(DXMAIN==3,1,0),]

covariates<-c("AGE","GENDER","EDUYR","SMOKE3","HXHTDX","HXDMDX","HXHLDX","HXHEARTDX","HXSTRDX","BMI","FMDM","APOEe4")
tb2_vars<-c("DXMAIN_dementia",covariates)
tb2_factor_vars<-tb2_vars[tb2_vars %in% factor_vars]
rdtb2<-rd[,.SD,.SDcols=c(tb2_vars)]
label.rd.tb2<-label.rd[variable %in% tb2_vars]

model.logistic<-glm(DXMAIN_dementia ~ AGE + GENDER + EDUYR + SMOKE3 + HXHTDX + HXDMDX + HXHLDX + HXHEARTDX + HXSTRDX + BMI + FMDM + APOEe4, data=rdtb2)
res.logistic<-glmshow.display(model.logistic)
out.logistic<-LabelepiDisplay(res.logistic)
# tb2<-datatable(out.logistic,caption = res.logistic$first.line)

tb2.out<-as.data.frame(out.logistic)
tb2.out<-tb2.out[,1:(ncol(tb2.out)-1)]
rownames(tb2.out)<-NULL
tb2.out<-cbind(rownames(out.logistic),tb2.out)
colnames(tb2.out)[1]<-" "

addWorksheet(wb=output, sheetName="Table2")
writeData(wb=output, startRow=1, sheet="Table2", x=res.logistic$first.line)
writeData(wb=output, colNames=T, startRow=2, sheet="Table2", x=tb2.out)


#-----------------------------
saveWorkbook(output, "output.xlsx", overwrite = T)
