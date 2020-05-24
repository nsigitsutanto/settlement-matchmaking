#call libraries
rm(list=ls())
library(bit64)
library(tidyverse)
library(readxl)
library(openxlsx)
library(data.table)
library(lubridate)
library(gtools)
library(zoo)
library(RDCOMClient)

#tokopedia loading
##convert to xlsx
setwd("//nfi-data-01/QEA-QEB$/06. Report/Settlement/Tokopedia")
filestp <- list.files(path = "//nfi-data-01/QEA-QEB$/06. Report/Settlement/Tokopedia", 
                      pattern = "*.xls",
                      full.names = TRUE)
nfilestp<-length(filestp)
tpresult<-list()
for (z in 1:nfilestp){
  xls <- COMCreate("Excel.Application")	
  wb = xls[["Workbooks"]]$Open( normalizePath(filestp[z]))
  xls[["DisplayAlerts"]]=FALSE
  wb$SaveAs(suppressWarnings( normalizePath(paste0('//nfi-data-01/QEA-QEB$/06. Report/Settlement/Tokopedia/XLsX/tp',z,'.xlsx'))))
  wb$Close()
  xls$Quit()
}
##tokped loading
setwd("//nfi-data-01/QEA-QEB$/06. Report/Settlement/Tokopedia/XLSX")
filestpx <- list.files(path = "//nfi-data-01/QEA-QEB$/06. Report/Settlement/Tokopedia/XLSX", 
                       pattern = "*.xlsx",
                       full.names = TRUE)
nfilestpx<-length(filestpx)
tpxresult<-list()
for (z in 1:nfilestpx){  
  phase1=read_excel(filestpx[z])
  judul=colnames(phase1)
  judul=tolower(judul)
  judul=gsub(' ','',judul)
  judul=gsub("[(]rp[)]",'',judul)
  colnames(phase1)=judul
  tpxresult[[z]]<-phase1
}
rawtp<-do.call(plyr::rbind.fill,tpxresult)
#tokopedia processing
toped=rawtp
invcode <- list(".*dana atas\\s*\\s*|\\s*barang.*",".*Ongkir-",".*dari invoice)\\s*|\\s*Harga.*",".*- ", ".*invoice ", ".*Invoice ")
invcode <- paste(unlist(invcode), collapse = "|")
toped$idorder<-gsub(invcode,"",toped$description)
#transaksi di luar orderid tokopedia
outsidetoped=toped %>% 
  filter(!grepl("INV", idorder)) %>% 
  filter(!grepl("Withdrawal",description)) %>% 
  select(1,2,3)
topads=outsidetoped %>% filter(grepl("tokopedia ads", description, ignore.case = TRUE))
write.xlsx(topads,file = "//nfi-data-01/QEA-QEB$/06. Report/Settlement/Topads.xlsx")
write.xlsx(outsidetoped,file = "//nfi-data-01/QEA-QEB$/06. Report/Settlement/Topads dll.xlsx")

#transaksi orderid tokopedia
toped=toped %>% 
  filter(grepl("INV", idorder)) %>% 
  select(c(1,2,3,5))
minustoped<-toped %>% 
  filter(grepl("kelebihan ongkos|resolusi|asuransi|service fee|ongkir|voucher|lebih murah", description, ignore.case = TRUE)) %>% 
  mutate(nominal=nominal*(-1))
#cost separation tokopedia
commfeetoped<-minustoped %>% 
  filter(grepl("service fee", description, ignore.case = TRUE))
btlcosttoped<-minustoped %>% 
  filter(grepl("voucher", description, ignore.case = TRUE))
ongkirtoped<-minustoped %>% 
  filter(grepl("ongkir|lebih murah", description, ignore.case = TRUE))
other1toped<-minustoped %>% 
  filter(grepl("kelebihan ongkos|resolusi|asuransi", description, ignore.case = TRUE))
other1toped$status<-ifelse(grepl('resolusi',other1toped$description,ignore.case=T),"Refund", NA_character_)
stattoped<-other1toped[, -c(2,3)]
stattoped<-stattoped %>% 
  filter(!is.na(status))
other2toped<-toped %>% 
  filter(grepl("selisih ongkos kirim|subsidi dari tokopedia", description, ignore.case = TRUE))
otherstoped<-plyr::rbind.fill(other1toped,other2toped)
trsctoped<-toped %>% 
  filter(grepl("transaksi penjualan", description, ignore.case = TRUE))
#sum duplicate 
ongkir<-ongkirtoped %>% 
  group_by(date,idorder) %>% 
  summarise(ongkir=sum(nominal)) %>% 
  ungroup %>% 
  select(c(2,3))
commfee<-commfeetoped %>% 
  group_by(date,idorder) %>% 
  summarise(commfee=sum(nominal)) %>% 
  ungroup %>% 
  select(c(2,3))
btlcost<-btlcosttoped%>% 
  group_by(date,idorder) %>% 
  summarise(btlcost=sum(nominal)) %>% 
  ungroup %>% 
  select(c(2,3))
others<-otherstoped %>% 
  group_by(date,idorder) %>% 
  summarise(others=sum(nominal)) %>% 
  ungroup %>% 
  select(c(2,3))
#combining tokopedia
trsctoped<- dplyr::left_join(trsctoped, ongkir, by=c("idorder" = "idorder"))
trsctoped<- dplyr::left_join(trsctoped, commfee, by=c("idorder" = "idorder"))
trsctoped<- dplyr::left_join(trsctoped, btlcost, by=c("idorder" = "idorder"))
trsctoped$fullfee<- NA_character_
trsctoped<- dplyr::left_join(trsctoped, others, by=c("idorder" = "idorder"))
trsctoped<- dplyr::left_join(trsctoped, stattoped, by=c("idorder" = "idorder"))
trsctoped[is.na(trsctoped)] <- 0
trsctoped=trsctoped %>% 
  select(3:9,11) %>% 
  mutate(fullfee=as.numeric(fullfee))

#bukalapak loading
setwd("//nfi-data-01/QEA-QEB$/06. Report/Settlement/Bukalapak")
filesbl <- list.files(path = "//nfi-data-01/QEA-QEB$/06. Report/Settlement/Bukalapak", 
                      pattern = "*.csv",
                      full.names = TRUE)
nfilesbl<-length(filesbl)
blresult<-list()
for (z in 1:nfilesbl){
  phase1=read_csv(filesbl[z])
  judul=colnames(phase1)
  judul=tolower(judul)
  judul=gsub(' ','',judul)
  colnames(phase1)=judul
  blresult[[z]]<-phase1
}
rawbl<-do.call(plyr::rbind.fill,blresult)
rawbl$waktu<-mdy_hm(rawbl$waktu)
#bukalapak processing
bl=rawbl
#transaksi orderid bukalapak
bl$idorder<-gsub(".*#","",bl$keterangan)
bl=bl %>% 
  filter(!is.na(idorder)) %>% select(1,2,4,5)
colnames(bl)<-c("date","nominal","note","idorder")
trscbl<-bl %>% 
  filter(grepl("remit untuk transaksi", note, ignore.case = TRUE)) %>% filter(!grepl("pembatalan", note, ignore.case = TRUE))
commfeebl<-bl %>% 
  filter(grepl("fee brand", note, ignore.case = TRUE))
othersbl<-bl %>% 
  filter(grepl("pembatalan remit|selisih ongkir", note, ignore.case = TRUE))
othersbl$status<-ifelse(grepl('pembatalan remit',othersbl$note,ignore.case=T),"Refund", NA_character_)
refbl<-othersbl[, -c(3)]
refbl<-refbl %>% 
  filter(!is.na(status))
refbl<-refbl %>% 
  rename(others = nominal) %>% select(2:4)
trscbl<-trscbl %>% 
  group_by(idorder) %>% 
  summarise(nominal=sum(nominal))
commfee<-commfeebl %>% 
  group_by(idorder) %>% 
  summarise(commfee=sum(nominal))
others<-othersbl %>% 
  filter(!grepl("pembatalan", note, ignore.case = TRUE)) %>% 
  group_by(idorder) %>% summarise(others=sum(nominal))
trscbl<- dplyr::left_join(trscbl, commfee, by=c("idorder" = "idorder"))
trscbl<- dplyr::left_join(trscbl, others, by=c("idorder" = "idorder"))
trscbl<-plyr::rbind.fill(trscbl,refbl)
trscbl[is.na(trscbl)] <- 0

#Elevenia loading
setwd("//nfi-data-01/QEA-QEB$/06. Report/Settlement/Elevenia")
filesele <- list.files(path = "//nfi-data-01/QEA-QEB$/06. Report/Settlement/Elevenia", 
                       pattern = "*.xls",
                       full.names = TRUE)
nfilesele<-length(filesele)
eleresult<-list()
for (z in 1:nfilesele){
  phase1=read_excel(filesele[z],skip = 4)
  judul=colnames(phase1)
  judul=tolower(judul)
  judul=gsub(' ','',judul)
  colnames(phase1)=judul
  eleresult[[z]]<-phase1
}
rawele<-do.call(plyr::rbind.fill,eleresult)
#elevenia processing
ele=rawele %>% 
  select(1,7,9:13)
ele[3:7] <- lapply(ele[3:7], as.numeric)
ele$biayatransaksi<-NULL
options(scipen=999)
ele$tanggalkonfirmasipembelian<-gsub("/","-",ele$tanggalkonfirmasipembelian)	
ele$tanggalkonfirmasipembelian<-as.POSIXct(ele$tanggalkonfirmasipembelian,format="%d-%m-%Y")
ongkirele<-ele %>% 
  filter(ele$hargabarangsetelahdiskon==0)
dateele=ele [, -c(3:6)]
dateele[duplicated(dateele[ , c("nomorpemesanan")]),"tanggalkonfirmasipembelian"]<-NA
dateele=dateele %>% 
  filter(!is.na(tanggalkonfirmasipembelian))
trscele=ele %>% 
  group_by(nomorpemesanan) %>% 
  summarise(nominal=sum(hargabarangsetelahdiskon),ongkir=sum(ongkoskirim)) %>% 
  mutate(commfee = 0.01*nominal*(-1))
trscele<- dplyr::left_join(dateele, trscele, by=c("nomorpemesanan" = "nomorpemesanan"))
colnames(trscele)<-c("idorder","date","nominal","ongkir","commfee")
trscele$date=NULL

#blibli loading
setwd("//nfi-data-01/QEA-QEB$/06. Report/Settlement/Blibli")
filesbli <- list.files(path = "//nfi-data-01/QEA-QEB$/06. Report/Settlement/Blibli", 
                    pattern = "*.xlsx",
                    full.names = TRUE)
nfilesbli<-length(filesbli)
bliresult<-list()
for (z in 1:nfilesbli){
  phase1=read_excel(filesbli[z])
  judul=colnames(phase1)
  judul=tolower(judul)
  judul=gsub(' ','',judul)
  colnames(phase1)=judul
  bliresult[[z]]<-phase1
}
rawbli<-do.call(plyr::rbind.fill,bliresult)
##blibli processing
bli=rawbli %>% 
  select(c(2,7:11,12,13,15,16)) %>%
  mutate(fee=fee*(-1)) #fee made negative
#adjustment and penalty
othersbli<-bli %>% 
  filter(grepl("adjustment|cancel|other", transaction, ignore.case = TRUE))
othersbli$status<-ifelse(grepl('cancel',othersbli$transaction,ignore.case=T),"Refund", ifelse(grepl('penalty',othersbli$transaction,ignore.case = T),"Penalty",ifelse(grepl('logistic',othersbli$transaction,ignore.case=T),"Adj. Logistic","Oth. Adjust.")))
statbli<-othersbli[, -c(2:9)]
statbli1=othersbli %>% 
  group_by(orderid) %>% 
  summarise(other=sum(total))
statbli2<- dplyr::left_join(statbli1, statbli, by=c("orderid" = "orderid")) %>% 
  select(1,2,4)
#extract initial of status
statbli2$status=toupper(sapply(regmatches(statbli2$status, gregexpr('(?<=^|\\s)[[:alpha:]]', statbli2$status, perl=TRUE)), paste0, collapse=''))
statbli2=statbli2 %>% 
  group_by(orderid) %>% 
  mutate(col=paste("status",1:n())) %>% 
  spread(col,status)
statbli2$status=do.call(paste,statbli2[,3:length(statbli2)])
statbli2$status=gsub("NA","",statbli2$status)
statbli2=statbli2 %>% select(1,2,length(statbli2))
colnames(statbli2)<-c("idorder","other","status")
#separation of transaction detail
trscbli=bli %>% 
  filter(grepl("sales", transaction, ignore.case = TRUE)) %>% 
  filter(!grepl("cancel", transaction, ignore.case = TRUE)) %>% 
  select(-c(2,4,5))
nombli=trscbli %>% select(c(1:3))
prombli=trscbli %>% select(c(1,2,4))
feebli=trscbli %>% select(c(1,2,5))
taxbli=trscbli %>% select(c(1,2,6))
nombli1<-nombli %>% group_by(orderid) %>% summarise(nominal=sum(jumlah))
prombli1=prombli %>% group_by(orderid) %>% summarise(btlcost=sum(promosimerchant))
feebli1=feebli %>% group_by(orderid) %>% summarise(commfee=sum(fee))
taxbli1=taxbli %>% group_by(orderid) %>% summarise(others=sum(pph23))
datebli=trscbli %>% select(1,2)
datebli[duplicated(datebli[ , c("orderid")]),"delivereddate"]<-NA
datebli=datebli %>% filter(!is.na(delivereddate))
#combining transaction data blibli
trscbli2<- dplyr::left_join(datebli, nombli1, by=c("orderid" = "orderid"))
trscbli2<- dplyr::left_join(trscbli2, prombli1, by=c("orderid" = "orderid"))
trscbli2<- dplyr::left_join(trscbli2, feebli1, by=c("orderid" = "orderid"))
trscbli2<- dplyr::left_join(trscbli2, taxbli1, by=c("orderid" = "orderid"))
refbli<-statbli2%>% filter(grepl("R|AL", status, ignore.case = TRUE))
colnames(trscbli2)<-c("idorder","date","nominal","btlcost","commfee","others")	
trscbli2<- dplyr::left_join(trscbli2, refbli, by=c("idorder" = "idorder"))
trscbli2[is.na(trscbli2)] <- 0
trscbli2$others=trscbli2$others+trscbli2$other
trscbli2<-trscbli2[, -c(2,7)]
colnames(statbli2)<-c("idorder","others","status")
statbli2$total=statbli2$others
statbli2<-statbli2 %>% filter(!grepl("R|AL", status, ignore.case = TRUE))
trscbli2<-plyr::rbind.fill(trscbli2,statbli2)
trscbli2$total<-NULL

#lazada loading
setwd("//nfi-data-01/QEA-QEB$/06. Report/Settlement/Lazada")
fileslz <- list.files(path = "//nfi-data-01/QEA-QEB$/06. Report/Settlement/Lazada", 
                      pattern = "*.csv",
                      full.names = TRUE)
nfileslz<-length(fileslz)
lzresult<-list()
for (z in 1:nfileslz){
  phase1=read_csv(fileslz[z])
  judul=colnames(phase1)
  judul=tolower(judul)
  judul=gsub(' ','',judul)
  colnames(phase1)=judul
  lzresult[[z]]<-phase1
}
rawlz<-do.call(plyr::rbind.fill,lzresult)
##lazada processing
laz=rawlz
laz<-laz[, -c(4:7,10:13,15:22)]
trsclaz<-laz %>% filter(grepl("item price credit", feename, ignore.case = TRUE))
commfeelaz<-laz %>% filter(grepl("payment fee", feename, ignore.case = TRUE))
fullfeelaz<-laz %>% filter(grepl("handling fee", feename, ignore.case = TRUE))
btlcostlaz<-laz %>% filter(grepl("promotional charges", feename, ignore.case = TRUE))
otherslaz<-laz %>% filter(!grepl("payment fee|handling fee|promotional charges", feename, ignore.case = TRUE))
trsclazx<-trsclaz %>% group_by(orderno.) %>% summarise(nominal=sum(amount))
commfeelazx<-commfeelaz %>% group_by(orderno.) %>% summarise(commfee=sum(amount),others1=sum(vatinamount))
fullfeelazx<-fullfeelaz %>% group_by(orderno.) %>% summarise(fullfee=sum(amount))
btlcostlazx<-btlcostlaz %>% group_by(orderno.) %>% summarise(btlcost=sum(amount))
otherslazx<-otherslaz %>% group_by(orderno.) %>% summarise(others2=sum(amount))
laz1<- dplyr::left_join(trsclazx, commfeelazx, by=c("orderno." = "orderno."))
laz2<- dplyr::left_join(laz1, fullfeelazx, by=c("orderno." = "orderno."))
laz3<- dplyr::left_join(laz2, btlcostlazx, by=c("orderno." = "orderno."))
laz4<- dplyr::left_join(laz3, otherslazx, by=c("orderno." = "orderno."))
laz4$x=laz4$nominal-laz4$others2
laz4$status<-ifelse(laz4$nominal==laz4$others2,"Success","Need Follow Up")
statlaz<-laz4[, -c(2:8)]
colnames(statlaz)<-c("idorder","status")
laz4=laz4[, -c(7,8)]
datelaz=trsclaz[, -c(2:5)]
datelaz[duplicated(datelaz[ , c("orderno.")]),"transactiondate"]<-NA
datelaz=datelaz %>% filter(!is.na(transactiondate))
trsclaza<- dplyr::left_join(datelaz, laz4, by=c("orderno." = "orderno."))
trsclaza[is.na(trsclaza)] <- 0
colnames(trsclaza)<-c("date","idorder","nominal","commfee","others","fullfee","btlcost","status")
trsclaza=trsclaza %>% group_by(date,idorder) %>% summarise(nominal=sum(nominal),fullfee=sum(fullfee),btlcost=sum(btlcost),commfee=sum(commfee),others=sum(others))
trsclaza<- dplyr::left_join(trsclaza, statlaz, by=c("idorder" = "idorder")) %>% select(-c(1))
trsclaza=trsclaza[c("idorder","nominal","commfee","btlcost","fullfee","others","status")]
#recalculate commfee, fulfee


#jdid loading
setwd("//nfi-data-01/QEA-QEB$/06. Report/Settlement/JD.ID")
filesjd <- list.files(path = "//nfi-data-01/QEA-QEB$/06. Report/Settlement/JD.ID", 
                       pattern = "*.xls",
                       full.names = TRUE)
nfilesjd<-length(filesjd)
jdresult<-list()
#jdid transaction
for (z in 1:nfilesjd){
  phase1=read_excel(filesjd[z],sheet = 1)
  judul=colnames(phase1)
  judul=tolower(judul)
  judul=gsub(' ','',judul)
  colnames(phase1)=judul
  jdresult[[z]]<-phase1
}
rawjd<-do.call(plyr::rbind.fill,jdresult)
##jdid processing
jd<-rawjd[, -c(2,3,6:10,12:18)]
jd$pay_time=ymd(jd$pay_time)
jd$discount[is.na(jd$discount)] <- 0
jd[2:4] <- lapply(jd[2:4], as.numeric)
colnames(jd)[4]="fee"
trscjd<-jd %>% group_by(order_id,pay_time) %>% summarise(nominal=sum(price),btlcost=sum(discount),fee=sum(fee))
colnames(trscjd)<-c("idorder","date","nominal","btlcost","commfee")
trscjd$commfee=trscjd$commfee*(-1)
trscjd$btlcost=trscjd$btlcost*(-1)
trscjd=trscjd[c("date","idorder","nominal","commfee","btlcost")]
#jdid refund
for (z in 1:nfilesjd){
  phase1=read_excel(filesjd[z],sheet = 2)
  judul=colnames(phase1)
  judul=tolower(judul)
  judul=gsub(' ','',judul)
  colnames(phase1)=judul
  jdresult[[z]]<-phase1
}
rawrefjd<-do.call(plyr::rbind.fill,jdresult)
refundjd<-rawrefjd %>% select(2,9)
colnames(refundjd)=c("order_id","paymenttoplatform")
refundjd$paymenttoplatform <- as.numeric(as.character(refundjd$paymenttoplatform))
refundjd$paymenttoplatform = refundjd$paymenttoplatform * (-1)
refundjd<-refundjd %>% group_by(order_id) %>% summarise(others=sum(paymenttoplatform))
trscjdid<-dplyr::left_join(trscjd, refundjd, by=c("idorder" = "order_id"))
trscjdid$others[is.na(trscjdid$others)] <- 0
trscjdid$status<-ifelse(trscjdid$nominal+trscjdid$others-trscjdid$commfee-trscjdid$btlcost==0,"Refund","Success")
trscjdid=trscjdid[c("date","idorder","nominal","commfee","btlcost","others","status")]
trscjdid$date=NULL

#jdid adjustment
#for (z in 1:nfilesjd){
  #phase1=read_excel(filesjd[z],sheet = 3)
  #judul=colnames(phase1)
  #judul=tolower(judul)
  #judul=gsub(' ','',judul)
  #colnames(phase1)=judul
  #jdresult[[z]]<-phase1
#}
#rawadjjd<-do.call(plyr::rbind.fill,jdresult)

#combining all marketplace data
trscall<-plyr::rbind.fill(trsctoped,trscjdid,trscbl,trscbli2,trscele,trsclaza)
trscall$status[trscall$status=="0"]<-"Success"
trscall$status[is.na(trscall$status)] = "Success"
trscall$btlcost=NULL

#Nutrimart loading
setwd("//nfi-data-01/QEA-QEB$/06. Report/Settlement/Nutrimart")

#forstok loading
setwd("//nfi-data-01/QEA-QEB$/13. Knowledge\\Project eCom\\Proposal Andra\\Masterdata Program Final\\Clean Data ")
rawfs=fread("data_all.csv")
fsx=rawfs
fsx[is.na(fsx)] <- 0
judul=colnames(fsx)
judul=tolower(judul)
judul=gsub(' ','',judul)
colnames(fsx)=judul
fs=fsx %>% 
  filter(year==2020) %>% 
  filter(channel!="Nutrimart") %>% 
  filter(channel!="Shopee") %>% 
  select(c(1,3,5,12:20,39,115,116,118,120:121,138,141,142)) %>% 
  mutate(regularprice=ifelse(regularprice==0,total/qty.invoiced,regularprice))
colnames(fs)[1]="channelorderid"
#bundle price
fs$bundlename=ifelse(fs$bundlename=="",NA_character_,fs$bundlename)
bun=fs %>% group_by(channelorderid,bundlename) %>% summarise(pricebundle=sum(totalnet),total1=sum(total),regularprice1=sum(regularprice),sellingprice1=sum(sellingprice)) %>% filter(!is.na(bundlename))
fs=dplyr::left_join(fs, bun, by=c("channelorderid" = "channelorderid","bundlename"="bundlename"))
#fs$totalnet=ifelse(fs$totalnet!=fs$pricebundle & !is.na(fs$bundlename),fs$pricebundle,fs$totalnet)
#fs$total=ifelse(fs$total!=fs$pricebundle & !is.na(fs$bundlename),fs$total1,fs$total)
#fs$regularprice=ifelse(fs$regularprice!=fs$regularprice1 & !is.na(fs$bundlename),fs$regularprice1,fs$regularprice)
#fs$sellingprice=ifelse(fs$sellingprice!=fs$sellingprice1 & !is.na(fs$bundlename),fs$sellingprice1,fs$sellingprice)
#calculate gross sales
grs=fs %>% 
  filter(brand!="Bundle") %>%
  filter(brand!="Gimmick") %>% 
  mutate(sellingprice=ifelse(sellingprice==0,regularprice,sellingprice)) %>% 
  group_by(channelorderid) %>% 
  summarise(grosssales=sum(qty.invoiced*sellingprice-sellerdiscount+shipping),sumtotalnet=sum(totalnet))
match=merge(fs,grs)
#panggil JNE
setwd("//nfi-data-01/QEA-QEB$/06. Report/Settlement/JNE")
filesjne <- list.files(path = "//nfi-data-01/QEA-QEB$/06. Report/Settlement/JNE", 
                      pattern = "*.xlsx",
                      full.names = TRUE)
nfilesjne<-length(filesjne)
jneresult<-list()
for (z in 1:nfilesjne){
  phase1=read_xlsx(filesjne[z],sheet = "10894900")
  judul=colnames(phase1)
  judul=tolower(judul)
  judul=gsub(' ','',judul)
  colnames(phase1)=judul
  jneresult[[z]]<-phase1
}
rawjne<-do.call(plyr::rbind.fill,jneresult)
##JNE processing
jne=rawjne %>% select(2,14)
jne=unique(jne)
match=dplyr::left_join(match, jne, by=c("awb" = "awb"))
#gimmick extraction
gim=match %>% filter(brand=="Gimmick") %>% select(1,21) %>% group_by(channelorderid) %>% summarise(gimcost=sum(totalhargacost))
match=match %>% filter(brand!="Gimmick")
match=dplyr::left_join(match, gim, by=c("channelorderid" = "channelorderid"))
#forstok processing
#flagging bundle and single
#fs$bundlename=ifelse(fs$bundlename=="","S","B")
#ratio calculation per item in one orderid
#match$shipping[match$shipping==0]<-NA
#seller,shipping & discount forstok
seldis=match %>% group_by(channelorderid) %>% summarise(seldis=sum(sellerdiscount))
ship=match %>% filter(brand!="Bundle") %>% group_by(channelorderid) %>% summarise(ship=sum(shipping))

tot=match %>% filter(brand!="Bundle") %>% group_by(channelorderid) %>% summarise(gs=sum(total))
match=dplyr::left_join(match, seldis, by=c("channelorderid" = "channelorderid"))
match=dplyr::left_join(match, ship, by=c("channelorderid" = "channelorderid"))
match=dplyr::left_join(match, tot, by=c("channelorderid" = "channelorderid"))
match$shipping[is.na(match$shipping)] <- 0
match$sellerdiscount[is.na(match$sellerdiscount)] <- 0
#ratio calculation
match$ratio=match$total/match$gs
fs=match %>% filter(brand!="Bundle")
#first match
match<-dplyr::left_join(fs, trscall, by=c("channelorderid" = "idorder"))
match$paidstatus<-ifelse(is.na(match$status),'Unpaid',ifelse(match$status=='Success','Paid',ifelse(match$status=='P','Not Paid',NA_character_)))
match[is.na(match)] <- 0
#replacing JNE
match$ongkir=ifelse(match$amount>0,match$amount,match$ongkir)
#real cost separation
match$btlcost=ifelse(match$amount>0,match$gs+match$ship-match$sumtotalnet-match$amount,match$gs-match$sumtotalnet)
match$contri=match$gs*match$ratio
match$commfee1=match$commfee*match$ratio
match$btlcost1=match$btlcost*match$ratio
match$fullfee1=match$fullfee*match$ratio

#renov on lazada and fbl
lzdfbl=match %>% filter(grepl("Lazada|FBL",channel))
lzdfbl1=lzdfbl[, -c(27,32,33)]
stnlzd=lzdfbl1 %>% group_by(channelorderid,channel) %>% summarise(sumtotalnet=sum(totalnet))
lzdfbl2=merge(lzdfbl1,stnlzd)
#recalculate ratio lzdfbl
totlzdfbl=lzdfbl2 %>% group_by(channelorderid,channel) %>% summarise(gs=sum(total))
lzdfbl2=merge(lzdfbl2,totlzdfbl)
lzdfbl2$ratio=lzdfbl2$total/lzdfbl2$gs
#crazy price
lzdfbl2$ratio=ifelse((grepl("crazy price", lzdfbl2$bundlename, ignore.case = TRUE)) & lzdfbl2$total==0,1,lzdfbl2$ratio)

#recalculate btlcost lzdfbl
lzdfbl2$btlcost=ifelse(lzdfbl2$amount>0,lzdfbl2$gs+lzdfbl2$ship-lzdfbl2$sumtotalnet-lzdfbl2$amount,lzdfbl2$gs-lzdfbl2$sumtotalnet)
lzdfbl2$contri=lzdfbl2$gs*lzdfbl2$ratio
lzdfbl2$commfee1=lzdfbl2$commfee*lzdfbl2$ratio
lzdfbl2$btlcost1=lzdfbl2$btlcost*lzdfbl2$ratio
lzdfbl2$fullfee1=lzdfbl2$fullfee*lzdfbl2$ratio

#delete former lzdfbl on match
match1=match %>% filter(!grepl("Lazada|FBL",channel))
match1<-plyr::rbind.fill(match1,lzdfbl2)

#renov on tokped bl
tpbl=match1 %>% filter(grepl("Tokopedia|Bukalapak",channel))
tpbl=unique(tpbl)
#delete former lzdfbl on match
match2=match1 %>% filter(!grepl("Tokopedia|Bukalapak",channel))
match2<-plyr::rbind.fill(match2,tpbl)

#jd condition
match2$btlcost=ifelse(match2$total==0 & match2$brand!="Bonus Produk" & match2$channel=="JD Indonesia" & match2$btlcost1==0,0,match2$btlcost)

write.xlsx(match2,file = "//nfi-data-01/QEA-QEB$/06. Report/Settlement/Test.xlsx")

#unpaid
unpaid=match2 %>% filter(paidstatus=="Unpaid")
write.xlsx(unpaid,file = "//nfi-data-01/QEA-QEB$/06. Report/Settlement/Unpaid Test.xlsx")