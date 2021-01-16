packages<-c('optparse','xlsx','readxl','data.table','stringr')
message(paste('Loading packages:',paste(packages,collapse=', ')))
lapply(packages, require, character.only = TRUE)

dat <- data.frame(read_excel('Rota.xlsx',sheet=paste('Rota',1)) )
times <- data.frame(read_excel('Rota.xlsx',sheet='Breakdown') )[,1:8]

ward<- TRUE 


uniq.weeks<-unique(dat[,2])
uniq.weeks<- trim(gsub(uniq.weeks,pattern=" {2,}",replacement=' ')) 

#as.Date( as.Date(Sys.time()),"%d-%B-%Y") - as.Date( as.Date(Sys.time()),"%d-%B-%Y") 

nb.weeks<-nrow(dat)

tra<-data.frame(matrix(nrow=nb.weeks*7,ncol=10))
colnames(tra)<-c('Subject','Start Date','Start Time','End Date','End Time','All Day Event','Description','Location','Private','Day')
#tra[,2]<- gsub(dat[i,1],pattern='\\.',replacement='-') 

date<- unlist(strsplit(dat[1,1],'\\.') )
#start_date<-  as.Date(paste(date[2], date[1],paste(paste0('20',date[3])),sep='/')) 
start_date<- as.Date(paste(date[2], date[1],paste(paste0('20',date[3])),sep='/'),"%m/%d/%y")

date.formatted<- gsub(dat[,1],pattern='\\.',replacement='/') 
date.dat<-data.frame(matrix(nrow=nb.weeks*7,ncol=2))
dts<-c()
for(p in 1:length(date.formatted))if(p==1)dts<- rep(date.formatted[p],7) else dts<- c(dts, rep(date.formatted[p],7))
date.dat[,1]<- dts
for(i in 1:nrow(tra))
{
	date.dat[i,2]<- as.character(start_date+i-1) 
	tra[i,2]<-as.character(start_date+i-1) 
}

tweakDate<- function(date)
{
	y<-  unlist(strsplit(tra[i,2],'/') ) 
	y<-  paste(y[3],y[1],y[2],sep='-') 
	y
}


create_ical <- function(start, end, summary, domain="example.com") {

  require(uuid, quietly = TRUE, warn.conflicts = FALSE)

  sprintf(
    "BEGIN:VCALENDAR
VERSION:2.0
PRODID:-//rstats//NONSGML v1.0//EN
BEGIN:VEVENT
UID:%s@%s
DTSTAMP:%s
DTSTART:%s
DTEND:%s
SUMMARY:%s
END:VEVENT
END:VCALENDAR
", uuid::UUIDgenerate(),
    domain, 
    format(Sys.time(), "%Y%m%dT%H%M%SZ", tz="GMT"), 
    format(start, "%Y%m%dT%H%M%SZ", tz="GMT"), 
    format(end, "%Y%m%dT%H%M%SZ", tz="GMT"), 
    summary
)

}



days<- c('Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday') 
tra$Day <- rep(days,nb.weeks)
tra[,6:9]<-'FALSE'

ppl<-dat[,2:ncol(dat)]


for(rota in 1:5)
#for(rota in 5)
{
	dat <- data.frame(read_excel('Rota.xlsx',sheet=paste('Rota',rota)) )

	for(ppl in 2:ncol(dat)) 
	{
		outICS<- paste0(gsub(colnames(dat)[ppl],pattern='\\.',replacement='_') ,'.ics') 
		if(file.exists(outICS))file.remove(outICS)
		
		write.table('BEGIN:VCALENDAR\n',outICS,col.names=F,row.names=F,quote=F,sep=',',append=F)

		for(i in 1:nrow(tra))
		{
			date<- unlist(strsplit(dat[i,1],'\\.') )
			#tra[i,2]<-paste(date[2], date[1],paste(paste0('20',date[3]),sep='/'),sep='/')
			# tra[i,2]<-as.character(start_date+i-1)

	 		date2<- unlist(strsplit(as.character(start_date+i-1) ,'-')) 
			date3 <-paste( date2[2], date2[3],date2[1],sep='/')
			date4<- unlist(strsplit(as.character(start_date+i) ,'-')) 
			date5 <-paste( date4[2], date4[3],date4[1],sep='/')
			tra[i,2]<- date3

			week.monday <- date.dat[date.dat[,2]==as.character(start_date+i-1) ,1] 
			dat.row<- dat[,1]==gsub(week.monday,pattern='/',replacement='\\.') 
			dat.row<-dat[dat.row,ppl] 

			#job<- trim(gsub(dat[i,ppl],pattern=" {2,}",replacement=' '))
			week<-as.numeric(unlist(strsplit(dat.row,' ') ) [2])
			week.row<- which(times[,1]==week)
			week.hours<- times[week.row+2,2:8]
			week.hours[is.na(week.hours)]<-'OFF'
			colnames(week.hours)<- days 

			tra[i,4]<-date3
			day.hours<- grep(tra$Day[i],colnames(week.hours) ) 
			day.hours<- week.hours[,day.hours]
			tra[i,1]<- day.hours[1]
			day.hours<- trim(unlist(strsplit(day.hours,'-') ) )
			tra[i,3]<- day.hours[1]
			tra[i,5]<- day.hours[2]
			if(tra[i,3]=='20:30') tra[i,4]<-date5

			if(day.hours[1]=='OFF')tra[i,6]<- TRUE

			tra[is.na(tra[,3]),3]<- '09:00'
			tra[is.na(tra[,5]),5]<- '09:00'

			tra[tra[,5]=='24:00',5]<- '00:00'
			tra[tra[,3]=='OFF',3]<- '09:00'
			tra[,6]<-FALSE

			create_ical(
			  as.POSIXct(paste(tweakDate(tra[i,2]),tra[i,3]) , origin="1970-01-01 00:00:00"),
			  as.POSIXct(paste(tweakDate(tra[i,4]),tra[i,5]), origin="1970-01-01 00:00:00"),
			  tra[i,1],
			  "somedom.org"
			) -> ics_event

			ics_event<-  gsub(ics_event,pattern='END:VCALENDAR\n',replacement='') 
			ics_event<-  gsub(ics_event,pattern='BEGIN:VCALENDAR\n',replacement='') 

			if(ward)
			{
				nam<- gsub(colnames(dat)[ppl],pattern='\\.',replacement=' ')
				initials<- gsub('\\b(\\pL)\\pL{2,}|.','\\U\\1',nam,perl = TRUE)
				subject<- paste(initials,tra[i,1])
				tra[i,1]<- subject
			} else tra[i,1]<- day.hours[1]

			write.table(ics_event,outICS,col.names=F,row.names=F,quote=F,sep=',',append=T)
		}
		write.table('\nEND:VCALENDAR',outICS,col.names=F,row.names=F,quote=F,sep=',',append=T)

		write.table(tra[,1:9], paste0(gsub(colnames(dat)[ppl],pattern='\\.',replacement='_') ,'.csv'),col.names=T,row.names=F,quote=F,sep=',')

	}
}

files<-list.files(pattern='csv')

wardFile <-  data.frame(read_excel('Wards.xlsx',sheet='Sheet1',col_names=F) )
wardFile<- wardFile[!is.na(wardFile[,1]),]
colnames <- c('Subject','Start Date','Start Time','End Date','End Time','All Day Event','Description','Location','Private') 
wards<-unique(wardFile[,1])

oDir<-'Wards'
if(!file.exists(oDir))dir.create(oDir)

for(w in 1:length(wards))
{
	oFile<-  paste0(gsub(wards[w],pattern=' ',replacement='_') ,'.csv') 
	if(file.exists(oFile))file.remove(oFile)

	for(cols in 2:ncol(wardFile))
	{
		if(!is.na(wardFile[w,cols]))
		{
			firstname<- gsub(wardFile[w,cols],pattern=' .*',replacement='')
			secondname<- gsub(wardFile[w,cols],pattern='.* ',replacement='')

			if(length(grep(firstname, files) )>0)
			{
				if(length(grep(firstname, files) )==1)dat<- read.csv(files[grep(firstname, files) ] ) 
				if(length(grep(firstname, files) )>1)dat<- read.csv(files[grep(firstname,files)][grep(secondname,files[ grep(firstname,files)]) ]) 

				write.table(dat, oFile,col.names=F,row.names=F,quote=T,sep=',',append=T)

			}
		}
	}

	tra<-read.csv(oFile,header=F,stringsAsFactors=F)
	colnames(tra)<- colnames

	t2<- tra[-grep('OFF',tra[,1]),]
	t2<- t2[!t2[,3]=='17:00',]
	t2<- t2[!t2[,3]=='20:30',]
	t2<- t2[!t2[,3]=='08:00',]
	write.table(t2,file.path(oDir,oFile),col.names=T,row.names=F,quote=T,sep=',',append=F)

	oFileICS<-  paste0(gsub(wards[w],pattern=' ',replacement='_') ,'.ics') 
	if(file.exists(oFileICS))file.remove(oFileICS)
	write.table('BEGIN:VCALENDAR\n',file.path(oDir,oFileICS),col.names=F,row.names=F,quote=F,sep=',',append=F)

	for(row in 1:nrow(t2))
	{
		create_ical(
			  as.POSIXct(paste(tweakDate(t2[row,2]),t2[row,3]) , origin="1970-01-01 00:00:00"),
			  as.POSIXct(paste(tweakDate(t2[row,4]),t2[row,5]), origin="1970-01-01 00:00:00"),
			  t2[row,1],
			  "somedom.org"
			) -> ics_event

		ics_event<-  gsub(ics_event,pattern='END:VCALENDAR\n',replacement='') 
		ics_event<-  gsub(ics_event,pattern='BEGIN:VCALENDAR\n',replacement='') 
		write.table(ics_event,file.path(oDir,oFileICS),col.names=F,row.names=F,quote=F,sep=',',append=T)
	}
	write.table('\nEND:VCALENDAR',file.path(oDir,oFileICS),col.names=F,row.names=F,quote=F,sep=',',append=T)
}

file.remove(list.files(pattern='csv'))
file.remove(list.files('Wards/',pattern='csv',full.names=T))


