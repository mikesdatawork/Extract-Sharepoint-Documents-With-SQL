![MIKES DATA WORK GIT REPO](https://raw.githubusercontent.com/mikesdatawork/images/master/git_mikes_data_work_banner_01.png "Mikes Data Work")        



## Contents    
- [About Process](##About-Process)  
- [SQL Logic](#SQL-Logic)  
- [Build Info](#Build-Info)  
- [Author](#Author)  
- [License](#License)       

## About-Process

  
![Extract Sharepoint Documents With SQL]( https://mikesdatawork.files.wordpress.com/2018/05/image0013.png "Extract Sharepoint Documents With SQL")
 
<p>Here's another example on how you can both search for, and extract individual Sharepoint Documents with SQL. With the image above you can see how on the left pane I'm doing a basic query to find the documents I'm looking for.
All I needed was the Database Name, List Name, File Name, and the URL so I could positively locate the file. I wanted to get extra information about the file so I added the [alldocs].[timecreated] and [docstreams].[size] just to get an idea of the time when the documents were created, and how large the files are.</p>  

I then simply copied and pasted the 4 values I needed to extract the file (using the script on the right).
One could always go through Sharepoint and get the files that way; however these are peppered across a variety of different sites. I created the two (Find & Extract) scripts to sure up the process, but ultimately I combined them together into one massive automation so that many thousands of files could be identified, and extracted by just one click.
One of the more common questions I get around this process is what happens if you run it more than once, and forget to remove the file?
Anyway; hope you find this helpful. Both scripts are below.
SEARCH FOR FILES



## SQL-Logic
```SQL
use [WSS_Content_Database];
set nocount on
 
select
    'database'  = db_name()
,   'time_created'  = left(alldocs.timecreated, 19)
,   'kb'        = (convert(bigint,alldocstreams.size))/1024
,   'mb'        = (convert(bigint,alldocstreams.size))/1024/1024
,   'list_name' = alllists.tp_title
,   'file_name' = alldocs.leafname
,   'url'       = alldocs.dirname
,   'last_url_folder' = right(alldocs.dirname, charindex('/', reverse('/' + alldocs.dirname)) - 1)
from
    alldocs join alldocstreams  on alldocs.id=alldocstreams.id 
    join alllists           on alllists.tp_id = alldocs.listid
where
    --alldocstreams.[size] > 2048
    right([alldocs].[leafname], 2) in ('oc', 'cx', 'df', 'sg', 'xt')
    and alllists.tp_title like '%FY12 Documents%'
order by
    alldocs.timecreated desc
,   alldocs.dirname 
```


EXTRACT THE FILE



## SQL-Logic
```SQL
use master;
set nocount on
 
declare @ole_automation int
set     @ole_automation = (select cast([value_in_use] as int) from sys.configurations where [configuration_id] = '16388')
if      @ole_automation = 0
    begin
    exec sp_configure 'Ole Automation Procedures', 1; reconfigure with override;
    end;
go
 
use tempdb;
set nocount on
 
declare @url            varchar(1000)
declare @list           varchar(255)
declare @file           varchar(255)
declare @database       varchar(255)
declare @extension      varchar(5)
declare @destination_path   varchar(255)
/********************************************************************/
set @database   = 'WSS_Content_Database'
set @list   = 'Archive FY12 Documents'
set @file   = '7684_HiringPacket.pdf'
set @url    = 'sites/Archive of Hiring Docs FY2012'
/********************************************************************/
set @extension = (select reverse(left(reverse(@file),charindex('.',reverse(@file))-1)))
set @destination_path   = '\\sps1\w$\' + @file
 
declare @extract_file   varchar(max)
set @extract_file   = 
'use [' + @database + '];
set nocount on;
 
declare @object_token int
declare @content_binary varbinary(max)
select  @content_binary = alldocstreams.content from alldocs join alldocstreams on alldocs.id = alldocstreams.id join alllists on alllists.tp_id = alldocs.listid
where  
    alllists.tp_title   = ''' + @list + '''
    and alldocs.leafname    = ''' + @file + '''
    and alldocs.dirname = ''' + @url  + '''
 
exec sp_oacreate ''adodb.stream'', @object_token output
exec sp_oasetproperty @object_token, ''type'', 1
exec sp_oamethod @object_token, ''open''
exec sp_oamethod @object_token, ''write'', null, @content_binary
exec sp_oamethod @object_token, ''savetofile'', null, ''' + @destination_path + ''', 2
exec sp_oamethod @object_token, ''close''
exec sp_oadestroy @object_token
'
exec    (@extract_file)
```


[![WorksEveryTime](https://forthebadge.com/images/badges/60-percent-of-the-time-works-every-time.svg)](https://shitday.de/)

## Build-Info

| Build Quality | Build History |
|--|--|
|<table><tr><td>[![Build-Status](https://ci.appveyor.com/api/projects/status/pjxh5g91jpbh7t84?svg?style=flat-square)](#)</td></tr><tr><td>[![Coverage](https://coveralls.io/repos/github/tygerbytes/ResourceFitness/badge.svg?style=flat-square)](#)</td></tr><tr><td>[![Nuget](https://img.shields.io/nuget/v/TW.Resfit.Core.svg?style=flat-square)](#)</td></tr></table>|<table><tr><td>[![Build history](https://buildstats.info/appveyor/chart/tygerbytes/resourcefitness)](#)</td></tr></table>|

## Author

[![Gist](https://img.shields.io/badge/Gist-MikesDataWork-<COLOR>.svg)](https://gist.github.com/mikesdatawork)
[![Twitter](https://img.shields.io/badge/Twitter-MikesDataWork-<COLOR>.svg)](https://twitter.com/mikesdatawork)
[![Wordpress](https://img.shields.io/badge/Wordpress-MikesDataWork-<COLOR>.svg)](https://mikesdatawork.wordpress.com/)

     
## License
[![LicenseCCSA](https://img.shields.io/badge/License-CreativeCommonsSA-<COLOR>.svg)](https://creativecommons.org/share-your-work/licensing-types-examples/)

![Mikes Data Work](https://raw.githubusercontent.com/mikesdatawork/images/master/git_mikes_data_work_banner_02.png "Mikes Data Work")

