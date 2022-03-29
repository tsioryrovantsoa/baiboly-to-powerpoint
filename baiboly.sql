-- cmd misy an'ilay fichier

sqlite3 bible_baiboly.sqlite

.tables

--requete
select * from android_metadata;
select * from chapters;
select * from texts;
select * from version;

--mombamomba an'ilay table
pragma table_info('chapters');
--id
--title
--num
--mode
--short-title
--ntitle

--ohatra : 8220 - Genesisy - 50 - 1 - -

pragma table_info('android_metadata');
--locale

pragma table_info('texts');
--chapter_id
--chapter_num
--position
--text
--rank
--head
--bookmark
--higlight
--note
--bookmark_date
--highlight_date
--note_date
--_id
--ntext

--ohatra : 8220-1-0-Malagasy Bible (1865) Ny namoronan\'Andriamanitra izao tontolo izao-1-1-||||||1|

--longue 
select chapter_id,chapter_num,position,text from texts where chapter_id=8220 and chapter_num=1 and  position between 1 and 5; 
--court
select chapter_id,chapter_num,position,text from texts where chapter_id=8265 and chapter_num=15 and  position between 10 and 10; 

pragma table_info('version');

--code

--AllBoky
select _id,title from chapters;
--ordre alphabetique
select _id,title from chapters ORDER BY title ASC;

--getVerser
select chapter_id,chapter_num,position,text from texts where chapter_id=8265 and chapter_num=15 and  position between 10 and 10; 

