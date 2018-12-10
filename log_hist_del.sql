select 
  'arc'||lpad(to_char(SEQUENCE#),5,'0')||'.log' 
from v$log_history 
where 
   trunc(FIRST_TIME) < trunc(sysdate-5)    -- то что удалить
and 
  SEQUENCE# <= (select max(SEQUENCE#) from v$log where ARCHIVED='YES' )
 