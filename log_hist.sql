select 
  'arc'||lpad(to_char(SEQUENCE#),5,'0')||'.log' 
from v$log_history 
where 
   trunc(FIRST_TIME) between trunc(sysdate-2) and trunc(sysdate)   -- �������� �� ������� � �� ��� ��������� ���
and 
  SEQUENCE# <= (select max(SEQUENCE#) from v$log where ARCHIVED='YES' )
 