import psutil as ps
from time import sleep


PROCNAME = ['TEAMS','POSTGRE','SQL SERVER','TEAMV','SKYPE','RUNTIME BROKER','PG_CTL','ONEDRIVE']
total, killed, failed = 0,0,0
for proc in ps.process_iter():
    total+=1
    info = proc.as_dict(attrs=['pid', 'name'])
    for PROC in PROCNAME:
        if PROC in str(info['name']).upper() :
            #print('Processo: {} (PID: {})'.format(info['name'], info['pid']))
            try:
                proc.kill()
                killed +=1
            except:
                #print('NÃO MORREU - Processo: {} (PID: {})'.format(info['name'], info['pid']))
                failed +=1

print('Total: ' + str(total))
print('Killed: ' + str(killed))
print('Failed: ' + str(failed))
sleep(5)