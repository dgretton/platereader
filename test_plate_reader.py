import win32com.client

d = win32com.client.Dispatch('BMG_ActiveX.BMGRemoteControl')
d.OpenConnectionV('CLARIOstar')
db_dir = 'C:\\Program Files (x86)\\BMG\\CLARIOstar\\User\\Definit'
d.ExecuteAndWait(['PlateOut'])
d.ExecuteAndWait(['PlateIn'])
d.ExecuteAndWait(['Run', 'chemostat_od_abs', db_dir, 'C:\\Users\\Hamilton\\Dropbox\\plate_reader_results\\mars_database'])
d.CloseConnection()
