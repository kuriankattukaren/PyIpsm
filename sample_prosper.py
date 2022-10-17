import src.ipsm as Ipsm

prosper_exe = "C:\\Program Files\\Petroleum Experts\\IPM 12\\prosper.exe"

# Establish openserver connection...
openserver_file = "PX32.OpenServer.1"
open_server = Ipsm.OpenServer(openserver_file)

# Start executables...
prosper = Ipsm.Program(prosper_exe, open_server)

# Do prosper command...
open_file_command = "PROSPER.OPENFILE=C:\\Apps\\src\\Src\\IPSM\\Openserver\\Openserver\\Code\\Test\\Test_v2.Out"
prosper_error = prosper.commands.doSlow(open_file_command)
prosper_error = prosper.commands.do("PROSPER.SHUTDOWN")
print(f"Prosper error: {prosper_error.number}")