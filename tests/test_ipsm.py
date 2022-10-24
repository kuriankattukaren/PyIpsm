import pytest
from time import sleep
import win32com.client as wincom
from ..src.PyIpsm.ipsm import OpenServer, Program

openserver_file = "PX32.OpenServer.1"

prosper_exe = "C:\\Program Files\\Petroleum Experts\\IPM 12\\prosper.exe"
mbal_exe = "C:\\Program Files\\Petroleum Experts\\IPM 12\\mbal.exe"
gap_exe = "C:\\Program Files\\Petroleum Experts\\IPM 12\\gap.exe"

comError = wincom.pywintypes.com_error #type: ignore


def test_openserver_valid():
    open_server = OpenServer(openserver_file)
    assert open_server != None


def test_openserver_invalid():
    with pytest.raises(comError):
        OpenServer("random_name")


@pytest.mark.parametrize(
    "shutdown_command, program_exe",
    [
        (
            "PROSPER.SHUTDOWN", 
            prosper_exe
        ),
        (
            "MBAL.SHUTDOWN", 
            mbal_exe
        ),
        (
            "GAP.SHUTDOWN", 
            gap_exe
        ),
    ],
)
def test_program_start_exit(shutdown_command, program_exe):
    open_server = OpenServer(openserver_file)
    prosper = Program(program_exe, open_server)

    error = prosper.commands.do(shutdown_command)
    assert error.number == 0


@pytest.mark.parametrize(
    "command, program_exe", 
    [
        (
            "PROSPER.SHUTDOWN", 
            prosper_exe
        )
    ],
)
def test_command_do(command, program_exe):
    open_server = OpenServer(openserver_file)
    prosper = Program(prosper_exe, open_server)

    error = prosper.commands.do(command)
    assert error.number == 0


@pytest.mark.parametrize(
    "command, program_exe",
    [
        (
            "PROSPER.OPENFILE=C:\\Apps\\src\\Src\\IPSM\\Openserver\\Openserver\\Code\\Test\\Test_v2.Out",
            prosper_exe,
        )
    ],
)
def test_command_doSlow(command, program_exe):
    open_server = OpenServer(openserver_file)
    prosper = Program(program_exe, open_server)

    prosper_error = prosper.commands.doSlow(command)
    assert prosper_error.number == 0

    prosper.commands.do("PROSPER.SHUTDOWN")


@pytest.mark.parametrize(
    "program_exe, openserver_variable, value",
    [
        (
            prosper_exe, 
            "PROSPER.SIN.SUM.Fluid", 
            1
        )
    ],
)
def test_command_set(program_exe, openserver_variable, value):
    open_server = OpenServer(openserver_file)
    prosper = Program(program_exe, open_server)

    open_file = "PROSPER.OPENFILE=C:\\Apps\\src\\Src\\IPSM\\Openserver\\Openserver\\Code\\Test\\Test_v2.Out"
    prosper.commands.doSlow(open_file)

    error = prosper.commands.set(openserver_variable, value)
    assert error.number == 0

    prosper.commands.do("PROSPER.SHUTDOWN")


@pytest.mark.parametrize(
    "program_exe, openserver_variable, expected_value",
    [
        (
            prosper_exe, 
            "PROSPER.SIN.SUM.Fluid", 
            "1"
        )
    ],
)
def test_command_get(program_exe, openserver_variable, expected_value):
    open_server = OpenServer(openserver_file)
    prosper = Program(program_exe, open_server)

    open_file = "PROSPER.OPENFILE=C:\\Apps\\src\\Src\\IPSM\\Openserver\\Openserver\\Code\\Test\\Test_v2.Out"
    prosper.commands.doSlow(open_file)

    error = prosper.commands.set(openserver_variable, "1")
    assert error.number == 0

    error, value = prosper.commands.get(openserver_variable)
    assert value == expected_value

    prosper.commands.do("PROSPER.SHUTDOWN")


@pytest.mark.parametrize(
    "program_exe, error_number, error_description",
    [
        (
            prosper_exe, 
            0, 
            ""
        ), 
        (
            prosper_exe, 
            3004, 
            "Variable name was not found"
        )
    ],
)
def test_error_description(program_exe, error_number, error_description):
    open_server = OpenServer(openserver_file)
    prosper = Program(program_exe, open_server)

    expected_description = prosper.commands.description(error_number)
    assert error_description == expected_description

    prosper.commands.do("PROSPER.SHUTDOWN")


@pytest.mark.parametrize(
    "program_exe, openserver_variable, application_name",
    [
        (
            prosper_exe, 
            "PROSPER.SHUTDOWN", 
            "PROSPER"
        )
    ],
)
def test_application_name(program_exe, openserver_variable, application_name):
    open_server = OpenServer(openserver_file)
    prosper = Program(program_exe, open_server)

    application = prosper.commands._application_name(openserver_variable)
    assert application == application_name

    prosper.commands.do("PROSPER.SHUTDOWN")
