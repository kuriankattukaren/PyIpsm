import os as os
from time import sleep
from typing import Literal, List
import win32com.client as com_conn

Application = Literal["PROSPER", "MBAL", "GAP"]
ApplicationsList: List[Application] = ["PROSPER", "MBAL", "GAP"]


class OpenServer:
    def __init__(self, openserver_name: str):
        self.name = openserver_name
        self.connect(self.name)

    def connect(self, openserver_name: str):
        print(f"Establishing connection to open_server: {openserver_name}")
        self.connection = com_conn.Dispatch(openserver_name)
        print(f"Openserver connection established...")

    def __str__(self):
        return f"Openserver: {self.name}"

class Error:
    def __init__(self, number: int, description: str):
        self.number = number
        self.description = description

    def __str__(self):
        return f"Error number: {self.number}: {self.description}"

class Commands:
    def __init__(self, openserver: OpenServer):
        self.openserver = openserver

    def _error(self, error_number: int):
        text = self.description(error_number)
        error = Error(error_number, text)
        return error

    def _application_name(self, variable_name: str):
        components = variable_name.split(".")
        program_name = components[0]

        if program_name in ApplicationsList:
            return program_name
        else:
            return ""

    def description(self, error_number: int):
        """ Returns the description for error number."""
        conn = self.openserver.connection
        description: str = conn.GetErrorDescription(error_number)
        return description

    def do(self, command: str):
        """ Executes commands that return quickly..."""
        lerr: int = self.openserver.connection.DoCommand(command)
        error = self._error(lerr)
        return error

    def doSlow(self, command: str):
        """ Executes commands that take long time..."""
        lerr: int = self.openserver.connection.DoCommandAsync(command)
        error = self._error(lerr)

        if error.number > 0:
            return error

        appname = self._application_name(command)
        while self.openserver.connection.IsBusy(appname):
            sleep(2)

        lerr = self.openserver.connection.GetLastError(appname)
        error = self._error(lerr)
        return error

    def set(self, variable: str, value: str):
        lerr = self.openserver.connection.SetValue(variable, value)
        error = self._error(lerr)
        return error

    def get(self, variable_name: str):
        value: str = self.openserver.connection.GetValue(variable_name)

        application = self._application_name(variable_name)
        lerr = self.openserver.connection.GetLastError(application)
        error = self._error(lerr)
        return error, value


class Program:
    def __init__(self, program_name: str, openserver: OpenServer):
        print(f"Starting up: {program_name}")
        os.startfile(program_name)

        self.openserver = openserver
        self.commands = Commands(self.openserver)
