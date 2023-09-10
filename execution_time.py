import time
from colorama import Fore, init
init()
class ExecutionTime:
    def getSomeTime(self):
        some_time = time.time()
        return  some_time

    def setExecutionTime(self, end_time, start_time):
        execution_time = int(end_time - start_time)
        return execution_time

    def getSeconds(self, execution_time):
        seconds = int(execution_time % 60)
        return seconds

    def getMinutes(self, execution_time):
        minutes = int(execution_time // 60)
        return minutes

    def getExecutionTime(self, set_execution_time):
        minutes = self.getMinutes(set_execution_time)
        seconds = self.getSeconds(set_execution_time)
        print(Fore.CYAN + "Tiempo de ejecución:", minutes, "minutos", seconds, "segundos")
