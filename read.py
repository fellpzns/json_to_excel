import  jpype     
import  asposecells
       
jpype.startJVM() 
from asposecells.api import Workbook
workbook = Workbook("json_test.json")
workbook.save("json_test.xlsm")
jpype.shutdownJVM()