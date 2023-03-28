import numpy as np
from openpyxl import *
from openpyxl.styles import *    
import copy

class port: #attri: 0: in; 1: out. 
    def __init__(self,portName,attri):
        self.key=str(portName)
        self.flag=False #traverse factor, no real meaning
        self.flag2=False #traverse factor, no real meaning
        self.calibrated= True #always true
        self.nextDevice= None
        self.portAttribute=int(attri)

    def attach(self,nextDevice):
        if self.nextDevice != None:
            print("Error!",self.key,"has more than one attached devices!")
            exit(1)
        self.nextDevice=nextDevice

    
class MMI: #"left","right": heater is always on the right, or look back against the light direction 
    def __init__(self,MMIname):
        self.key=str(MMIname)
        self.flag=False #traverse factor, no real meaning
        self.flag2=False #traverse factor, no real meaning
        self.calibrated= True #always true
        self.leftIn= None
        self.rightIn= None
        self.leftOut=None
        self.rightOut=None
    
    def attachLeftIn(self,nextDevice):
        if self.leftIn != None:
            print("Error!",self.key,"leftIn has more than one attached devices!")
            exit(1)
        self.leftIn=nextDevice

    def attachRightIn(self,nextDevice):
        if self.rightIn != None:
            print("Error!",self.key,"RightIn has more than one attached devices!")
            exit(1)
        self.rightIn=nextDevice

    def attachLeftOut(self,nextDevice):
        if self.leftOut != None:
            print("Error!",self.key,"LeftOut has more than one attached devices!")
            exit(1)
        self.leftOut=nextDevice

    def attachRightOut(self,nextDevice):
        if self.rightOut != None:
            print("Error!",self.key,"RightOut has more than one attached devices!")
            exit(1)
        self.rightOut=nextDevice


class MZIheater:
    def __init__(self,heaterName):
        self.key=str(heaterName)
        self.flag=False #traverse factor, no real meaning
        self.flag2=False #traverse factor, no real meaning
        self.calibrated= False
        self.leftIn= None
        self.rightIn= None
        self.leftOut=None
        self.rightOut=None
    
    def attachLeftIn(self,nextDevice):
        if self.leftIn != None:
            print("Error!",self.key,"leftIn has more than one attached devices!")
            exit(1)
        self.leftIn=nextDevice

    def attachRightIn(self,nextDevice):
        if self.rightIn != None:
            print("Error!",self.key,"RightIn has more than one attached devices!")
            exit(1)
        self.rightIn=nextDevice

    def attachLeftOut(self,nextDevice):
        if self.leftOut != None:
            print("Error!",self.key,"LeftOut has more than one attached devices!")
            exit(1)
        self.leftOut=nextDevice

    def attachRightOut(self,nextDevice):
        if self.rightOut != None:
            print("Error!",self.key,"RightOut has more than one attached devices!")
            exit(1)
        self.rightOut=nextDevice


class Phaseheater:
    def __init__(self,heaterName):
        self.key=str(heaterName)
        self.flag=False #traverse factor, no real meaning
        self.flag2=False #traverse factor, no real meaning
        self.calibrated= False
        self.In= None
        self.Out=None
    
    def attachIn(self,nextDevice):
        if self.In != None:
            print("Error!",self.key,"In has more than one attached devices!")
            exit(1)
        self.In=nextDevice

    def attachOut(self,nextDevice):
        if self.Out != None:
            print("Error!",self.key,"LeftOut has more than one attached devices!")
            exit(1)
        self.Out=nextDevice


class chip:
    def __init__(self,chipName):
        self.key=str(chipName)
        self.inPort=[]
        self.outPort=[]
        self.MMI=[]
        self.MZIheater=[]
        self.Phaseheater=[]
        self.Longlist=[]

    def pushinPort(self,nextDevice):
        for u in self.inPort:
            if u.key==nextDevice.key:
                print("Error! Different devices cannot have the same name: ",u.key)
                exit(1)
        self.inPort.append(nextDevice)

    def pushoutPort(self,nextDevice):
        for u in self.outPort:
            if u.key==nextDevice.key:
                print("Error! Different devices cannot have the same name: ",u.key)
                exit(1)
        self.outPort.append(nextDevice)

    def pushMMI(self,nextDevice):
        for u in self.MMI:
            if u.key==nextDevice.key:
                print("Error! Different devices cannot have the same name: ",u.key)
                exit(1)
        self.MMI.append(nextDevice)

    def pushMZIheater(self,nextDevice):
        for u in self.MZIheater:
            if u.key==nextDevice.key:
                print("Error! Different devices cannot have the same name: ",u.key)
                exit(1)
        self.MZIheater.append(nextDevice)

    def pushPhaseheater(self,nextDevice):
        for u in self.Phaseheater:
            if u.key==nextDevice.key:
                print("Error! Different devices cannot have the same name: ",u.key)
                exit(1)
        self.Phaseheater.append(nextDevice)

def initialize(Chip):
    for u in Chip.Longlist:
        u.flag=False
        u.flag2=False


def CheckIn(list,name):
    for u in list:
        if u.key==name:
            return u
    return None


def creatChip(chipName,path=r"devices.xlsx"): #path=r"devices.xlsx"
    wb1 = load_workbook(path)
    ws1 = wb1['port']
    port_inform=ws1['A2':'C4']

    ws2 = wb1['MMI']
    MMI_inform=ws2['A2':'E2']

    ws3 = wb1['MZIheater']
    MZIheater_inform=ws3['A2':'E5']

    ws4 = wb1['Phaseheater']
    Phaseheater_inform=ws4['A2':'C2']

    Chip=chip(chipName)

    for u in port_inform:
        port_name=str(u[0].value)
        port_attri=int(u[1].value)
        port_attach=str(u[2].value)
        if port_attri==0:
            if CheckIn(Chip.inPort,port_name)==None:
                newPort=port(port_name,0)
                Chip.inPort.append(newPort)
                Chip.Longlist.append(newPort)
                
        else:
            if CheckIn(Chip.outPort,port_name)==None:
                newPort=port(port_name,1)
                Chip.outPort.append(newPort)
            
    for u in MMI_inform:
        MMI_name=str(u[0].value)
        if CheckIn(Chip.MMI,MMI_name)==None:
            newMMI=MMI(MMI_name)
            Chip.MMI.append(newMMI)

    for u in MZIheater_inform:
        MZIheater_name=str(u[0].value)
        if CheckIn(Chip.MZIheater,MZIheater_name)==None:
            newMZIheater=MZIheater(MZIheater_name)
            Chip.MZIheater.append(newMZIheater)

    for u in Phaseheater_inform:
        Phaseheater_name=str(u[0].value)
        if CheckIn(Chip.Phaseheater,Phaseheater_name)==None:
            newPhaseheater=Phaseheater(Phaseheater_name)
            Chip.Phaseheater.append(newPhaseheater)

    #create attachments
    Chip.Longlist.extend(Chip.outPort)
    Chip.Longlist.extend(Chip.MMI)
    Chip.Longlist.extend(Chip.MZIheater)
    Chip.Longlist.extend(Chip.Phaseheater)

    for u in port_inform:
        port_name=str(u[0].value)
        port_attri=int(u[1].value)
        port_attach=str(u[2].value)
        newPort=CheckIn(Chip.Longlist,port_name)
        attachedDevice=CheckIn(Chip.Longlist,port_attach)
        if newPort==None or (attachedDevice==None and port_attach!="None"):
            print("Error! Unknown device",port_name,"->",port_attach)
            exit(1)
        if port_attach!="None":
            newPort.attach(attachedDevice)

    for u in MMI_inform:
        MMI_name=str(u[0].value)
        MMI_leftIn=str(u[1].value)
        MMI_rightIn=str(u[2].value)
        MMI_leftOut=str(u[3].value)
        MMI_rightOut=str(u[4].value)

        newMMI=CheckIn(Chip.Longlist,MMI_name)
        newleftIn=CheckIn(Chip.Longlist,MMI_leftIn)
        newrightIn=CheckIn(Chip.Longlist,MMI_rightIn)
        newleftOut=CheckIn(Chip.Longlist,MMI_leftOut)
        newrightOut=CheckIn(Chip.Longlist,MMI_rightOut)
        if MMI_leftIn!="None":
            if newMMI==None or newleftIn ==None:
                print("Error! Unknown device",MMI_name,"->",MMI_leftIn)
                exit(1)
            newMMI.attachLeftIn(newleftIn)
        if MMI_rightIn!="None":
            if newMMI==None or newrightIn ==None:
                print("Error! Unknown device",MMI_name,"->",MMI_rightIn)
                exit(1)
            newMMI.attachRightIn(newrightIn)
        if MMI_leftOut!="None":
            if newMMI==None or newleftOut ==None:
                print("Error! Unknown device",MMI_name,"->",MMI_leftOut)
                exit(1)
            newMMI.attachLeftOut(newleftOut)
        if MMI_rightOut!="None":
            if newMMI==None or newrightOut ==None:
                print("Error! Unknown device",MMI_name,"->",MMI_rightOut)
                exit(1)
            newMMI.attachRightOut(newrightOut)

    for u in MZIheater_inform:
        MZI_name=str(u[0].value)
        MZI_leftIn=str(u[1].value)
        MZI_rightIn=str(u[2].value)
        MZI_leftOut=str(u[3].value)
        MZI_rightOut=str(u[4].value)

        newMZI=CheckIn(Chip.Longlist,MZI_name)
        newleftIn=CheckIn(Chip.Longlist,MZI_leftIn)
        newrightIn=CheckIn(Chip.Longlist,MZI_rightIn)
        newleftOut=CheckIn(Chip.Longlist,MZI_leftOut)
        newrightOut=CheckIn(Chip.Longlist,MZI_rightOut)
        if MZI_leftIn!="None":
            if newMZI==None or newleftIn ==None:
                print("Error! Unknown device",MZI_name,"->",MZI_leftIn)
                exit(1)
            newMZI.attachLeftIn(newleftIn)
        if MZI_rightIn!="None":
            if newMZI==None or newrightIn ==None:
                print("Error! Unknown device",MZI_name,"->",MZI_rightIn)
                exit(1)
            newMZI.attachRightIn(newrightIn)
        if MZI_leftOut!="None":
            if newMZI==None or newleftOut ==None:
                print("Error! Unknown device",MZI_name,"->",MZI_leftOut)
                exit(1)
            newMZI.attachLeftOut(newleftOut)
        if MZI_rightOut!="None":
            if newMZI==None or newrightOut ==None:
                print("Error! Unknown device",MZI_name,"->",MZI_rightOut)
                exit(1)
            newMZI.attachRightOut(newrightOut)

    for u in Phaseheater_inform:
        Pha_name=str(u[0].value)
        Pha_In=str(u[1].value)
        Pha_Out=str(u[2].value)
        newPha=CheckIn(Chip.Longlist,Pha_name)
        newIn=CheckIn(Chip.Longlist,Pha_In)
        newOut=CheckIn(Chip.Longlist,Pha_Out)

        if Pha_In!="None":
            if newPha==None or newIn ==None:
                print("Error! Unknown device",Pha_name,"->",Pha_In)
                exit(1)
            newPha.attachIn(newIn)
        if Pha_Out!="None":
            if newPha==None or newOut ==None:
                print("Error! Unknown device",Pha_name,"->",Pha_Out)
                exit(1)
            newPha.attachOut(newOut)
    return Chip


def findExistLoop(Chip,endMZI): #used for MZIcalibration
    initialize(Chip)#initialize map

    #BFS find one loop end at MZI
    Route=[]
    step=[]
    father=[]
    head=0
    if endMZI.leftIn==None:
        return False
    Route.append(endMZI.leftIn)
    step.append(0)
    father.append(-1)
    while (True): #left branch full searching
        if (type(Route[head])==Phaseheater and Route[head].In!=None and Route[head].In.flag==False):
            Route[head].In.flag=True
            Route.append(Route[head].In)
            step.append(step[head]+1)
            father.append(head)
        elif(type(Route[head])==MZIheater or type(Route[head])==MMI):
            if(Route[head].leftIn!=None and Route[head].leftIn.flag==False):
                Route[head].leftIn.flag=True
                Route.append(Route[head].leftIn)
                step.append(step[head]+1)
                father.append(head)
            if(Route[head].rightIn!=None and Route[head].rightIn.flag==False):
                Route[head].rightIn.flag=True
                Route.append(Route[head].rightIn)
                step.append(step[head]+1)
                father.append(head)
        head+=1
        if len(Route)<=head:
            break


    if endMZI.rightIn==None:
        return False
    Route.append(endMZI.rightIn)
    step.append(0)
    father.append(-1)
    while (True):
        if (type(Route[head])==Phaseheater and Route[head].In!=None and Route[head].In.flag==True):
            if ((type(Route[head].In)==MZIheater and Route[head].In.calibrated==False) or (type(Route[head].In)==MMI)):
                temp=0
                for temp in range(len(Route)):
                    if Route[temp]==Route[head].In:
                        break
                if(Route[father[temp]]!=Route[head]): #from different routes
                    return True
        elif (type(Route[head])==MZIheater or type(Route[head])==MMI):
            if(Route[head].leftIn!=None and Route[head].leftIn.flag==True):
                if ((type(Route[head].leftIn)==MZIheater and Route[head].leftIn.calibrated==False) or (type(Route[head].leftIn)==MMI)):
                    temp=0
                    for temp in range(len(Route)):
                        if Route[temp]==Route[head].leftIn:
                            break
                    if(Route[father[temp]]!=Route[head]): #from different routes
                        return True
            elif(Route[head].rightIn!=None and Route[head].rightIn.flag==True):
                if ((type(Route[head].rightIn)==MZIheater and Route[head].rightIn.calibrated==False) or (type(Route[head].rightIn)==MMI)):
                    temp=0
                    for temp in range(len(Route)):
                        if Route[temp]==Route[head].rightIn:
                            break
                    if(Route[father[temp]]!=Route[head]): #from different routes
                        return True

        if (type(Route[head])==Phaseheater and Route[head].In!=None and Route[head].In.flag2==False): #For MZI heater, the entrance heaters' calibration is not required. 
            Route[head].In.flag2=True
            Route.append(Route[head].In)
            step.append(step[head]+1)
            father.append(head)
        elif(type(Route[head])==MZIheater or type(Route[head])==MMI):
            if(Route[head].leftIn!=None and Route[head].leftIn.flag2==False):
                Route[head].leftIn.flag2=True
                Route.append(Route[head].leftIn)
                step.append(step[head]+1)
                father.append(head)
            if(Route[head].rightIn!=None and Route[head].rightIn.flag2==False):
                Route[head].rightIn.flag2=True
                Route.append(Route[head].rightIn)
                step.append(step[head]+1)
                father.append(head)
        head+=1
        if len(Route)<=head:
            return False #failed to find a loop
        

def calibrateOneMZIheater(Chip,heater_address_to_be_calibrated):
    if heater_address_to_be_calibrated.calibrated==True:
        print(heater_address_to_be_calibrated.key,"has already be calibrated!")
        exit(1)

    if findExistLoop(Chip,heater_address_to_be_calibrated)==True:
        return None

    initialize(Chip)#initialize map

    #BFS find one exit
    Route=[]
    step=[]
    father=[]
    head=0
    Route.append(heater_address_to_be_calibrated)
    step.append(0)
    father.append(-1)
    while (type(Route[head])!=port or Route[head].portAttribute!=1):
        if (type(Route[head])==Phaseheater and Route[head].Out!=None and Route[head].Out.flag==False and (type(Route[head].Out)==Phaseheater or Route[head].Out.calibrated==True)):
            Route[head].Out.flag=True
            Route.append(Route[head].Out)
            step.append(step[head]+1)
            father.append(head)
        elif(type(Route[head])!=Phaseheater):
            if(Route[head].leftOut!=None and Route[head].leftOut.flag==False and (type(Route[head].leftOut)==Phaseheater or  Route[head].leftOut.calibrated==True)):
                Route[head].leftOut.flag=True
                Route.append(Route[head].leftOut)
                step.append(step[head]+1)
                father.append(head)
            if(Route[head].rightOut!=None and Route[head].rightOut.flag==False and (type(Route[head].rightOut)==Phaseheater or Route[head].rightOut.calibrated==True)):
                Route[head].rightOut.flag=True
                Route.append(Route[head].rightOut)
                step.append(step[head]+1)
                father.append(head)
        head+=1
        if len(Route)<=head:
            return None #failed to find a route to exit
    exitRoute=[]
    temp=head
    while(father[temp]!=-1):
        exitRoute.insert(0,Route[temp])
        temp=father[temp]

    #BFS find one entrance
    Route=[]
    step=[]
    father=[]
    head=0
    Route.append(heater_address_to_be_calibrated)
    step.append(0)
    father.append(-1)
    while (type(Route[head])!=port or Route[head].portAttribute!=0):
        if (type(Route[head])==Phaseheater and Route[head].In!=None and Route[head].In.flag==False): #For MZI heater, the entrance heaters' calibration is not required. 
            Route[head].In.flag=True
            Route.append(Route[head].In)
            step.append(step[head]+1)
            father.append(head)
        elif(type(Route[head])!=Phaseheater):
            if(Route[head].leftIn!=None and Route[head].leftIn.flag==False):
                Route[head].leftIn.flag=True
                Route.append(Route[head].leftIn)
                step.append(step[head]+1)
                father.append(head)
            if(Route[head].rightIn!=None and Route[head].rightIn.flag==False):
                Route[head].rightIn.flag=True
                Route.append(Route[head].rightIn)
                step.append(step[head]+1)
                father.append(head)
        head+=1
        if len(Route)<=head:
            return None #failed to find a route to entrance
    entranceRoute=[]
    temp=head
    while(father[temp]!=-1):
        entranceRoute.append(Route[temp])
        temp=father[temp]

    Mode=0
    formerOne=entranceRoute[-1]
    laterOne=exitRoute[0]
    if (heater_address_to_be_calibrated.leftIn==formerOne and heater_address_to_be_calibrated.leftOut==laterOne)or(heater_address_to_be_calibrated.rightIn==formerOne and heater_address_to_be_calibrated.rightOut==laterOne):
        Mode=1
    else:
        Mode=2
    
    RouteAll=copy.deepcopy(entranceRoute)
    RouteAll.append(heater_address_to_be_calibrated)
    RouteAll.extend(exitRoute)
    return [RouteAll,Mode]


def calibrateOnePhaseheater(Chip,heater_address_to_be_calibrated):
    if heater_address_to_be_calibrated.calibrated==True:
        print(heater_address_to_be_calibrated.key,"has already be calibrated!")
        exit(1)
    if type(heater_address_to_be_calibrated.Out)!=MZIheater:
        print(heater_address_to_be_calibrated.key,"Warning! Unknown cases.")
        exit(1)
    
    initialize(Chip)#initialize map
    endMZI=heater_address_to_be_calibrated.Out
    #BFS find nearest loop end at MZI
    Route=[]
    step=[]
    father=[]
    head=0
    if endMZI.leftIn==None:
        print(heater_address_to_be_calibrated.key,"Error! Cannot find a loop. ")
        exit(1)
    Route.append(endMZI.leftIn)
    step.append(0)
    father.append(-1)
    while (True): #left branch full searching
        if (type(Route[head])==Phaseheater and Route[head].In!=None and Route[head].In.flag==False and not(type(Route[head].In)==Phaseheater and Route[head].In.calibration==False)):
            Route[head].In.flag=True
            Route.append(Route[head].In)
            step.append(step[head]+1)
            father.append(head)
        elif(type(Route[head])==MZIheater or type(Route[head])==MMI):
            if(Route[head].leftIn!=None and Route[head].leftIn.flag==False and not(type(Route[head].leftIn)==Phaseheater and Route[head].leftIn.calibration==False)):
                Route[head].leftIn.flag=True
                Route.append(Route[head].leftIn)
                step.append(step[head]+1)
                father.append(head)
            if(Route[head].rightIn!=None and Route[head].rightIn.flag==False and not(type(Route[head].rightIn)==Phaseheater and Route[head].rightIn.calibration==False)):
                Route[head].rightIn.flag=True
                Route.append(Route[head].rightIn)
                step.append(step[head]+1)
                father.append(head)
        head+=1
        if len(Route)<=head:
            break


    if endMZI.rightIn==None:
        print(heater_address_to_be_calibrated.key,"Error! Cannot find a loop. ")
        exit(1)
    Route.append(endMZI.rightIn)
    step.append(0)
    father.append(-1)
    frontPoint=None
    while (True):
        if (type(Route[head])==Phaseheater and Route[head].In!=None and Route[head].In.flag==True):
            if ((type(Route[head].In)==MZIheater and Route[head].In.calibrated==True) or (type(Route[head].In)==MMI)):
                temp=0
                for temp in range(len(Route)):
                    if Route[temp]==Route[head].In:
                        break
                if(Route[father[temp]]!=Route[head]): #from different routes
                    frontPoint=Route[head].In
                    break
        elif (type(Route[head])==MZIheater or type(Route[head])==MMI):
            if(Route[head].leftIn!=None and Route[head].leftIn.flag==True):
                if ((type(Route[head].leftIn)==MZIheater and Route[head].leftIn.calibrated==True) or (type(Route[head].leftIn)==MMI)):
                    temp=0
                    for temp in range(len(Route)):
                        if Route[temp]==Route[head].leftIn:
                            break
                    if(Route[father[temp]]!=Route[head]): #from different routes
                        frontPoint=Route[head].leftIn
                        break
            elif(Route[head].rightIn!=None and Route[head].rightIn.flag==True):
                if ((type(Route[head].rightIn)==MZIheater and Route[head].rightIn.calibrated==True) or (type(Route[head].rightIn)==MMI)):
                    temp=0
                    for temp in range(len(Route)):
                        if Route[temp]==Route[head].rightIn:
                            break
                    if(Route[father[temp]]!=Route[head]): #from different routes
                        frontPoint=Route[head].rightIn
                        break

        if (type(Route[head])==Phaseheater and Route[head].In!=None and Route[head].In.flag2==False and not(type(Route[head].In)==Phaseheater and Route[head].In.calibration==False)): 
            Route[head].In.flag2=True
            Route.append(Route[head].In)
            step.append(step[head]+1)
            father.append(head)
        elif(type(Route[head])==MZIheater or type(Route[head])==MMI):
            if(Route[head].leftIn!=None and Route[head].leftIn.flag2==False and not(type(Route[head].leftIn)==Phaseheater and Route[head].leftIn.calibration==False)):
                Route[head].leftIn.flag2=True
                Route.append(Route[head].leftIn)
                step.append(step[head]+1)
                father.append(head)
            if(Route[head].rightIn!=None and Route[head].rightIn.flag2==False and not(type(Route[head].rightIn)==Phaseheater and Route[head].rightIn.calibration==False)):
                Route[head].rightIn.flag2=True
                Route.append(Route[head].rightIn)
                step.append(step[head]+1)
                father.append(head)
        head+=1
        if len(Route)<=head:
            return None #failed to find a loop

    
    midRoute1=[]
    temp=head
    while temp!=-1:
        midRoute1.append(Route[temp])
        temp=father[temp]
    midRoute2=[]
    temp=0
    for temp in range(len(Route)):
        if Route[temp]==frontPoint:
            break
    while temp!=-1:
        midRoute2.append(Route[temp])
        temp=father[temp]
    
    initialize(Chip)#initialize map
    
    #begin to find entrance route and exit route
    #BFS find one exit
    Route=[]
    step=[]
    father=[]
    head=0
    Route.append(endMZI)
    step.append(0)
    father.append(-1)
    while (type(Route[head])!=port or Route[head].portAttribute!=1):
        if (type(Route[head])==Phaseheater and Route[head].Out!=None and Route[head].Out.flag==False and (type(Route[head].Out)==Phaseheater or Route[head].Out.calibrated==True)):
            Route[head].Out.flag=True
            Route.append(Route[head].Out)
            step.append(step[head]+1)
            father.append(head)
        elif(type(Route[head])!=Phaseheater):
            if(Route[head].leftOut!=None and Route[head].leftOut.flag==False and (type(Route[head].leftOut)==Phaseheater or  Route[head].leftOut.calibrated==True)):
                Route[head].leftOut.flag=True
                Route.append(Route[head].leftOut)
                step.append(step[head]+1)
                father.append(head)
            if(Route[head].rightOut!=None and Route[head].rightOut.flag==False and (type(Route[head].rightOut)==Phaseheater or Route[head].rightOut.calibrated==True)):
                Route[head].rightOut.flag=True
                Route.append(Route[head].rightOut)
                step.append(step[head]+1)
                father.append(head)
        head+=1
        if len(Route)<=head:
            return None #failed to find a route to exit
    exitRoute=[]
    temp=head
    while(father[temp]!=-1):
        exitRoute.insert(0,Route[temp])
        temp=father[temp]

    #BFS find one entrance
    Route=[]
    step=[]
    father=[]
    head=0
    Route.append(frontPoint)
    step.append(0)
    father.append(-1)
    while (type(Route[head])!=port or Route[head].portAttribute!=0):
        if (type(Route[head])==Phaseheater and Route[head].In!=None and Route[head].In.flag==False): #For MZI heater, the entrance heaters' calibration is not required. 
            Route[head].In.flag=True
            Route.append(Route[head].In)
            step.append(step[head]+1)
            father.append(head)
        elif(type(Route[head])!=Phaseheater):
            if(Route[head].leftIn!=None and Route[head].leftIn.flag==False):
                Route[head].leftIn.flag=True
                Route.append(Route[head].leftIn)
                step.append(step[head]+1)
                father.append(head)
            if(Route[head].rightIn!=None and Route[head].rightIn.flag==False):
                Route[head].rightIn.flag=True
                Route.append(Route[head].rightIn)
                step.append(step[head]+1)
                father.append(head)
        head+=1
        if len(Route)<=head:
            return None #failed to find a route to entrance
    entranceRoute=[]
    temp=head
    while(father[temp]!=-1):
        entranceRoute.append(Route[temp])
        temp=father[temp]

    Mode=0
    formerOne=entranceRoute[-1]
    laterOne=exitRoute[0]
    if (frontPoint.leftIn==formerOne and endMZI.leftOut==laterOne)or(frontPoint.rightIn==formerOne and endMZI.rightOut==laterOne):
        Mode=1
    else:
        Mode=2

    return [entranceRoute,frontPoint,midRoute1,midRoute2,endMZI,exitRoute,Mode]



    


def calibrateAllMZIheater(Chip):
    for u in Chip.MZIheater:
        u.calibrated=False
    for u in Chip.Phaseheater:
        u.calibrated=False

    flag0=True
    Count=0
    while(flag0):
        flag0=False
        for u in Chip.MZIheater:
            if u.calibrated==True:
                continue
            
            back=calibrateOneMZIheater(Chip,u)
            if back != None:
                [Route,Mode]=back
                flag0=True
                Count+=1
                u.calibrated=True


                print("\n","step",Count,":", u.key,", Mode =",Mode)
                tag=0
                for u in Route:
                    if tag!=0:
                        print("->",end="")
                    tag=1
                    print(u.key,end="")  
                print() 


    if Count != len(Chip.MZIheater):
        print("\n","Waring! MZI heaters can not all be calibrated!")
        print()



def calibrateAllPhaseheater(Chip):
    for u in Chip.Phaseheater:
        u.calibrated=False
        
    flag0=True
    Count=0
    while(flag0):
        flag0=False
        for u in Chip.Phaseheater:
            if u.calibrated==True:
                continue
            
            back=calibrateOnePhaseheater(Chip,u)
            if back != None:
                [entranceRoute,frontPoint,midRoute1,midRoute2,endMZI,exitRoute,Mode]=back
                flag0=True
                Count+=1
                u.calibrated=True


                print("\n","step",Count,":", u.key,", Mode =",Mode)

                tag=0
                for u in entranceRoute:
                    if tag!=0:
                        print("->",end="")
                    tag=1
                    print(u.key,end="")

                print()
                print(frontPoint.key,"\{")

                tag=0
                for u in midRoute1:
                    if tag!=0:
                        print("->",end="")
                    tag=1
                    print(u.key,end="")
                print()

                tag=0
                for u in midRoute2:
                    if tag!=0:
                        print("->",end="")
                    tag=1
                    print(u.key,end="")
                print()
                print("\}",endMZI.key)

                tag=0
                for u in exitRoute:
                    if tag!=0:
                        print("->",end="")
                    tag=1
                    print(u.key,end="")
                print()


    if Count != len(Chip.Phaseheater):
        print("\n","Waring! Phase heaters can not all be calibrated!")







if __name__=="__main__":
    path=r"toyModel.xlsx"
    Chip=creatChip("OAMChip",path)
    print(Chip.key)
    # for u in Chip.Longlist:
    #     print(u.key)

    # print(Chip.inPort[0] is Chip.MMI[0].rightIn)     
    # print(Chip.inPort[0].key, Chip.MMI[0].rightIn.key)

    calibrateAllMZIheater(Chip)
    calibrateAllPhaseheater(Chip)







        

        




        

    

        










