import core.tdm as tdm
import core.tdm.trace
import core.eam as eam
from core.tdm.trace import TraceResults, TraceItemEntities
import core.gui as gui
import core.gdm
import gdm
from core.gdm.lookuptables import *
from core.gui import *
from core.gui.editpanel import *
import datetime
import os
import itertools
import sys
import re
import core.jms
#import win32com.client
from reports.excel import ExcelApplication

# Two lines below generated from the command:
#python makepy.py -i VISLIB.DLL
#from win32com.client import gencache
#gencache.EnsureModule('{00021A98-0000-0000-C000-000000000046}', 0, 4, 11)

#===========================================================================================================================
# Return true if the passed in entity is an ISP equipment, otherwise return false
#===========================================================================================================================
def is_isp_class(ent):
	isp_classes = ["ISP_RACK","ISP_PORT_AND_OWNER_mixin","ISP_CABLE", "TERM_PORTGR","FIBER_CABLE_SEG_ISP","COUPLER_PORTGR"]
	if ent is None:
		return False
	for isp_class in isp_classes:
		if ent.is_class(isp_class):
			return True
	return False

#===========================================================================================================================
# Return False if we reach a "stop" class entity or if the entity is null, otherwise return true
#===========================================================================================================================
def is_stop_class(ent):
	stop_classes = ["SPLICE_ENCLOSURE","RF_NODE","fdm_storage_loop"] #Removed SITE from list
	if ent is None:
		return False
	for stop_class in stop_classes:
		if ent.is_class(stop_class):
			return False
	return True
	
#===========================================================================================================================
# Helper function for retreving attributes
#===========================================================================================================================
def checkValue(value):
	if value is None:
		return ""
	else:
		return str(value)



def checkValue(value):
	if value is None:
		return ""
	else:
		return str(value)

def getLocation(ent):
	locCode = str(ent.ISPA_SECTION_F_CODE)
	locate = locCode.split(".")

	for i in range(len(locate)):
		if i == 0:
			loc = "Floor:" + locate[i]
		elif i == 1:
			loc = loc + "; " + "Row: " + locate[i]
		elif i == 2:			
			if not (re.search("rack", locate[i].lower())):
				loc = loc + "; " + "Rack: " +locate[i]
			else:
				loc = loc + "; " +locate[i]
		elif i == 3:
			loc = loc + "; " + "RU: " + locate[i]
		elif i ==4:
			loc = loc + "; " + "Slot: " + locate[i]
		elif i > 4 and i < len(locate)-2:
			loc = loc + "/" + locate[i]
		elif i == len(locate)-2 and i > 4:
			loc = loc + "; " + "Port: " + locate[i]
		elif i == len(locate)-1 and len(locate)-2 == 4:
			loc = loc + "; " + "Port: " + locate[i]
		elif i == len(locate)-1 and i > 4:
			loc = loc + " " + locate[i]  
	return loc

def getChassis(ent):

	parent = ent
	while True:
		try:
			if parent.is_class("ISP_PORT"):
				parent = parent.ISPA_PORT_OWNER_FK
			else:
				parent = parent.PARENT_NODEHOUSING
		except:
			break

		if parent.is_class("ISP_CHASSIS"):
			break

	return parent
def getProject(entity):
	if entity.is_class("ISP_PORT"):
		parent = entity.ISPA_PORT_OWNER_FK
		if parent.is_class("ISP_CARD"):
			project =checkValue(parent.gdm_ea_attr_29)
		else:
			project = ""
	elif entity.is_class("ISP_CARD"):
		project = checkValue(entity.gdm_ea_attr_29)
	else:
		project = ""
	return project


def findSlave(entityList):
	slave = []
	element = []
	for ent in entityList: 
		if ent.is_class("ISP_PORT"):
			slave.append(ent)

		else:
			s =  SPATIALnet.service("eam$find_slaves", ent , "*", "*")
			next = []
			for elem in s:
				next.append(elem[0])

			element = findSlave(next)
			for elem in element:
				slave.append(elem)

	return slave

def main():


	port_info = ConfigurationDictionary("PORT_DICT")
	equip_dict = ConfigurationDictionary("EQDICT")
	
	trace_Reports_Desc = []
	trace_Reports = []
	trace_Reports_Desc.append("A Site")	
	trace_Reports_Desc.append("A name")
	trace_Reports_Desc.append("A CLLI")
	trace_Reports_Desc.append("A location")
	trace_Reports_Desc.append("A address")
	trace_Reports_Desc.append("A Chassis")
	trace_Reports_Desc.append("Ports")
	trace_Reports_Desc.append("Equip")
	trace_Reports_Desc.append("Project")
	trace_Reports_Desc.append("Z Site")
	trace_Reports_Desc.append("Z name")
	trace_Reports_Desc.append("Z CLLI")
	trace_Reports_Desc.append("Z location")
	trace_Reports_Desc.append("Z address")
	trace_Reports_Desc.append("Circuit ID")
	trace_Reports_Desc.append("Date")
	try:
		sel = gdm.selected_entity()
		child = SPATIALnet.service("eam$find_slaves",sel,"*","*")
		leaf = []
		slave = []
		project = ""
		for elem in child:
			leaf.append(elem[0])
		slave=findSlave(leaf)

		slave = list(set(slave))
		slave.sort()
		ports = []
		for port in slave:
			card = port.ISPA_PORT_OWNER_FK
			print "Port: ", card.ISPA_NAME,port.ISPA_PORT_NAME

		print "Total Number of Ports: ", len(slave)
		print ""
		for port in slave:
			card = port.ISPA_PORT_OWNER_FK
			cables = SPATIALnet.service("eam$find_slaves",port,"*","*")
			trace_Report = attributes = [""]*16
			project = getProject(port)

			trace_Report[8] = project

			print "Project: ", project
			if sel.is_class("ISP_CHASSIS") or sel.is_class("ISP_RACK"):
						site = sel.ISPA_BUILDING_FK
						trace_Report[0] = str(site.NETWORK_KEY)
						trace_Report[1] = str(site.fdm_designation)
						trace_Report[2] = str(site.ISPA_CLLI)
						trace_Report[3] = str(site.fdm_nh_location)
						trace_Report[4] = "%s ; %s ; %s ; %s" % (site.fdm_address1, site.fdm_town,site.fdm_state,site.fdm_zipcode)
						if sel.is_class("ISP_CHASSIS"):
								trace_Report[5] = getLocation(sel) + ": " + str(sel.ISPA_NAME)
			if len(cables) > 0 :
				try:
					print "Try Tracing"
					print cables[0]
					start = [core.tdm.trace.TraceStartPoint(cables[0][0],1)]
					print start[0]
					trace = core.tdm.trace.Trace(start)
					print ""
					print "-----------------------------------------------------------------"
					print "Port: ", card.ISPA_NAME," ", port.ISPA_PORT_NAME 
					
					
					if card.ISPA_NAME is not None: 
						trace_Report[6] = "Port: " + card.ISPA_NAME + " "+port.ISPA_PORT_NAME
					else:
						trace_Report[6] = "Port: " +  port.ISPA_PORT_NAME
					results = trace.run()
					master_circuit = ""
					if port.fdm_ringmaster_fk is not None:
						if port.fdm_ringmaster_fk.fdm_ringmaster_name is not None:
								master_circuit = checkValue(port.fdm_ringmaster_fk.fdm_ringmaster_name)
								print master_circuit
					for result in results.getTraceResults():

						entity_list = []
						#print "Try printing"

						def storeTraceResult(node,direction,parent):
							entity_list.append(node) #append the nodes to the end of the entity list

						#result.trace_tree.applyBidirectional(core.tdm.trace.TraceNode.printCallback, walk_type = "bidirectional")
						print "-----------------------------------------------------------------"
						result.trace_tree.applyBidirectional(storeTraceResult)

					
						#master circuit details

						if len(entity_list)>1: #if the length of the entity list is more than 1 then
							 #declare trace_reports with 23 empty attributes
							z_end_nh = entity_list[len(entity_list)-1].upstream_osp_nh
							z_end_key = z_end_nh.NETWORK_KEY
							z_end_address = "%s ; %s ; %s ; %s" % (z_end_nh.fdm_address1, z_end_nh.fdm_town,z_end_nh.fdm_state,z_end_nh.fdm_zipcode)
							z_end_name = z_end_nh.fdm_designation
							z_end_clli = z_end_nh.gdm_ea_attr_01
							z_end_type = z_end_nh.fdm_site_type_code
							z_end_location = z_end_nh.fdm_nh_location 

							equip = []
							project = ""
							correct_order = True
							first_port = None

							for i in range(len(entity_list)): #loop through the entity list
								ent2 = entity_list[i].entity #take the ith entity and put it into ent2
								if ent2.is_class("ISP_PORT"): #if the current entity is an ISP_PORT
									#print str(entity_list[i].entity)+ " : depth("+ str(entity_list[i].depth)+"), branch("+str(entity_list[i].branch_number)+")"
									if first_port is None: #and if the first port has not yet been found then
										first_port = ent2  #make the ith entity the first port
									parent = ent2.ISPA_PORT_OWNER_FK #find the parent of the ent2
									chassis = getChassis(ent2)
									if parent.fdm_interface_fk is not None: #if the parent has an osp interface then
										#found patch panel
										  #get the chassis of ent2
										pnl =  checkValue(getLocation(ent2))+ " ; " +checkValue(chassis.ISPA_NAME)+ " ; " + checkValue(parent.ISPA_NAME) 
										equip.append("Patch Panel: "+pnl) #get the name of the panel and all information
									else:     #if the parent has no osp interface
											#Dictionary look up for isp equipment
										try:
											equip_type_details = equip_dict.values(parent.ISPA_EQUIP_DICT_FK.NETWORK_KEY)  #get the details of the type of equipment that is ent2
											desc = checkValue(equip_type_details.DESC1) #get the description of the chassis from the dictionary
									
											#	#found true end
											if ent2 != first_port and entity_list[i].branch_number==1: #if ent2 is not the first port and the ith element in the entity list's branch # = 1 then
												correct_order=False #the entities are not in the correct order
											#add all information for the isp a end
											equip.append("End Equipment"+": "+checkValue(getLocation(ent2))  + " ; " + checkValue(chassis.ISPA_NAME)+ " ; " + checkValue(parent.ISPA_NAME)) 
										except Exception as e:
											#lov conversion not found
											print e
								elif ent2.is_class("ISP_PATCH_CORD"):
									desc = "Length: "+ str(ent2.LE_LENGTH) + " ; "
					
									#Dictionary look up for isp equipment
									try:
										equip_type_details = equip_dict.values(ent2.ISPA_EQUIP_DICT_FK.NETWORK_KEY)
										desc = desc + equip_type_details.MODEL
									except Exception as e:
										#lov conversion not found
										#pass
										print (str(e))
					
									equip.append("Patch Cable: "+desc)
						if first_port != port:
							equip.reverse()
						for e in equip:
							print e 
					
					trace_Report[7] = equip
					trace_Report[9] = z_end_nh
					trace_Report[10] = z_end_name
					trace_Report[11] = z_end_clli
					trace_Report[12] = z_end_location
					trace_Report[13] = z_end_address
					trace_Report[14] = master_circuit
					trace_Report[15] = str(datetime.datetime.now().date())
					
				except Exception as e:
					print e
					continue
				
			else:
				print ""
				print "-----------------------------------------------------------------"
				print card.ISPA_NAME, port.ISPA_PORT_NAME, ": Not Connected"
				if card.ISPA_NAME is not None:
					trace_Report[6] = "Port: " + card.ISPA_NAME +  " " +port.ISPA_PORT_NAME + ": Not Connected"
				else:
					trace_Report[6] = "Port: " + port.ISPA_PORT_NAME + ": Not Connected"
				print "-----------------------------------------------------------------"
				print ""
			trace_Reports.append(trace_Report)
	except Exception as e:
		print e
	for report in trace_Reports:
		print report
	return trace_Reports

class RunlistGenerator:
	def __init__(self):
		self.exl = None
		self.Workbook = None
		self.WorkSheet = None
		self.com = None

	def createReport(self, RunlistData):
		try:
			version = ExcelApplication.getExcelVersion()
			if version is not None:
				self.exl = ExcelApplication()
				self.Workbook = self.exl.new_workbook()
				self.WorkSheet = self.Workbook.addsheet("Runlist")
				self.WorkSheet.activate()
		except:
			sys.exit("Please ensure Microsft Excel is installed.")
		try:
			self.exl.show()
			self.Workbook.removedefaultsheets()
			length = len(RunlistData[0])
			self.WorkSheet.setlocation(1,1)
			row = 2
			col = 1
			startrow = 0
			endrow = 0 
			portnum = 0
			self.WorkSheet.COM().Range("A1:J1").Merge()
			self.WorkSheet.COM().Cells(1,1).value = RunlistData[0][0].AAddress + " "+  RunlistData[0][0].ACLLI + " "+ RunlistData[0][0].AName + " "+  RunlistData[0][0].ALocation + " " + RunlistData[0][0].AChassis.replace("; ", " ")
			self.WorkSheet.COM().Cells(2,1).value = "Port"
			self.WorkSheet.COM().Cells(2,2).value = "Far Side Address"
			self.WorkSheet.COM().Cells(2,3).value = "Circuit ID"
			self.WorkSheet.COM().Cells(2,4).value = "Project"
			self.WorkSheet.COM().Cells(2,5).value = "Patch Cord Length"
			self.WorkSheet.COM().Cells(2,6).value = "Patch Cord Type"
			self.WorkSheet.COM().Cells(2,7).value = "A Side Floor"
			self.WorkSheet.COM().Cells(2,8).value = "A Side Row"
			self.WorkSheet.COM().Cells(2,9).value = "A Side Rack"
			self.WorkSheet.COM().Cells(2,10).value = "A Side RU"
			self.WorkSheet.COM().Cells(2,11).value = "A Side Slot"
			self.WorkSheet.COM().Cells(2,12).value = "A Side Port"
			self.WorkSheet.COM().Cells(2,13).value = "A Side Chassis"
			self.WorkSheet.COM().Cells(2,14).value = "A Side Full Name"
			self.WorkSheet.COM().Cells(2,15).value = "Z Side Floor"
			self.WorkSheet.COM().Cells(2,16).value = "Z Side Row"
			self.WorkSheet.COM().Cells(2,17).value = "Z Side Rack"
			self.WorkSheet.COM().Cells(2,18).value = "Z Side RU"
			self.WorkSheet.COM().Cells(2,19).value = "Z Side Slot"
			self.WorkSheet.COM().Cells(2,20).value = "Z Side Port"
			self.WorkSheet.COM().Cells(2,21).value = "Z Side Chassis"
			self.WorkSheet.COM().Cells(2,22).value = "Z Side Full Name"
			self.WorkSheet.setbold(1,1,1,2,24)
			for i in range(0,length-1):
				portnum = portnum + 1
				row = row + 1
				startrow = row
				col  = 1
				nonetype =  "%s ; %s ; %s ; %s" %(None,None,None,None)
				self.WorkSheet.COM().Cells(row,"A").value = RunlistData[0][i].Port
				if RunlistData[0][i].ZAddress != RunlistData[0][i].AAddress and RunlistData[0][i].ZAddress != nonetype:
					self.WorkSheet.COM().Cells(row,"B").value = RunlistData[0][i].ZName+ " "+ RunlistData[0][i].ZAddress 
				if RunlistData[0][i].CircuitID != "":
					self.WorkSheet.COM().Cells(row,"C").value = RunlistData[0][i].CircuitID
				if RunlistData[0][i].Project != "":
					self.WorkSheet.COM().Cells(row,"D").value = RunlistData[0][i].Project
				self.WorkSheet.setborder([4,8],1,3,1,row,col,row,col)
				self.WorkSheet.setbold(1,row,col,row,col)
				col = 7
				equipnum = 0
				for eq in RunlistData[0][i].Equip:
					
					if equipnum == 3:
						row = row + 1
						col = 7
						equipnum = 0
					elif equipnum == 2:
						col = 15
					equipnum = equipnum + 1
					equip = []
					types = []
					types = eq.split(": ")
					type = types[0]
					equipment = eq.split("; ")
					temp = []
					temp = equipment[0].split(": ")
					equipment[0] = temp[1]
					
					if type == "Patch Cable":
						col2 = 5
						equipment[0] = temp[1] + ": " + temp [2]
						for equip in equipment:
							self.WorkSheet.COM().Cells(row,col2).value = equip
							col2 = col2 + 1
					else:
						for equip in equipment:
							self.WorkSheet.COM().Cells(row,col).value = equip
							col = col + 1
						
					
					
				endrow = row
				if portnum%2 == 0:
					self.WorkSheet.setcolorindex(15, startrow,1,endrow,99)
			self.WorkSheet.COM().Columns("A:Z").AutoFit()
			self.WorkSheet.sethorizontalalignment(2, 1,1, 1000,26)
			self.WorkSheet.activate()
			self.exl.COM().Range("B3").Select()
			self.exl.COM().ActiveWindow.FreezePanes = True
		except Exception as e:
			 print e
		filename = RunlistData[0][0].AChassis
		filename = filename.split(": ")
		leng = len(filename) - 1
		self.Workbook.saveas("C:\\D-Drive\\" + filename[leng] + ".xls")





class RunlistData:
	ASite = None
	AName = None
	ACLLI = None
	ALocation = None
	AAddress = None
	AChassis = None
	Port = None
	Equip = None
	Project = None
	ZSite = None
	ZName = None
	ZCLLI = None
	ZLocation = None
	ZAddress = None
	CircuitID = None
	Date = None
	def parseArray(self, dataArray):
		result = []
		data = []
		for i in range(0, len(dataArray)):
			data.append(RunlistData())
			
			data[i].ASite = dataArray[i][0]
			data[i].AName= dataArray[i][1]
			data[i].ACLLI = dataArray[i][2]
			data[i].ALocation = dataArray[i][3]
			data[i].AAddress = dataArray[i][4]
			data[i].AChassis = dataArray[i][5]
			data[i].Port = dataArray[i][6]
			data[i].Equip = dataArray[i][7]
			data[i].Project = dataArray[i][8]
			data[i].ZSite = dataArray[i][9]
			data[i].ZName = dataArray[i][10]
			data[i].ZCLLI = dataArray[i][11]
			data[i].ZLocation = dataArray[i][12]
			data[i].ZAddress = dataArray[i][13]
			data[i].CircuitID = dataArray[i][14]
			data[i].Date= dataArray[i][15]
		result.append(data)
		return result

if __name__ == '__main__':
	exl = RunlistGenerator()
	exl.createReport(RunlistData().parseArray(main()))