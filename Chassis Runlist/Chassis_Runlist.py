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
import core.jms
import win32com.client
from win32com.client import constants
from core.reports.excel import ExcelApplication, constants

# Two lines below generated from the command:
#python makepy.py -i VISLIB.DLL
from win32com.client import gencache
gencache.EnsureModule('{00021A98-0000-0000-C000-000000000046}', 0, 4, 11)

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
	trace_Reports = []
	trace_Report = attributes = [""]*15
	trace_Reports_Desc = []

	trace_Reports_Desc.append("A Site")	
	trace_Reports_Desc.append("A name")
	trace_Reports_Desc.append("A CLLI")
	trace_Reports_Desc.append("A location")
	trace_Reports_Desc.append("A address")
	trace_Reports_Desc.append("A Chassis")
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
		for elem in child:
			leaf.append(elem[0])
		slave=findSlave(leaf)

		slave = list(set(slave))
		slave.sort()

		#for port in slave:
		#	card = port.ISPA_PORT_OWNER_FK
		#	print "Port: ", card.ISPA_NAME,port.ISPA_PORT_NAME

		print "Total Number of Ports: ", len(slave)
		print ""
		equipList = []
		projectList = []
		ZSite = []
		ZName = []
		ZCLLI = []
		ZLocation = []
		ZAddress = []
		Date = []
		CirucuitID = []
		for port in slave:
			card = port.ISPA_PORT_OWNER_FK
			cables = SPATIALnet.service("eam$find_slaves",port,"*","*")
			if len(cables) > 0 :
				try:
					#print "Try Tracing"
					#print cables[0]
					start = [core.tdm.trace.TraceStartPoint(cables[0][0],1)]
					#print start[0]
					trace = core.tdm.trace.Trace(start)
					print ""
					print "-----------------------------------------------------------------"
					print "Port: ", card.ISPA_NAME, port.ISPA_PORT_NAME 
					results = trace.run()
					if port.fdm_ringmaster_fk is not None:
						if port.fdm_ringmaster_fk.fdm_ringmaster_name is not None:
								trace_Report[13] = checkValue(port.fdm_ringmaster_fk.fdm_ringmaster_name)
								print trace_Report[13]
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

							end_isp_design = []
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
										pnl = checkValue(chassis.ISPA_NAME) + " ; "+checkValue(chassis.ISPA_SECTION_F_CODE) + " ; " + checkValue(ent2.ISPA_PORT_NAME) 
										equip.append("Patch Panel"+": "+pnl) #get the name of the panel and all information
									else:     #if the parent has no osp interface
											#Dictionary look up for isp equipment
										try:
											equip_type_details = equip_dict.values(parent.ISPA_EQUIP_DICT_FK.NETWORK_KEY)  #get the details of the type of equipment that is ent2
											desc = checkValue(equip_type_details.DESC1) #get the description of the chassis from the dictionary
									
											#	#found true end
											if ent2 != first_port and entity_list[i].branch_number==1: #if ent2 is not the first port and the ith element in the entity list's branch # = 1 then
												correct_order=False #the entities are not in the correct order
											#add all information for the isp a end
											equip.append( "End Equipment"+": "+checkValue(ent2.ISPA_SECTION_F_CODE)  + " ; " + checkValue(chassis.ISPA_NAME)+ " ; " + checkValue(parent.ISPA_NAME)) 
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
						for e in equip:
							print e 
						equipList.append(equip)
				except Exception as e:
					print e
					continue

			else:
				print ""
				print "-----------------------------------------------------------------"
				print card.ISPA_NAME, port.ISPA_PORT_NAME, ": Not Connected"
				print "-----------------------------------------------------------------"
				print "" 

	except Exception as e:
		print e

class RunlistGenerator():
	def createReport(self, RunlistData):
		version = ExcelApplication.getExcelVersion()
		if version is not None:
			xl = ExcelApplication()
			wb = xl.new_workbook()
			sheet = wb.addsheet("Runlist")
class RunlistData():
	ASite = None
	AName = None
	ACLLI = None
	ALocation = None
	AAddress = None
	AChassis = None
	Equip = []
	Project = []
	ZSite = []
	ZName = []
	ZCLLI = []
	ZLocation = []
	ZAddress = []
	CircuitID = []
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
			data[i].Equip.append(dataArray[i][6])
			data[i].Project.append(dataArray[i][7])
			data[i].ZSite.append(dataArray[i][8])
			data[i].ZName.append(dataArray[i][9])
			data[i].ZCLLI.append(dataArray[i][10])
			data[i].ZLocation.append(dataArray[i][11])
			data[i].ZAddress.append(dataArray[i][12])
			data[i].CircuitID.append(dataArray[i][13])
			data[i].Date= dataArray[i][14]
		result.append(data)
		return result


if __name__ == '__main__':
	exl = RunlistGenerator()
	exl.createReport(RunlistData().parseArray(main()))