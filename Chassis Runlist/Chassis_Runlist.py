import core.tdm as tdm
import core.tdm.trace
import core.eam as eam
from core.tdm.trace import TraceResults, TraceItemEntities
import core.gui as gui
import core.gdm
from core.gdm.lookuptables import *
from core.gui import *
from core.gui.editpanel import *
import datetime
import os
import sys
import core.jms
import win32com.client
from win32com.client import constants

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
port_info = ConfigurationDictionary("PORT_DICT")
equip_dict = ConfigurationDictionary("EQDICT")
trace_Reports = []
trace_Reports_Desc = []
master_circuits = []

trace_Reports_Desc.append("A Site")
trace_Reports_Desc.append("A Site name")
trace_Reports_Desc.append("A Site CLLI")
trace_Reports_Desc.append("A Site type")
trace_Reports_Desc.append("A Site location")
trace_Reports_Desc.append("A Site address")
trace_Reports_Desc.append("A Site End Equip")
trace_Reports_Desc.append("A Site Equip")
trace_Reports_Desc.append("Usage")
trace_Reports_Desc.append("Z Site")
trace_Reports_Desc.append("Z name")
trace_Reports_Desc.append("Z CLLI")
trace_Reports_Desc.append("Z type")
trace_Reports_Desc.append("Z location")
trace_Reports_Desc.append("Z address")
trace_Reports_Desc.append("Z End Equip")
trace_Reports_Desc.append("Z Equip")
trace_Reports_Desc.append("Date")


try:
	sel = gdm.selected_entity()

	child = SPATIALnet.service("eam$find_slaves",sel,"*","*")

	leaf = []

	slave = []
	slave2 = []
	for elem in child:
		leaf.append(elem[0])
	slave=findSlave(leaf)


	slave = list(set(slave))
	slave.sort()

	for port in slave:
		card = port.ISPA_PORT_OWNER_FK
		print "Port: ", card.ISPA_NAME,port.ISPA_PORT_NAME
		#if port.is_class("ISP_TRACEABLE") or port.is_class("_fdm_traceable"):
		#	slave2.append(port)

	print "Total Number of Ports: ", len(slave)

	#slave3 = []

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
				print "Try running on", card.ISPA_NAME, port.ISPA_PORT_NAME 
				results = trace.run()

				for result in results.getTraceResults():

					entity_list = []
					print "Try printing"

					def storeTraceResult(node,direction,parent):
						entity_list.append(node) #append the nodes to the end of the entity list

					result.trace_tree.applyBidirectional(core.tdm.trace.TraceNode.printCallback, walk_type = "bidirectional")
					print "----------"
					result.trace_tree.applyBidirectional(storeTraceResult)

					if len(entity_list)>1: #if the length of the entity list is more than 1 then
						trace_Report = attributes = [""]*18 #declare trace_reports with 23 empty attributes
						a_end_isp_design = ""
						a_end_equip = []
						a_end_osp_cable = ""
						osp_equip = None
						correct_order = True
						first_port = None

						for i in range(len(entity_list)): #loop through the entity list
							ent2 = entity_list[i].entity #take the ith entity and put it into ent2
							if ent2.is_class("ISP_PORT"): #if the current entity is an ISP_PORT
								print str(entity_list[i].entity)+ " : depth("+ str(entity_list[i].depth)+"), branch("+str(entity_list[i].branch_number)+")"
								if first_port is None: #and if the first port has not yet been found then
									first_port = ent2  #make the ith entity the first port
								parent = ent2.ISPA_PORT_OWNER_FK #find the parent of the ent2
								chassis = getChassis(ent2)
								if parent.fdm_interface_fk is not None: #if the parent has an osp interface then
									#found patch panel
									  #get the chassis of ent2
									pnl = checkValue(chassis.ISPA_NAME) + " ; "+checkValue(chassis.ISPA_SECTION_F_CODE) + " ; " + checkValue(ent2.ISPA_PORT_NAME) 
									a_end_equip.append(addedInJob(chassis,"Patch Panel")+": "+pnl) #get the name of the panel and all information
							else:     #if the parent has no osp interface
						#Dictionary look up for isp equipment
								try:
									equip_type_details = equip_dict.values(parent.ISPA_EQUIP_DICT_FK.NETWORK_KEY)  #get the details of the type of equipment that is ent2
									desc = checkValue(equip_type_details.DESC1) #get the description of the chassis from the dictionary
									a_end_isp_design = ent2,"End Equipment")+": "+checkValue(ent2.ISPA_SECTION_F_CODE)  + " ; " + checkValue(chassis.ISPA_NAME)+ " ; " + checkValue(parent.ISPA_NAME) + " - "+ checkValue(chassis.gdm_ea_attr_21) + " | " + checkValue(chassis.gdm_ea_attr_20)
									if portmc == True:
										trace_Report[19] = checkValue(ent2.fdm_ringmaster_fk.fdm_ringmaster_name)
									#	#found true end
									if ent2 != first_port and entity_list[i].branch_number==1: #if ent2 is not the first port and the ith element in the entity list's branch # = 1 then
										correct_order=False #the entities are not in the correct order
									#add all information for the isp a end
									a_end_isp_design = ent2,"End Equipment")+": "+checkValue(ent2.ISPA_SECTION_F_CODE) + " ; " + checkValue(parent.ISPA_NAME) + " ; "+equip_type_details.MODEL + " ; " + equip_type_details.DESC1

								except Exception as e:
									#lov conversion not found
									print e
			except Exception as e:
				print e
				continue
		else:
			print card.ISPA_NAME, port.ISPA_PORT_NAME, ": Not Connected"

	#for port in slave3:
	#	card = port.ISPA_PORT_OWNER_FK
	#	print "Port: ", card.ISPA_NAME,port.ISPA_PORT_NAME, "Not Connected"
except Exception as e:
	print e