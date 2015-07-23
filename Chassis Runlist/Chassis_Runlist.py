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
	sel = gdm.selected_entity()
	if sel.is_class("ISP_CHASSIS"):
		sel = gdm.selected_entity()
		child = SPATIALnet.service("eam$find_slaves",sel,"*","*")
		leaf = []
		slave = []
		for elem in child:
			leaf.append(elem[0])
		slave=findSlave(leaf)

		slave = list(set(slave))
		slave.sort()
		for port in slave:
			card = port.ISPA_PORT_OWNER_FK
			print "Port: ", card.ISPA_NAME,port.ISPA_PORT_NAME

		print "Total Number of Ports: ", len(slave)
