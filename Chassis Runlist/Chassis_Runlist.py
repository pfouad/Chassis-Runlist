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

def main():
	sel = gdm.selected_entity()
