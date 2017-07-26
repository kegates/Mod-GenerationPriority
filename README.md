# Mod-GenerationPriority
Excel Add-in: Modifies the current ISO New England Generation Dispatch Priority List (Excel Document) based on inputted TTCs (Total Transfer Capabilities of NE Interfaces) as well as the current generation outages, taken from CROW exported excel document.  This project is done for ISO New England.


=============================================================================================
@filename: README.txt
@author: Kevin Gates
@version: 1.0

@date: 07/17/2017 15:14:00

@description: Contains information on the current version of CROW compare.
	      Will be updated as program is updated. (Description, Instructions, and Instillation)
=============================================================================================

DESCRIPTION: This addin contains a Macro intended to be used in modifying the Generator Priority
	     list.  The idea is that CROW generator outages can be imported into the Priority List
 	     Workbook and that the eco. maxes of units under outage can be updated (using a date under
	     under study inputted by the user).  In addition, TTCs can be inputted for the given
	     date under study.  The NE demand tab will then use these values to update itself on
	     how much demand it can sustain based on the Transfers coming in/out and the new
	     eco maxes of the generators.  

---------------------------------------------------------------------------------------------

INSTRUCTIONS: 
		
		1. If you wish to import CROW outages, open CROW and go to Generator Outages.
		   Make sure your Definition is on '<default>' (Can view by clicking 'View Definition')

		2. Unders 'Dates:' Change the first Date to the date you wish to study and
		   the second box to the date under study if selecting dates or '24 hours' if
		   selecting hours.  It should look like the example bellow:

			ex. Dates: mm/dd/yyyy to: mm/dd/yyyy
				
					--or--
			    Dates: mm/dd/yyyy Next 24 hours  

		   (Note: you can always expand these dates but make sure they are at least
			  this long and revolve around the date under study)

		3. Click on an outage, then hit Cntrl-A and click the 'Export to Excel' Button.
		   It should look like a small excel Icon.  Save this file.

		4. Open the Generator Priority List and go to Add-Ins -> Modify Generator Priority List
		   -> Modify Priority List on your top tab. If you decide to use CROW outage, the file
		   explorer will open and you will need to find and select the CROW excel file you saved.

		5. If you wish to use TTCs note that they represent the planned transfer from outside
		   locations TO New England (e.g Positive values increase our sustainable load)



---------------------------------------------------------------------------------------------

INSTILLATION:
		
		1. Open the zipped folder and move the folder inside your Microsoft AddIns
		   folder

			Example: C:\Users\<username>\AppData\Roaming\Microsoft\AddIns

		2. Open Excel and go to File -> Options -> Add-Ins -> Go... (Next to Manage Dropdown)
		   NOTE: Make sure 'Manage:' is set to 'Excel Add-Ins' before clcking 'Go...'

		3. Click Browse and find the folder you moved in step 1.  Within that folder, select the
		   file ending in '.xlam', it should be called 'Mod_GenPrior'.  Make sure 'Mod_GenPrior' Appears
		   checked in the 'Add-Ins available' list.

		4. Open MODInstallToMenu.xlsm from the folder and click 'Install Compare Menu', it should
		   appear under or above the CROW Add-In in the Add-Ins Tab.  If not, contact Kevin Gates
		   at short term outage (Supervisor: Norman Sproehnle)
