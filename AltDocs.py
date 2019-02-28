import cx_Oracle
import Tkinter,tkFileDialog
from Tkinter import *
import ttk
from ttkthemes import ThemedStyle
import Tkinter as tk
import re;

#AltDocs is a program that connects to the a database, pulls all relevant information about our apex
#applications, and displays them for the user. 
#Once the user has chosen a specific component, they have the option of creating documentation for this item. 
#This documentation is stored within the same database, currently inside the table apex_docs.

#TODO: 
#Create Page comment, in both application and documentation table.
#Clean code, a lot of it is redundant and can be vastly improved.

class Application(Frame):

	def createWidgets(self):  
		#Initiate connection with the database
		self.my_dsn = cx_Oracle.makedsn("*****",1521,service_name="*****")
		self.connection = cx_Oracle.connect(user="*****", password="*****", dsn=self.my_dsn)

		#Done before rendering, makes sure documentation table is cleaned up
		self.cleanTable()

		#Renders all of the necessary application objects
		self.renderApplicationMenu()
		self.renderPageMenu()
		self.renderItemMenu()
		self.renderTextFields()
		self.renderComponentAtts()

		#Quits the app completely
		self.QUIT = Button(self)
		self.QUIT["text"] = "QUIT"
		self.QUIT["command"] =  self.quit
		self.QUIT.grid(row=11, column=3)
		self.QUIT.config(width="15")
		self.allObjects.append(self.QUIT)

		#When pressed, calls updateDocs() 
		self.SUBMIT = Button(self)
		self.SUBMIT["text"] = "SUBMIT"
		self.SUBMIT["command"] =  self.updateDocs #how to call function with button
		self.SUBMIT.grid(row=11, column=2)
		self.SUBMIT.config(width="15")
		self.allObjects.append(self.SUBMIT)


###################################################################################################################################
#Parsing and Transactions
###################################################################################################################################

	#When submit is pressed,
	#updates the apex_docs table with what is currently in the comments field
	def updateDocs (self):
			
		#If null item is selected, we cannot update the table. Exits function immediately.
		if (self.pageItemList.get() == 'Select Item' and self.pageItemList.get() == 'Select Page'):
			print 'No Item or Page Selected, Cannot Submit.'
			return;

		if (self.pageItemList.get() == 'Select Item'):
			self.updatePageDocs()
			return;

		self.updatePageDocs()

		updateCursor = self.connection.cursor()

		#Getting all neccessary values for the update.
		app_id = self.parseID(self.applicationList.get())[0]
		app_name = self.parseID(self.applicationList.get())[1]
		page_id = self.parseID(self.applicationPageList.get())[0]
		page_name = self.parseID(self.applicationPageList.get())[1]
		component_type = self.parseID(self.pageItemList.get())[0]
		component_name = self.parseID(self.pageItemList.get())[1]

		subtype = self.parseAttribute(self.subtypeLabel["text"])
		region = self.parseAttribute(self.regionLabel["text"])
		label = self.parseAttribute(self.labelLabel["text"])

		comments = self.commentsEntry.get('1.0', 'end-1c')
		description = self.descEntry.get('1.0', 'end-1c')

		#Here we update the entry; if no entry exists, we insert a new entry.
		updateQuery =  ('UPDATE apex_docs SET '  
						' component_type = \'' 		+ subtype + '\','
						' component_region = \'' 	+ region + '\','
						' component_label = '		+ label + ','
						' comments = \'' 			+ comments + '\','
						' description = \'' 		+ description + '\'' 
						' WHERE application_id = ' 	+ app_id + 
						' AND page_id = ' 			+ page_id +
						' AND component_name = \'' 	+ component_name + '\'')

		insertQuery = ('INSERT INTO apex_docs ('
						'application_id, '
						'application_name, '
						'page_id, '
						'page_name, '
						'component_name, '
						'component_type, '
						'component_subtype, '
						'component_region, '
						'component_label, '
						'comments, '
						'changes, '
						'description'
						') VALUES '
						'(' 		+ app_id + 
						', \'' 		+ app_name + 
						'\',' 		+ page_id + 
						', \'' 		+ page_name + 
						'\', \'' 	+ component_name + 
						'\', \'' 	+ component_type + 
						'\', \'' 	+ subtype + 
						'\', \'' 	+ region + 
						'\', ' 		+ label + 
						', \'' 		+ comments + 
						'\', \''
						'\', \'' 	+ description + '\')')

		updateCursor.execute(updateQuery)

		if (updateCursor.rowcount == 0):
			print ('New row inserted into componenets table!')
			updateCursor.execute(insertQuery)

		self.connection.commit()
		print ('Component documentation updated!')

		print ('Documentation successfully updated!')
		updateCursor.close()
		#END updateDocs()

	def updatePageDocs(self):

		updateCursor = self.connection.cursor()

		app_id = self.parseID(self.applicationList.get())[0]
		app_name = self.parseID(self.applicationList.get())[1]
		page_id = self.parseID(self.applicationPageList.get())[0]
		page_name = self.parseID(self.applicationPageList.get())[1]

		pageComments = self.pageCommentsEntry.get('1.0', 'end-1c')
		pageChanges = self.pageChangesEntry.get('1.0', 'end-1c')
		pageDescription = self.pageDescEntry.get('1.0', 'end-1c')

		updatePageQuery = ( 'UPDATE apex_docs_pages SET '
							' comments = \'' 			+ pageComments + '\',' 
							' changes = \'' 			+ pageChanges + '\','
							' description = \'' 		+ pageDescription + '\'' 
							' WHERE application_id = ' 	+ app_id + 
							' AND page_id = ' 			+ page_id)

		insertPageQuery = ( 'INSERT INTO apex_docs_pages ('
							'application_id, '
							'page_id, '
							'page_name, '
							'comments, '
							'changes, '
							'description'
							') VALUES '
							'(' 		+ app_id +
							', ' 		+ page_id + 
							', \'' 		+ page_name +
							'\', \'' 	+ pageComments + 
							'\', \'' 	+ pageChanges + 
							'\', \'' 	+ pageDescription + '\')')

		print(updatePageQuery)
		updateCursor.execute(updatePageQuery)

		if (updateCursor.rowcount == 0):
			print ('New row inserted into pages table!')
			updateCursor.execute(insertPageQuery)

		self.connection.commit()
		print ('Page Documentation updated!')
		updateCursor.close()
		#END updatePageDocs()

	#Deletes all entries in the docs table that no longer exist in the apex database.
	def cleanTable (self):
		deleteCursor = self.connection.cursor()

		#TODO: Combine into single query, is definitely do-able.
		#Queries for each necessary apex table
		deleteItemQuery =  ('SELECT component_name ' 
						'FROM apex_docs docs ' 
						'WHERE NOT EXISTS   (SELECT item_name ' 
              								'FROM apex_application_page_items apex ' 
              								'WHERE apex.item_name = docs.component_name ' 
                							'AND apex.application_id = docs.application_id ' 
                							'AND apex.page_id = docs.page_id) ' 
                							'AND docs.component_type = \'ITEM\'')
		####################################################################################
		deleteButtonQuery =  ('SELECT component_name ' 
						'FROM apex_docs docs ' 
						'WHERE NOT EXISTS   (SELECT button_name ' 
              								'FROM apex_application_page_buttons apex ' 
              								'WHERE apex.button_name = docs.component_name ' 
                							'AND apex.application_id = docs.application_id ' 
                							'AND apex.page_id = docs.page_id) ' 
                							'AND docs.component_type = \'BUTTON\'')
	    ####################################################################################
		deleteRegionQuery =  ('SELECT component_name ' 
						'FROM apex_docs docs ' 
						'WHERE NOT EXISTS   (SELECT region_name ' 
              								'FROM apex_application_page_regions apex ' 
              								'WHERE apex.region_name = docs.component_name ' 
                							'AND apex.application_id = docs.application_id ' 
                							'AND apex.page_id = docs.page_id) ' 
                							'AND docs.component_type = \'REGION\'')
	    ####################################################################################
		deleteBranchQuery =  ('SELECT component_name ' 
						'FROM apex_docs docs ' 
						'WHERE NOT EXISTS   (SELECT branch_name ' 
              								'FROM apex_application_page_branches apex ' 
              								'WHERE apex.branch_name = docs.component_name ' 
                							'AND apex.application_id = docs.application_id ' 
                							'AND apex.page_id = docs.page_id) ' 
                							'AND docs.component_type = \'BRANCH\'')
	    ####################################################################################
		deleteActionQuery =  ('SELECT component_name ' 
						'FROM apex_docs docs ' 
						'WHERE NOT EXISTS   (SELECT dynamic_action_name ' 
              								'FROM apex_application_page_da apex ' 
              								'WHERE apex.dynamic_action_name = docs.component_name ' 
                							'AND apex.application_id = docs.application_id ' 
                							'AND apex.page_id = docs.page_id) ' 
                							'AND docs.component_type = \'DYNAMIC ACTION\'')
		####################################################################################
		deleteProcessQuery =  ('SELECT component_name ' 
						'FROM apex_docs docs ' 
						'WHERE NOT EXISTS   (SELECT process_name ' 
              								'FROM apex_application_page_proc apex ' 
              								'WHERE apex.process_name = docs.component_name ' 
                							'AND apex.application_id = docs.application_id ' 
                							'AND apex.page_id = docs.page_id) ' 
                							'AND docs.component_type = \'PROCESS\'')

		#TODO: Throw following 6 blocks into single loop, by having the above queries be placed in an interable list.
		#Next we run through each of these queries and throw the results into the deletion list.
		deletionList = []
		#######################################
		print(deleteItemQuery)
		deleteCursor.execute(deleteItemQuery)
		for result in deleteCursor:
			deletionList.append(result[0])
		#######################################
		print(deleteButtonQuery)
		deleteCursor.execute(deleteButtonQuery)
		for result in deleteCursor:
			deletionList.append(result[0])
		#######################################
		print(deleteRegionQuery)
		deleteCursor.execute(deleteRegionQuery)
		for result in deleteCursor:
			deletionList.append(result[0])
		#######################################
		print(deleteBranchQuery)
		deleteCursor.execute(deleteBranchQuery)
		for result in deleteCursor:
			deletionList.append(result[0])
		#######################################
		print(deleteActionQuery)
		deleteCursor.execute(deleteActionQuery)
		for result in deleteCursor:
			deletionList.append(result[0])
		#######################################
		print(deleteProcessQuery)
		deleteCursor.execute(deleteProcessQuery)
		for result in deleteCursor:
			deletionList.append(result[0])
		#######################################

		print('The following components and their documentation are about to be deleted: ')
		for component_name in deletionList:
			print(component_name + '\n')

		#Finally, we take the results from those queries and delete all of the matching entries in apex_docs
		#This is tricky, but basically we can append component_name as many times as we need and still have
		#all of the necessary entries deleted.
		
		#Pull component_names from deletionList, append it to the query. We only do this if deletionList is not empty.
		cleaningQuery = ''
		if deletionList:
			cleaningQuery = 'DELETE FROM apex_docs WHERE component_name IN (\'' + deletionList[0] + '\''
			deletionList.pop(0)
			for component_name in deletionList:
				cleaningQuery = cleaningQuery + ', \'' + component_name + '\''
			cleaningQuery = cleaningQuery + ')'

		print(cleaningQuery)
		
		if (cleaningQuery != ''):
			deleteCursor.execute(cleaningQuery)
		self.connection.commit();

		print('Deleted components successfully cleaned!')
		#END cleanTable()

	#Uses regex substitution to remove all of the commas, parantheses, brackets, and quotes from the result set.
	def cleanResultTextBody (self, result):

		return re.sub(r'(\'|\,|\(|\)|\{|\})', '', result);


	def parseID (self, selection):

		return selection.split(' - ');

	def parseAttribute (self, selection):

		return selection.split(': ')[1]

###################################################################################################################################
#Queries
###################################################################################################################################

	#Grabs all of the APEX applications from the DB
	def queryApexApplications (self):
		applicationsCursor = self.connection.cursor()

		applicationsQuery = 'SELECT application_id, application_name FROM apex_applications WHERE workspace_display_name = \'*****\' ORDER BY application_id'

		applicationsCursor.execute(applicationsQuery)

		returnList = [];
		for result in applicationsCursor:
			returnList.append(str(result[0]) + ' - ' + result[1])

		applicationsCursor.close()

		return returnList;


	#Grabs all of the APEX pages of associated application from the DB
	def queryApexPages (self):
		pagesCursor = self.connection.cursor()

		pagesQuery = 'SELECT page_id, page_name FROM apex_application_pages WHERE application_id = ' + self.parseID(self.applicationList.get())[0] 

		pagesCursor.execute(pagesQuery)

		returnList = [];
		for result in pagesCursor:
			returnList.append(str(result[0]) + ' - ' + result[1])

		pagesCursor.close()

		return returnList;

	#TODO: Find a way to improve runtime, pretty slow right now
	#Grabs all of the APEX items of associated page from the DB
	def queryApexItems (self):
		itemsCursor = self.connection.cursor()

		itemsQuery 		= 'SELECT item_name 			FROM apex_application_page_items 	WHERE application_id = ' + self.parseID(self.applicationList.get())[0]  + ' AND page_id = ' + self.parseID(self.applicationPageList.get())[0] 
		buttonsQuery 	= 'SELECT button_name 			FROM apex_application_page_buttons 	WHERE application_id = ' + self.parseID(self.applicationList.get())[0]  + ' AND page_id = ' + self.parseID(self.applicationPageList.get())[0] 
		regionsQuery 	= 'SELECT region_name 			FROM apex_application_page_regions 	WHERE application_id = ' + self.parseID(self.applicationList.get())[0]  + ' AND page_id = ' + self.parseID(self.applicationPageList.get())[0] 
		branchesQuery 	= 'SELECT branch_name 			FROM apex_application_page_branches WHERE application_id = ' + self.parseID(self.applicationList.get())[0]  + ' AND page_id = ' + self.parseID(self.applicationPageList.get())[0] 
		daQuery 		= 'SELECT dynamic_action_name 	FROM apex_application_page_da 		WHERE application_id = ' + self.parseID(self.applicationList.get())[0]  + ' AND page_id = ' + self.parseID(self.applicationPageList.get())[0] 
		processesQuery 	= 'SELECT process_name 			FROM apex_application_page_proc 	WHERE application_id = ' + self.parseID(self.applicationList.get())[0]  + ' AND page_id = ' + self.parseID(self.applicationPageList.get())[0] 

		returnList = [];
		#############################################
		itemsCursor.execute(itemsQuery)

		for result in itemsCursor:
			returnList.append('ITEM - ' + result[0])
		#############################################
		itemsCursor.execute(buttonsQuery)

		for result in itemsCursor:
			returnList.append('BUTTON - ' + result[0])
		#############################################
		itemsCursor.execute(regionsQuery)

		for result in itemsCursor:
			returnList.append('REGION - ' + result[0])
		#############################################
		itemsCursor.execute(branchesQuery)

		for result in itemsCursor:
			returnList.append('BRANCH - ' + result[0])
		#############################################
		itemsCursor.execute(daQuery)

		for result in itemsCursor:
			returnList.append('DYNAMIC ACTION - ' + result[0])
		#############################################
		itemsCursor.execute(processesQuery)

		for result in itemsCursor:
			returnList.append('PROCESS - ' + result[0])
		#############################################

		itemsCursor.close()

		return returnList;

	#Queries the pages table for documentation columns
	def queryPageDocs (self):
		docsCursor = self.connection.cursor()

		pageID = self.parseID(self.applicationPageList.get())[0]
		applicationID = self.parseID(self.applicationList.get())[0]

		docsQuery = (' SELECT comments, changes, description' 
					 ' FROM apex_docs_pages' 
					 ' WHERE application_id = ' + applicationID + 
					 ' AND page_id = ' + pageID)

		docsCursor.execute(docsQuery)

		returnList = []
		for result in docsCursor:
			for data in result:
				if data == None:
					returnList.append('')
				else:
					returnList.append(data)

		if (len(returnList) == 0):
				returnList.append('')
				returnList.append('')
				returnList.append('')



		return returnList;


	#Queries the given items subtype, region, and source as well as associated documentation
	#TODO: Add functionality for other types of items; easiest way to do this would be to query from single table apex_docs,
	#and update that table with info from the DB at program load.
	def queryItemData (self):
		dataCursor = self.connection.cursor()

		componentType = self.parseID(self.pageItemList.get())[0]
		
		#Queries for component attributes
		if componentType == 'ITEM':
			dataQuery = ('SELECT display_as, region, label' 
					 ' FROM apex_application_page_items' 
					 ' WHERE application_id = ' + self.parseID(self.applicationList.get())[0] + 
					 ' AND page_id = ' + self.parseID(self.applicationPageList.get())[0] + 
					 ' AND item_name = \'' +  self.parseID(self.pageItemList.get())[1] + '\'')
		elif componentType == 'BUTTON':
			dataQuery = ('SELECT button_template, region, label' 
					 ' FROM apex_application_page_buttons' 
					 ' WHERE application_id = ' + self.parseID(self.applicationList.get())[0] + 
					 ' AND page_id = ' + self.parseID(self.applicationPageList.get())[0] + 
					 ' AND button_name = \'' +  self.parseID(self.pageItemList.get())[1] + '\'')
		elif componentType == 'REGION':
			dataQuery = ('SELECT template, parent_region_name, region_header_text' 
					 ' FROM apex_application_page_regions' 
					 ' WHERE application_id = ' + self.parseID(self.applicationList.get())[0] + 
					 ' AND page_id = ' + self.parseID(self.applicationPageList.get())[0] + 
					 ' AND region_name = \'' +  self.parseID(self.pageItemList.get())[1] + '\'')
		elif componentType == 'BRANCH':
			dataQuery = ('SELECT branch_type, when_button_pressed, condition_type' 
					 ' FROM apex_application_page_branches' 
					 ' WHERE application_id = ' + self.parseID(self.applicationList.get())[0] + 
					 ' AND page_id = ' + self.parseID(self.applicationPageList.get())[0] + 
					 ' AND branch_name = \'' +  self.parseID(self.pageItemList.get())[1] + '\'')
		elif componentType == 'DYNAMIC ACTION':
			dataQuery = ('SELECT \'NA\', \'NA\', \'NA\'' 
					 ' FROM apex_application_page_da' 
					 ' WHERE application_id = ' + self.parseID(self.applicationList.get())[0] + 
					 ' AND page_id = ' + self.parseID(self.applicationPageList.get())[0] + 
					 ' AND dynamic_action_name = \'' +  self.parseID(self.pageItemList.get())[1] + '\'')
		elif componentType == 'PROCESS':
			dataQuery = ('SELECT process_point, region_name, process_type' 
					 ' FROM apex_application_page_proc' 
					 ' WHERE application_id = ' + self.parseID(self.applicationList.get())[0] + 
					 ' AND page_id = ' + self.parseID(self.applicationPageList.get())[0] + 
					 ' AND process_name = \'' +  self.parseID(self.pageItemList.get())[1] + '\'')

		#Query for component documentation
		docsQuery = ('SELECT comments, description' 
					 ' FROM apex_docs' 
					 ' WHERE application_id = ' + self.parseID(self.applicationList.get())[0] + 
					 ' AND page_id = ' + self.parseID(self.applicationPageList.get())[0] + 
					 ' AND component_name = \'' +  self.parseID(self.pageItemList.get())[1] + '\'')

		dataCursor.execute(dataQuery)

		#We have to account for potentially null values here, which is why the nested conditional replaces all None values with empty string
		returnList = []
		for result in dataCursor:
			for data in result:
				if data == None:
					returnList.append('')
				else:
					returnList.append(data)

		dataCursor.execute(docsQuery)
		print(docsQuery);
		for result in dataCursor:
			for data in result:
				if data == None:
					returnList.append('')
				else:
					returnList.append(data)

		dataCursor.close()

		print(returnList);

		return returnList;

###################################################################################################################################
#Rendering
###################################################################################################################################

	#Creates objects for applications dropdown box
	def renderApplicationMenu (self):
		self.allObjects = []
		self.allTextAreas = []

		#Application entry field label
		self.applicationLabel = Label(self, text = "Application: ", background='#424242', fg='white')
		self.applicationLabel.grid(columnspan=1, row=0, column=2)
		self.allObjects.append(self.applicationLabel)

		#Function call to get list of applications from DB
		apexApplications = self.queryApexApplications();
		print('Applications successfully populated!')

		#StringVar for dropdown menu
		self.applicationList = StringVar(self)
		self.applicationList.set('Select Application') #Want the first element of the dropdown to be blank
		self.applicationList.trace('w', self.populatePageMenu)

		self.applicationMenu = OptionMenu(self, self.applicationList, *apexApplications)
		self.applicationMenu.grid(columnspan=1, row=0, column=3)
		self.allObjects.append(self.applicationMenu)


	#Creates objects for pages dropdown box
	def renderPageMenu (self):

		#APEX index entry field label
		self.pageLabel = ttk.Label(self, text = "Page: ")
		self.pageLabel.grid(columnspan=1, row=1, column=2)
		self.allObjects.append(self.pageLabel)
		
		applicationPages = ['']

		#StringVar for dropdown menu
		self.applicationPageList = StringVar(self)
		self.applicationPageList.set('Select Page') #Want the first element of the dropdown to be blank
		self.pageTraceID = self.applicationPageList.trace('w', self.populateItemMenu)

		self.applicationPageMenu = OptionMenu(self, self.applicationPageList, *applicationPages)
		self.applicationPageMenu.grid(columnspan=1, row=1, column=3)
		self.allObjects.append(self.applicationPageMenu)


	#Creates objects for items dropdown box
	def renderItemMenu (self):

		#APEX index entry field label
		self.itemLabel = ttk.Label(self, text = "Component: ")
		self.itemLabel.grid(columnspan=1, row=1, column=5)
		self.allObjects.append(self.itemLabel)

		#Function call to get list of apex indices from DB
		pageItems = ['']
		
		#StringVar for dropdown menu
		self.pageItemList = StringVar(self)
		self.pageItemList.set('Select Component') #Want the first element of the dropdown to be blank
		self.itemTraceID = self.pageItemList.trace('w', self.populateItemFields)

		self.pageItemMenu = OptionMenu(self, self.pageItemList, *pageItems)
		self.pageItemMenu.grid(columnspan=1, row=1, column=6)
		self.allObjects.append(self.pageItemMenu)

	#Creates objects for the comments, changes, and description text fields
	def renderTextFields (self):
		#####################################################
		#Comments entry field label
		self.pageCommentsLabel = Label(self, text="Page Comments: ")
		self.pageCommentsLabel.grid(columnspan=1, row=2, column=1)
		self.allObjects.append(self.pageCommentsLabel)

		#Entry field for the comments
		self.pageCommentsEntry = Text(self, width = 40, height = 10)
		self.pageCommentsEntry.grid(columnspan=2, row=2, column=2)
		self.allTextAreas.append(self.pageCommentsEntry)

		#####################################################
		#Changes entry field label
		self.pageChangesLabel = Label(self, text="Page Changes: ")
		self.pageChangesLabel.grid(columnspan=1, row=3, column=1)
		self.allObjects.append(self.pageChangesLabel)

		#Entry field for the changes
		self.pageChangesEntry = Text(self, width = 40, height = 10)
		self.pageChangesEntry.grid(columnspan=2, row=3, column=2)
		self.allTextAreas.append(self.pageChangesEntry)

		#####################################################
		#Description entry field label
		self.pageDescLabel = Label(self, text="Page Description: ")
		self.pageDescLabel.grid(columnspan=1, row=4, column=1)
		self.allObjects.append(self.pageDescLabel)

		#Entry field for the Description
		self.pageDescEntry = Text(self, width = 40, height = 10)
		self.pageDescEntry.grid(columnspan=2, row=4, column=2)
		self.allTextAreas.append(self.pageDescEntry)

		#####################################################


		#####################################################
		#Comments entry field label
		self.commentsLabel = Label(self, text="Component Comments: ")
		self.commentsLabel.grid(columnspan=1, row=2, column=4)
		self.allObjects.append(self.commentsLabel)

		#Entry field for the comments
		self.commentsEntry = Text(self, width = 40, height = 10)
		self.commentsEntry.grid(columnspan=2, row=2, column=5)
		self.allTextAreas.append(self.commentsEntry)

		#####################################################
		#Description entry field label
		self.DescLabel = Label(self, text="Component Description: ")
		self.DescLabel.grid(columnspan=1, row=3, column=4)
		self.allObjects.append(self.DescLabel)

		#Entry field for the Description
		self.descEntry = Text(self, width = 40, height = 10)
		self.descEntry.grid(columnspan=2, row=3, column=5)
		self.allTextAreas.append(self.descEntry)

		#####################################################

	#Creates objects for the component attributes
	def renderComponentAtts (self):

		#Label for subtype attribute
		self.subtypeLabel = Label(self, text="Component Subtype: ")
		self.subtypeLabel.grid(columnspan=1, row=5, column=5)
		self.allObjects.append(self.subtypeLabel)

		#Label for subtype region
		self.regionLabel = Label(self, text="Component Region: ")
		self.regionLabel.grid(columnspan=1, row=6, column=5)
		self.allObjects.append(self.regionLabel)

		#Label for subtype label
		self.labelLabel = Label(self, text="Component Label: ")
		self.labelLabel.grid(columnspan=1, row=7, column=5)
		self.allObjects.append(self.labelLabel)

###################################################################################################################################
#Populating
###################################################################################################################################

	#Called whenever application OptionMenu is changed, resets the pages OptionMenu
	def populatePageMenu (self, *args):

		#Temporarily delete traces so that we can reset value to default without triggering function
		self.applicationPageList.trace_vdelete("w", self.pageTraceID)
		self.pageItemList.trace_vdelete("w", self.itemTraceID)


		applicationPages = self.queryApexPages();

		#Empty menu and reset selected value
		self.applicationPageMenu["menu"].delete(0, "end")
		self.applicationPageList.set('Select Page')
		self.pageTraceID = self.applicationPageList.trace('w', self.populateItemMenu)
		self.pageItemList.set('Select Component')
		self.itemTraceID = self.pageItemList.trace('w', self.populateItemFields)

		#Reset if application changed
		self.pageCommentsEntry.delete('1.0', END)
		self.pageChangesEntry.delete('1.0', END)
		self.pageDescEntry.delete('1.0', END)
		self.commentsEntry.delete('1.0', END)
		self.descEntry.delete('1.0', END)

		#Refill menu with updated list
		for page in applicationPages:
			self.applicationPageMenu["menu"].add_command(label = page, command = tk._setit(self.applicationPageList, page))

		print('Pages successfully populated!')

	#Called whenever pages OptionMenu is changed, resets the items OptionMenu, also populates the page documentation fields
	def populateItemMenu (self, *args):

		#Temporarily delete trace so that we can reset value to default without triggering function
		self.pageItemList.trace_vdelete("w", self.itemTraceID)

		pageDocs = self.queryPageDocs()
		pageItems = self.queryApexItems()

		#Empty menu and reset selected value
		self.pageItemMenu["menu"].delete(0, "end")
		self.pageItemList.set('Select Item')
		self.itemTraceID = self.pageItemList.trace('w', self.populateItemFields)

		#Deleting current data in page documentation fields, replacing with new data
		self.pageCommentsEntry.delete('1.0', END)
		self.pageChangesEntry.delete('1.0', END)
		self.pageDescEntry.delete('1.0', END)
		self.pageCommentsEntry.insert(END, pageDocs[0])
		self.pageChangesEntry.insert(END, pageDocs[1])
		self.pageDescEntry.insert(END, pageDocs[2])

		self.commentsEntry.delete('1.0', END)
		self.descEntry.delete('1.0', END)

		#Refill menu with updated list
		for item in pageItems:
			self.pageItemMenu["menu"].add_command(label = item, command = tk._setit(self.pageItemList, item))

		print('Items successfully populated!')


	#Called whenever items OptionMenu is changed, resets the items fields
	def populateItemFields (self, *args):

		itemData = self.queryItemData()

		self.subtypeLabel['text'] = 'Component Subtype: ' + itemData[0]
		self.regionLabel['text'] = 'Component Region: ' + itemData[1]
		self.labelLabel['text'] = 'Component Label: \'' + itemData[2] + '\''

		#Replace current text inside the text area with data
		self.commentsEntry.delete('1.0', END)
		self.descEntry.delete('1.0', END)
		self.commentsEntry.insert(END, itemData[3])
		self.descEntry.insert(END, itemData[4])

		print('Item data successfully populated!')


	def __init__(self, master=None):
		Frame.__init__(self, master, background='#424242')
		#self.style = ThemedStyle(self)
		#self.style.set_theme('black')
		color = '#424242'
		self.pack()
		self.createWidgets()
		for wid in self.allObjects:
			wid.configure(background = color)
			wid.configure(foreground = 'white')
		for wid in self.allTextAreas:
			wid.configure(background = '#232323')
			wid.configure(foreground = '#42ff58')


root = Tk()
app = Application(master=root)
app.mainloop()
#path = "C:\Users\\t0rmsanz\Documents\Work Docs\FirstConversion\wirtrnrg.txt";
root.destroy()