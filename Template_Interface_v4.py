# Imports --------------------------------------------------------------------------------------------------------------------------------------------------
import os,sys
from PyQt4 import QtCore, QtGui
from PyQt4.QtGui import QFileDialog
import xlwings as xw

# Formatting -----------------------------------------------------------------------------------------------------------------------------------------------
def bold(txt):
	return "<b>" + txt + "</b>"
def br(txt, new=""): # Line break
	return txt + "<br>" + new
def label(txt):
	return bold(txt + ": ")
def link(link, txt="(link)"):
	return '<a href="http://www.' + link + '">' + txt + "</a>"	
def check_list(txt): # Check if given value is a list, returns list
	if type(txt) != list:
		txt = [txt]
	return txt	
def res_ID(id): # Respondent ID
	id = str(id).replace("-", "").replace(" ", "").split(".")
	if type(id) == list:
		id = id[0]
	while len(id) < 10:
		id = "0" + id
	return id

# Button ----------------------------------------------------------------------------------------------------------------------------------------------------
class Button(QtGui.QPushButton):
	def __init__(self, txt="", fn=None, args=None, parent=None):
		super(Button, self).__init__(txt, parent)
		self.true = self.text
		if fn != None:
			self.fn = fn
		else:
			self.fn = print
		if args != None:
			self.args = args
		else:
			self.args = [self.true]
		self.clicked.connect(self.call)
	def call(self): # Default function on click
		self.fn(*self.args)

# Dropdown -------------------------------------------------------------------------------------------------------------------------------------------------
class Dropdown(QtGui.QComboBox):
	def __init__(self, ops=[], fn=print, parent=None):
		super(Dropdown, self).__init__(parent)
		self.addItems(ops)
		self.currentIndexChanged.connect(fn)
	def value(self):
		return str(self.currentText())
	def update(self, ops=[]):
		self.clear()
		self.addItems(ops)
		
# Widget --------------------------------------------------------------------------------------------------------------------------------------------------
class Widget(QtGui.QWidget):
	def __init__(self, parent=None):
		super().__init__(parent)
		self.layout = QtGui.QGridLayout()
		self.row = 0
		self.add_wide()
	def add_widget(self, widget, column_width=1, column=0, row_width=1, row=None, increment_row=True):
		if widget != None:
			if row == None:
				row = self.row
			if type(widget) == str:
				widget = QtGui.QLabel(widget)
			self.layout.addWidget(widget, self.row, column, row_width, column_width)
			if increment_row:
				self.row += row_width
	def add_row(self, options):
		row_width = 1
		for i in options:
			if type(i) == list and len(i) > row_width:
				row_width = len(i)
		for i in range(0, len(options)):
			if type(options[i]) == list:
				for j in range(0, len(options[i])):
					self.add_widget(options[i][j], 1, i, 1, self.row + j, False)
			else:
				self.add_widget(options[i], 1, i, row_width, None, False)
		self.row += row_width
	def add_wide(self, label=""): # Add a single widget that takes up all 3 columns
		self.add_widget(label, 3)
	def add_labeled(self, options, txt="", extra=""): # Add labeled widget(s)
		if type(txt) == str and txt != "":
			txt = label(txt)
		options = check_list(options)
		for i in range(0, len(options)):
			if i == 0:
				self.add_row([txt, options[i], extra])
			else:
				self.add_widget(options[i], 1, 1)
	def finalize(self):
		self.add_wide()
		self.setLayout(self.layout)
		
class Widget_Home(Widget):
	def __init__(self, parent=None):
		super().__init__(parent)
		self.filepath = parent.filepath
		# Intro
		txt = br(bold("Welcome to the NCRC HMDA Template Interface! ") + "(version 4.0)", "This interface exists to help simplify the process of filling out NCRC's HMDA template.")
		txt = br(txt, "It was designed for use with template version " + parent.template_ver + ", and will not work if there isn't a file named '" + parent.template_filename + "' found within the same folder as this program.")
		txt = br(txt, "In order to successfully fill out the template, please follow all steps below.")
		self.add_wide(txt)
		self.add_wide()
		# HMDA data
		txt = "Before proceeding, please make sure that you have downloaded the HMDA data you wish to use from the " + link("consumerfinance.gov/data-research/hmda/explore", "CFPB website")
		txt += ", and that this file is contained within the same folder as this program."
		txt = br(txt, br(label("HMDA data download instructions"), bold("1. ") + "Select the year(s) of data you wish to process, then select a suggested filter."))
		txt = br(txt + "Keep in mind that 'All Records' may return a very large amount of data.", bold("2. ") + "Select the geography you wish to examine under Location. ")
		txt = br(txt + "If you are interested in comparing multiple geographies, include all of them. Verify the kinds of property you wish to look at.", bold("3. ") + "Select which kinds of loans you are interested in.") 
		txt = br(txt, bold("4. ") + "Do not filter for lender; you will select a lender later.")
		txt = br(txt, bold("5. ") + "Under 'Preview the results', verify the number of loans is a manageable volume given your expertise, computer, and needs. ")
		txt = br(txt + "If it is too big, go back and further limit the data you are selecting.", bold("6. ") + "Click the radio button to include labels and codes, and then select 'download'.")
		txt = br(txt, bold("7. ") + "Wait for the entire file to download, and then move the file to the same folder as this program. If the file's name is different from the one below, please enter the correct filename.")
		self.add_wide(txt)
		self.add_wide()
		self.data_filename = QtGui.QLineEdit("hmda_lar.csv")
		self.add_labeled(self.data_filename, "HMDA data filename", Button("Browse...", self.browse, [], self))
		self.add_wide("When using the 'Browse...' option, be sure to add the file's type to the end (for example: '.csv'). " + bold("Do not ") + "select files outside of the current folder.")
		self.add_wide()
		# Target data
		self.add_wide("By default, this program will go through the data provided and create the Target data for you. However, you can also choose to provide seperate file(s) that are pre-filtered:")
		self.target = QtGui.QCheckBox()
		self.add_labeled(self.target, "Get Target Data From a Seperate File")
		self.add_wide()
		# MSA
		txt = "In addition to the above files, you also need to make sure there is a file named '" + parent.MSA_filename + "' included in the same folder as this program."
		self.add_wide(txt)
		# Continue
		self.add_wide("If you have read and followed all the instructions above, please select the continue button below:")
		self.add_wide()
	def get_target(self):
		return self.target.isChecked()
	def browse(self):
		new = QFileDialog.getOpenFileName()
		new = new.split("/")
		new = new[len(new) - 1]
		self.data_filename.setText(new)
	
class Widget_Basic(Widget):
	def __init__(self, parent=None):
		super().__init__(parent)
		# Intro
		self.add_wide("Please enter some basic information about your data.")
		# Year(s)
		self.years = QtGui.QLineEdit("2014")
		self.add_labeled(self.years, "Year(s)", br("Seperate multiple years by commas;", "for example: 2014, 2012, 2011"))
		# Geography Type
		self.geo_type = Dropdown(["MSA", "State", "County"]) # Add back in census tract later? implemented below, but too many results for the program to handle.
		self.add_labeled(self.geo_type, "Geography Type")
		self.add_wide(br("It may take a moment for geography data to load.", "Do not interfere with any Excel windows that may open during this time."))

class Widget_Geo(Widget):
	def __init__(self, parent=None):
		super().__init__(parent)
		# Intro
		self.add_wide("Please enter the name(s) of your selected geography.")
		# Variables (to be populated in update)
		self.filepath = parent.filepath
		self.filename = parent.MSA_filename
		self.data_MSAs = []
		self.data_other = []
		self.data_MSAs_names = []
		self.data_states = []
		self.data_counties = []
		# Geography Selection
		self.geo_name = Dropdown(self.data_MSAs_names, self.update_name, self)
		self.geo_name_label = QtGui.QLabel(label("None"))
		self.add_labeled(self.geo_name, self.geo_name_label)
		self.geo_name_select = []
		self.geo_name_select_label = QtGui.QLabel()
		self.add_labeled(self.geo_name_select_label, QtGui.QLabel(label("Selected")))
		self.add_labeled(Button("Reset Selected", self.reset_select, [], self))
		self.add_wide()
	def reset_select(self):
		self.geo_name_select = []
		self.update_select()
	def update_select(self): # Update the displayed selected values
		txt = ""
		for i in range(0, len(self.geo_name_select)):
			if i != 0:
				txt += "<br>"
			txt += self.geo_name_select[i]
		if txt == "":
			txt = "(None)"
		self.geo_name_select_label.setText(txt)
	def update_sub(self, data, txt): # Helper function for update
		self.geo_name.update(data)
		self.geo_name_label.setText(label(txt))
	def update(self, geo_type):
		if self.data_MSAs == [] or self.data_other == []:
			MSA = xw.Book(self.filepath + self.filename)
			xw.apps.active.screen_updating = False
			MSA_list = MSA.sheets['MSAList']
			self.data_MSAs = MSA_list.range('A1:E1').expand().value[1:]
			MSA_tract = MSA.sheets['TractDemData']
			self.data_other = MSA_tract.range("A1:S1").expand().value[1:]
			xw.apps.active.screen_updating = True
			xw.apps.active.quit()
		if geo_type not in str(self.geo_name_label.text()):
			if geo_type == "MSA":
				if self.data_MSAs_names == []:
					for i in self.data_MSAs:
						if i[0] not in self.data_MSAs_names:
							self.data_MSAs_names.append(i[0])
				self.update_sub(self.data_MSAs_names, "MSA(s)")
			else:
				if self.data_states == []:
					for i in self.data_other:
						if i[0] not in self.data_states:
							self.data_states.append(i[0])
						if i[5] not in self.data_counties:
							self.data_counties.append(i[5])
				if geo_type == "State":
					self.update_sub(self.data_states, "State(s)")
				else: # County
					self.update_sub(self.data_counties, "Count(y/ies)")
			self.geo_name_select = []
			self.update_select()
	def update_name(self): # Update geo_name_select; triggered whenever dropdown is changed
		name = self.geo_name.value()
		if name in self.geo_name_select:
			self.geo_name_select.remove(name)
		else:
			self.geo_name_select.append(name)
		self.update_select()

class Widget_Lender(Widget):
	def __init__(self, parent=None):
		super().__init__(parent)
		# Intro
		txt = br("Please enter the ID(s) of your selected lender(s) below.", "If you don't know a lender's ID, you can look it up using the ")
		txt += link("ffiec.gov/hmdaadwebreport/DisWelcome.aspx", "FFIEC's disclosure tool") + "."
		self.add_wide(txt)
		# Lender(s)
		self.lenders = QtGui.QLineEdit("")
		self.add_labeled(self.lenders, "Lender ID(s)", br("Seperate multiple IDs by commas;", "for example: 0000014025, 0000451965, 0000029141"))
		# Target Data
		self.target_intro = QtGui.QLabel(br("Enter the filename(s) of the target data below.", "Seperate multiple filenames by commas; for example: hmda_lar_1.csv, hmda_lar_2.csv, hmda_lar_3.csv"))
		self.add_wide(self.target_intro)
		self.target_label = QtGui.QLabel(label("Target Filename(s)"))
		self.target = QtGui.QLineEdit("")
		self.browse_button = Button("Browse...", self.browse, [], self)
		self.add_row([self.target_label, self.target, self.browse_button])
		self.browse_txt = QtGui.QLabel("When using the 'Browse...' option, be sure to add the file's type to the end (for example: '.csv'). " + bold("Do not ") + "select files outside of the current folder.")
		self.add_wide(self.browse_txt)
		# Finish Info
		txt = "When you select the Finish button below, the template will be filled out according to the information you have provided."
		txt = br(txt, "A loading screen will appear while the program does this; once the process is complete, the program will return the the home screen.")
		txt = br(txt, "A version of the template with your information will remain open; Don't forget to save, and make sure not to save it over the original!")
		self.add_wide(txt)
	def update(self, show):
		if show:
			self.target_intro.show()
			self.target_label.show()
			self.target.show()
			self.browse_button.show()
			self.browse_txt.show()
		else:
			self.target_intro.hide()
			self.target_label.hide()
			self.target.hide()
			self.browse_button.hide()
			self.browse_txt.hide()
	def get(self):
		if self.lenders.text() == "":
			return []
		result_initial = check_list(self.lenders.text().split(", "))
		result = []
		for i in result_initial:
			i = i.split(",")
			if type(i) == list:
				result += i
			else:
				result.append(i)
		for i in range(0, len(result)):
			result[i] = res_ID(result[i])
		return result
	def get_target(self):
		if self.target.text() == "":
			return []
		result_initial = check_list(self.target.text().split(", "))
		result = []
		for i in result_initial:
			i = i.split(",")
			if type(i) == list:
				result += i
			else:
				result.append(i)
		return result
	def browse(self):
		new = QFileDialog.getOpenFileName()
		new = new.split("/")
		new = new[len(new) - 1]
		if self.get_target() == []:
			self.target.setText(new)
		else:
			self.target.setText(self.target.text() + ", " + new)
		
class Widget_Bar(Widget):
	def __init__(self, txt="", min=0, max=100, parent=None):
		super().__init__(parent)
		# Info
		self.add_wide(br("Data is currently being processed.", "Do not interfere with any Excel windows that may open during this time."))
		self.label = QtGui.QLabel(txt)
		self.add_wide(self.label)
		# Progress Bar
		self.bar = QtGui.QProgressBar(self)
		self.add_wide(self.bar)
		self.min = min
		self.max = max
		self.bar.setRange(min, max)
		self.current = min
		self.bar.setValue(min)
		# Cancel
		self.add_wide(Button("Cancel", parent.widget_set, [parent.home], self))
	def add_step(self, increment=1):
		self.current += increment
		self.bar.setValue(self.current)
	def reset(self, txt="", min=None, max=None):
		if txt != "":
			self.label.setText(txt)
		if min != None:
			self.min = min
		if max != None:
			self.max = max
		self.current = min
		self.bar.setRange(self.min, self.max)
		self.bar.reset()
		
# Main Window ---------------------------------------------------------------------------------------------------------------------------------------------
class MainWindow(QtGui.QMainWindow):
	def __init__(self, parent=None):
		# Set up window
		super().__init__(parent)
		self.central_widget = QtGui.QStackedWidget()
		self.setCentralWidget(self.central_widget)
		self.showMaximized()
		# Variables
		self.filepath = str(os.path.dirname(os.path.realpath(__file__))) + "\\"
		self.template_ver = '12'
		self.template_filename = 'HMDA Template u.v.' + self.template_ver + '.xltm'
		self.MSA_filename = "DemDataMSA.xlsx"
		# Make Widgets
		self.home = Widget_Home(self)
		self.lender = Widget_Lender(self)
		self.basic = Widget_Basic(self)
		self.geo = Widget_Geo(self)
		self.end = Widget_Bar("Opening template...", 0, 5, self)
		# Finish Widgets
		self.home.add_wide(Button("Continue", self.widget_set, [self.basic, "Basic Information"], self))
		self.widget_finish(self.home)
		self.basic.add_labeled(Button("Continue", self.widget_set, [self.geo, "Geography"], self), Button("Back", self.widget_set, [self.home, "Home"], self))
		self.widget_finish(self.basic)
		self.geo.add_labeled(Button("Continue", self.widget_set, [self.lender, "Lender"], self), Button("Back", self.widget_set, [self.basic, "Basic Information"], self))
		self.widget_finish(self.geo)
		self.lender.add_labeled(Button("Finish", self.finish, [], self), Button("Back", self.widget_set, [self.geo, "Geography"], self))
		self.widget_finish(self.lender)
		self.widget_finish(self.end)
		# Set Home
		self.widget_set(self.home)
	
	# Widget ----------------------------------------------------------------------------------------------------------------------------------------------
	def widget_set(self, widget, txt=None):
		if widget == self.home:
			txt = "Home"
		elif widget == self.geo:
			self.geo.update(self.basic.geo_type.value())
		elif widget == self.lender:
			self.lender.update(self.home.get_target())
		if txt != None:
			self.setWindowTitle("NCRC HMDA Template Interface - " + txt)
		self.central_widget.setCurrentWidget(widget)
	def widget_finish(self, widget):
		widget.finalize()
		self.central_widget.addWidget(widget)
	
	# Steps -----------------------------------------------------------------------------------------------------------------------------------------------
	def finish(self):
		# Open up the template
		self.widget_set(self.end, "Finish")
		template = xw.Book(self.filepath + self.template_filename)
		xw.apps.active.screen_updating = False
		self.end.add_step()
		template_charts = template.sheets['Charts']
		self.end.add_step()
		template_HMDA_all = template.sheets['HMDAAll']
		self.end.add_step()
		template_HMDA_target = template.sheets['HMDATarget']
		self.end.add_step()
		template_tract = template.sheets['TractData']
		self.end.add_step()
		# Set the geograph(y/ies)
		self.end.reset("Setting geography and year...", 0, 2)
		template_charts.range("E2").value = self.geo.geo_name_select_label.text()
		self.end.add_step()
		# Set the year(s)
		template_charts.range("E3").value = self.basic.years.text()
		self.end.add_step()
		# Get and set the MFI, Families below FPL
		geo_type = self.basic.geo_type.value()
		geos = self.geo.geo_name_select
		MFI = 0
		MFI_count = 0
		FPL = 0
		FPL_count = 0
		geos_new = []
		self.end.reset("Getting MFI and Families Below FPL...", 0, len(geos)*len(self.geo.data_MSAs))
		if geos != []:
			for i in geos:
				for j in self.geo.data_MSAs:
					if (geo_type == "MSA" and j[0] == i) or (geo_type == "State" and j[2] == i) or (geo_type == "County" and j[1] == i):
						MFI += int(j[3])
						MFI_count += 1
						FPL += int(j[4])
						FPL_count += 1
						if geo_type == "MSA":
							county = j[1] + " "
							state = j[2]
							if state == "PR":
								county += "Municipio"
							else:
								county += "County"
							geos_new.append(county + "," + state)
					self.end.add_step()
		if geo_type == "MSA":
			geos = geos_new
			geo_type = "County"
		self.end.reset("Setting MFI, Families Below FPL, and Lender...", 0, 3)
		if MFI_count != 0:
			MFI = MFI/MFI_count
		template_charts.range("E4").value = MFI
		self.end.add_step()
		if FPL_count != 0:
			FPL = FPL/FPL_count
		template_charts.range("E6").value = FPL
		self.end.add_step()
		# Set the lender(s) - currently ID only and not checked
		template_charts.range("E7").value = self.lender.lenders.text()
		self.end.add_step()
		# Fill out TractData
		row = 2
		self.end.reset("Getting and setting Tract Data...", 0, len(geos)*len(self.geo.data_other))
		if geos != []:
			for i in geos:
				for j in self.geo.data_other:
					if ((geo_type == "County" and "!" in j[5].replace(i, "!")) or (geo_type == "State" and j[0] == i)):
						template_tract.range("A" + str(row) + ":S" + str(row)).value = j
						template_tract.range("T" + str(row)).value = "=IF(S" + str(row) + '>(Charts!$E$5),"MUI","LMI")'
						template_tract.range("U" + str(row)).value = "=($H" + str(row) + "-$I" + str(row) + ")/$H" + str(row)
						template_tract.range("V" + str(row)).value = "=IF($U" + str(row) + '>(Charts!$F$18),"Minority","Non-Minority")'
						row += 1
					self.end.add_step()
		else:
			self.end.add_step()
		# Fill out the HMDA All data
		data_filename = self.home.data_filename.text()
		HMDA = xw.Book(self.filepath + data_filename)
		HMDA_sheet = HMDA.sheets.active
		row_max = HMDA_sheet.range("A1:BZ1").expand().end('down')
		row_max = str(row_max).split("$")
		row_max = int(row_max[len(row_max) - 1].replace(">", ""))
		row_target = 2
		lenders = self.lender.get()
		row_max += 1
		self.end.reset("Getting and setting HMDA All Data...", 0, row_max + 1)
		template_HMDA_all.range("A2" + ":BZ" + str(row_max)).value = HMDA_sheet.range("A2" + ":BZ" + str(row_max)).value
		self.end.add_step()
		xw.books[data_filename.split(".")[0]].close()
		extra = ['=((INDIRECT("H" & ROW()))*1000)/(INDIRECT("BS" & ROW()))', '=IF((INDIRECT("CA" & ROW()))<0.8,1,0)', '=IF((INDIRECT("CA" & ROW()))<0.79,1,0)', '=IF((INDIRECT("BZ" & ROW()))<80,1,0)']
		extra += ['=IF((INDIRECT("BZ" & ROW()))<79,1,0)', '=IF((INDIRECT("BW" & ROW()))>((Charts!F18)*100),1,0)', '=IF((INDIRECT("BW" & ROW()))>((Charts!F18)*100),1,0)']
		extra += ['=CONCATENATE(TEXT((INDIRECT("BP" & ROW())), "00"), TEXT((INDIRECT("AL" & ROW())),"000"),TEXT(((INDIRECT("W" & ROW()))*100),"000000"))', '=(INDIRECT("H" & ROW()))*1000']
		extra_all = []
		for i in range(2, row_max):
			extra_all.append(extra)
			self.end.add_step()
		template_HMDA_all.range("CA2:CH" + str(row_max - 1)).value = extra_all
		# Fill out the HMDA Target data
		if lenders != []:
			target_data = []
			if self.home.get_target():
				targets = self.lender.get_target()
				self.end.reset("Getting HMDA Target Data...", 0, len(targets))
				if targets != []:
					for i in targets:
						Target = xw.Book(self.filepath + i)
						Target_sheet = Target.sheets.active
						row_max = Target_sheet.range("A1:BZ1").expand().end('down')
						row_max = str(row_max).split("$")
						row_max = int(row_max[len(row_max) - 1].replace(">", ""))
						target_data += Target_sheet.range("A2" + ":BZ" + str(row_max)).value
						xw.books[i.split(".")[0]].close()
			else:
				self.end.reset("Getting HMDA Target Data...", 0, row_max)
				for row in range(2, row_max):
					id = res_ID(template_HMDA_all.range("BN" + str(row)).value)
					for j in lenders:
						if j == id:
							target_data += [template_HMDA_all.range("A" + str(row) + ":CH" + str(row)).value]
					self.end.add_step()
			self.end.reset("Setting HMDA Target Data...", 0, len(target_data)*2)
			row_max = len(target_data) + 1
			if row_max > 1:
				if self.home.get_target():
					template_HMDA_target.range("A2:BZ" + str(row_max)).value = target_data
					extra_all = []
					for i in range(2, row_max + 1):
						extra_all.append(extra)
						self.end.add_step()
					template_HMDA_target.range("CA2:CH" + str(row_max)).value = extra_all
				else:
					template_HMDA_target.range("A2:CH" + str(row_max)).value = target_data
		# Refresh pivot tables
		refresh = template.macro('refresh')
		refresh()
		xw.apps.active.screen_updating = True
		# Return home
		self.widget_set(self.home)
	
# Main Code -----------------------------------------------------------------------------------------------------------------------------------------------
if __name__ == "__main__":
	app = QtGui.QApplication(sys.argv)
	main = MainWindow()
	main.show()
	sys.exit(app.exec_())