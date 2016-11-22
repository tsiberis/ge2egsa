#!/usr/bin/env python
# -*- coding: utf-8 -*-

'''
'ge2egsa'
Converts Google Earth coordinates (lon, lat) to native Greek 'EGSA' coordinates and vice versa.
It makes use of the Proj4 library through the pyproj module.
It is associated with excel files so it makes heavy use of the xlrd & xlwt modules.
It also produces a dxf drawing and a kml file as an output with dxfwrite & simplekml modules.
GUI of choice is the superb wxpython kit.

Developed at October of 2014.
'''

#-------------------Imports-------------------
import codecs
import webbrowser
from os import path, getcwd
from math import ceil, floor
import wx
from wx.lib.wordwrap import wordwrap
import pyproj
import xlrd
import xlwt
import simplekml
from dxfwrite import DXFEngine as dxf

#-------------------Constants-------------------
VERSION = '0.1'
_a_=pyproj.Proj(init='epsg:4326') #wgs84
_b_=pyproj.Proj(init='epsg:2100') #egsa
style = xlwt.easyxf('pattern: pattern solid, fore_colour white;''align: vertical center, horizontal center;''borders: left thin, right thin, top thin, bottom thin;')
style2 = xlwt.easyxf('pattern: pattern solid, fore_colour white;''align: vertical center, horizontal center;''borders: left thin, right thin;')

#-------------------Classes-------------------

class MyApp(wx.App):
	def OnInit(self):
		self.frame = MyFrame(None)
		self.SetTopWindow(self.frame)
		self.frame.Show()
		return True

class MyFrame(wx.Frame):
	def __init__(self, parent, id=-1, title='ge2egsa v'+VERSION, pos=wx.DefaultPosition, size=(450,200), style=wx.DEFAULT_FRAME_STYLE^(wx.RESIZE_BORDER | wx.MINIMIZE_BOX | wx.MAXIMIZE_BOX), name='_my_frame'):
		super(MyFrame, self).__init__(parent, id, title, pos, size, style, name)
		self.panel = wx.Panel(self,-1)
		self.panel.SetBackgroundColour('White')
		self.CreateStatusBar()
		menubar = wx.MenuBar()
		menu1 = wx.Menu()
		_wgs2egsa = menu1.Append(-1, u'Google Earth σε ΕΓΣΑ', u'Μετατρέψτε γεωγραφικό μήκος(λ) και πλάτος(φ) σε ΕΓΣΑ')
		_egsa2wgs = menu1.Append(-1, u'ΕΓΣΑ σε Google Earth', u'Μετατρέψτε ΕΓΣΑ σε γεωγραφικό μήκος(λ) και πλάτος(φ)')
		menu1.AppendSeparator()
		_exit = menu1.Append(wx.ID_EXIT, u'Έξοδος', u'Τερματισμός του προγράμματος')
		menubar.Append(menu1, u'Ενέργεια')
		menu2 = wx.Menu()
		_tutorial = menu2.Append(wx.ID_HELP, u'Οδηγίες', u'Ανοίγει ένα tutorial στον browser σας')
		_about = menu2.Append(wx.ID_ABOUT, u'Περί', u'Πληροφορίες για το πρόγραμμα')
		menubar.Append(menu2, u'Βοήθεια')
		self.SetMenuBar(menubar)
		# Bindings
		self.Bind(wx.EVT_MENU, self.close, _exit)
		self.Bind(wx.EVT_MENU, self.tutorial, _tutorial)
		self.Bind(wx.EVT_MENU, self.about, _about)
		self.Bind(wx.EVT_MENU, self.wgs_to_egsa, _wgs2egsa)
		self.Bind(wx.EVT_MENU, self.egsa_to_wgs, _egsa2wgs)

	def close(self,event):
		self.Close()
		
	def tutorial(self,event):
		webbrowser.open_new_tab("file:///" + path.join(getcwd(),'ge2egsa_tutorial.html'))
		#webbrowser.open_new(r'ge2egsa_tutorial.pdf')

	def about(self,event):
		# Add an about dialog as a child frame
		self.AboutWindow = AboutFrame(self)
		self.AboutWindow.Show()

	def wgs_to_egsa(self,event):
		self.SetStatusText('')
		# Open an existing workbook
		book = xlrd.open_workbook('1.xls')
		wgs84_sheet = book.sheet_by_index(0)
		nr = wgs84_sheet.nrows
		# Create a new workbook
		new_book = xlwt.Workbook()
		s = new_book.add_sheet('WGS-84')
		# Write the headers
		s.write_merge(0,0,0,2,u'Γεωγραφικό Μήκος (λ)',style)
		s.write_merge(0,0,3,5,u'Γεωγραφικό Πλάτος (φ)',style)
		s.write_merge(0,1,6,7,u'Δεκαδικές Μοίρες',style)
		s.write_merge(0,0,8,9,u'ΕΓΣΑ',style)
		s.write_merge(1,1,8,8,u'Χ',style)
		s.write_merge(1,1,9,9,u'Υ',style)
		s.write_merge(0,1,10,10,u'ΕΓΣΑ ΓΙΑ AutoCAD',style)
		s.write_merge(0,1,11,11,u'ΕΓΣΑ ΓΙΑ WORD',style)
		s.write_merge(1,1,0,0,u'Μοίρες',style)
		s.write_merge(1,1,1,1,u'Πρώτα',style)
		s.write_merge(1,1,2,2,u'Δεύτερα',style)
		s.write_merge(1,1,3,3,u'Μοίρες',style)
		s.write_merge(1,1,4,4,u'Πρώτα',style)
		s.write_merge(1,1,5,5,u'Δεύτερα',style)
		# Adjust column width
		width_list = [10,10,10,10,10,10,15,15,20,20,30,30]
		for c in range(12):
			s.col(c).width = width_list[c]*256
		# Copy the existing workbook to the new one
		for row in range(2,nr):
			for col in range(6):
				a = wgs84_sheet.cell(row,col).value
				s.write(row,col,a,style2)
		# Add the decimal wgs-84 coords to the new workbook
		wgs_coords = []
		for row in range(2, nr):
			a0 = wgs84_sheet.cell(row,0).value
			a1 = wgs84_sheet.cell(row,1).value
			a2 = wgs84_sheet.cell(row,2).value
			a = abs(a0) + a1/60.0 + a2/3600.0
			if a0 < 0: a *= -1
			s.write(row,6,a,style2)
			b0 = wgs84_sheet.cell(row,3).value
			b1 = wgs84_sheet.cell(row,4).value
			b2 = wgs84_sheet.cell(row,5).value
			b = abs(b0) + b1/60.0 + b2/3600.0
			if b0 < 0: b *= -1
			s.write(row,7,b,style2)
			wgs_coords.append((a,b))
		# Save the new workbook
		new_book.save('2.xls')
		# Open the new workbook for reading
		_book = xlrd.open_workbook('2.xls')
		_wgs84_sheet = _book.sheet_by_index(0)
		_nr = _wgs84_sheet.nrows
		# Compute the egsa coords and write them in various forms
		egsa_coords = []
		for row in range(2,_nr):
			x1 = _wgs84_sheet.cell(row,6).value
			y1 = _wgs84_sheet.cell(row,7).value
			x2,y2 = pyproj.transform(_a_, _b_, x1, y1) #egsa coordinates
			egsa_coords.append((x2,y2))
			s.write(row,8,x2,style2)
			s.write(row,9,y2,style2)
			x3 = str(x2).split('.')
			y3 = str(y2).split('.')
			z = x3[0]+'.'+x3[1][:4]+','+y3[0]+'.'+y3[1][:4]
			s.write(row,10,z,style2)
			t = 'X='+x3[0]+','+x3[1][:4]+' '+'Y='+y3[0]+','+y3[1][:4]
			s.write(row,11,t,style2)
		# Save the new workbook
		new_book.save('2.xls')
		# Create a kml & a dxf file
		self.create_kml_and_dxf(wgs_coords,egsa_coords,'2')

	def egsa_to_wgs(self,event):
		self.SetStatusText('')
		# Open an existing workbook
		book = xlrd.open_workbook('1.xls')
		egsa_sheet = book.sheet_by_index(1)
		nr = egsa_sheet.nrows
		# Create a new workbook
		new_book = xlwt.Workbook()
		s = new_book.add_sheet('EGSA')
		# Write the headers
		s.write_merge(0,0,0,1,u'ΕΓΣΑ',style)
		s.write_merge(1,1,0,0,u'Χ',style)
		s.write_merge(1,1,1,1,u'Υ',style)
		s.write_merge(0,1,2,3,u'Δεκαδικές Μοίρες',style)
		s.write_merge(0,0,4,6,u'Γεωγραφικό Μήκος (λ)',style)
		s.write_merge(0,0,7,9,u'Γεωγραφικό Πλάτος (φ)',style)
		s.write_merge(1,1,4,4,u'Μοίρες',style)
		s.write_merge(1,1,5,5,u'Πρώτα',style)
		s.write_merge(1,1,6,6,u'Δεύτερα',style)
		s.write_merge(1,1,7,7,u'Μοίρες',style)
		s.write_merge(1,1,8,8,u'Πρώτα',style)
		s.write_merge(1,1,9,9,u'Δεύτερα',style)
		s.write_merge(0,1,10,10,u'ΕΓΣΑ ΓΙΑ AutoCAD',style)
		s.write_merge(0,1,11,11,u'ΕΓΣΑ ΓΙΑ WORD',style)
		# Adjust column width
		width_list = [15,15,15,15,10,10,10,10,10,10,30,30]
		for c in range(12):
				s.col(c).width = width_list[c]*256
		# Copy the existing workbook to the new one
		for row in range(2,nr):
				for col in range(2):
						a = egsa_sheet.cell(row,col).value
						s.write(row,col,a,style2)
		# Compute the wgs-84 coords and write them
		wgs_coords = []
		egsa_coords = []
		for row in range(2,nr):
				x1 = egsa_sheet.cell(row,0).value
				y1 = egsa_sheet.cell(row,1).value
				egsa_coords.append((x1,y1))
				x2,y2 = pyproj.transform(_b_, _a_, x1, y1) #wgs-84 coordinates
				s.write(row,2,x2,style2)
				s.write(row,3,y2,style2)
				wgs_coords.append((x2,y2))
				# write some egsa also
				x3 = str(x1).split('.')
				y3 = str(y1).split('.')
				for u in range(1,4):
						if len(x3[1]) == u:
								x3[1] += (4-u) * '0'
								break
				if len(x3[1]) > 4:
					x3[1] = x3[1][:4]
				for u in range(1,4):
						if len(y3[1]) == u:
								y3[1] += (4-u) * '0'
								break
				if len(y3[1]) > 4:
					y3[1] = y3[1][:4]
				z = x3[0]+'.'+x3[1]+','+y3[0]+'.'+y3[1]
				s.write(row,10,z,style2)
				t = 'X='+x3[0]+','+x3[1]+' '+'Y='+y3[0]+','+y3[1]
				s.write(row,11,t,style2)
		# Save the new workbook
		new_book.save('3.xls')
		# Open the new workbook for reading
		_book = xlrd.open_workbook('3.xls')
		_egsa_sheet = _book.sheet_by_index(0)
		_nr = _egsa_sheet.nrows
		# Compute the decimal wgs-84 coords and write them
		for row in range(2,_nr):
				x = _egsa_sheet.cell(row,2).value
				y = _egsa_sheet.cell(row,3).value
				if x < 0:
						x1 = int(ceil(x))
				else:
						x1 = int(floor(x))
				s.write(row,4,x1,style2)
				x2 = (x%1*60)//1
				s.write(row,5,x2,style2)
				x3 = ((x%1*60)%1)*60
				s.write(row,6,x3,style2)
				if y < 0:
						y1 = int(ceil(y))
				else:
						y1 = int(floor(y))
				s.write(row,7,y1,style2)
				y2 = (y%1*60)//1
				s.write(row,8,y2,style2)
				y3 = ((y%1*60)%1)*60
				s.write(row,9,y3,style2)
		# Save the new workbook
		new_book.save('3.xls')
		# Create a kml & a dxf file
		self.create_kml_and_dxf(wgs_coords,egsa_coords,'3')

	def create_kml_and_dxf(self,wgs_coords,egsa_coords,filename):
		try:
			# Create a kml file
			k = simplekml.Kml()
			ls = k.newlinestring(name='A LineString')
			ls.tessellate=1
			ls.altitudemode=simplekml.AltitudeMode.clamptoground
			ls.coords = wgs_coords
			k.save(filename+'.kml')
			# Create a dxf file
			drawing = dxf.drawing(filename+'.dxf')
			drawing.add_layer('Polyline')
			polyline= dxf.polyline(flags=0, color=7, layer='Polyline')
			polyline.add_vertices(egsa_coords)
			drawing.add(polyline)
			drawing.save()
			# All ok
			self.SetStatusText('Ok') 
		except:
			self.SetStatusText(u'Κλείστε όλα τα σχετικά αρχεία') 

class AboutFrame(wx.Frame):
	def __init__(self, parent, id=-1, title=u'Περί', pos=wx.DefaultPosition, size=(550,550), style=wx.DEFAULT_FRAME_STYLE^(wx.RESIZE_BORDER | wx.MINIMIZE_BOX | wx.MAXIMIZE_BOX), name='_about_frame'):
		super(AboutFrame, self).__init__(parent, id, title, pos, size, style, name)
		sc = wx.ScrolledWindow(self)
		sc.SetBackgroundColour('White')
		sc.SetScrollbars(0,20,0,60)
		wx.StaticText(sc, -1, wordwrap(u'''\nΤο πρόγραμμα ge2egsa μετατρέπει τις συντεταγμένες του Google Earth σε συντεταγμένες ΕΓΣΑ και το αντίστροφο. Οι μετατροπές γίνονται με τη βοήθεια της βιβλιοθήκης Proj4.\n\nΠρογραμματισμός\nΠαναγιώτης Τσιμπέρης, Δασολόγος\nΤμήμα Δασικών Χαρτογραφήσεων\nΔιευθύνσεως Δασών Ν. Σάμου\n\nΛογισμικό ανάπτυξης\nUbuntu 12.04, Python 2.7.3, wxPython 2.8.12.1, pyproj 1.8.9, xlrd 0.6.1-2, xlwt 0.7.2, simplekml 1.2.3, dxfwrite 1.2.0, Geany 0.21, LibreOffice Calc/Writer/Web 3.5.7.2\n\nΔοκιμάστηκε σε:\nUbuntu Linux 12.04, Windows XP && 7\n\nΒιβλιογραφία\n1. "Python Geospatial Development 2nd edition", Erik Westra, ISBN 978-1-78216-152-3\n2. "wxPython in Action", Noel Rappin && Robin Dunn, ISBN 1-932394-62-1\n3. "KML 2.2 – An OGC Best Practice", Google Company\n4. "Working with Excel files in Python", Chris Withers with help from John Machin\n5. "wxpython 2.8 Application Development Cookbook", Cody Precord, ISBN 978-1-849511-78-0\n\nΔιαδικτυακές πηγές\nhttp://trac.osgeo.org/proj/\nhttp://zetcode.com/wxpython/\nhttps://developers.google.com/kml/\nhttp://simplekml.readthedocs.org/en/latest/\nhttp://dxfwrite.readthedocs.org/en/latest/\nhttp://www.python-excel.org/\n\nCopyright (C) 2014 Παναγιώτης Τσιμπέρης\nΑυτό το πρόγραμμα είναι ελεύθερο λογισμικό. Μπορείτε να το αναδιανείμετε και/ή τροποποιήσετε υπό τους όρους της Ελάσσονος Γενικής Άδειας Δημόσιας Χρήσης GNU, όπως εκδόθηκε από το Ίδρυμα Ελεύθερου Λογισμικού (Free Software Foundation), την έκδοση 2.1 της Άδειας, ή (κατ’ επιλογή σας) οποιαδήποτε μεταγενέστερη έκδοση.\nΑυτό το πρόγραμμα διανέμεται με την ελπίδα ότι θα είναι χρήσιμο, αλλά ΧΩΡΙΣ ΚΑΜΙΑ ΑΠΟΛΥΤΩΣ ΕΓΓΥΗΣΗ. Χωρίς ακόμη και την σιωπηρή εγγύηση ΕΜΠΟΡΕΥΣΙΜΟΤΗΤΑΣ ή ΚΑΤΑΛΛΗΛΟΤΗΤΑΣ ΓΙΑ ΣΥΓΚΕΚΡΙΜΕΝΗ ΧΡΗΣΗ.\nΓια περισσότερες λεπτομέρειες δείτε την Ελάσσονα Γενική Άδεια Δημόσιας Χρήσης GNU. Πρέπει να έχετε λάβει ένα αντίγραφο της Ελάσσονος Γενικής Άδειας Δημόσιας Χρήσης μαζί με αυτό το πρόγραμμα. Αν  όχι, επικοινωνήστε γραπτώς με το Ίδρυμα Ελεύθερου Λογισμικού (Free Software Foundation), Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.''', 500, wx.ClientDC(self),0,5))
		self.Bind(wx.EVT_MENU, self.close)
	def close(self,event):
		self.Close()
		self.Destroy()

if __name__ == '__main__':
	app = MyApp(False)
	app.MainLoop()
