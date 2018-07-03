import tkinter as tk
from tkinter import filedialog
import tkinter.messagebox
from tkinter.filedialog import askopenfilename
import pandas as pd
import sqlite3


def sharepoint(filename):
    try:
        #""" This function takes in the data from sharepoint and cleans it, it returns the cleaned dataframe """
        data = pd.read_excel(filename, sheet_name='SP')
        rows_to_drop = ['Current SharePoint 2010 sites','     Not migrating to 365 (locked)','     Migration to 365 complete (locked)','     Stale (may not need to be migrated)','Total Active Site collections (2010 & O365 sites)','Sites Created by Group creation process','O365 sites - shell - informational']
        data.drop(labels = rows_to_drop, inplace=True)
        data.dropna(inplace = True, how = 'all')
        data=data.transpose()
        return data

    except:
        tkinter.messagebox.showinfo('Error!', 'Problem with Sharepoint tab in workbook.')


def exchangecleanup(file):
    """ This function takes in the exchange dataset and cleans it up, it returns the dataframe"""
    try:

        data = pd.read_excel(file, sheet_name='Exchange')
        data.dropna(axis = 1,how = 'all' , inplace = True)
        data.fillna(value =0, inplace = True)
        for column in data.columns:
            if column == 'Department' or column == 'Total Mailbox (2007 & 365)':
                continue
            else:
                data[column] = 100*data[column]
        return data

    except:
        tkinter.messagebox.showinfo('Error!', 'Problem with Exchange tab in workbook.')

def OD4Bclean(filename):
    """ This function takes in the data fro the Onedrive usage and performs data cleaning on it, it also splits up the data into Students and faculty dataset."""

    try:

        data = pd.read_csv(filename)
        usefulcols = [ 'File Count', 'Active File Count','Storage Used (Byte)', 'College Name', 'Facutly', 'Staff', 'Student']
        data = data[usefulcols]
        data.fillna(value = {'College Name':'Unknown'}, inplace = True)
        data.dropna(inplace = True)
        data=data.astype(dtype = {'File Count': 'int64','Active File Count':'int64','Storage Used (Byte)':'int64' },inplace = True)
        Fac_Studf = data.loc[data['Facutly'] == 'yes']
        Faconly=Fac_Studf.loc[Fac_Studf['Student'] == 'no']
        Studf = data.loc[data['Student'] == 'yes']
        Stu_Staff_only=Studf.loc[Studf['Facutly'] == 'no']
        Dropthis = ['Facutly', 'Staff', 'Student']
        Stu_Staff_only = Stu_Staff_only.drop(labels = Dropthis, axis = 1)
        Faconly = Faconly.drop(labels = Dropthis, axis = 1)
        Faconly=Faconly.groupby(by ='College Name').sum()
        Stu_Staff_only = Stu_Staff_only.groupby(by ='College Name').sum()
        Stu_Staff_only['Storage in gigabyte by Students']= Stu_Staff_only['Storage Used (Byte)']/1000000000
        Faconly['Storage in gigabyte by Faculty']= Faconly['Storage Used (Byte)']/1000000000
        return Faconly,Stu_Staff_only

    except:

        tkinter.messagebox.showinfo('Error!', 'Problem with Onedrive data workbook.')



def databasesetup(sharepoint,exchange,onedrivefaculty,onedrivestudent):
    """ This function takes in the cleaned dataframes and sends it to a sqlite database"""
    try:

        conn = sqlite3.connect("UH_Office365_Migration_Database.db")
        cur = conn.cursor()
        sharepoint.to_sql('Sharepoint', conn, if_exists = 'replace')

        exchange.to_sql('Exchange', conn, if_exists = 'replace')
        onedrivefaculty.to_sql('Onedrive_Faculty_data', conn, if_exists = 'replace')
        onedrivestudent.to_sql('Onedrive_Student_data', conn, if_exists = 'replace')

        conn.commit()
        conn.close()
        tkinter.messagebox.showinfo('Success!', 'Database has been updated.')

    except:
        tkinter.messagebox.showinfo('Error!', 'Database not updated, check files.')


class Parent(tk.Tk):
    def __init__(self, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)
        container = tk.Frame(self)
        container.pack()
        container.grid_rowconfigure(0, weight = 1)
        container.grid_columnconfigure(0, weight = 1)
        self.frames = {}
        # allows the different screens to be raised to the front
        for F in (First_page,Image_view):
            frame = F(container, self)
            self.frames[F]= frame
            frame.grid(row=0,column=0,sticky = 'nsew')
        self.show_screen(First_page)

    def show_screen(self, screenpage):
        # selects the appropriate screen to be pushed to the front of the application
        frame = self.frames[screenpage]
        frame.tkraise()

class First_page(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller
        # Front page label
        Front_page_label = tk.Label(self,text = 'Please select the excel and csv files.')
        Front_page_label.pack(side='left',fill = 'x',expand =1)
        # Button used to select an image from the file system
        BrowseButton1 = tk.Button(self,text='Broswe for Excel.', command= self.browse_excel)
        BrowseButton1.pack(side= 'left',fill = 'x',expand =1,padx = 20)
        BrowseButton2 = tk.Button(self,text='Broswe for CSV.', command= self.browse_csv)
        BrowseButton2.pack(side= 'left',fill = 'x',expand =1,padx = 20)

        # Button used to view the selected image and to proceed to the next page in the app
        ViewButton1 = tk.Button(self,text='Clean the data files .', command= self.view)
        ViewButton1.pack(side= 'left',fill = 'y',expand =1,padx = 20)

    def view(self):

        try:
            #print(filename)
            global SharePointdf,Onedrivefaculty, OnedriveStudent,Exchangedf
            SharePointdf = sharepoint(filename= filename)
            Exchangedf = exchangecleanup(file=filename)
            Onedrivefaculty, OnedriveStudent = OD4Bclean(filename= csvfile)
            SharePointdf.to_excel('SharepointData.xlsx')
            Exchangedf.to_excel('ExchangeData.xlsx')
            Onedrivefaculty.to_excel('OD4BfacultyData.xlsx')
            OnedriveStudent.to_excel('OD4BStudentData.xlsx')
            #print('cleaned')
            tkinter.messagebox.showinfo('Success!', 'The files have been cleaned, you may now upload to the database.')
            self.controller.show_screen(Image_view)

        except:
            tkinter.messagebox.showinfo('Error!', 'Clean up failed. Please check the files.')


    def browse_excel(self):
        """ This function opens up a tkinter dialog box which allows the user to select the needed excel file from the file directory. """
        global filename
        #filename =  filedialog.askopenfilename(initialdir = "/",title = "Select image file",filetypes = (("jpeg files","*.jpg"),("png files","*.png"),("all files","*.*")))
        filename =  askopenfilename(initialdir = "E:/Images",title = "choose your file",filetypes = (("excel files","*.xlsx"),("all files","*.*")))
        #print(filename)
        if len(filename) == 0:
            tkinter.messagebox.showinfo('Error', 'Please select an excel file for Sharepoint and Exchange data.')
        else:
            tkinter.messagebox.showinfo('Success!', 'The Sharepoint and Exchange data excel file has been uploaded.')




    def browse_csv(self):
        """ This function opens up a tkinter dialog box which allows the user to select the needed csv file from the file directory. """

        global csvfile
        csvfile =  askopenfilename(initialdir = "E:/Images",title = "choose your file",filetypes = (("excel files","*.csv"),("all files","*.*")))
        #print(csvfile)
        if len(csvfile) == 0:
            tkinter.messagebox.showinfo('Error', 'Please select a csv file for Onedrive Data')
        else:
            tkinter.messagebox.showinfo('Success!', 'The Onedrive data csv file has been uploaded.')

    def upload():
        """ This function passes the dataframes into an sqlite database  """
        databasesetup(sharepoint=SharePointdf ,exchange=Exchangedf,onedrivefaculty=Onedrivefaculty,onedrivestudent=OnedriveStudent)


class Image_view(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller
        uploadButton = tk.Button(self,text = 'Upload to SQLite database.',command = First_page.upload)
        uploadButton.pack(side = 'bottom',fill = 'x',expand =1,padx = 20)
        processButton = tk.Button(self,text = 'Go back to homepage.',command= self.homepage)
        processButton.pack(side = 'bottom',fill = 'x',expand =1,padx = 20)

    def homepage(self):
        self.controller.show_screen(First_page)


App = Parent()
App.title('Technology Support Services')
App.geometry('800x450')
App.mainloop()
