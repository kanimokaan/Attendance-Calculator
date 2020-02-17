### importing necessary libraries

from Tkinter import *
from tkFileDialog import askopenfilename
from xlrd import open_workbook

### main program

class Mpone(Frame):
    def __init__(self, parent, color, frame1, frame2, frame3, frame4, frame5, frame6, frame7):  ### defining initials
        self.parent = parent    ### will be main frame
        self.frame1 = frame1    ### will be 1st frame from the top
        self.frame2 = frame2    ### will be 2nd frame from the top
        self.frame3 = frame3    ### will be 3th frame from the top
        self.frame4 = frame4    ### will be 4th frame from the top
        self.frame5 = frame5    ### will be 5th frame from the top
        self.frame6 = frame6    ### will be 6th frame from the top
        self.frame7 = frame7    ### will be 7th frame from the top

        ### All background colors of frames are same as the input 'color'

        self.color = color
        self.parent['bg'] = self.color
        self.frame1['bg'] = self.color
        self.frame2['bg'] = self.color
        self.frame3['bg'] = self.color
        self.frame4['bg'] = self.color
        self.frame5['bg'] = self.color
        self.frame6['bg'] = self.color
        self.frame7['bg'] = self.color

        self.initGUI()    ### runs the GUI funtion

        ### file choosing

        self.browse = {}    ### storage for file type info
        self.browse['defaultextension'] = ".xls"    ### defining default file type
        self.browse['filetypes'] = [('all files', '.*'), ('excel files', '.xls')]    ### file type info

    def font(self, size=15, bold="bold"):    ### default font size is 15
        return 'Calibri %d %s' % (size, bold)

    def initGUI(self):    ### GUI function

        Label(self.frame1, text='Attendance Calculator', font=self.font(16), fg='white', bg='deepskyblue4').pack(fill=X)    ### header

        Canvas(self.frame2, height=5, bg='orange').pack(fill=X)    ### line under the header

        Label(self.frame3, text='Enter Input File Name', font=self.font(13), fg='yellow', bg='deepskyblue4').pack(side=LEFT, pady=5)    ### explanation for file choosing button
        Button(self.frame3, bg='lightgoldenrod', text='Browse', command=self.load_data).pack(side=LEFT, padx=20, pady=5)    ### file choosing button runs load_data() func

        Label(self.frame4, text='Pass Treshold (%)', font=self.font(10), fg='yellow', bg='deepskyblue4').pack(side=LEFT)    ### explanation for treshold input text box
        entry1_var = IntVar()    ### default input type for treshold input text box
        self.entry1 = Entry(self.frame4, textvariable=entry1_var, width=3)    ### defining treshold input text box as 'self.' to use it later
        self.entry1.pack(side=LEFT, padx=20)    ### packing treshold input text box


        Canvas(self.frame5, height=2, bg='orange').pack(fill=X, pady=2)    ### other line

        Button(self.frame6, bg='lightgoldenrod', text='Calculate', command=self.calculator).pack(side=LEFT, padx=95, pady=4)    ### calculating button runs calculator() func
        Button(self.frame6, bg='lightgoldenrod', text='Clear', command=self.clearence).pack(side=LEFT, padx=20, pady=4)    ### clearance button runs clearence() func

        self.text1 = Text(self.frame7, height=10)    ### text box as '.self' to use it later, it shows results
        self.text1.tag_configure("center", justify='center')    ### centered text
        self.text1.pack(side=LEFT, padx=10)    ### packing text box

    def load_data(self):    ### file reading function
        global data_sheet    ### defining as global to use anywhere
        data_sheet = open_workbook(askopenfilename()).sheet_by_index(0)    ### asks for file choosing and reads the first page

    def calculator(self):    ### attendance calculating function
        failures_monday = 0    ### to store total number of failures in monday section
        failures_tuesday = 0    ### to store total number of failures in tuesday section
        treshold = int(self.entry1.get())    ### takes the number entered by user to treshold input text box and makes it integer


        for i in range(data_sheet.nrows-2):    ### takes number of info filled rows as range, first two rows have header and titles which are unnessesary
            absence = 0    ### to storage number of absences
            if str(data_sheet.cell(i+2,4))[6::] == "'Monday'":    ### checks if data is from monday section
                for k in range(14):    ### there are 14 weeks
                    if str(data_sheet.cell(i + 2, k + 5))[6::] == "'NO'":    ### check if there is any absence
                        absence += 1    ### if there is, increase the number by 1
                if float(absence)*100/14 > 100-treshold:    ### checks if absence percentage is more than upper limit
                    failures_monday += 1    ### if it is, increase the number by 1
            elif str(data_sheet.cell(i+2,4))[6::] == "'Tuesday'":    ### checks if data is from tuesday section
                for k in range(14):    ### there are 14 weeks
                    if str(data_sheet.cell(i + 2, k + 5))[6::] == "'NO'":    ### checks if there is any absence
                        absence += 1    ### if there is, increase the number by 1
                if float(absence)*100/14 > 100-treshold:    ### checks if absence percentage is more than upper limit
                    failures_tuesday += 1    ### if it is, increase the number by 1

        failures_total = failures_monday + failures_tuesday    ### total number of failures

        ### some string variables to use in showing results
        result1 = 'Total Number of Students due to inadequate attendance:'
        result2 = 'Total number of failures in Session 1 (Monday):'
        result3 = 'Total number of failures in Session 2 (Tuesday):'

        ### inserting results to result showing text box
        ### inserts at the end of current text
        self.text1.insert(END, result1+' '+str(failures_total)+'\n')
        self.text1.insert(END, result2+' '+str(failures_monday)+'\n')
        self.text1.insert(END, result3+' '+str(failures_tuesday)+'\n')

        self.text1.tag_add("center", "1.0", "end")    ### makes texts centered

    def clearence(self):    ### clears the result showing text box
        self.text1.delete('1.0', END)    ### deletes character from first one to last one




def main():    ### main function
    global root    ### making 'root' global to use everywhere
    root = Tk()    ### 'root' is Tk() func
    root.geometry('500x315')    ### initial window size

    ### defining frames, they will used as inputs of Mpone func

    kutu1 = Frame(root)
    kutu2 = Frame(root)
    kutu3 = Frame(root)
    kutu4 = Frame(root)
    kutu5 = Frame(root)
    kutu6 = Frame(root)
    kutu7 = Frame(root)

    app = Mpone(root,"deepskyblue4", kutu1, kutu2, kutu3, kutu4, kutu5, kutu6, kutu7)

    ### packing frames like they are lines

    kutu1.pack(fill=X)
    kutu2.pack(fill=X)
    kutu3.pack(fill=X)
    kutu4.pack(fill=X)
    kutu5.pack(fill=X)
    kutu6.pack(fill=X)
    kutu7.pack(fill=X)

    root.mainloop()    ### shows the window continuously

if __name__ == '__main__':    ### runs the program
    main()