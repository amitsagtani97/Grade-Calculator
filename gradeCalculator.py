from Tkinter import *
from ScrolledText import *
import tkMessageBox
import xlrd
import xlwt
import tkFileDialog
import os.path


#This class has methods to generate final grade report.
class calculateGrade():
    #These variables stores the weightage of assignments, initiliased to 0.
    MP1 = 0
    MP2 = 0
    MP3 = 0
    MP4 = 0
    MP5 = 0
    MIDTERM = 0
    FINAL = 0
    ATTENDENCE = 0
    data = {} 
    maxMarks = 100
    students = []
    attendenceSheet = ""#"ENGR102 Attendance.xlsx"
    gradingSheet = ""#"ENGR 102 Class Gradings.xlsx"

    #Checks if the total weightage equals to 100.
    def checkTotal(self):
        if self.MP1 + self.MP2 + self.MP3 + self.MP4 + self.MP5 + self.FINAL + self.MIDTERM + self.ATTENDENCE != 100:
            return False
        else:
            return True

    #Imports data from given excel sheets and stores them into a map.
    def importData(self):
        book1 = xlrd.open_workbook(self.gradingSheet)
        book2 = xlrd.open_workbook(self.attendenceSheet)
        grades = book1.sheet_by_index(0)
        attendence1 = book2.sheet_by_index(0)
        attendence2 = book2.sheet_by_index(1)
        
        i = 1
        while True:
            try:
                firstName = grades.cell(i, 2)
                lastName = grades.cell(i, 3)
                name = str(firstName.value) + " " + str(lastName.value)
                self.data[name] = []
                self.students.append(name)

                for j in xrange(6, 13):
                    self.data[name].append(float(grades.cell(i, j).value))
                totalAttendence = attendence1.row_values(i+1).count('Y') + attendence2.row_values(i+1).count('Y')
                self.data[name].append(totalAttendence)
                i+=1
            except:
                break
        return(self.finalGrade())

    #calculates final grades based on the weightage of different assignments.
    # This uses simple unitary method to calculte relative grades.
    def finalGrade(self):
        print self.data
        self.weightage = [self.MP1, self.MP2, self.MP3, self.MP4, self.MP5, self.MIDTERM, self.FINAL, self.ATTENDENCE]
        for student in self.data:
            print student
            for i in xrange(len(self.data[student])):
                self.data[student][i] = (self.weightage[i] * self.data[student][i]) / self.maxMarks

        for student in self.data:
            self.data[student] = sum(self.data[student])

        print self.data
        self.printData()
        return (self.data, self.students)


    def printData(self):
        book = xlwt.Workbook()
        sheet1 = book.add_sheet("grade")

        for num in xrange(len(self.data)):
            name = self.students[num]
            marks = self.data[self.students[num]]

            row = sheet1.row(num)
            row.write(0, name)
            row.write(1, marks)

        book.save("grades.xlsx")


#This class creates GUI for the application. 
class createGUI:
    calcComp = calculateGrade()
    def __init__(self):
        self.window = Tk()
        self.window.title("ENGR102 Numerical Grade Calculator")

        self.window.geometry('900x600')
        
    #Draws the Gui
    def drawElements(self):
        self.lbl1 = Label(self.window, text="MP1 % ")
        self.lbl1.grid(column=0, row=0)
        self.txt1 = Entry(self.window,width=10)
        self.txt1.grid(column=1, row=0)

        self.lbl2 = Label(self.window, text="MP2 % ")
        self.lbl2.grid(column=2, row=0)
        self.txt2 = Entry(self.window,width=10)
        self.txt2.grid(column=3, row=0)

        self.lbl3 = Label(self.window, text="MP3 % ")
        self.lbl3.grid(column=4, row=0)
        self.txt3 = Entry(self.window,width=10)
        self.txt3.grid(column=5, row=0)

        self.lbl4 = Label(self.window, text="MP4 % ")
        self.lbl4.grid(column=6, row=0)
        self.txt4 = Entry(self.window,width=10)
        self.txt4.grid(column=7, row=0)

        self.lbl5 = Label(self.window, text="MP5 % ")
        self.lbl5.grid(column=8, row=0)
        self.txt5 = Entry(self.window,width=10)
        self.txt5.grid(column=9, row=0)

        self.midterm = Label(self.window, text="Midterm % ", height=2)
        self.midterm.grid(column=0, row=1)
        self.txt6 = Entry(self.window,width=10)
        self.txt6.grid(column=1, row=1)

        self.final = Label(self.window, text="Final % ", height=2)
        self.final.grid(column=0, row=2)
        self.txt7 = Entry(self.window,width=10)
        self.txt7.grid(column=1, row=2)

        self.attendence = Label(self.window, text="Attendence % ", height=2)
        self.attendence.grid(column=0, row=3)
        self.txt8 = Entry(self.window,width=10)
        self.txt8.grid(column=1, row=3)

        self.gradingFile = Label(self.window, text="Grading File ", height=2)
        self.gradingFile.grid(column=4, row=2)
        self.uploadGradingFile = Button(self.window, text="Browse", bg="orange", fg="black", command=self.openGradeSheet)
        self.uploadGradingFile.grid(column=5, row=2)
        self.gradingFileName = Label(self.window, text="", height=2)
        self.gradingFileName.grid(column=6, row=2)

        self.attendenceFile = Label(self.window, text="Attendence File ", height=2)
        self.attendenceFile.grid(column=4, row=3)
        self.uploadattendenceFile = Button(self.window, text="Browse", bg="orange", fg="black", command=self.openAttendenceSheet)
        self.uploadattendenceFile.grid(column=5, row=3)
        self.attendenceFileName = Label(self.window, text="", height=2)
        self.attendenceFileName.grid(column=6, row=3)


        self.calculateBtn = Button(self.window, text="Calculate", height=2, bg="orange", fg="black", command=self.initWeightage) #createGUI.calcComp.importData)
        self.calculateBtn.grid(column=4, row=8)

        self.saveBtn = Button(self.window, text="Save", height=2, bg="orange", fg="black")
        self.saveBtn.grid(column=5, row=8)

        self.txt = ScrolledText(self.window,width=100,height=20)
        self.txt.place(x=100, y=200)
        self.printData()
        self.window.mainloop()  

    # This method is called on clicking calculate button. And stores the weightage assigned to diferent assessments.
    def initWeightage(self):
        createGUI.calcComp.MP1 = float(self.txt1.get())
        createGUI.calcComp.MP2 = float(self.txt2.get())
        createGUI.calcComp.MP3 = float(self.txt3.get())
        createGUI.calcComp.MP4 = float(self.txt4.get())
        createGUI.calcComp.MP5 = float(self.txt5.get())
        createGUI.calcComp.MIDTERM = float(self.txt6.get())
        createGUI.calcComp.FINAL = float(self.txt7.get())
        createGUI.calcComp.ATTENDENCE = float(self.txt8.get())

        #When total assignments weightage doesn't sum to 100, this error is popped.
        if createGUI.calcComp.checkTotal() == False:
            tkMessageBox.showerror('Error', 'The assessment components do NOT sum up to 100')

        #Else grades are calculated and displayed on Text area.
        else:
            self.data, self.students = createGUI.calcComp.importData()
            self.updateResult()

    #This method is called to open a file Dialog for uploading grade sheet.
    def openGradeSheet(self):
        filename =  tkFileDialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("all files","*.*"),("jpeg files","*.jpg")))
        print filename
        createGUI.calcComp.gradingSheet = filename
        #self.gradingFileName.configure(text=filename)

    #This method is called to open a file Dialog for uploading Attendence sheet.
    def openAttendenceSheet(self):
        filename =  tkFileDialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("all files","*.*"),("jpeg files","*.jpg")))
        print filename
        createGUI.calcComp.attendenceSheet = filename
        #self.attendenceFileName.configure(text=filename)

    #This method prints the grades on the scrollable text region.
    def printData(self):
        if(os.path.exists('grades.xlsx')):
            self.txt.delete(1.0,END)
            book1 = xlrd.open_workbook('grades.xlsx')
            grades = book1.sheet_by_index(0)
            i = 0
            result = ""
            while True:
                try:
                    name = grades.cell(i, 0).value
                    result += (name + "\t\t")
                    score = grades.cell(i, 1).value
                    result += (str(score) + "\n")
                    i+=1
                except:
                    break
            print result
            self.txt.insert(INSERT,result)

    #update result on the text area.
    def updateResult(self):
        self.txt.delete(1.0,END)
        result = ""
        for i in self.students:
            result += i + "\t\t"
            result += (str(self.data[i]) + "\n")
        self.txt.insert(INSERT,result)


if __name__ == '__main__': 
    # Creates an instance of the GUI class.
    main = createGUI()
    main.drawElements()