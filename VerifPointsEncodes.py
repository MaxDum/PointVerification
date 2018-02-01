"""
Small program to check if the points of the original Exell file are the same ase the encoded one.
For this, you need to specify the orginal file name (the one you made before encoding it) with the column letter for the name (or ID) of the student and the column letter of the points.
The same informations has to be given for the encoded one (of course, if you chose the name on one side, chose it also on the other side, else it won't work).


Enjoy,


Maxim Dumortier
Orginal program 25/01/2018

Licence : GPL2
"""

from xlrd import open_workbook # pip install xlrd



original_file = "Original Sample.xls" # source file
original_name_column = "C" # column in wich the code must search after de names or ID's
original_points_column = "I" # column in wich the code must search after de points of the student

encoded_file = "Encoded Sample - NOK.xlsx"   # file returned from proeco
encoded_name_column = "B" # column in wich the code must search after de names or ID's
encoded_points_column = "G" # column in wich the code must search after de points of the student

allowedNonPoints = ["PP", "PR", "Z", "CM", "V", "ML", "FR", "SO", "D"] # sometimes instead of numbers, you can have these letters with a special meaning


class Student(object):
    def __init__(self, name, original_points, encoded_points):
        self.name = name
        self.orginal_points = original_points
        self.encoded_points = encoded_points
        
    
    def check(self): # check the different cases (float or str) and check if they are the same.
        # Return True is both are the same, else return False
        try:
            float(self.orginal_points)
            try:
                float(self.encoded_points)
                return abs(self.orginal_points-self.encoded_points) <= 0.05
            except ValueError:
                return False # original is a number but not encoded
            
        except ValueError: # in the case it's not a float the juste check if the letters are the same
            try:
                float(self.encoded_points)
                return False # original is a str but not encoded
            except ValueError:
                return self.orginal_points.upper() == self.encoded_points.upper()
        

    def __str__(self):
        return ("{0} \t original : {1} \t encoded : {2}").format(self.name, self.orginal_points, self.encoded_points)


class Students(object): # list of students
    def __init__(self, original_file, original_name_column, original_points_column, encoded_file, encoded_name_column, encoded_points_column):
        self.student_list = []
        self.__original_dict = {}
        self.__encoded_dict = {}
        self.original_file = original_file
        self.original_name_column = original_name_column
        self.original_points_column = original_points_column
        self.encoded_file = encoded_file
        self.encoded_name_column = encoded_name_column
        self.encoded_points_column = encoded_points_column
        
        self.load_original()
        self.load_encoded()
        
        if len(self.__original_dict) != len(self.__encoded_dict):
            raise ValueError("The number of students must be the same in the two files")
        else:
            self.number_students = len(self.__original_dict)
            
        self.create_list()

    def load_file(self, file, name_column, points_column):
        values = {} #dict of name : points
        wb = open_workbook(file)
        for sheet in wb.sheets():
            number_of_rows = sheet.nrows
            name_col = ord(name_column.upper())- ord("A") # integer of the expected column
            point_col = ord(points_column.upper()) - ord("A") # integer of the expected column
            
            for row in range(1, number_of_rows):
                name  = (sheet.cell(row, name_col).value)
                value = (sheet.cell(row, point_col).value)
                try:
                    value = float(value)
                    if name != '':
                        values[name] = value
                except ValueError:
                    if value in allowedNonPoints:
                        values[name] = value
                    else:
                        pass
        return values

    def load_original(self):
        self.__original_dict = self.load_file(self.original_file, self.original_name_column, self.original_points_column)

    def load_encoded(self):
        self.__encoded_dict = self.load_file(self.encoded_file, self.encoded_name_column, self.encoded_points_column)

    def create_list(self):
        for original_name in self.__original_dict.keys():
            self.student_list.append(Student(original_name, self.__original_dict[original_name], self.__encoded_dict[original_name])) 

    def print(self):
        for student in self.student_list:
            print(student)

    def check(self):
        checked = True
        incorectStudents = []
        for student in self.student_list:
            if not student.check():
                incorectStudents.append(student)
                checked = False

        if checked :
            print("==> The list is OK, well encoded :)")
        else:
            print("The list isn't OK, you must check your encoding :(")
            #self.print()
            print()
            print("\tList of the incorect encodings :")
            for student in incorectStudents:
                print(student)



s = Students(original_file, original_name_column, original_points_column, encoded_file, encoded_name_column, encoded_points_column)
#s.print() # can be used to pretty print the results
s.check()









        
