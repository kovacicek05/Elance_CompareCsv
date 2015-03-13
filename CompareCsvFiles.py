'''
Created on 28.02.2015.

@author: Milan Kovacic
@e-mail: kovacicek@hotmail.com
@skype: kovacicek0508988
'''

from sys import exit
from os import path, remove
import pandas as pd
import codecs
from pandas import ExcelWriter

file1 = 'file1.csv'
file2 = 'file2.csv'
IdentifyDuplicates = False
WriteProcessedFiles = False

class CompareCsvFile:
    ColumnMapping = "ColumnMapping.xlsx"
    Results = "Results.csv"
    
    def __init__(self):       
        self.CheckFiles()
        self.ReadFiles()
        self.NormalizeKeyword()
        self.NormalizeCustomerId()
        self.Sort()        
        self.Compare()
        self.WriteFiles()
        print ("End")
    
    def CheckFiles(self):
        """Check if files are there"""
        
        if path.exists(file1) and path.exists(file2):
            print ("Input files are there")
        else:           
            print ("Input file(s) are not there")
            exit()
        
    def ReadFiles(self):
        """Read files and fill NA values because of sorting. 
        Sorting will report error if finds different types or empty fields"""  
             
        self.col_map = pd.read_excel(self.ColumnMapping)
        self.file_1 = pd.read_csv(file1, usecols=self.col_map["File1"], dtype=self.col_map["Data Type"], header=5, delimiter=",", encoding='latin-1', low_memory=False)
        self.file_1.fillna("N/A", inplace=True)
        self.file_2 = pd.read_csv(file2, usecols=self.col_map["File2"], dtype=self.col_map["Data Type"], header=5, delimiter=",", encoding='latin-1', low_memory=False)
        self.file_2.fillna("N/A", inplace=True)

    def NormalizeKeyword(self):
        print ("\nNormalize Keyword") 
        if "Keyword" in self.col_map["File1"].values:
            self.ReplaceKeyword(self.file_1)
        if "Keyword" in self.col_map["File2"].values:    
            self.ReplaceKeyword(self.file_2)
            
    def ReplaceKeyword(self, x):
        """ x is data frame object with Keyword column"""
        
        for i, item in enumerate (x["Keyword"]):
            if "+" in item[:1]:
                item = item.replace("+", " +", 1)
            if '"' in item:
                item = item.replace('"', '')
            if "[" in item:
                item = item.replace('[' , "")                    
            if "]"  in item:
                item = item.replace("]" , "")
            x["Keyword"][i] = item            
    
    def NormalizeCustomerId(self):
        print ("\nNormalize Customer ID") 
        if "Customer ID" in self.col_map["File1"].values:
            self.ReplaceCustomerId(self.file_1)
        if "Customer ID" in self.col_map["File2"].values:
            self.ReplaceCustomerId(self.file_2)
    
    def ReplaceCustomerId(self, x):
        """ x is data frame object with Customer ID  column.
        Format without dashes chosen because of easier converting"""
        
        for i, item in enumerate(x["Customer ID"]):
            if type(item) == str:
                if "-" in item:
                    item = item.replace("-", "")
                    x["Customer ID"][i] = item
            else:
                pass                        
                    
    def Sort(self):
        """Sorting by all columns, ascending by default.
        NA values must not be empty and data in one column must be of same type"""
        
        print ("\nSorting files")
        print ("\t Sorting file 1...")      
        self.file_1.sort(list(self.col_map["File1"].values),  inplace=True)
        print ("\t Sorting file 2...")
        self.file_2.sort(list(self.col_map["File2"].values), inplace=True)
    
    def Compare(self):
        print ("\nCompare")
        self.PrepareResultFile()
        self.FindAndRemoveDuplicates()
        self.GetTotals()
        self.CompareTotals()
        self.RemoveTail()
        
    def PrepareResultFile(self):
        """Create Output file and write header to it"""
        
        with open(CompareCsvFile.Results, "w") as f:
            cols = ["Error Description", "Row"]
            for item in self.col_map["File1"].values:
                cols.append(item)
            cols.append("")
            for item in self.col_map["File2"].values:
                cols.append(item)
            cols.append("")
            for item in self.col_map["File2"].values:
                cols.append(item)
            f.write("{}\n".format(";".join(cols)))
    
    def FindAndRemoveDuplicates(self):
        """Find duplicates using duplicated method of Data Frame class. 
        New column Duplicated has been added and if row is duplicated
        then True has been written in the proper field. Based on that 
        duplicates are printed to Output file and column has been removed.
        Adding that column is necessary, after sorting indexes in Data Frame
        are modified so this is the easiest way to print duplicates.
        After printing duplicates are removed with drop_duplicates method"""
        
        print ("\t Find And Remove Duplicates")
        for x, item in enumerate([self.file_1, self.file_2]):          
            item["Duplicate"] = item.duplicated()
            duplicates = list()
            for row in item.iterrows():
                if row[1]["Duplicate"] == True:
                    duplicates.append(row[0])
            item.drop("Duplicate", axis=1, inplace=True)
            
            if IdentifyDuplicates == True:
                with codecs.open(CompareCsvFile.Results, "a", "utf-8") as f:
                    for row in item.iterrows():
                        if row[0] in duplicates:
                            if not row[1].tolist()[0] == "N/A": #skip initially blank lines, i.e ,,,,,,,,,,,,,,,,,
                                f.write("Duplicate row in file{};{};{}\n".format(x+1, row[0]+7, ";".join(row[1].tolist())))
            item.drop_duplicates(inplace=True) 
    
    def GetTotals(self):
        """After sorting, data with --, i.e. rows with Total values are 
        in first rows, so first 4 rows are read in Series object.
        Data from 14th column further is considered, from Impressions column.
        After reading data frames, columns with totals are removed so data frames
        are clean and with unnecessary rows, ready for comparing"""
        
        print ("\t Get Totals")
        self.totals = {"File1":dict(), "File2": dict()}
        for x,item in enumerate([self.file_1, self.file_2]):
            if x == 0:
                filename = "File1"
            else:
                filename = "File2"
            counter = 0
            for row in item.head(4).iterrows():
                data = row[1].tolist()
                if "--" in data[0]:
                    counter += 1
                    if "Total" in data[1] or "Total" in data[2]:
                        #print ("Total")
                        self.totals[filename]["Total"] = data[14:]
                    elif "Search" in data[1]:
                        #print ("Search")
                        self.totals[filename]["Search Network"] = data[14:]
                    elif "Display" in data[1]:
                        #print ("Display")
                        self.totals[filename]["Display Network"] = data[14:]
                    elif "Other" in data[6]:
                        #print ("Other")
                        self.totals[filename]["Other Search Terms"] = data[14:]
                    else:
                        pass
            #remove n top rows
            item.drop(item.head(counter).index, inplace=True)
            counter = 0
        
    def CompareTotals(self):
        """Comparing Totals and printing to file"""
        
        print ("\t Compare Totals")
        with open(CompareCsvFile.Results, "a") as f:
            f.write("\nTotals\n")
            totals = ['Search Network', 'Display Network', 'Other Search Terms', 'Total']
            cols = self.col_map["File1"].values.tolist()
            index = self.col_map["File1"].values.tolist().index("Impressions")
            header = ["File", "Network"] + cols[index:]
            
            #comparing
            key = "Total Comparison"
            self.totals[key] = {
                                "Total" : list(),
                                "Other Search Terms" : list(),
                                "Search Network" : list(),
                                "Display Network" : list()
                                }
            lenght = len(self.col_map["File1"].values.tolist()) - index
            for item in self.totals[key].keys():
                if (item in self.totals["File1"].keys()) and (item in self.totals["File2"].keys()):
                    for index in range(0,lenght):
                        if self.totals["File1"][item][index] == self.totals["File2"][item][index]:
                            self.totals[key][item].append("True")
                        else:
                            self.totals[key][item].append("False")
                elif(item not in self.totals["File1"].keys()) and (item not in self.totals["File2"].keys()):
                    for index in range(0,lenght):
                        self.totals[key][item].append("True")
                else:
                    for index in range(0,lenght):
                        self.totals[key][item].append("False")
                        
            #print to file
            keys = ["File1", "File2", "Total Comparison"]
            for key in keys:
                data = ""
                #write header
                if key == "Total Comparison":
                    header[0] = key
                else:
                    header[0] = "File"
                f.write("{}\n".format(";".join(header)))
                for item in totals: 
                    if item in self.totals[key].keys():
                        data += "{};{};{}\n".format(key, item, ";".join(self.totals[key][item]))
                    else:
                        data += "{};{}\n".format(key, item)
                f.write("{}\n\n".format(data))
                data = ""
   
    def RemoveTail(self):
        """From file 1 remove: One empty line, Approximate rows per second and Time to create report"""
        
        print("\t Remove Tail From File 1")   
        #remove Empty line, Approximate rows per second and Time to create report, from file 1
        self.file_1.drop(self.file_1.tail(3).index, inplace = True)      

    def WriteFiles(self):
        """Writing Output file in xlsx format and also filtered Data Frames for file 1 and file 2.
        Second one is used for debugging"""
        
        print ("\nWriting output files in xlsx format...")
        
        #Convert Results.csv to Results.xlsx
        results = pd.read_csv(self.Results, delimiter=";")
        print ("\t Convert Results file")
        writer = ExcelWriter("Results.xlsx", engine='xlsxwriter')
        results.to_excel(writer, "Sheet1", index = False)
        writer.save()
        remove(self.Results)
        
        if WriteProcessedFiles == True:
            #Write file 1 - processed
            print ("\t Writing file 1 - It take some time")
            writer = ExcelWriter('ProcessedFile1.xlsx', engine='xlsxwriter')
            self.file_1.to_excel(writer,'Sheet1', index=False)
            writer.save()
             
            #Write file 2 - processed 
            print ("\t Writing file 2 - It take some time also")
            writer = ExcelWriter("ProcessedFile2.xlsx", engine='xlsxwriter')
            self.file_2.to_excel(writer,'Sheet1', index=False)
            writer.save()

   
def main():
    CompareCsvFile()

if __name__ == "__main__":
    main()