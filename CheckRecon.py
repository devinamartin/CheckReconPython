#Script has a set of functions to match a set of imported data between a list and a dictionary
#The list comprises of a set of check dollar Ammounts
#The dictionary comprises of a "Key" 4 digit Check Number, and a list of possiblly whole, or split dollar Ammounts
#The dict can have "Cross contamination" as there may be duplicate check # shared between different checks potentially lead
#Order in the aproach minimizes false negatives but as of now, with limited data there will be error

#returns Dict Checks, Sum of Ammounts
#Dict Checks{"check number": Check ammount}
def pullImport():
    #Read In Excel Data
    #Eventually I would like to query SF directly
    import xlrd
    book=xlrd.open_workbook("SFDeposit.xls")
    sheet=book.sheet_by_index(0)

    #Salesforce Deposit Sheet
    DSSF = {}
    DSSFsum = float(0)

    for row in range(sheet.nrows):
        #Skips Adding Header
        if row == 0:
            continue
        try:
            DSSFsum+=float(sheet.cell_value(row,5))
            DSSF.setdefault(sheet.cell_value(row,10), [])
            DSSF[sheet.cell_value(row,10)].append(float(sheet.cell_value(row,5)))
        except:
            continue
    return(DSSF, round(DSSFsum, 2))

#returns a list of check ammounts
def pullManual():
    #Read In Excel Data
    #Eventually I would like a GUI to enter in the data manaully, then exporting a report to excel.

    import xlrd
    book2=xlrd.open_workbook("ManualDeposit.xlsx")
    sheet2=book2.sheet_by_index(0)
    #Deposit Sheet Manual Entry, this is entered by hand currently but would like to impliment a optical reader tool for check scans. Check # and Ammount.
    DSME = []

    for row in range(sheet2.nrows):
        #skips adding Header
        if row == 0:
            continue
        if sheet2.cell_value(row,0) == '':
            break
        try:
            DSME.append(float(sheet2.cell_value(row,0)))
        except:
            break

    return(DSME)

#Returns (List Manual, Dict Import) unmatched items via
def removeMatches(Manual, Imported):
    ManualUnmatched = list(Manual)
    ImportUnmatched = dict(Imported)

    for entry in Manual:
        for key in Imported:
            if entry in ImportUnmatched[key]:
                #pop uses index, remove uses first matching value
                ImportUnmatched[key].remove(entry)
                ManualUnmatched.remove(entry)
                break
            else:
                continue
            break

    return (ManualUnmatched, ImportUnmatched)

#Returns a merging of split checks OR creates a No Check # Entry for formatting
def combineChecks(Imported):
    ImportedCombined = {}
    for key in Imported:
        if key == '':
            for item in Imported[key]:
                ImportedCombined.setdefault('No Check #', [])
                ImportedCombined['No Check #'].append(item)
        else:
            #use this if checksum to remove unneeded data in method 1
            keysum=(round(sum(Imported[key]),2))
            if keysum > 0:
                ImportedCombined.setdefault(key, [])
                ImportedCombined[key].append(keysum)

    return(ImportedCombined)

#Returns a Dict that has removed Keys with empty lists
def dictCleaner(Imported):
    CleanDict = dict(Imported)
    for key in Imported:
        if not Imported[key]:
            del CleanDict[key]
    return(CleanDict)


if __name__ == '__main__':
    while True:
        #TODO add a catch n/N to exit program.
        ManualEntries = pullManual()
        SalesForceEntries = pullImport()
        
        
        #Removes 1 to 1 matches from dict entries and the list entries
        FirstPass = removeMatches(ManualEntries, SalesForceEntries[0])
        #Combines the remaining sums of checks and creates the "no Check #" entries
        ConsolidateRemaining = combineChecks(FirstPass[1])
        #Removes 1 to 1 matches from the combined entries
        SecondPass = removeMatches(FirstPass[0], ConsolidateRemaining)
        #Clean Second Pass of empty lists (check values were matched)

        ManualNotFound = SecondPass[0]
        ManualSum = round(sum(ManualEntries), 2)
        
        ImportNotFound = dictCleaner(SecondPass[1])

        
        print("\n\n----------------------------------------------")
        print("Manual Sum: ", ManualSum)
        print("Import Sum: ",SalesForceEntries[1])
        #for Accounting formatting
        Difference = round(ManualSum-SalesForceEntries[1],2)
        if Difference < 0:
            print("Total Difference (Manual Under): (",-Difference,")" )
        elif Difference == 0:
            print("Manual Entry and SalesForce MATCHES!")
            input("Press Enter to continue...")
            exit()
        else:
            print("Total Difference: (Manual Over)", Difference)

        print("\n\n----------------------------------------------")
        print("Not found in the Manual Deposit Sheet")
        print("----------------------------------------------")
        ManualNotFound.sort()
        if not ManualNotFound:
            print("All entries accounted for.")
        for item in ManualNotFound:
            #TODO Add Excel Line numbers.
            print(item)

        print("\n----------------------------------------------")
        print("Not found in the SalesForce Deposits")
        print("----------------------------------------------")
        for key in ImportNotFound:
            print("Check Number: ", key,"-- Ammount: ", ImportNotFound[key][0])
        input("Press Enter to continue...")
    exit()
