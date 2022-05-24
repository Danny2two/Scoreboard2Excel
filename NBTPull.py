import python_nbt.nbt as nbt
import pandas as pd
from os.path import exists

def exportscore(search_terms, path_to_scoreboard_dat = "scorboard.dat", exportfilename = "none"):
    """
    :param search_terms:
    list of your search terms
    :param path_to_scoreboard_dat:
    point to your scoreboard.dat file
    :param exportfilename:
    Optional export to excel, give file name as string with extention xlsx
    :return:
    """
    file = nbt.read_from_nbt_file(path_to_scoreboard_dat) #Read Data, pointed to a scoreboard.dat normaly
    ExcelOutputName = exportfilename #Export file name, must end with .xlsx

    Search_Terms_List = search_terms #Give the names of the objectives we want to pull

    for term in Search_Terms_List:
        lookfor = term #This is the data we want to get
        objective = "{'type_id': 8, 'value': '" + lookfor + "'}" #create the fill string we are looking for, sometimes this might have to change
        print("Serching for: " + objective)

        datawewant = []
        stodata = file.get("data") #Getting the "data" section of the file

        scores = stodata.get("PlayerScores") #get all player scores for ALL score objectives
        for i in scores:
            if str(i.get("Objective")) == objective: #look for what matches our serch term

                name = i.get("Name") #get all of the names from the data that matched our objective
                count = 0
                finalname = ""
                for charic in str(name): #This extracts the name from the NBT data in the dunmbest way possible
                    if charic =="'": #looks for ' in the string then after certan amoint starts recording name
                        count += 1
                        #print(count)
                    else:
                        if count == 5:
                            #print(charic)
                            finalname += str(charic)
                name = finalname # set the name we pass to the list

                score = i.get("Score") #get all of the scores from the data that matched our objective
                finalscore = ""
                count = 0
                for charic in str(score):
                    if charic ==":": #looks for ' in the string then after certan amoint starts recording score
                        count += 1
                        #print(count)
                    else:
                        if count == 2 and charic != "}" and charic != " ": #extracts the number as a string and reconstructs it
                            finalscore += str(charic)
                score = finalscore

                #print(name)
                #print(score)

                namescore = (name, score)
                datawewant.append(namescore) #make all of the name and scores we just got into one list
        print("Found data matching search: \n" + str(datawewant))


        if exportfilename != "none":
            namelist = []
            scorelist = []
            for i in datawewant: #make Names and scores to be exported to excel
                namelist.append(i[0])
                scorelist.append(int(i[1])) #make the score we got in a string into an int

            col1 = "Names"
            col2 = "Score"

            if namelist != []: #check that some names where found
                excelsheet = pd.DataFrame({col1:namelist,col2:scorelist}) #set up our data frame
                #with pd.ExcelWriter(ExcelOutputName, mode= "a", if_sheet_exists="replace",engine="openpyxl") as writer: #export it to an excel sheet
                    #excelsheet.to_excel(writer, sheet_name=lookfor, index=False)

                if not exists(ExcelOutputName):
                    with pd.ExcelWriter(ExcelOutputName,engine="openpyxl") as writer:  # export it to an excel sheet
                        excelsheet.to_excel(writer, sheet_name=lookfor, index=False)
                    print("Error, Expected file "+ ExcelOutputName + " was not found! Creating it!.")
                else:
                    with pd.ExcelWriter(ExcelOutputName, mode="a", if_sheet_exists="replace",engine="openpyxl") as writer:  # export it to an excel sheet
                        excelsheet.to_excel(writer, sheet_name=lookfor, index=False)
                    print("File " + ExcelOutputName + " was edited succesfully. \n")
            else:
                print("Error: Names list empty, perhaps your serch term does not exist or has a typo.")


exportscore(search_terms=("ts_Walk","ts_Jump","ts_Sprint"),path_to_scoreboard_dat="scoreboard.dat",exportfilename = "test.xlsx")