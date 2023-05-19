##backend database manipulation
#utilize database for quick search, sort
#delete database ans save to spreadsheet to save space when not used
import xlrd
import sqlite3
"""
removes object with given name from db
"""
def remove_from_database(name, db):
    con = sqlite3.connect("fooddb.db")  # create db
    cur = con.cursor()  # create cursor
    command = "DELETE FROM {} WHERE name = \'{}\'".format(db, name)
    cur.execute(command)
    con.commit() #commit changes
    print("{} removed from table {}".format(name, db))
    
    """ **test
    res = cur.execute("SELECT name FROM comps")
    print(res.fetchall())
    """
    con.close() #save
    
    #update excel?
    
    """
    add to comps the given object
    """
def add_to_comps(name, cals, carbs, protein, fat, serving):
    #utilize sql for database
    con = sqlite3.connect("fooddb.db")  # create db
    cur = con.cursor()  # create cursor
    
    #insert
    command = "INSERT INTO comps VALUES (\'{}\', {}, {}, {}, {}, {})".format(name, cals, carbs, protein, fat, serving)
    cur.execute(command)
    con.commit()
    print("{} added to comps database".format(name))
    con.close()
    
    #update excel?
    
"""
search db with given condition - cals, name, etc
"""
def search_range(condition):
    
"""
takes from component excel sheet and populates db
"""
def create_comp_database():
    
    # utilize sql for database
    con = sqlite3.connect("fooddb.db")  # create db
    cur = con.cursor()  # create cursor

    # clean
    cur.execute("DROP TABLE comps")
    # create components table, don't run multiple times?
    cur.execute("CREATE TABLE comps(name, cals, carbs, protein, fat, serving_size)")

    # # insert components into database
    row = 2  # start on row 3
    col = 0  # start on first col
    #numrows = worksheet
    for x in range (11):
        name = worksheet.cell_value(row, col)
        # correct spaces
        name = name.replace(" ", "_")      
        cals = worksheet.cell_value(row, col + 1)
        carbs = worksheet.cell_value(row, col + 2)
        protein = worksheet.cell_value(row, col + 3)
        fat = worksheet.cell_value(row, col + 4)
        serving = worksheet.cell_value(row, col + 5)
        # price = worksheet.cell_value(row, col+6)
        command = "INSERT INTO comps VALUES (\'{}\', {}, {}, {}, {}, {})".format(name, cals, carbs, protein, fat, serving)
        #print(command)
        cur.execute(command) #execute insert
        row = row + 1 #iterate to next row
    con.commit() #commit additions
    res = cur.execute("SELECT name FROM comps")
    print(res.fetchall())
    cur.close() #write to disk



# create foods table
# cur.execute("CREATE TABLE foods(name, components)")


# open workbook
workbook = xlrd.open_workbook("Macros.xlsx")
# open component worksheet
worksheet = workbook.sheet_by_index(0)  # has one sheet

create_comp_database() #populate components
remove_from_database("pasta", "comps")