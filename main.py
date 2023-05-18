import xlrd
import sqlite3
# utilize sql for database
con = sqlite3.connect("fooddb.db")  # create db
cur = con.cursor()  # create cursor

# clean
cur.execute("DROP TABLE comps")
# create components table, don't run multiple times?
cur.execute("CREATE TABLE comps(name, cals, carbs, protein, fat, serving_size)")

# create foods table
# cur.execute("CREATE TABLE foods(name, components)")

# test
res = cur.execute("SELECT name FROM sqlite_master")
print(res.fetchone())

# open workbook
workbook = xlrd.open_workbook("Macros.xlsx")
# open worksheet
worksheet = workbook.sheet_by_index(0)  # has one sheet

# # insert components into database
row = 2  # start on row 3
col = 0  # start on first col
numrows = worksheet
for x in range (2):
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
    print(command)
    cur.execute(command)
    row = row + 1
res = cur.execute("SELECT * FROM comps")
print(res.fetchone())
