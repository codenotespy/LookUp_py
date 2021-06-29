import sqlite3
import pandas as pd

import os
import xlsxwriter





# To connect to the database:
connection = sqlite3.connect('db.sqlite3')
# To create cursor for sql queries:
cursor = connection.cursor()

# To delete everything in myapp_hb_price_list:
cursor.execute("DELETE FROM table1")
cursor.execute("DELETE FROM table2")
cursor.execute("DELETE FROM match1")
cursor.execute("DELETE FROM match2")
cursor.execute("DELETE FROM match3")
cursor.execute("DELETE FROM match4")
cursor.execute("DELETE FROM match5")
cursor.execute("DELETE FROM match6")
cursor.execute("DELETE FROM output")

#To insert table1.xlsx in database from excel:
wb = pd.read_excel('table1.xlsx',sheet_name = None)
for sheet in wb:
    wb[sheet].to_sql('table1', connection,index=False, if_exists="append")
#    wb[sheet].to_sql(sheet,connection,index=False, if_exists="append") - This create a list in database with same name with sheet name in excel.
connection.commit()

# TO copy field1 to field2
cursor.execute("UPDATE table1 SET field3 = field1")
connection.commit()
# TO REMOVE UNREQUIRED CHARACTERS
cursor.execute("UPDATE table1 SET field3 = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(field3, '-', ''), ',', ''), ' ', ''), '.', ''), '/', ''), '\\', ''), '|', ''), '*', ''), '+', ''), ':', ''), ';', ''), '?', '')")
connection.commit()





#To insert table2.xlsx in database from excel:
wb = pd.read_excel('table2.xlsx',sheet_name = None)
for sheet in wb:
    wb[sheet].to_sql('table2', connection,index=False, if_exists="append")
connection.commit()

# TO copy field1 to field2
cursor.execute("UPDATE table2 SET field2 = field1")
connection.commit()
# TO REMOVE UNREQUIRED CHARACTERS
cursor.execute("UPDATE table2 SET field2 = REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(field2, '-', ''), ',', ''), ' ', ''), '.', ''), '/', ''), '\\', ''), '|', ''), '*', ''), '+', ''), ':', ''), ';', ''), '?', '')")
connection.commit()



### To join for match1:
cursor.execute("INSERT INTO match1 SELECT Field1, Field2, Field3 FROM (SELECT table2.Field1, table1.Field2, table1.Field3, ROW_NUMBER() OVER (PARTITION BY table2.Field2 ORDER BY table2.Field2) AS RAWNUM FROM table2 LEFT JOIN table1 ON table1.field3 LIKE '%' || table2.Field2 || '%') WHERE RAWNUM=1")
connection.commit()

### To join for match2:
cursor.execute("INSERT INTO match2 SELECT Field1, Field2, Field3 FROM (SELECT table2.Field1, table1.Field2, table1.Field3, ROW_NUMBER() OVER (PARTITION BY table2.Field2 ORDER BY table2.Field2) AS RAWNUM FROM table2 LEFT JOIN table1 ON table1.field3 LIKE '%' || table2.Field2 || '%') WHERE RAWNUM=2")
connection.commit()

### To join for match3:
cursor.execute("INSERT INTO match3 SELECT Field1, Field2, Field3 FROM (SELECT table2.Field1, table1.Field2, table1.Field3, ROW_NUMBER() OVER (PARTITION BY table2.Field2 ORDER BY table2.Field2) AS RAWNUM FROM table2 LEFT JOIN table1 ON table1.field3 LIKE '%' || table2.Field2 || '%') WHERE RAWNUM=3")
connection.commit()

### To join for match4:
cursor.execute("INSERT INTO match4 SELECT Field1, Field2, Field3 FROM (SELECT table2.Field1, table1.Field2, table1.Field3, ROW_NUMBER() OVER (PARTITION BY table2.Field2 ORDER BY table2.Field2) AS RAWNUM FROM table2 LEFT JOIN table1 ON table1.field3 LIKE '%' || table2.Field2 || '%') WHERE RAWNUM=4")
connection.commit()

### To join for match5:
cursor.execute("INSERT INTO match5 SELECT Field1, Field2, Field3 FROM (SELECT table2.Field1, table1.Field2, table1.Field3, ROW_NUMBER() OVER (PARTITION BY table2.Field2 ORDER BY table2.Field2) AS RAWNUM FROM table2 LEFT JOIN table1 ON table1.field3 LIKE '%' || table2.Field2 || '%') WHERE RAWNUM=5")
connection.commit()

### To join for match6:
cursor.execute("INSERT INTO match6 SELECT Field1, Field2, Field3 FROM (SELECT table2.Field1, table1.Field2, table1.Field3, ROW_NUMBER() OVER (PARTITION BY table2.Field2 ORDER BY table2.Field2) AS RAWNUM FROM table2 LEFT JOIN table1 ON table1.field3 LIKE '%' || table2.Field2 || '%') WHERE RAWNUM=6")
connection.commit()







### To join for Output:
cursor.execute("INSERT INTO output SELECT table2.Field1, match1.Field2, match2.Field2, match3.Field2, match4.Field2, match5.Field2, match6.Field2 FROM table2 LEFT JOIN match1 ON match1.Field3=table2.Field2 LEFT JOIN match2 ON match2.Field3=table2.Field2 LEFT JOIN match3 ON match3.Field3=table2.Field2 LEFT JOIN match4 ON match4.Field3=table2.Field2 LEFT JOIN match5 ON match5.Field3=table2.Field2 LEFT JOIN match6 ON match6.Field3=table2.Field2")
connection.commit()









### TO EXPORT FROM SQLITE TO EXCEL:
cursor.execute("SELECT Field1, Field2, Field3, Field4, Field5, Field6 FROM output")
# TO fetch SELECTED DATA IN THE DATABASE
rows = cursor.fetchall()
# TO CREATE EXCEL FILE
workbook = xlsxwriter.Workbook('output.xlsx')
worksheet = workbook.add_worksheet()
# TO WRITE IN THE CREATED EXCEL
worksheet.write('A1', 'Field1')
worksheet.write('B1', 'Field2')
worksheet.write('C1', 'Field3')
worksheet.write('D1', 'Field4')
worksheet.write('E1', 'Field5')
worksheet.write('F1', 'Field6')

row = 1
col = 0
for module in rows:
    worksheet.write_row(row, col, module)
    row += 1

workbook.close()
connection.close()
# TO OPEN THE SAVED EXCEL FILE
os.system("start EXCEL.EXE output.xlsx")








