import sqlite3

conn = sqlite3.connect("DataEntryDatabase.db")
cur = conn.cursor()

def create_tb():
    cur.execute("create table if not exists DataEntryTable ( first_name text, last_name text, title text, gender text, age text, address text, course_name text, course_duration text, course_fees text, terms text)")
    print("Table has been created successfully...")
    
# create_tb()

def fetchdata():
    cur.execute("select * from DataENtryTable")
    result = cur.fetchall()
    for i in result:
        print(i)
    
fetchdata()