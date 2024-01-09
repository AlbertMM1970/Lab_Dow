
import sqlite3

class Sql_Data():

    def __init__(self):
        
        try:
            self.sqliteConnection = sqlite3.connect('SQLite_Lab.db')            
            print("Successfully Connected to SQLite and Database created")     
       
        except sqlite3.Error as error:
            print("Error while connecting to SQLite", error)
   
    def create_db(self):

        try: 
            cursor = self.sqliteConnection.cursor()           
            sqlite_create_table_query = '''CREATE TABLE IF NOT EXISTS db_resin (
                                id INTEGER PRIMARY KEY,                                
                                date datetime NOT NULL, 
                                resin VARCHAR(100) NOT NULL UNIQUE,                               
                                mfi DECIMAL NOT NULL,
                                density DECIMAL NOT NULL);'''            
           
            cur = cursor.execute(sqlite_create_table_query)
            self.sqliteConnection.commit()
            if cur == True:
                print("SQLite table created")
            else:
                print("SQLite table already exists")
            cursor.close()

        except sqlite3.Error as error:
            print("Error while creating a SQLite table", error)
        finally:
            if self.sqliteConnection:
                self.sqliteConnection.close()
                print("SQLite connection is closed")

    def insert_data(self, date,name, mfi, density):

        try:
            sqliteConnection = sqlite3.connect('SQLite_Lab.db')
            cursor = sqliteConnection.cursor()
            print("Successfully Connected to SQLite")

            sqlite_insert_query = '''INSERT INTO db_resin(
                                id,                                  
                                date,
                                resin, 
                                mfi, 
                                density) 
                                VALUES (null,{},{},{},{})'''.format(date,name,mfi, density)

            count = cursor.execute(sqlite_insert_query)
            sqliteConnection.commit()
            print("Record inserted successfully into db_resin table ",
                  cursor.rowcount)
            cursor.close()

        except sqlite3.Error as error:
            print("Failed to insert data into sqlite table", error)  
        finally:
            if sqliteConnection:
                sqliteConnection.close()
                print("The SQLite connection is closed")   
    
    def remove_data(self,data_to_remove):

        try:
            sqliteConnection = sqlite3.connect('SQLite_Lab.db')
            cursor = sqliteConnection.cursor()
            print("Successfully Connected to SQLite")

            sqlite_insert_query = '''DELETE FROM db_resin WHERE name = {}'''.format(data_to_remove)                         

            count = cursor.execute(sqlite_insert_query)
            sqliteConnection.commit()
            print("Record deleted successfully into db_resin table ",
                cursor.rowcount)
            cursor.close()

        except sqlite3.Error as error:
            print("Failed to insert data into sqlite table", error)
        finally:
            if sqliteConnection:
                sqliteConnection.close()
                print("The SQLite connection is closed")
    
    def update_data(self,Id, date,name, mfi, density):

        try:
            sqliteConnection = sqlite3.connect('SQLite_Lab.db')
            cursor = sqliteConnection.cursor()
            print("Successfully Connected to SQLite")

            sqlite_insert_query = '''UPDATE db_resin SET
                                     date = '{}',
                                     resin = '{}', 
                                     mfi = '{}',
                                     density = '{}',
                                     WHERE Id = '{}' '''.format(date,name, mfi, density,Id)                               
       
            count = cursor.execute(sqlite_insert_query)
            sqliteConnection.commit()
            print("Record inserted successfully into db_resin table ",
                    cursor.rowcount)
            cursor.close()

        except sqlite3.Error as error:
            print("Failed to insert data into SQLite table", error)
        finally:
            if sqliteConnection:
                sqliteConnection.close()
                print("The SQLite connection is closed")