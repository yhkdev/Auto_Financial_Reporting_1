import sqlite3

class SqliteHelper:

    def __init__(self, name=None):
        self.conn = None
        self.cursor = None

        if name:
            self._create_connection(name)

    def _create_connection(self, name):
        try:
            self.conn = sqlite3.connect(name)
            self.cursor = self.conn.cursor()
            print(sqlite3.version)
        except sqlite3.Error as e:
            print(e)

    def create_table(self):
        """ Note: 'admin' should be boolean but sqlite don't support bool so declare as integer
        """
        macro_table = """CREATE TABLE Macros (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        title TEXT NOT NULL,
                        description TEXT NOT NULL)"""

        sheet_table = """CREATE TABLE Sheets (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        macro_id INTEGER,
                        copy_from TEXT NOT NULL,
                        paste_to TEXT NOT NULL,
                        FOREIGN KEY(macro_id) REFERENCES Macros(id))"""

        # cell_table = """CREATE TABLE Cells (
        #                 id INTEGER PRIMARY KEY AUTOINCREMENT,
        #                 sheet_id INTEGER,
        #                 copy_from TEXT NOT NULL,
        #                 paste_to TEXT NOT NULL,
        #                 FOREIGN KEY(sheet_id) REFERENCES Sheets(id))"""

        try:
            c = self.cursor
            c.execute(macro_table)
            c.execute(sheet_table)
        except sqlite3.Error as e:
            print(e)

    def save_macro(self):
        pass

    def edit(self, query):  # Update
        c = self.cursor
        c.execute(query)
        self.conn.commit()

    def delete(self, query):  # Delete
        c = self.cursor
        c.execute(query)
        self.conn.commit()

    def insert(self, query, inserts):  # Insert
        c = self.cursor
        c.execute(query, inserts)
        self.conn.commit()

    # def _add_item(self, conn, data_to_insert, sql):
    #     cur = conn.cursor()
    #     cur.execute(sql, data_to_insert)
    #     return cur.lastrowid  # To be used as Foreign key for other tables

    def select(self, query):  # Select
        c = self.cursor
        c.execute(query)
        return c.fetchall()



# test = SqliteHelper("test.db")
# test.create_table()
#
# test.edit("INSERT INTO users (name,year,admin) VALUES ('john',1992,0)")