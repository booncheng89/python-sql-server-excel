import pyodbc

class connection:
    def __init__(self, conn_string, query, params):
        # refer to documentation
        # https://github.com/mkleehammer/pyodbc/wiki/Connecting-to-SQL-Server-from-Windows
        driver = r'DRIVER={ODBC Driver 17 for SQL Server};'
        self.conn =pyodbc.connect( driver + conn_string)
        self.query = query
        self.params = params

    # function to fetch data
    def fetch_query(self):
        c = self.conn.cursor()
        c.execute(self.query, self.params)
        results = c.fetchall()
        self.conn.close()
        return results

    # function without data return
    def exec_query(self):
        c = self.conn.cursor()
        c.execute(self.query, self.params)
        self.conn.commit()
        self.conn.close()

    def close(self):
        self.conn.close()
