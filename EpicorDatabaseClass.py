import pypyodbc

class EpicorDatabase:
    #conn = pypyodbc.Connection()
    
    def Connect(self,dsn,pwd):
        self.conn = pypyodbc.connect('DSN='+dsn+';PWD='+pwd)

    def Sql(self,statement):
        cur = self.conn.cursor()
        cur.execute(statement)

    def  Commit(self):
        self.conn.commit()

    def Close(self):
        self.conn.close()

    def Rollback(self):
        self.conn.rollback()
