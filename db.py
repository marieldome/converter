import sqlite3


class Database:
    def __init__(self, db):
        # print(db)
        self.conn = sqlite3.connect(db,check_same_thread=False)
        self.cur = self.conn.cursor()
        self.cur.execute("CREATE TABLE IF NOT EXISTS items (id INTEGER PRIMARY KEY AUTOINCREMENT,item_code TEXT NOT NULL, sup_id TEXT NOT NULL)")
        self.cur.execute("CREATE TABLE IF NOT EXISTS suppliers (sup_id INTEGER, supplier_name TEXT NOT NULL)")
        self.conn.commit()
        self.insert_initial_data()

    def insert_initial_data(self):
        suppliers = [(2,'JS UNITRADE MDSE., INC.'),
                     (10,'SUYEN CORPORATION')]
                    #  [(2,'JS UNITRADE MDSE., INC.'),
                    #  (3,'COSMETIQUE ASIA CORPORATION'),
                    #  (10,'SUYEN CORPORATION'),
                    #  (13,'ACS MANUFACTURING CORPORATION'),
                    #  (14,'VALIANT DISTRIBUTION, INC.'),
                    #  (15,'MCKENZIE DISTRIBUTION CO., INC.'),
                    #  (16,'SCPG ASIA PACIFIC'),
                    #  (18,'ALECO ENTERPRISE')]
        initial_items = [('8CP',2),('10C',2),('CAD',2),('CAF',2),('CAN',2),('CDM',2),('CDN',2),('CF',2),('CHP',2),('CHD',2),('CHT',2),('CLF',2),('CM',2),('CMO',2),('CN',2),('CP',2),('CS',2),('EQ',2),('FD',2),('GP',2),('HD',2),('HH',2),('HY',2),('KM',2),('NB',2),('SAD',2),('SAW',2),('T17',2), #JS
                         ('AZ',10),('CP',10),('FM',10),('HP',10),('PX',10),('TC',10),('TD',10)] #SUYEN
        self.cur.execute("SELECT COUNT(*) FROM suppliers")
        row1 = self.cur.fetchone()
        if row1[0] == 0 :
            self.cur.executemany("INSERT INTO suppliers VALUES (?,?)",suppliers)
        self.cur.execute("SELECT COUNT(*) FROM items")
        row2 = self.cur.fetchone()
        if row2[0] == 0 :
            self.cur.executemany("INSERT INTO items VALUES (NULL,?,?)",initial_items)

        self.conn.commit()

    def fetch_supplierids(self):
        self.cur.execute("SELECT sup_id FROM suppliers")
        rows = self.cur.fetchall()
        return rows

    def fetch_supplier_by_id(self,id):
        self.cur.execute("SELECT supplier_name FROM suppliers WHERE sup_id=?",(id,))
        row = self.cur.fetchone()
        return row

    def insert_item(self,itemcode,supid):
        self.cur.execute("SELECT * FROM items WHERE item_code = ? AND sup_id = ?",(itemcode,supid))
        row = self.cur.fetchone()
        if row == None:
            self.cur.execute("INSERT INTO items VALUES (NULL,?,?)",(itemcode, supid))
            self.conn.commit()
            if self.cur.rowcount == 1:
                return "success"
        else:
            return "duplicate"
    
    def delete_item(self,itemid):
        self.cur.execute("DELETE FROM items WHERE id=?",(itemid,))
        self.conn.commit()
        if self.cur.rowcount > 0:
            return "success"
        else:
            return "failed"
    
    def update_item(self,itemid,itemcode,supid):
        self.cur.execute("UPDATE items SET item_code =?, sup_id = ? WHERE id = ?",(itemcode,supid,itemid))
        self.conn.commit()
        if self.cur.rowcount == 1:
            return "success" 
        else:
            return "failed"

    def fetch_items_by_supid(self,supid):
        self.cur.execute("SELECT i.id,i.item_code,s.sup_id,s.supplier_name FROM items i INNER JOIN suppliers s ON s.sup_id = i.sup_id WHERE i.sup_id = ?",(supid,))
        rows = self.cur.fetchall()
        return rows

    def fetch_items_convert(self,supid):  
        ilist = list()
        self.cur.execute("SELECT item_code FROM items WHERE sup_id = ?",(supid,))
        rows = self.cur.fetchall()
        for r in rows :
            ilist.append(r[0])

        return ilist

    def fetch_all_suppliers(self):
        self.cur.execute("SELECT * FROM suppliers")
        rows = self.cur.fetchall()
        return rows

    def insert_supplier(self,supid,supname):
        self.cur.execute("SELECT * FROM suppliers WHERE sup_id = ?",(supid,))
        row = self.cur.fetchone()
        if row == None:
            self.cur.execute("INSERT INTO suppliers VALUES (?,?)",(supid, supname))
            self.conn.commit()
            if self.cur.rowcount == 1:
                return "success"
        else:
            return "duplicate"

    def update_supplier(self,supid,supname):
        self.cur.execute("UPDATE suppliers SET supplier_name =? WHERE sup_id = ?",(supname,supid))
        self.conn.commit()
        if self.cur.rowcount == 1:
            return "success" 
        else:
            return "failed"

    def delete_supplier(self,supid):
        self.cur.execute("DELETE FROM suppliers WHERE sup_id=?",(supid,))
        self.conn.commit()
        if self.cur.rowcount > 0:
            return "success"
        else:
            return "failed"
        
    def __del__(self):
        self.conn.close()


