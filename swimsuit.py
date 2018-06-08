#!/usr/bin/python
# -*- coding: gbk -*-

import json, sqlite3, pyodbc, time, codecs, math, os, logging, re
import datetime
import xlrd

def UpdateLpnFnskuDict():
    conn = sqlite3.connect("swimsuit.db")
    c = conn.cursor()

    for f in os.listdir("LPN_FNSKU_Dict"):
        c.execute("SELECT NULL FROM DictFileUsed WHERE `FileName`=?", (f,))
        if None != c.fetchone():
            continue

        wb = xlrd.open_workbook("LPN_FNSKU_Dict/{}".format(f))
        sh = wb.sheet_by_index(0)

        skuIdx = -1
        lpnIdx = -1
        fRow = 0
        while fRow < 10:
            for i in range(fRow, sh.ncols):
                if "FNSKU" == sh.cell_value(0, i).upper():
                    skuIdx = i
                if sh.cell_value(0, i).upper() in ["LPN", "LICENSE-PLATE-NUMBER"]:
                    lpnIdx = i
            if skuIdx != -1 and lpnIdx != -1:
                break
            fRow += 1
        if 10 == fRow:
            print "First Row Error, Please check the name of each column"
            return

        fRow += 1
        for i in range(fRow, sh.nrows):
            sku = sh.cell_value(i, skuIdx).upper()
            lpn = sh.cell_value(i, lpnIdx).upper()
            c.execute("INSERT OR REPLACE INTO `LpnFnskuDict` (`LPN`,`FNSKU`) VALUES (?, ?)", (lpn, sku))
            conn.commit()
        c.execute("INSERT INTO `DictFileUsed` (`FileName`) VALUES (?)", (f, ))
        conn.commit()

def sq(x):
    return "N'" + x.replace("'","''") + "'"

def formatted(x):
    return x.replace("'", "").replace("-", "").replace(".", "").replace(",", "").replace(" & ", "-").replace("&", "").replace(" / ", "-").replace(" ", "-")

    
if __name__ == "__main__":
    UpdateLpnFnskuDict()
