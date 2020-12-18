# -*- coding: utf-8 -*-
import csv
import time
from datetime import datetime
from operator import itemgetter
from collections import Counter
from os.path import splitext
import wx
import wx.grid
import wx.lib.agw.aui as aui
from pandas import ExcelFile
from pandas import read_csv


#WILDCARD
readwildcard = "Excel- oder CSV-Datei (*.xlsx, *.csv)|*.xlsx;*.csv|"     \
              "TXT-Datei (*.txt)|*.txt|"        \
              "All files (*.*)|*.*"

writewildcard = "CSV-Datei (*.csv)|*.csv"

ID_IMPORT_PLAYER_1 = 100
ID_IMPORT_PLAYER_2 = 102
ID_COMPARE_PLAYER = 101
ID_EXPORT_PLAYER = 103

def diff_header():
    return ("Grund", "ID", "TTR-Diff", "TTR", "QTTR", "Name", "Vorname", "Jahrgang", "Geschlecht", "Verband", "Verein", "Letztes Spiel", "Anzahl Spiele", "Init-TTR",
        "Init-Art", "Init-Datum", "Init-Gruppe", "Init-Position", "QTTR Name", "QTTR Vorname", "QTTR Letztes Spiel", "QTTR Anzahl Spiele", 
        "QTTR Init-TTR", "QTTR Init-Art", "QTTR Init-Datum", "QTTR Init-Gruppe", "QTTR Init-Position")

def player_header():
    return ("ID", "Name", "Vorname", "Jahrgang", "Geschlecht", "Verband", "Verein", "TTR", "Letztes Spiel", 
    "Spiele", "Init-Art", "Init-TTR", "Init-Datum", "Init-Gruppe", "Init-Position")

def IsValid(range_a, range_b, diff_min, diff_max):
    try:
        return abs(range_b - range_a) < diff_max and abs(range_b - range_a) > diff_min
    except:
        return False
    
def IsValidOne(value, diff_min, diff_max):
    try:
        return value < diff_max and value > diff_min
    except:
        return False

def GetInteger(value):
    try:
        return int(value)
    except:
        return 0

def GetYear(value):
    try:
        return datetime.strptime(value, '%d.%m.%Y').strftime('%Y')
    except:
        return ""

def GetDate(value):
    try:
        if type(value) is str:
            return value
        else:
            return value.strftime('%d.%m.%Y')
    except:
        return ""

class NUSpieler(object):
    def __init__(self, nu_import):
        self.id = nu_import[0]
        self.name = nu_import[1]
        self.vorname = nu_import[2]
        self.verband = nu_import[3]
        self.verein = nu_import[4]
        self.ttr = GetInteger(nu_import[5])
        self.spiele = GetInteger(nu_import[6])
        self.letztes_spiel_datum = GetDate(nu_import[7])
        self.einstufung = nu_import[8]
        self.einstufung_ttr = GetInteger(nu_import[9])
        self.einstufung_datum = GetDate(nu_import[10])
        self.einstufung_gruppe = nu_import[11]
        self.einstufung_position = GetInteger(nu_import[12])
        self.jahrgang = GetYear(nu_import[13])
        self.geschlecht = nu_import[14]

    def tuple_all(self):
        return (self.id, self.name, self.vorname, self.jahrgang, self.geschlecht, self.verband, self.verein, self.ttr, self.letztes_spiel_datum, self.spiele,\
                self.einstufung, self.einstufung_ttr, self.einstufung_datum, self.einstufung_gruppe, self.einstufung_position)

class CompareTable(wx.grid.GridTableBase):
    def __init__(self, spieler):
        wx.grid.GridTableBase.__init__(self)
        self.data = spieler
        self.colLabels = diff_header()

    def GetNumberRows(self):
        return len(self.data)
    def GetNumberCols(self):
        return len(self.data[0])
    def GetColLabelValue(self, col):
        if len(self.colLabels) > 0:
            return self.colLabels[col]
    def IsEmptyCell(self, row, col):
        return self.data[row][col] == ''
    def GetValue(self, row, col):
        return self.data[row][col]
    def SetValue(self, row, col, value):
        self.data[row][col] = value
    def GetAttr(self, row, col, kind):
        attr = wx.grid.GridCellAttr()
        attr.SetBackgroundColour('#f0f8ff' if col > 17 or col == 4 else wx.WHITE)
        return attr

class NewPlayerTable(wx.grid.GridTableBase):
    def __init__(self, spieler):
        wx.grid.GridTableBase.__init__(self)
        self.data = spieler
        self.colLabels = player_header()

    def GetNumberRows(self):
        return len(self.data)
    def GetNumberCols(self):
        return len(self.data[0])
    def GetColLabelValue(self, col):
        if len(self.colLabels) > 0:
            return self.colLabels[col]
    def IsEmptyCell(self, row, col):
        return self.data[row][col] == ''
    def GetValue(self, row, col):
        return self.data[row][col]
    def SetValue(self, row, col, value):
        self.data[row][col] = value

class PlayerTable(wx.grid.GridTableBase):
    def __init__(self, spieler):
        wx.grid.GridTableBase.__init__(self)
        self.data = []
        for s in spieler:
            self.data.append(spieler[s].tuple_all())
        self.colLabels = player_header()

    def GetNumberRows(self):
        return len(self.data)
    def GetNumberCols(self):
        return len(self.data[0])
    def GetColLabelValue(self, col):
        if len(self.colLabels) > 0:
            return self.colLabels[col]
    def IsEmptyCell(self, row, col):
        return self.data[row][col] == ''
    def GetValue(self, row, col):
        return self.data[row][col]
    def SetValue(self, row, col, value):
        self.data[row][col] = value

class ResultPanel(wx.Panel):
    def __init__(self, parent, ttrimport):
        wx.Panel.__init__(self, parent=parent)
        self.SetBackgroundColour('#DDE2E6')
        ttrimport.resultNotebook = aui.AuiNotebook(self, style=wx.NB_MULTILINE)

        ttrimport.resultBox = wx.BoxSizer(wx.VERTICAL)
        ttrimport.resultBox.Add(ttrimport.resultNotebook, 1, wx.EXPAND|wx.ALL, border=0)
        self.SetAutoLayout(True)
        self.SetSizer(ttrimport.resultBox)
        self.Layout()

class ImportPanel(wx.Panel):
    def __init__(self, parent, ttrimport):
        wx.Panel.__init__(self, parent=parent)
        self.SetBackgroundColour('#CCE2E6')
        ttrimport.importNotebook = aui.AuiNotebook(self, style=wx.NB_MULTILINE)

        ttrimport.importBox = wx.BoxSizer(wx.VERTICAL)
        ttrimport.importBox.Add(ttrimport.importNotebook, 1, wx.EXPAND|wx.ALL, border=0)
        self.SetAutoLayout(True)
        self.SetSizer(ttrimport.importBox)
        self.Layout()

class PropPanel(wx.Panel):
    def __init__(self, parent):
        wx.Panel.__init__(self, parent=parent)
        self.SetBackgroundColour('#D4E1ED')

class LogPanel(wx.Panel):
    def __init__(self, parent, ttrimport):
        wx.Panel.__init__(self, parent=parent)

        ttrimport.log = wx.TextCtrl(self, wx.ID_ANY, size=(200,300),
            style = wx.TE_MULTILINE|wx.TE_READONLY)
        font = ttrimport.log.GetFont()
        font.SetFamily(wx.FONTFAMILY_MODERN)
        ttrimport.log.SetFont(font)

        ttrimport.logBox = wx.BoxSizer(wx.VERTICAL)
        ttrimport.logBox.Add(ttrimport.log, 1, wx.EXPAND|wx.ALL, border=0)
        self.SetAutoLayout(True)
        self.SetSizer(ttrimport.logBox)
        self.Layout()

class TTR_Vergleich(wx.Frame):
    def __init__(self, title):
        wx.Frame.__init__(self, None, wx.ID_ANY, title, pos = (100,100), size = (800,500))
        self.player_file_opened = False
        self.player_grid_init = False
        self.player_file_opened2 = False
        self.player_grid_init2 = False

        self.InitLayout()
        self.InitButtons()
        self.InitBindings()
        self.Maximize()
        self.Show(True)

    def LogLine(self, text_a):
        self.log.AppendText(text_a + "\n")

    def DefineStatus(self, text_a, text_b):
        self.sb.SetStatusText(text_a, 0)
        self.sb.SetStatusText(text_b, 1)
        self.LogLine(text_a + " (" + text_b + ")")

    def DefineStatStatus(self, text_a, text_b):
        self.sb.SetStatusText(text_a, 0)
        self.sb.SetStatusText(text_b, 1)
        self.LogLine(text_a + " " + text_b)

    def DefineStatistics(self, box, text_label, text_statistics):
        text = wx.StaticText(self.propPanel, label=text_label)
        text.SetFont(wx.Font(10, wx.DEFAULT, wx.NORMAL, wx.BOLD))
        if box:
            box.Add(text, 0, wx.EXPAND)
            box.Add(wx.StaticText(self.propPanel, label=text_statistics), 0, wx.EXPAND|wx.ALIGN_RIGHT, 20)
        self.DefineStatStatus("  {0:21s}".format(text_label), "{0:20s}".format(text_statistics))

    def DefineStatistics2(self, box, text_label, text_statistics1, text_statistics2):
        text = wx.StaticText(self.propPanel, label=text_label)
        text.SetFont(wx.Font(10, wx.DEFAULT, wx.NORMAL, wx.BOLD))
        box.Add(text, 0, wx.EXPAND)
        box.Add(wx.StaticText(self.propPanel, label=text_statistics1), 0, wx.EXPAND|wx.ALIGN_RIGHT)
        box.Add(wx.StaticText(self.propPanel, label=text_statistics2), 0, wx.EXPAND|wx.ALIGN_RIGHT)
        self.DefineStatStatus("  {0:25s}".format(text_label), "{0:12s}{1:12s}".format(text_statistics1, text_statistics2))

    def InitLayout(self):
        vSplitter = wx.SplitterWindow(self, wx.ID_ANY)
        hSplitter = wx.SplitterWindow(vSplitter, wx.ID_ANY)
        topvSplitter = wx.SplitterWindow(hSplitter, wx.ID_ANY)

        self.propPanel = PropPanel(vSplitter)
        self.resultPanel = ResultPanel(hSplitter, self)
        self.importPanel = ImportPanel(topvSplitter, self)
        self.logPanel = LogPanel(topvSplitter, self)

        topvSplitter.SplitVertically(self.importPanel, self.logPanel, 1000)
        hSplitter.SplitHorizontally(topvSplitter, self.resultPanel, 300)
        vSplitter.SplitVertically(self.propPanel, hSplitter, 300)

        self.sb = wx.StatusBar(self, wx.ID_ANY)
        self.sb.SetFieldsCount(2)
        self.sb.SetStatusWidths([400, 500])
        self.SetStatusBar(self.sb)

    def InitButtons(self):
        self.Freeze()

        self.propSizer = wx.BoxSizer(wx.VERTICAL)

        self.gauge = wx.Gauge(self.propPanel, -1, 50, size=(180, 25))
        self.propSizer.Add(self.gauge, 0, wx.LEFT|wx.UP, 10)

        self.propImportPlayerButton = wx.Button(self.propPanel, ID_IMPORT_PLAYER_1, label='TTR-Player-Import', size=(150, 25))
        self.propSizer.Add(self.propImportPlayerButton, 0, wx.LEFT|wx.UP, 20)
        self.propImportPlayerButton2 = wx.Button(self.propPanel, ID_IMPORT_PLAYER_2, label='QTTR-Player-Import', size=(150, 25))
        self.propSizer.Add(self.propImportPlayerButton2, 0, wx.LEFT|wx.UP, 20)
        self.playerSizer = wx.BoxSizer(wx.HORIZONTAL)
        self.propCompareButton = wx.Button(self.propPanel, ID_COMPARE_PLAYER, label='Player-Vergleich', size=(150, 25))
        self.playerSizer.Add(self.propCompareButton, 0)
        self.propSizer.Add(self.playerSizer, 0, wx.LEFT|wx.UP, 20)

        self.propImportPlayerButton2.Disable()
        self.playerSizer.Hide(self.propCompareButton)

        self.propPanel.SetAutoLayout(True)
        self.propPanel.SetSizer(self.propSizer)
        self.propPanel.Layout()
        self.Thaw()
        self.propPanel.Refresh()

    def InitBindings(self):
        # Bindings fuer die Menue-Felder
        self.Bind(wx.EVT_BUTTON, self.OnImportPlayer1, id=ID_IMPORT_PLAYER_1)
        self.Bind(wx.EVT_BUTTON, self.OnImportPlayer2, id=ID_IMPORT_PLAYER_2)
        self.Bind(wx.EVT_BUTTON, self.OnComparePlayer, id=ID_COMPARE_PLAYER)
        self.Bind(wx.EVT_BUTTON, self.OnExportPlayer, id=ID_EXPORT_PLAYER)

    def FillPlayerGrid(self):
        self.allplayerGrid = wx.grid.Grid(self.importNotebook)
        self.allplayerGrid.CreateGrid(100, 26)
        self.allplayerGrid.SetColLabelSize(20)
        self.importNotebook.AddPage(self.allplayerGrid, "Importierte TTR-Player")
        self_table = PlayerTable(self.allplayer)
        self.allplayerGrid.SetTable(self_table, True)
        self.allplayerGrid.EnableEditing(False)
        
    def ReadWorksheet(self, ws, player, i, start):
        self.DefineStatus("{0:26s} {1:7d} ".format("Excel-Import: ", len(ws.index)),
                          "Laufzeit: %1.2fs" % (time.process_time() - start))
        self.gauge.SetRange(len(ws.index))
        try:
            for row in zip(ws['InterneNr'], ws['Nachname'], ws['Vorname'], ws['Verband'], ws['Verein'], ws['TTR'],
                           ws['Anzahl Einzel gesamt'], ws['Letztes Spiel'], ws['Einstufungsart zuletzt'], 
                           ws['Einstufungswert zuletzt'], ws['Einstufungsdatum zuletzt'],
                           ws['Einstufungsgruppe zuletzt'], ws['Einstufungsposition zuletzt'],
                           ws['Geburtsdatum'], ws['Geschlecht'], ws['Status'], ws['Kumulierte Inaktivitätsabzüge']):
                if True or row[15] != 'i' and row[6] > 10 and row[14] == "M" and int(row[7][-4:])>2015 and int(row[16]) == 0 \
                and 2019-int(row[13]) > 18:
                    player.update({row[0]:NUSpieler(row)})
                    if i%10000 == 0:
                        self.gauge.SetValue(i)
                    i += 1
        except:
            for row in zip(ws['InterneNr'], ws['Nachname'], ws['Vorname'], ws['Verband'], ws['Verein'], ws['TTR aus Snapshot'],
                           ws['Anzahl Einzel aus Snapshot'], ws['Letztes Spiel aus Snapshot'], ws['Einstufungsart zuletzt'], 
                           ws['Einstufungswert zuletzt'], ws['Einstufungsdatum zuletzt'],
                           ws['Einstufungsgruppe zuletzt'], ws['Einstufungsposition zuletzt'],
                           ws['Geburtsdatum'], ws['Geschlecht'], ws['Status'], ws['Kumulierte Inaktivitätsabzüge aus Snapshot']):
                if True or row[15] != 'i' and row[6] > 10 and row[14] == "M" and int(row[7][-4:])>2015 and int(row[16]) == 0 \
                and 2019-int(row[13]) > 18:
                    player.update({row[0]:NUSpieler(row)})
                    if i%10000 == 0:
                        self.gauge.SetValue(i)
                    i += 1
        self.gauge.SetValue(i)
        
    def Open(self, path, player):
        i = 1
        start = time.process_time()
        self.gauge.SetValue(0)
        ext = splitext(path)[1]
        if ext == ".csv":
            ws = read_csv(path, sep=';', encoding='utf8')
            self.ReadWorksheet(ws, player, i, start)
        elif ext == ".xlsx":
            wb = ExcelFile(path)
            self.DefineStatus("{0:26s} {1:7d} ".format("Excel-Open: ", 0),
                                  "Laufzeit: %1.2fs" % (time.process_time() - start))
            for ws_name in wb.sheet_names:
                ws = wb.parse(ws_name)
                self.ReadWorksheet(ws, player, i, start)

    def OpenPlayer(self):
        if self.player_file_opened == False:
            d = wx.FileDialog(None,"W\xe4hle eine Player-Datei",".","",readwildcard,wx.FD_OPEN)
            self.allplayer = dict()
            if d.ShowModal() == wx.ID_OK:
                start = time.process_time()
                self.Open(d.GetPath(), self.allplayer)
                self.FillPlayerGrid()
                self.sb.SetStatusText(d.GetPath(), 1)
                self.Refresh()
                self.player_file_opened = True            
            d.Destroy()
            self.DefineStatus("{0:26s} {1:7d} ".format("Importierte TTR-Spieler: ", len(self.allplayer)),
                    "Laufzeit: %1.2fs" % (time.process_time() - start))
        else:
            wx.MessageBox('Du kannst nur eine Player-Datei laden!', 'Info', wx.OK | wx.ICON_ERROR)

    def OnImportPlayer1(self, e):
        self.OpenPlayer()

        if self.player_file_opened == True:
            self.propImportPlayerButton2.Enable()
            self.propImportPlayerButton.Disable()
            self.propPanel.Layout()

    def FillPlayerGrid2(self):
        self.allplayerGrid2 = wx.grid.Grid(self.importNotebook)
        self.allplayerGrid2.CreateGrid(100, 26)
        self.allplayerGrid2.SetColLabelSize(20)
        self.importNotebook.AddPage(self.allplayerGrid2, "Importierte QTTR-Player")
        self_table2 = PlayerTable(self.allplayer2)
        self.allplayerGrid2.SetTable(self_table2, True)
        self.allplayerGrid2.EnableEditing(False)

    def OpenPlayer2(self):
        if self.player_file_opened2 == False:
            self.allplayer2 = dict()
            d = wx.FileDialog(None,"W\xe4hle eine Player-Datei",".","",readwildcard,wx.FD_OPEN)
            if d.ShowModal() == wx.ID_OK:
                start = time.process_time()
                self.Open(d.GetPath(), self.allplayer2)
                self.FillPlayerGrid2()
                self.sb.SetStatusText(d.GetPath(), 1)
                self.Refresh()
                self.player_file_opened2 = True
            d.Destroy()
            self.DefineStatus("{0:26s} {1:7d} ".format("Importierte QTTR-Spieler: ", len(self.allplayer2)),
                    "Laufzeit: %1.2fs" % (time.process_time() - start))
        else:
            wx.MessageBox('Du kannst nur eine Player-Datei laden!', 'Info', wx.OK | wx.ICON_ERROR)

    def OnImportPlayer2(self, e):
        self.OpenPlayer2()

        if self.player_file_opened2 == True:
            self.playerSizer.Show(self.propCompareButton)
            self.propImportPlayerButton2.Disable()
            self.propPanel.Layout()

    def WriteResultFile(self, fullpath, spielerliste):
        f = open(fullpath, "w", encoding="utf-8", newline="")
        writer = csv.writer(f, delimiter=';')
        writer.writerow(diff_header())
        writer.writerows(spielerliste)
        f.close()

    def OnExportPlayer(self, e):
        if len(self.ttr_spieler) > 0:
            d = wx.FileDialog(None,"W\xe4hle einen Datei-Namen f\xfcr die Export-Datei!",".","",writewildcard,wx.FD_SAVE)
            if d.ShowModal() == wx.ID_OK:
                path = d.GetPath()
                prepath = path.split('.')[0]

                self.WriteResultFile(prepath + "_geaenderter_ttr.csv", self.ttr_spieler)
                self.WriteResultFile(prepath + "_geaenderter_nachname.csv", self.name_spieler)
                self.WriteResultFile(prepath + "_geaenderter_vorname.csv", self.vorname_spieler)
                self.WriteResultFile(prepath + "_geaenderte_initialisierung.csv", self.init_spieler)
                self.WriteResultFile(prepath + "_geaenderter_init_wert.csv", self.init_wert_spieler)
                self.WriteResultFile(prepath + "_geaendertes_init_datum.csv", self.init_datum_spieler)
                self.WriteResultFile(prepath + "_geaendertes_init_gruppe.csv", self.init_gruppe_spieler)
                self.WriteResultFile(prepath + "_geaendertes_init_position.csv", self.init_position_spieler)
                self.WriteResultFile(prepath + "_geaenderte_spieleanzahl.csv", self.spiele_spieler)
                self.WriteResultFile(prepath + "_geaendertes_letztes_spiel.csv", self.letztes_spiel_spieler)
                self.WriteResultFile(prepath + "_neue_spieler.csv", self.neue_spieler)
                self.WriteResultFile(prepath + "_entfernte_spieler.csv", self.entfernte_spieler)
            d.Destroy()

    def diff_tuple(self, id):
        reason_string = ""
        if self.allplayer2[id].ttr != self.allplayer[id].ttr != 0:
            reason_string += "Ge\xe4nderter TTR-Wert, "
        if self.allplayer2[id].einstufung != self.allplayer[id].einstufung != 0:
            reason_string += "Ge\xe4nderte Init-Art, "
        if self.allplayer2[id].einstufung_ttr != self.allplayer[id].einstufung_ttr != 0:
            reason_string += "Ge\xe4nderter Init-Wert, "
        if self.allplayer2[id].einstufung_datum != self.allplayer[id].einstufung_datum != 0:
            reason_string += "Ge\xe4ndertes Init-Datum, "
        if self.allplayer2[id].einstufung_gruppe != self.allplayer[id].einstufung_gruppe != 0:
            reason_string += "Ge\xe4nderte Init-Gruppe, "
        if self.allplayer2[id].einstufung_position != self.allplayer[id].einstufung_position != 0:
            reason_string += "Ge\xe4nderte Init-Position, "
        if self.allplayer2[id].name != self.allplayer[id].name != 0:
            reason_string += "Ge\xe4nderter Name, "
        if self.allplayer2[id].vorname != self.allplayer[id].vorname != 0:
            reason_string += "Ge\xe4nderter Vorname, "
        if self.allplayer2[id].spiele != self.allplayer[id].spiele != 0:
            reason_string += "Ge\xe4nderte Spiele-Anzahl, "
        if self.allplayer2[id].letztes_spiel_datum != self.allplayer[id].letztes_spiel_datum != 0:
            reason_string += "Ge\xe4ndertes letztes Spiel-Datum, "
        return (reason_string, #0
    id, #1
    abs(int(self.allplayer2[id].ttr) - int(self.allplayer[id].ttr)), #2
    self.allplayer[id].ttr, #3
    self.allplayer2[id].ttr if self.allplayer[id].ttr != self.allplayer2[id].ttr else "", #4
    self.allplayer[id].name, #5
    self.allplayer[id].vorname, #6
    self.allplayer[id].jahrgang,
    self.allplayer[id].geschlecht,
    self.allplayer[id].verband, #7
    self.allplayer[id].verein,#8
    self.allplayer[id].letztes_spiel_datum, #9
    self.allplayer[id].spiele, #10
    self.allplayer[id].einstufung_ttr, #11
    self.allplayer[id].einstufung, #12
    self.allplayer[id].einstufung_datum, #13
    self.allplayer[id].einstufung_gruppe,
    self.allplayer[id].einstufung_position,
    self.allplayer2[id].name if self.allplayer[id].name != self.allplayer2[id].name else "", #14
    self.allplayer2[id].vorname if self.allplayer[id].vorname != self.allplayer2[id].vorname else "", #15
    self.allplayer2[id].letztes_spiel_datum if self.allplayer[id].letztes_spiel_datum != self.allplayer2[id].letztes_spiel_datum else "", #16
    self.allplayer2[id].spiele if self.allplayer[id].spiele != self.allplayer2[id].spiele else "", #17
    self.allplayer2[id].einstufung_ttr if self.allplayer[id].einstufung_ttr != self.allplayer2[id].einstufung_ttr else "", #18
    self.allplayer2[id].einstufung if self.allplayer[id].einstufung != self.allplayer2[id].einstufung else "", #19
    self.allplayer2[id].einstufung_datum if self.allplayer[id].einstufung_datum != self.allplayer2[id].einstufung_datum else "", #20
    self.allplayer2[id].einstufung_gruppe if self.allplayer[id].einstufung_gruppe != self.allplayer2[id].einstufung_gruppe else "", #21
    self.allplayer2[id].einstufung_position if self.allplayer[id].einstufung_position != self.allplayer2[id].einstufung_position else "" #22
    )

    def AddResultGrid(self, spielerliste, name, use_compare_table):
        if len(spielerliste):
            result_grid = wx.grid.Grid(self.resultNotebook)
            result_grid.CreateGrid(0, 9)
            result_grid.SetColLabelSize(15)
            self.resultNotebook.AddPage(result_grid, name)
            if use_compare_table:
                result_grid.SetTable(CompareTable(spielerliste), True)
            else:
                result_grid.SetTable(NewPlayerTable(spielerliste), True)
            if len(spielerliste) < 1000:
                result_grid.AutoSize()
            result_grid.EnableEditing(False)

    def OnComparePlayer(self, e):
        start = time.process_time()
        self.gauge.SetValue(0)
        self.gauge.SetRange(len(self.allplayer) + len(self.allplayer2))

        ttr_list = list()
        ttr2_list = list()
        self.ttr_spieler = list()
        self.letztes_spiel_spieler = list()
        self.spiele_spieler = list()
        self.name_spieler = list()
        self.vorname_spieler = list()
        self.neue_spieler = list()
        self.entfernte_spieler = list()
        self.init_spieler = list()
        self.init_wert_spieler = list()
        self.init_datum_spieler = list()
        self.init_gruppe_spieler = list()
        self.init_position_spieler = list()
        ttr_sum = 0
        ttr2_sum = 0
        ttr_min = 3000
        ttr_max = 0
        ttr2_min = 3000
        ttr2_max = 0
        i = 0
        for id in self.allplayer2:
            ttr2_list.append(self.allplayer2[id].ttr)
            if self.allplayer2[id].ttr != '':
                ttr2_sum += int( self.allplayer2[id].ttr )
                ttr2_min = min(ttr2_min, int(self.allplayer2[id].ttr))
                ttr2_max = max(ttr2_max, int(self.allplayer2[id].ttr))
            if id in self.allplayer:
                tuple = self.diff_tuple(id)
                if self.allplayer[id].ttr != self.allplayer2[id].ttr:
                    self.ttr_spieler.append(tuple)
                if self.allplayer[id].name != self.allplayer2[id].name:
                    self.name_spieler.append(tuple)
                if self.allplayer[id].vorname != self.allplayer2[id].vorname:
                    self.vorname_spieler.append(tuple)
                if self.allplayer[id].spiele != self.allplayer2[id].spiele:
                    self.spiele_spieler.append(tuple)
                if self.allplayer[id].letztes_spiel_datum != self.allplayer2[id].letztes_spiel_datum:
                    self.letztes_spiel_spieler.append(tuple)
                if self.allplayer[id].einstufung != self.allplayer2[id].einstufung:
                    self.init_spieler.append(tuple)
                if self.allplayer[id].einstufung_ttr != self.allplayer2[id].einstufung_ttr:
                    self.init_wert_spieler.append(tuple)
                if self.allplayer[id].einstufung_datum != self.allplayer2[id].einstufung_datum:
                    self.init_datum_spieler.append(tuple)
                if self.allplayer[id].einstufung_gruppe != self.allplayer2[id].einstufung_gruppe:
                    self.init_gruppe_spieler.append(tuple)
                if self.allplayer[id].einstufung_position != self.allplayer2[id].einstufung_position:
                    self.init_position_spieler.append(tuple)
            else:
                self.neue_spieler.append((id, self.allplayer2[id].name, self.allplayer2[id].vorname,
                    self.allplayer2[id].jahrgang, self.allplayer2[id].geschlecht,
                    self.allplayer2[id].verband, self.allplayer2[id].verein, self.allplayer2[id].ttr,
                    self.allplayer2[id].letztes_spiel_datum, self.allplayer2[id].spiele,
                    self.allplayer2[id].einstufung, self.allplayer2[id].ttr, self.allplayer2[id].einstufung_datum,
                    self.allplayer2[id].einstufung_gruppe, self.allplayer2[id].einstufung_position))
            if i%10000 == 0:
                self.gauge.SetValue(i)
            i += 1
        for id in self.allplayer:
            ttr_list.append(self.allplayer[id].ttr)
            ttr_sum += int(self.allplayer[id].ttr)
            ttr_min = min(ttr_min, int(self.allplayer[id].ttr))
            ttr_max = max(ttr_max, int(self.allplayer[id].ttr))
            if id not in self.allplayer2:
                self.entfernte_spieler.append((id, self.allplayer[id].name, self.allplayer[id].vorname,
                    self.allplayer[id].jahrgang, self.allplayer[id].geschlecht,
                    self.allplayer[id].verband, self.allplayer[id].verein, self.allplayer[id].ttr,
                    self.allplayer[id].letztes_spiel_datum, self.allplayer[id].spiele,
                    self.allplayer[id].einstufung, self.allplayer[id].ttr, self.allplayer[id].einstufung_datum,
                    self.allplayer[id].einstufung_gruppe, self.allplayer[id].einstufung_position))
            if i%10000 == 0:
                self.gauge.SetValue(i)
            i += 1
        self.gauge.SetValue(i)
        self.ttr_spieler.sort(key=itemgetter(2), reverse=True)
        self.init_spieler.sort(key=itemgetter(2), reverse=True)
        self.init_wert_spieler.sort(key=itemgetter(2), reverse=True)
        self.init_datum_spieler.sort(key=itemgetter(2), reverse=True)
        self.init_gruppe_spieler.sort(key=itemgetter(2), reverse=True)
        self.init_position_spieler.sort(key=itemgetter(2), reverse=True)
        self.name_spieler.sort(key=itemgetter(2), reverse=True)
        self.vorname_spieler.sort(key=itemgetter(2), reverse=True)
        self.spiele_spieler.sort(key=itemgetter(2), reverse=True)
        self.letztes_spiel_spieler.sort(key=itemgetter(2), reverse=True)
        self.neue_spieler.sort(key=itemgetter(8), reverse=True)
        self.entfernte_spieler.sort(key=itemgetter(8), reverse=True)

        self.AddResultGrid(self.ttr_spieler, "Ge\xe4nderter TTR", True)
        self.AddResultGrid(self.init_spieler, "Ge\xe4nderte Init-Art", True)
        self.AddResultGrid(self.init_wert_spieler, "Ge\xe4nderter Init-Wert", True)
        self.AddResultGrid(self.init_datum_spieler, "Ge\xe4ndertes Init-Datum", True)
        self.AddResultGrid(self.init_gruppe_spieler, "Ge\xe4nderte Init-Gruppe", True)
        self.AddResultGrid(self.init_position_spieler, "Ge\xe4nderte Init-Position", True)
        self.AddResultGrid(self.name_spieler, "Ge\xe4nderter Nachname", True)
        self.AddResultGrid(self.vorname_spieler, "Ge\xe4nderter Vorname", True)
        self.AddResultGrid(self.spiele_spieler, "Ge\xe4nderte Spieleanzahl", True)
        self.AddResultGrid(self.letztes_spiel_spieler, "Ge\xe4ndertes letztes Spiel", True)
        self.AddResultGrid(self.neue_spieler, "Neue Spieler", False)
        self.AddResultGrid(self.entfernte_spieler, "Entfernte Spieler", False)

        self.propSaveButton = wx.Button(self.propPanel, ID_EXPORT_PLAYER, label='Speichere Resultat', size=(150, 25))
        self.propSizer.Add(self.propSaveButton, 0, wx.LEFT|wx.UP, 20)

        self.propSizer.Add(wx.StaticLine(self.propPanel, size=(10,5)), 0, wx.EXPAND|wx.TOP|wx.BOTTOM, 10)
        boldfont = wx.Font(10, wx.DEFAULT, wx.NORMAL, wx.BOLD)

        text1 = wx.StaticText(self.propPanel, label='\xC4nderungen')
        text1.SetFont(boldfont)
        self.propSizer.Add(text1, 0, wx.EXPAND|wx.LEFT, 20)
        self.propSizer.Add(wx.StaticLine(self.propPanel, size=(2,1)), 0, wx.EXPAND|wx.TOP|wx.BOTTOM, 3)

        box = wx.GridSizer(12, 2, 3, 3)
        self.propSizer.Add(box, 0, wx.EXPAND|wx.LEFT, 20)

        self.LogLine("")
        self.LogLine("\xC4nderungen:")
        self.DefineStatistics(box, 'TTR:', str(len(self.ttr_spieler)))
        self.DefineStatistics(box, 'Init-Art:', str(len(self.init_spieler)))
        self.DefineStatistics(box, 'Init-Wert:', str(len(self.init_wert_spieler)))
        self.DefineStatistics(box, 'Init-Datum:', str(len(self.init_datum_spieler)))
        self.DefineStatistics(box, 'Init-Gruppe:', str(len(self.init_gruppe_spieler)))
        self.DefineStatistics(box, 'Init-Position:', str(len(self.init_position_spieler)))
        self.DefineStatistics(box, 'Nachnamen:', str(len(self.name_spieler)))
        self.DefineStatistics(box, 'Vornamen:', str(len(self.vorname_spieler)))
        self.DefineStatistics(box, 'Anzahl Spiele:', str(len(self.spiele_spieler)))
        self.DefineStatistics(box, 'Datum letztes Spiel:', str(len(self.letztes_spiel_spieler)))
        self.DefineStatistics(box, 'Neue Spieler:', str(len(self.neue_spieler)))
        self.DefineStatistics(box, 'Entfernte Spieler:', str(len(self.entfernte_spieler)))
        self.LogLine("")


        self.propSizer.Add(wx.StaticLine(self.propPanel, size=(10,5)), 0, wx.EXPAND|wx.TOP|wx.BOTTOM, 10)
        headbox = wx.GridSizer(1, 3, 3, 3)
        self.propSizer.Add(headbox, 0, wx.EXPAND|wx.LEFT, 20)
        text = wx.StaticText(self.propPanel, label='Statistiken')
        text.SetFont(boldfont)
        headbox.Add(text, 0, wx.EXPAND|wx.LEFT)
        text = wx.StaticText(self.propPanel, label='TTR')
        text.SetFont(boldfont)
        headbox.Add(text, 0, wx.EXPAND|wx.RIGHT)
        text = wx.StaticText(self.propPanel, label='QTTR')
        text.SetFont(boldfont)
        headbox.Add(text, 0, wx.EXPAND|wx.RIGHT)
        self.propSizer.Add(wx.StaticLine(self.propPanel, size=(2,1)), 0, wx.EXPAND|wx.TOP|wx.BOTTOM, 3)

        statbox = wx.GridSizer(6, 3, 3, 3)
        self.propSizer.Add(statbox, 0, wx.EXPAND|wx.LEFT, 20)
        self.LogLine("Statistiken:")
        self.DefineStatistics2(statbox, 'Anzahl Spieler:', str(len(self.allplayer)), str(len(self.allplayer2)))
        self.DefineStatistics2(statbox, 'TTR-Schnitt:', str(round(float(ttr_sum)/float(len(self.allplayer)), 3)), \
                                                   str(round(float(ttr2_sum)/float(len(self.allplayer2)), 3)))
        self.DefineStatistics2(statbox, 'TTR-Modal (Wert/Anzahl):', str(Counter(ttr_list).most_common(1)[0]), \
                                                 str(Counter(ttr2_list).most_common(1)[0]))
        self.DefineStatistics2(statbox, 'TTR-Median:', str(sorted(ttr_list)[len(ttr_list)//2]), \
                                                  str(sorted(ttr2_list)[len(ttr2_list)//2]))
        self.DefineStatistics2(statbox, 'TTR-Minimum:', str(ttr_min), str(ttr2_min))
        self.DefineStatistics2(statbox, 'TTR-Maximum:', str(ttr_max), str(ttr2_max))
        self.LogLine("")

        self.LogLine("TTR-\xC4nderungen:")
        self.DefineStatistics(None, 'Zwischen 51 und ...:', str(sum(IsValidOne(x[2], 50, 9999) for x in self.ttr_spieler)))
        self.DefineStatistics(None, 'Zwischen 41 und 50 :', str(sum(IsValidOne(x[2], 40, 51) for x in self.ttr_spieler)))
        self.DefineStatistics(None, 'Zwischen 31 und 40 :', str(sum(IsValidOne(x[2], 30, 41) for x in self.ttr_spieler)))
        self.DefineStatistics(None, 'Zwischen 21 und 30 :', str(sum(IsValidOne(x[2], 20, 31) for x in self.ttr_spieler)))
        self.DefineStatistics(None, 'Zwischen 11 und 20 :', str(sum(IsValidOne(x[2], 10, 21) for x in self.ttr_spieler)))
        self.DefineStatistics(None, 'Zwischen  6 und 10 :', str(sum(IsValidOne(x[2],  5, 11) for x in self.ttr_spieler)))
        self.DefineStatistics(None, 'Zwischen  1 und  5 :', str(sum(IsValidOne(x[2],  0,  6) for x in self.ttr_spieler)))
        self.DefineStatistics(None, 'Keine \xC4nderung     :', str(len(self.allplayer) - len(self.ttr_spieler)))
        self.LogLine("")

        self.LogLine("Anzahl-Einzel-\xC4nderungen:")
        self.DefineStatistics(None, 'Zwischen 21 und ...:', str(sum(IsValid(x[12], x[21], 20, 9999) for x in self.spiele_spieler)))
        self.DefineStatistics(None, 'Zwischen 11 und 20 :', str(sum(IsValid(x[12], x[21], 10, 21) for x in self.spiele_spieler)))
        self.DefineStatistics(None, 'Zwischen  6 und 10 :', str(sum(IsValid(x[12], x[21],  5, 11) for x in self.spiele_spieler)))
        self.DefineStatistics(None, 'Zwischen  1 und  5 :', str(sum(IsValid(x[12], x[21],  0,  6) for x in self.spiele_spieler)))
        self.DefineStatistics(None, 'Keine \xC4nderung     :', str(len(self.allplayer) - len(self.spiele_spieler)))
        self.LogLine("")

        self.LogLine("Init-TTR-\xC4nderungen:")
        self.DefineStatistics(None, 'Zwischen 101 und ...  :', str(sum(IsValid(x[13], x[22], 100,  9999) for x in self.init_wert_spieler)))
        self.DefineStatistics(None, 'Zwischen  51 und 100  :', str(sum(IsValid(x[13], x[22],  50, 101) for x in self.init_wert_spieler)))
        self.DefineStatistics(None, 'Zwischen  26 und  50  :', str(sum(IsValid(x[13], x[22],  25,  51) for x in self.init_wert_spieler)))
        self.DefineStatistics(None, 'Zwischen   1 und  25  :', str(sum(IsValid(x[13], x[22],   0,  26) for x in self.init_wert_spieler)))
        if sum(x[13] == 0 for x in self.init_wert_spieler) > 0:
            self.DefineStatistics(None, ' TTR-Spieler ohne Wert:', str(sum(x[13] == 0 for x in self.init_wert_spieler)))
        if sum(x[22] == 0 for x in self.init_wert_spieler) > 0:
            self.DefineStatistics(None, 'QTTR-Spieler ohne Wert:', str(sum(x[22] == 0 for x in self.init_wert_spieler)))
        self.DefineStatistics(None, 'Keine \xC4nderung        :', str(len(self.allplayer) - len(self.init_wert_spieler)))
        self.LogLine("")

        self.propCompareButton.Disable()
        self.propPanel.Layout()
        self.DefineStatus("{0:33s} ".format("Spielervergleich: "),
                    "Laufzeit: %1.2fs" % (time.process_time() - start))

app = wx.App(redirect=False)
window = TTR_Vergleich('TTR-Vergleich')
app.MainLoop()
