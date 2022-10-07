import xbmc
import xbmcgui
import xbmcplugin
import os, sys, linecache, json
import xbmcaddon
import xbmcvfs
import csv  
from datetime import datetime

addon = xbmcaddon.Addon()
addon_path = addon.getAddonInfo("path")
addon_icon = addon_path + '/resources/icon.png'

def settings(setting, value = None):
    # Get or add addon setting
    if value:
        xbmcaddon.Addon().setSetting(setting, value)
    else:
        return xbmcaddon.Addon().getSetting(setting)   

def translate(text):
    return xbmcaddon.Addon().getLocalizedString(text)

def get_installedversion():
    # retrieve current installed version
    json_query = xbmc.executeJSONRPC('{ "jsonrpc": "2.0", "method": "Application.GetProperties", "params": {"properties": ["version", "name"]}, "id": 1 }')
    json_query = json.loads(json_query)
    version_installed = []
    if 'result' in json_query and 'version' in json_query['result']:
        version_installed  = json_query['result']['version']['major']
    return str(version_installed)

def getDatabaseName():
    installed_version = get_installedversion()
    if installed_version == '10':
        return "MyVideos37.db"
    elif installed_version == '11':
        return "MyVideos60.db"
    elif installed_version == '12':
        return "MyVideos75.db"
    elif installed_version == '13':
        return "MyVideos78.db"
    elif installed_version == '14':
        return "MyVideos90.db"
    elif installed_version == '15':
        return "MyVideos93.db"
    elif installed_version == '16':
        return "MyVideos99.db"
    elif installed_version == '17':
        return "MyVideos107.db"
    elif installed_version == '18':
        return "MyVideos116.db"
    elif installed_version == '19':
        return "MyVideos119.db"
    elif installed_version == '20':
        return "MyVideos121.db"
       
    return ""  

def getmuDatabaseName():
    installed_version = get_installedversion()
    if installed_version == '10':
        return "MyMusic7.db"
    elif installed_version == '11':
        return "MyMusic18.db"
    elif installed_version == '12':
        return "MyMusic32.db"
    elif installed_version == '13':
        return "MyMusic46.db"
    elif installed_version == '14':
        return "MyMusic48.db"
    elif installed_version == '15':
        return "MyMusic52.db"
    elif installed_version == '16':
        return "MyMusic56.db"
    elif installed_version == '17':
        return "MyMusic60.db"
    elif installed_version == '18':
        return "MyMusic72.db"
    elif installed_version == '19':
        return "MyMusic82.db"
    elif installed_version == '20':
        return "MyMusic82.db"
       
    return ""  

def openKodiDB():                                   #  Open Kodi database
    try:
        from sqlite3 import dbapi2 as sqlite
    except:
        from pysqlite2 import dbapi2 as sqlite
                      
    DB = os.path.join(xbmcvfs.translatePath("special://database"), getDatabaseName())
    db = sqlite.connect(DB)

    return(db)    

def openKodiMuDB():                                   #  Open Kodi music database
    try:
        from sqlite3 import dbapi2 as sqlite
    except:
        from pysqlite2 import dbapi2 as sqlite
                      
    DB = os.path.join(xbmcvfs.translatePath("special://database"), getmuDatabaseName())
    db = sqlite.connect(DB)

    return(db)

def exportData(selectbl):                                        # CSV Output selected table

    try:

        #xbmc.log("Kodi selectable is: " +  str(selectbl), xbmc.LOGNINFO)

        folderpath = xbmcvfs.translatePath(os.path.join("special://home/", "output/"))
        if not xbmcvfs.exists(folderpath):
            xbmcvfs.mkdir(folderpath)
            xbmc.log("Kodi Export Output folder not found: " +  str(folderpath), xbmc.LOGINFO)

        for a in range(len(selectbl)):
            fpart = datetime.now().strftime('%H%M%S')
            selectindex = int(selectbl[a][:2])                   # Get list index to determine DB
            selectname = selectbl[a][2:]                         # Parse table name in DB

            #xbmc.log("Mezzmo selectable is: " +  str(selectindex) + ' ' + selectname, xbmc.LOGINFO)
            if selectindex < 11:
                dbexport = openKodiDB()
                dbase = 'videos_'
            else:
                dbexport = openKodiMuDB()
                dbase = 'music_'                


            outfile = folderpath + "kodi_" + dbase + selectname + "_" + fpart + ".csv"
            curm = dbexport.execute('SELECT * FROM '+selectname+'')
            recs = curm.fetchall()
            headers = [i[0] for i in curm.description]
            csvFile = csv.writer(open(outfile, 'w', encoding='utf-8'),
                             delimiter=',', lineterminator='\n',
                             quoting=csv.QUOTE_ALL, escapechar='\\')

            csvFile.writerow(headers)                       # Add the headers and data to the CSV file.
            for row in recs:
                recsencode = []
                for item in range(len(row)):
                    if isinstance(row[item], int) or isinstance(row[item], float):  # Convert to strings
                        rectemp = str(row[item])
                        try:
                            recitem = rectemp.decode('utf-8')
                        except:
                            recitem = rectemp
                    else:
                        rectemp = row[item]
                        try:
                            recitem = rectemp.decode('utf-8')
                        except:
                            recitem = rectemp
                    recsencode.append(recitem) 
                csvFile.writerow(recsencode)                
            dbexport.close()

        outmsg = folderpath
        dialog_text = translate(30300) + outmsg 
        xbmcgui.Dialog().ok(translate(30301), dialog_text)

    except Exception as e:
        printexception()
        dbexport.close()
        mgenlog = translate(30302) + selectname
        xbmcgui.Dialog().notification(translate(30303), mgenlog, addon_icon, 5000)            
        xbmc.log(mgenlog, xbmc.LOGINFO)

def printexception():

    exc_type, exc_obj, tb = sys.exc_info()
    f = tb.tb_frame
    lineno = tb.tb_lineno
    filename = f.f_code.co_filename
    linecache.checkcache(filename)
    line = linecache.getline(filename, lineno, f.f_globals)
    xbmc.log( 'EXCEPTION IN ({0}, LINE {1} "{2}"): {3}'.format(filename, lineno, line.strip(),     \
    exc_obj), xbmc.LOGINFO)

def selectExport():                                            # Select table to export

    try:
        while True:
            stable = []
            selectbl = []
            tables = ["Kodi Video DB - Actors","Kodi Video DB - Episodes","Kodi Video DB - Movies",              \
            "Kodi Video DB - TV Shows","Kodi Video DB - Artwork","Kodi Video DB - Path","Kodi Video DB - Files", \
            "Kodi Video DB - Streamdetails", "Kodi Video DB - Episode View", "Kodi Video DB - Movie View",       \
            "Kodi Video DB - Music Video View", "Kodi Music DB - Artist","Kodi Music DB - Album Artist View",    \
            "Kodi Music DB - Album View ","Kodi Music DB - Artist View", "Kodi Music DB - Song",                 \
            "Kodi Music DB - Song Artist View","Kodi Music DB - Song View"]
            ddialog = xbmcgui.Dialog()    
            stable = ddialog.multiselect(translate(30304), tables)
            if stable == None:                                 # User cancel
                break
            if 0 in stable:
                selectbl.append('00actor')
            if 1 in stable:
               selectbl.append('01episode')   
            if 2 in stable:
                selectbl.append('02movie')
            if 3 in stable:
                selectbl.append('03tvshow')
            if 4 in stable:
                selectbl.append('04art')    
            if 5 in stable:
                selectbl.append('05path')  
            if 6 in stable:
                selectbl.append('06files')  
            if 7 in stable:
                selectbl.append('07streamdetails')
            if 8 in stable:
                selectbl.append('08episode_view')
            if 9 in stable:
                selectbl.append('09movie_view')
            if 10 in stable:
                selectbl.append('10musicvideo_view')
            if 11 in stable:
                selectbl.append('11artist')
            if 12 in stable:
                selectbl.append('12albumartistview')
            if 13 in stable:
                selectbl.append('13albumview')
            if 14 in stable:
                selectbl.append('14artistview')
            if 15 in stable:
                selectbl.append('15song')
            if 16 in stable:
                selectbl.append('16songartistview')
            if 17 in stable:
                selectbl.append('17songview')
            if 18 in stable:
                selectbl.append('18mServers')
            exportData(selectbl)         

    except Exception as e:
        printexception()


selectExport()

