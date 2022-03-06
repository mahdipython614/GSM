import requests
import re
from cefpython3 import cefpython as cef
import pandas

try:
    import tkinter as tk
except ImportError:
    import Tkinter as tk
import sys
from tkinter.filedialog import askopenfilename
import platform
import logging as _logging

# Fix for PyCharm hints warnings
WindowUtils = cef.WindowUtils()

# Platforms
WINDOWS = (platform.system() == "Windows")
LINUX = (platform.system() == "Linux")
MAC = (platform.system() == "Darwin")

# Globals
logger = _logging.getLogger("tkinter_.py")


class MainFrame(tk.Frame):

    def __init__(self, root):
        self.browser_frame = None
        self.navigation_bar = None

        # Root
        pad = 3
        self._geom = '200x200+0+0'
        root.geometry("{0}x{1}+0+0".format(
            root.winfo_screenwidth() - pad, root.winfo_screenheight() - pad))
        root.bind('<Escape>', self.toggle_geom)

        tk.Grid.rowconfigure(root, 0, weight=1)
        tk.Grid.columnconfigure(root, 0, weight=1)

        # MainFrame
        tk.Frame.__init__(self, root)
        self.master.title("GSM")

        # NavigationBar
        self.navigation_bar = NavigationBar(self)
        self.navigation_bar.grid(row=0, column=0,
                                 sticky=(tk.N + tk.S + tk.E + tk.W))
        tk.Grid.rowconfigure(self, 0, weight=0)
        tk.Grid.columnconfigure(self, 0, weight=0)

        # BrowserFrame
        self.browser_frame = BrowserFrame(self, self.navigation_bar)
        self.browser_frame.grid(row=1, column=0,
                                sticky=(tk.N + tk.S + tk.E + tk.W))
        tk.Grid.rowconfigure(self, 1, weight=1)
        tk.Grid.columnconfigure(self, 0, weight=1)

        # Pack MainFrame
        self.pack(fill=tk.BOTH, expand=tk.YES)

    def toggle_geom(self, event):
        geom = self.master.winfo_geometry()
        print(geom, self._geom)
        self.master.geometry(self._geom)
        self._geom = geom

    def get_browser(self):
        if self.browser_frame:
            return self.browser_frame.browser
        return None

    def _on_focus_out(self, _):
        logger.debug("MainFrame.on_focus_out")
        self.master.master.focus_force()


class BrowserFrame(tk.Frame):

    def __init__(self, master, navigation_bar=None):
        self.navigation_bar = navigation_bar
        self.closing = False
        self.browser = None
        tk.Frame.__init__(self, master)
        self.bind("<Configure>", self.on_configure)

    def embed_browser(self):
        window_info = cef.WindowInfo()
        rect = [0, 0, self.winfo_width(), self.winfo_height()]
        window_info.SetAsChild(self.get_window_handle(), rect)
        self.browser = cef.CreateBrowserSync(window_info,
                                             url="file:///mymap.html")
        assert self.browser
        self.message_loop_work()

    def get_window_handle(self):
        if self.winfo_id() > 0:
            return self.winfo_id()
        elif MAC:
            # On Mac window id is an invalid negative value (Issue #308).
            # This is kind of a dirty hack to get window handle using
            # PyObjC package. If you change structure of windows then you
            # need to do modifications here as well.
            # noinspection PyUnresolvedReferences
            from AppKit import NSApp
            # noinspection PyUnresolvedReferences
            import objc
            # Sometimes there is more than one window, when application
            # didn't close cleanly last time Python displays an NSAlert
            # window asking whether to Reopen that window.
            # noinspection PyUnresolvedReferences
            return objc.pyobjc_id(NSApp.windows()[-1].contentView())
        else:
            raise Exception("Couldn't obtain window handle")

    def message_loop_work(self):
        cef.MessageLoopWork()
        self.after(10, self.message_loop_work)

    def on_configure(self, _):
        if not self.browser:
            self.embed_browser()


class NavigationBar(tk.Frame):
    def __init__(self, master):
        tk.Frame.__init__(self, master)
        self.htmlcode = '''<!DOCTYPE html>
                <html lang="en">
                <head>
                    <meta charset="UTF-8">
                    <title>map</title>
                    <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
                    <link href="https://static.neshan.org/sdk/leaflet/1.4.0/leaflet.css" rel="stylesheet" type="text/css">
                    <script src="https://static.neshan.org/sdk/leaflet/1.4.0/leaflet.js" type="text/javascript"></script>
                </head>
                <body>
                <div id="map" style="width: 1340px; height: 700px; background: #eee; border: 2px solid #aaa;"></div>
                <script type="text/javascript">
                    var myMap = new L.Map('map', {
                        key: 'web.36brwPxDnkFUqnfEXYWIbDM9gKN96y39FfBb48Pr',
                        maptype: 'dreamy',
                        poi: true,
                        traffic: false,
                        center: [32.4971364,54.0498515],
                        zoom: 5.75
                    });
                '''
        self.mnc_check = {}
        self.mnc = ['IR-MCI', 'TKC', 'MTCE', 'Taliya', 'Irancell', 'TCI', 'Iraphone']
        self.selected = tk.StringVar()
        self.selected.set('IR-MCI')
        self.code_mnc = {'IR-MCI': 11, 'TKC': 14, 'MTCE': 19, 'Taliya': 32, 'Irancell': 35, 'TCI': 70, 'Iraphone': 93}
        self.listlocations = []
        self.count_marker = 0
        self.add_menubar()

    def add_menubar(self):
        menubar = tk.Menu(root)
        filemenu = tk.Menu(menubar, tearoff=0)
        filemenu.add_command(label="Open...", command=self.openfile)
        filemenu.add_separator()
        filemenu.add_command(label="Exit", command=quit)
        menubar.add_cascade(label="File", menu=filemenu)

        helpmenu = tk.Menu(menubar, tearoff=0)

        for mnc in self.mnc:
            helpmenu.add_radiobutton(label=mnc, value=mnc, variable=self.selected)
        menubar.add_cascade(label="MNC", menu=helpmenu)
        root.config(menu=menubar)

    def openfile(self):
        filename = askopenfilename()  # show an "Open" dialog box and return the path to the selected file
        if filename != '':
            self.excelread(filename)
        for data in self.listdata:
            url_response = requests.get(
                'https://cellphonetrackers.org/gsm/classes/Cell.Search.php?mcc=432&mnc={0}&lac={1}&cid={2}'.format(
                    self.code_mnc[self.selected.get()], data[1], data[0]))
            self.lat = re.findall('(?<=Lat=)(.*)(?= Lon)', url_response.content.decode())
            self.lon = re.findall('(?<=Lon=)(.*)(?=</a><br/>)', url_response.content.decode())
            if len(self.lat)==0 or len(self.lon)==0:
                print('Sorry, but cell tower not found LAC: {0},CID: {1}.'.format(data[1],data[0]))
                continue
            self.listlocations.append(self.lat.extend(self.lon))
            information = r"MCC:432<br>MNC:{0}<br>LAC:{1}<br>CID:{2}<br><br>lat:{3}<br>lon:{4}".format(
                self.code_mnc[self.selected.get()], data[1], data[0], self.lat[0], self.lon[0])
            self.add_marker(self.lat, self.lon, information)
        self.writehtml()
        self.reload()

    def excelread(self, filename):
        fd = pandas.read_excel(filename.replace('\ '[0], r'\\'))
        self.listdata = fd.to_records(index=False, )

    def add_marker(self, lat, lon, info):
        self.count_marker += 1
        self.htmlcode += '''\tvar marker{3} = new L.marker([{0},{1}]).addTo(myMap);
                    marker{3}.bindPopup('{2}').openPopup();\n\t\t\t\t'''.format(lat[0], lon[0], info, self.count_marker)

    def writehtml(self):
        self.htmlcode += '''    
                        </script>
                        </body>
                        </html>'''
        htmlfile = open('mymap.html', 'w')
        htmlfile.truncate(0)
        htmlfile.write(self.htmlcode)
        htmlfile.close()

    def reload(self):
        self.master.get_browser().Reload()

    def on_button1(self, _):
        """Fix CEF focus issues (#255). See also FocusHandler.OnGotFocus."""
        logger.debug("NavigationBar.on_button1")
        self.master.master.focus_force()


if __name__ == '__main__':
    logger.setLevel(_logging.INFO)

    stream_handler = _logging.StreamHandler()
    formatter = _logging.Formatter("[%(filename)s] %(message)s")
    stream_handler.setFormatter(formatter)
    logger.addHandler(stream_handler)
    logger.info("CEF Python {ver}".format(ver=cef.__version__))
    logger.info("Python {ver} {arch}".format(
        ver=platform.python_version(), arch=platform.architecture()[0]))
    logger.info("Tk {ver}".format(ver=tk.Tcl().eval('info patchlevel')))
    assert cef.__version__ >= "55.3", "CEF Python v55.3+ required to run this"
    sys.excepthook = cef.ExceptHook  # To shutdown all CEF processes on error
    root = tk.Tk()
    app = MainFrame(root)
    # Tk must be initialized before CEF otherwise fatal error (Issue #306)
    cef.Initialize()

    app.mainloop()
    cef.Shutdown()
