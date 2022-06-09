from pynput.keyboard import Key,Controller, Listener, KeyCode
from pynput import keyboard
from playsound import playsound
keyboard = Controller()
import time, gspread, datetime, ntplib, os
from time import ctime

sa = gspread.service_account(filename="service_account.json")
sh = sa.open("ProjetIncendie")
wks = sh.worksheet("Dashboard")

ntp_client = ntplib.NTPClient()
response = ntp_client.request('it.pool.ntp.org')
print(ctime(response.tx_time))
liste = ctime(response.tx_time).split()
temps_officiel = liste[3]
heure = temps_officiel.split(":")[0]
minute = temps_officiel.split(":")[1]
seconde = temps_officiel.split(":")[2]
print("Heure:",heure,"\nMinutes:",minute,"\nSecondes:",seconde)

time.sleep(1)
emplacement = "attente"
status = "attente"
if wks.acell('F2').value == "Hors ligne":
  wks.update('F2', "En ligne")
  wks.update('H2', ctime(response.tx_time))
  wks.update('G2', "Stand By")
  time.sleep(1)
  emplacement = 1
  status = "pris"
elif wks.acell('F3').value == "Hors ligne":
  wks.update('F3', "En ligne")
  wks.update('H3', ctime(response.tx_time))
  wks.update('G3', "Stand By")
  time.sleep(1)
  emplacement = 2
  status = "pris"
elif wks.acell('F4').value == "Hors ligne":
  wks.update('F4', "En ligne")
  wks.update('H4', ctime(response.tx_time))
  wks.update('G4', "Stand By")
  emplacement = 3
  status = "pris"
else:
  wks.update('F5', "BUG.inscription!")
  emplacement = "Bug.inscription!"
  status = "Bug.inscription!"
print("Emplacement:", emplacement)
print("Status:", status)
temps_de_fonctionnement = 0
minutes = 0
loop = 0
#print("incendie:", incendie)
#print("sonnerie:", sonnerie)

ntp_client = ntplib.NTPClient()
response = ntp_client.request('time.windows.com')
liste = ctime(response.tx_time).split()
temps_officiel = liste[3]
heure_début = float(temps_officiel.split(":")[0])
minute_début = float(temps_officiel.split(":")[1])
seconde_début = float(temps_officiel.split(":")[2])

démarrage = wks.acell('B3').value.split()
print("Heure de démarrage:",démarrage)
while True:
  time.sleep(0.5)
  seconde_début+=0.5
  if seconde_début >= 60:
    minute_début += 1
    seconde_début = 0
  if minute_début == 60:
    heure_début+=1
    minute_début = 0
  if heure_début >= float(démarrage[0]) and minute_début >= float(démarrage[1]) and seconde_début >= float(démarrage[2]):
    break
  if heure_début >= float(démarrage[0]):
    break
print("Démarrage...")
if emplacement == 1:
  wks.update('G2', "En marche")
elif emplacement == 2:
  wks.update('G3', "En marche")
elif emplacement == 3:
  wks.update('G4',  "En marche")
else:
  pass

while True: 
  shell = wks.acell('B5').value
  if shell == "1": 
     print("Alarme incendie: on")
     keyboard.press(Key.media_volume_up)
     temps_de_fonctionnement += 15
     for i in range(500):
         keyboard.press(Key.media_volume_up)
     print("Son au max!")
     playsound('Realtek-cendiTone.mp3')
     wks.update('B5', 0)
  elif shell == "2":
     print("Sonnerie: on")
     keyboard.press(Key.media_volume_up)
     temps_de_fonctionnement += 15
     for i in range(500):
         keyboard.press(Key.media_volume_up)
     print("Son au max!")
     playsound('Windows-mediSonn.mp3')
     wks.update('B5', 0)
  elif shell == "3":
     print("Siuuu !")
     keyboard.press(Key.media_volume_up)
     temps_de_fonctionnement += 1
     for i in range(500):
         keyboard.press(Key.media_volume_up)
     print("Son au max!")
     playsound('Python_startup.mp3')
     wks.update('B5', 0)
  elif shell == "4":
     print("Rossiyaaa !")
     keyboard.press(Key.media_volume_up)
     temps_de_fonctionnement += 13
     for i in range(500):
         keyboard.press(Key.media_volume_up)
     print("Son au max!")
     playsound('Windows_Startup.mp3')
     wks.update('B5', 0)
  elif shell == "5":
     print("Russe vénère")
     keyboard.press(Key.media_volume_up)
     temps_de_fonctionnement += 8
     for i in range(500):
         keyboard.press(Key.media_volume_up)
     print("Son au max!")
     playsound('Cortana.mp3')
     wks.update('B5', 0)
  else:
    pass
  if shell == "10":
     temps_de_fonctionnement += 7200
     print("Terminating...")
     time.sleep(1)   
  if temps_de_fonctionnement >= 7200:
     if emplacement == 1:
       wks.update('F2', "Hors ligne")
       wks.update('G2', "Off")
       # os.remove("Incendie.mp3")
       # os.remove("Sonnerie.mp3")
       # os.remove("Token.json")
       time.sleep(1)
       break      
     elif emplacement == 2:
       wks.update('F3', "Hors ligne")
       wks.update('G3', "Off")
       time.sleep(1)
       break       
     elif emplacement == 3:
       wks.update('F4', "Hors ligne")
       wks.update('G4', "Off")
       time.sleep(1)
       break
     else:
       break
  time.sleep(10)
  temps_de_fonctionnement += 10
  minutes = temps_de_fonctionnement//60
  #print("En cours depuis:", temps_de_fonctionnement, "secondes")
  #print("Soit:", minutes, "minute(s)")