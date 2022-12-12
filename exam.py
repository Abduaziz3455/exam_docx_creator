import datetime
from pathlib import Path
import random
from docx import Document

import PySimpleGUI as sg
from docxtpl import DocxTemplate
sg.theme('DarkAmber')

# document = Document()
# for section in document.sections:
#     print(section.start_type)

document_path = Path(__file__).parent / "misol.docx"
doc = DocxTemplate(document_path)

today = datetime.datetime.today()
today_in_one_week = today + datetime.timedelta(days=7)


def make_win2():
   layout = []
   for i in range(1, int(values1['SON'])+1):
      i=str(i)
      if int(i)<15:
         layout.append([sg.Text(f"{i}-savolni kiriting:", size=(15,1), font='Montserrat 12', text_color='yellow'), sg.Input(key=f"savol{i}", do_not_clear=False)])
      else:
         layout.append([sg.Text(f"{i}-savolni kiriting:", size=(15,1), font='Montserrat 12', text_color='yellow'), sg.Input(key=f"savol{i}", do_not_clear=False)])
   layout.append([sg.Stretch(), sg.Button("Bilet yaratish", font='Montserrat 19'), sg.Exit(
            "Chiqish", font='Montserrat 17', button_color="red")])

   window = sg.Window('Bilet Yaratuvchi dastur', layout, location=(250,100), size=(1000,700), finalize=True)
   return window

def make_win1():
   layout = [
      [sg.Text("Fakultet nomi:", font='Montserrat 16', text_color='yellow'),
      sg.Input(key="FAKULTET", do_not_clear=False)],
      [sg.Text("Fan nomi:", font='Montserrat 16', text_color='yellow'),
      sg.Input(key="FAN", do_not_clear=False)],
      [sg.Text("Nechta savol kiritmoqchisiz:", font='Montserrat 16',
               text_color='yellow'), sg.Input(key="SON", do_not_clear=False)],
      [sg.Text("Tuzdi:", font='Montserrat 16', text_color='yellow'),
      sg.Input(key="TUZDI", do_not_clear=False)],
      [sg.Text("Kafedra mudiri:", font='Montserrat 16', text_color='yellow'),
      sg.Input(key="MUDIR", do_not_clear=False)],
      [sg.Button("Savollarni kiritish", font='Montserrat 16', button_color='red')],
      [sg.Push(), sg.Text(text="\n")],
      [sg.Push(),sg.Text(text="\n")],
      [sg.Push(),sg.Text(text="")],
      [sg.Stretch(), sg.Text('by Abduaziz and Nodirshox')],
   ]
   window = sg.Window('Bilet Yaratuvchi dastur', layout, location=(450,200), size=(600,400), element_justification="right", finalize=True)

   return window

window1, window2 = make_win1(), None
values1 = {}

while True:
   window, event, values = sg.read_all_windows()
   values1 = values1 | values # merge two dictionaries with |

   if event == sg.WIN_CLOSED or event == 'Chiqish':
      window.close()
      if window == window2:       # if closing win 2, mark as closed
         window2 = None
      elif window == window1:     # if closing win 1, exit program
         break

   elif event == 'Savollarni kiritish' and not window2:
      window2 = make_win2()

   if event == "Bilet yaratish":
      window2.close()
      # Add fields to our dict
      # values1["TODAY"] = today.strftime("%d-%m-%Y")
      # values1["TODAY_IN_ONE_WEEK"] = today_in_one_week.strftime("%d-%m-%Y")
      t=0
      while t<10:
         t+=1
         values1["savol1"], values1["savol2"], values1["savol3"] = random.sample(list(values1.values())[5:], 3)
         print(values1)
         # Render the template, save new word document & inform user
         for section in doc.:
            print(section.start_type)

         {%p for page in pages %}
            {{page}}
         {%p endfor %}


         doc.render(values1)
         output_path = Path(__file__).parent / \
            f"{values1['FAN'].upper()}-savollari.docx"
         doc.save(output_path)
      sg.popup("Imtihon savollari kiritildi",
               f"Shu yerda saqlandi: {output_path}", custom_text="Saqlandi", auto_close=True, auto_close_duration=5)

      window1.close()
      break


window.close()
