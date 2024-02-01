import os.path, gspread,yagmail
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from datetime import datetime, timedelta,date
from decouple import config


def get_calendar_service():    
    # If modifying these scopes, delete the file tokencalendar.json.
    SCOPES = ['https://www.googleapis.com/auth/calendar']
    creds = None
    # The file tokencalendar.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first time.
    if os.path.exists('tokencalendar.json'):
        creds = Credentials.from_authorized_user_file('tokencalendar.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('CredentialsCalendarioTutorias.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('tokencalendar.json', 'w') as token:
            token.write(creds.to_json())

    
    service = build('calendar', 'v3', credentials=creds)
    return service


def get_event(service,possible_date,possible_time):
    if type(possible_time)==str:
        day=int(possible_date.split("/")[0])
        month=int(possible_date.split("/")[1])
        year=int(possible_date.split("/")[2])
        hours=int(possible_time.split(":")[0])
        minutes=int(possible_time.split(":")[1])
        date_start=datetime(year,month,day,hours,minutes)
        date_end=date_start + timedelta(minutes=30)
        iso_date_start=date_start.isoformat()+"-05:00"
        iso_date_end=date_end.isoformat()+"-05:00"

        events_result = service.events().list(calendarId='primary', timeMin=iso_date_start, timeMax=iso_date_end).execute()
        event = events_result.get('items',"")
        if event:
            return event
        else:
            return (possible_date,possible_time)
    else:
        for time in possible_time:
            day=int(possible_date.split("/")[0])
            month=int(possible_date.split("/")[1])
            year=int(possible_date.split("/")[2])
            hours=int(time.split(":")[0])
            minutes=int(time.split(":")[1])
            date_start=datetime(year,month,day,hours,minutes)
            date_end=date_start + timedelta(minutes=30)
            iso_date_start=date_start.isoformat()+"-05:00"
            iso_date_end=date_end.isoformat()+"-05:00"

            events_result = service.events().list(calendarId='primary', timeMin=iso_date_start, timeMax=iso_date_end).execute()
            event = events_result.get('items',"")
            if event:
                continue
            else:
                return (possible_date,time)
        return event
    
def create_event(service,email,name,possible_date,possible_time):
    day=int(possible_date.split("/")[0])
    month=int(possible_date.split("/")[1])
    year=int(possible_date.split("/")[2])
    hours=int(possible_time.split(":")[0])
    minutes=int(possible_time.split(":")[1])
    date_start=datetime(year,month,day,hours,minutes)
    date_end=date_start + timedelta(minutes=30)
    iso_date_start=date_start.isoformat() 
    iso_date_end=date_end.isoformat() 
    info={
        "summary":f"Tutoria {name}",
        "description":"Nos vemos este dia en el salon 11B",
        "start":{"dateTime":iso_date_start,"timeZone":"America/Bogota"},
        "end":{"dateTime":iso_date_end,"timeZone":"America/Bogota"},
        "attendees":[{"email":email}]
        }
    event=service.events().insert(calendarId="primary",sendUpdates="all",body=info).execute()
    return info


def get_date(lugar,fecha):    

    if lugar=="colegio":
        sa=gspread.service_account(filename="C:/Users/gquintero/Desktop/python/pythontest-361720-6e048b13eddb.json")
    else:
        sa=gspread.service_account(filename="C:/Users/PC/Desktop/python/pythontest-361720-6e048b13eddb.json")

    google_sheet=sa.open("intento de automatizacion")
    hoja=google_sheet.worksheet("Copia de Third Term 2022")
    if fecha[0]=="0":
        fecha=fecha[1:]
    cell=hoja.find(fecha)
    info=hoja.cell(cell.row+1,cell.col).value
    col=cell.col
    i=0

    tutoring={"Day 1":["10:40","13:40"],"Day 2":"14:10","Day 3":"14:10","Day 4":"13:40","Day 5":"14:10"}
    while info not in tutoring:
        col=get_next([4,5,6,7,8],col)
        if col==4:
            i+=1            
        info=hoja.cell(cell.row+(3*i+1),col).value
    
    possible_date=hoja.cell(cell.row+(3*i),col).value
    possible_time=tutoring.get(info)
    return (possible_date,possible_time)


def get_next(list,value):
    try:
        index=list.index(value)
        return list[index+1]
    except (IndexError,ValueError) as e:
        return list[0]


def next_day_date(fecha):
    first_date=date.fromisoformat(fecha.split("T")[0])
    next_date=first_date + timedelta(days=1)
    return next_date.strftime("%d/%m/%Y")


def send_email(email,name,possible_date,possible_time):
    yag = yagmail.SMTP("gquintero@colegionuevayork.edu.co",config("gquintero_mail_password"))

    body ='''hola {nombre}, te queda agendada la tutoría para el {fecha} a las {hora}.  Nos vemos en el salon 11B.

Cordialmente,

Gustavo Orlando Quintero Quintero
Docente Matemáticas
Colegio Nueva York
Tel:601684890 ext. 134.'''
    sub="Tutoría Matemáticas"
    yag.send(to=email,subject=sub,contents=body.format(nombre=name,fecha=possible_date,hora=possible_time))


def main():
    lugar="colegio"

    email=input("Correo: ")
    name=input("Nombre: ")
    
    fecha=(date.today()+timedelta(days=1)).strftime("%d/%m/%Y")
    possible_date,possible_time=get_date(lugar,fecha)
    service=get_calendar_service()
    event=get_event(service,possible_date,possible_time)

    while type(event)!=tuple:
        start = event[0]['start'].get('dateTime', event[0]['start'].get('date'))
        fecha=next_day_date(start)
        possible_date,possible_time=get_date(lugar,fecha)
        event=get_event(service,possible_date,possible_time)

    possible_date,possible_time=event
    
    info=create_event(service,email,name,possible_date,possible_time)
    send_email(email,name,possible_date,possible_time)
    print(f"{info['summary']} fue agendada para el dia {possible_date} a las {possible_time}")

    


if __name__ == '__main__':
    main()