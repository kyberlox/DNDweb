from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse
from fastapi import FastAPI, Body
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from starlette.responses import RedirectResponse
from fastapi.staticfiles import StaticFiles

from openpyxl import load_workbook



path=""
PATH="public/imgs/location/"
wb = load_workbook('C:/Users/Александр/Desktop/DND/GameDB.xlsx')

app = FastAPI()
app.mount("/public", StaticFiles(directory="public", html=True))
#app.mount("/static", StaticFiles(directory="/", html=True))

origins = ["https://localhost", "http://127.0.0.1"]
app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["GET", "POST"],#, "OPTIONS", "DELETE", "PATH", "PUT"],
    allow_headers=["Content-Type", "Accept", "Location", "Allow", "Content-Disposition", "Sec-Fetch-Dest"],
)
 

@app.get("/")
def read_root():
    return RedirectResponse("http://127.0.0.1/public/menu.html")

@app.get("/make_person")
def make_person():
    return RedirectResponse("http://127.0.0.1/public/make_person.html")

@app.get("/game")
def make_person():
    return RedirectResponse("http://127.0.0.1/public/game.html")

@app.get("/take_quest")
def take_quest():
    return RedirectResponse("http://127.0.0.1/public/take_quest.html")



@app.get("/getMapByID/{ID}")
def getMapByID(ID: int):
    sheet = wb['Location']
    Map = []
    if ID == 0:
        path = PATH + "BackGround.jpg"
    elif ID in range(2, 14):
        path = PATH + "location2/" + sheet[f"B{ID}"].value
    elif ID in range(14, 86):
        path = PATH + "location3/" + sheet[f"B{ID}"].value
    else:
        path = PATH + "location2/" + sheet[f"C{ID}"].value
    return {"Map" : path}

@app.get("/getLocationByMapID/{ID}")
def getLocationByMapID(ID: int):
    sheet = wb['Location']
    print(ID)
    loc = []
    for i in range(2, 172):
        Id = int(sheet[f"C{i}"].value)
        
        if Id == ID:
            x = sheet[f"D{i}"].value
            y = sheet[f"E{i}"].value
            print(i, Id)
            if Id == 0:
                path = PATH + "Map.png"
            if Id in range(2, 14):
                file = PATH + "location1/" + str(sheet[f"B{i}"].value)
            elif Id in range(14, 86):
                file = PATH + "location2/" + str(sheet[f"B{i}"].value)
            elif Id in range(86, 173):
                file = PATH + "location2/" + str(sheet[f"B{i}"].value)
            
            loc.append([i, x, y, file])
        
    return {"location" : loc}