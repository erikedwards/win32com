import win32com.client
import math

# Constants
SCREEN_WIDTH = 5000
SCREEN_HEIGHT = 1000
size = (SCREEN_WIDTH, SCREEN_HEIGHT)

# init visualization values
circX = int(SCREEN_WIDTH / 6)
circY = int(SCREEN_HEIGHT / 2)
circScale = 250
wave = []
# waveX = int(SCREEN_WIDTH * (2/4))
waveX = circX
# waveY = int(SCREEN_HEIGHT / 2)
time = 0.00
timeStep = 0.005

# init Catia
catapp = win32com.client.Dispatch("CATIA.Application")
rootProd = catapp.ActiveDocument.Product

# grab reference to a catPart (shown here as 4th object, 3rd part) and add a new GeoSet
pyPart = catapp.Documents.Item(2).Part
hb1 = pyPart.HybridBodies.Add()
pyPart.Update()

# create a hybrid shape factory
hsf = pyPart.HybridShapeFactory

# max "N" to be evaluated. Not the same as number of terms to be evaluated
# maxN = 5, terms = 1, 3, and 5. Num terms = 3
# maxN = 11, terms = 1, 3, 5, 7, 9, 11. Num terms = 6
# num terms = (1/2)*(maxN + 1)
maxN = 1

# set frame rate
FR = 60


# --- draw axis system
rad = int(circScale * (4 / (1 * math.pi)))
pointYPos = hsf.AddNewPointCoord(waveX, circY + rad, 0)
pointYNeg = hsf.AddNewPointCoord(waveX, circY - rad, 0)
pointXNeg = hsf.AddNewPointCoord(waveX - 50, circY, 0)
pointXPos = hsf.AddNewPointCoord(SCREEN_WIDTH, circY, 0)
xAxisLine = hsf.AddNewLinePtPt(pyPart.CreateReferenceFromObject(pointXNeg),
                               pyPart.CreateReferenceFromObject(pointXPos))
yAxisLine = hsf.AddNewLinePtPt(pyPart.CreateReferenceFromObject(pointYNeg),
                               pyPart.CreateReferenceFromObject(pointYPos))
hb1.AppendHybridShape(xAxisLine)
hb1.AppendHybridShape(yAxisLine)
pyPart.Update()

# Main Loop
run = True

while run:

    # Game logic
    time += timeStep

    x = circX
    y = circY
    for n in range(1, maxN +1, 2):
        prevX = x
        prevY = y
        rad = int(circScale * ( 4 / (n * math.pi)))
        x += int(rad * math.cos(n * time))
        y += int(-1 * rad * math.sin(n * time))
        # --- draw circle
        # pygame.draw.circle(screen, BLACK, [prevX, prevY], rad, 1)
        # pygame.draw.line(screen, RED, [prevX, prevY], [x, y])
        #TODO --- add circle and line to Catia geoset "geom_circularMotion".



    # Add y value of smallest circle to wave
    wave.insert(0, y)
    if len(wave) > SCREEN_WIDTH:
        wave.pop(-1)

    # --- draw wave
    for i in range(0, len(wave)):
        pX = int(waveX + i)
        pY = int(wave[i])
        # pygame.draw.circle(screen, GREEN, [pX, pY], 1)
        #TODO --- add point to Catia geoset "geom_linearMotion-Sin(theta)"
        hb1.AppendHybridShape(hsf.AddNewPointCoord(pX, pY, 0))

    pyPart.Update()

    # --- draw connecting level line
    # pygame.draw.line(screen, GREY, [x, y], [waveX, y])


    # --- Update the screen
    # pygame.display.flip()

    # --- Limit to 60 frames per second
    # clock.tick(FR)


# ----------------------------------------------------------------
# ----------------------------------------------------------------
# ----------------------------------------------------------------


for i in range(10):
    point = hsf.AddNewPointCoord(i*10, 0.0, 0.0)
    hb1.AppendHybridShape(point)
    # pyPart.InWorkObject = point
    # pyPart.Update()
pyPart.Update()


