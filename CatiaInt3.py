import win32com.client
catapp = win32com.client.Dispatch("CATIA.Application")
rootProd = catapp.ActiveDocument.Product

'''
for i in range(rootProd.Products.Count):
    rootProd.Products.Item(i+1).PartNumber = "pyPart_numbero_" + str(10+i)
    # rootProd.Products.Item(i+1). = ""
'''

pyPart = catapp.Documents.Item(4).Part
hb1 = pyPart.HybridBodies.Add()
pyPart.Update()

hsf = pyPart.HybridShapeFactory
for i in range(10):
    point = hsf.AddNewPointCoord(i*10, 0.0, 0.0)
    hb1.AppendHybridShape(point)
pyPart.Update()


'''
documents1 = catapp.Documents
partDocument1 = documents1.Item(4)
part1 = partDocument1.Part

hybridBodies1 = part1.HybridBodies
hybridBody1 = hybridBodies1.Add()
part1.Update()

hsFactory1 = part1.HybridShapeFactory
hsPointCoord1 = hsFactory1.AddNewPointCoord(254.0, 254.0, 254.0)
hybridBody1.AppendHybridShape(hsPointCoord1)
part1.InWorkObject = hsPointCoord1
part1.Update()
'''