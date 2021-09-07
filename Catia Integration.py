import win32com.client
catapp = win32com.client.Dispatch("CATIA.Application")
rootProd = catapp.ActiveDocument.Product

for i in range(rootProd.Products.Count):
    rootProd.Products.Item(i+1).PartNumber = "pyPart_numbero_" + str(10+i)
    # rootProd.Products.Item(i+1). = ""

documents1 = catapp.Documents
partDocument1 = documents1.Item("pyPart_numbero_10.CATPart")
part1 = partDocument1.Part
hsFactory1 = part1.HybridShapeFactory
hsPointCoord1 = hsFactory1.AddNewPointCoord(254.0, 254.0, 254.0)
hybridBodies1 = part1.HybridBodies
hybridBody1 = hybridBodies1.Item("Geometry")
hybridBody1.AppendHybridShape(hsPointCoord1)
part1.InWorkObject = hsPointCoord1
part1.Update()
