import win32com.client

a = win32com.client.Dispatch('catia.application')

doc = a.ActiveDocument.Product

for i in range(doc.Products.Count):
    print('Part Number: ' + doc.Products.Item(i+1).PartNumber)
