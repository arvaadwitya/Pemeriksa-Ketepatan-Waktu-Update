import functionBase as fb

# Proses mengisi data target 
mainDataset = fb.importEmptyMainDataset()
listOfExcelFile = fb.exploreDirectory()
mainDataset = fb.fillEmptyMainDataset(mainDataset, listOfExcelFile)
mainDataset.to_excel('output.xlsx', index=False)