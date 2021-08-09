import functionBase as fb

# Proses kategorisasi SLA
mainDataset = fb.importFilledMainDataset()
mainDataset = mainDataset.astype(str)
mainDataset = fb.slaCategorization(mainDataset)
mainDataset.to_excel('output.xlsx', index=False)