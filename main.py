from excel import Excel

# create excel obj to translate
ex = Excel('latin.xlsx', to_latin=False)

# translate sheets one by one
# for sheet in ex.get_sheets_list():
#     ex.set_sheet(sheet)
#     ex.translate_sheet()

# translate all sheets at once
ex.translate_all_sheets()

# save file with new file name or old name (to save with old name use ex.save())
ex.save('cyrillic.xlsx')
