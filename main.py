import os
import comtypes.client

wdFormatPDF = 17
input_path = 'Published/'
output_path = 'pdfs/'

for in_file in os.listdir(input_path):
    if ((in_file.endswith('.docx') or in_file.endswith('.mht')) and not in_file.startswith('~$')):
        print('Converting', in_file, '...')
        word = comtypes.client.CreateObject('Word.Application')
        doc = word.Documents.Open(os.path.abspath(input_path + in_file))

        out_file = in_file.split(".")
        out_file[-1] = "pdf"
        out_file = ".".join(out_file)

        doc.SaveAs(os.path.abspath(output_path + out_file), FileFormat=wdFormatPDF)
        doc.Close()
        word.Quit()
        print(out_file, 'is done!')
