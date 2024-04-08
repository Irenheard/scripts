import os
import comtypes.client

def conversionFunction(x, y):
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1

    for filename in os.listdir(x):
        if filename.endswith(".pptx"):
            print("{0} found!!".format(filename))
            input_pptx_path = os.path.join(x, filename)
            output_pdf_path = os.path.join(y, os.path.splitext(filename)[0] + ".pdf")
            print("converting to .pdf")
            ppt = powerpoint.Presentations.Open(input_pptx_path)
            ppt.SaveAs(output_pdf_path, 32)
            ppt.Close()
    
    powerpoint.Quit()

inputFolder = r"D:\GIASemester2\GEOM73\Final Exam\pptx"
outputFolder = r"D:\GIASemester2\GEOM73\Final Exam"
conversionFunction(inputFolder, outputFolder)

print("\nslay")
