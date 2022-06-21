import os
import glob

cwd = os.getcwd()

for file in glob.glob(cwd+ "/**/750-*xlsx", recursive=True):

    if file.count("sng") == 1:
        # print('check complete')
        # print(file)
        splitNo = 4
        # pdfpath = file.rsplit('\\', 3)
    else:
        # pdfpath = file.rsplit('\\', 2)
        splitNo = 2

    pdfpath = file.rsplit('\\', splitNo)

# print(temp)
# print(cwd)
    print(pdfpath)


