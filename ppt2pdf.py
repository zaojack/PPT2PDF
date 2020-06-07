import comtypes.client
import os

def init_powerpoint():
   powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
   powerpoint.Visible = 1
   return powerpoint

def ppt_to_pdf(powerpoint, inputFileName, outputFileName, formatType = 32):
   if outputFileName[-3:] != 'pdf':
       outputFileName = outputFileName + ".pdf"
   deck = powerpoint.Presentations.Open(inputFileName)
   deck.SaveAs(outputFileName, formatType) # formatType = 32 for ppt to pdf
   deck.Close()

def convert_files_in_folder(powerpoint, folder):
   files = os.listdir(folder)
   pptfiles = [f for f in files if f.endswith((".ppt", ".pptx"))]
   for pptfile in pptfiles:
       print('正在转换',pptfile)
       fullpath = os.path.join(cwd, pptfile)
       fullpath2 = os.path.join(cwd+'/out', pptfile)
       ppt_to_pdf(powerpoint, fullpath, fullpath2)

if __name__ == "__main__":
   cwd=os.getcwd()
   a=os.path.exists('./out')
   if(a==False):
       os.mkdir(cwd+'./out')
   powerpoint = init_powerpoint()
   convert_files_in_folder(powerpoint, cwd)
   powerpoint.Quit()
   print('转换完成')