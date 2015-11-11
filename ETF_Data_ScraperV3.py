import Tkinter as tk
import time
   
top = tk.Tk()
   
def addText():
  tickerSymbols = ["VTIAX","PTTRX","PRFDX","DBLTX","TGBAX","FCNTX","CNSAX","ANZAX","FISCX","FACVX","PACIX","VCVSX","DEEAX","ACCBX","CLDAX"]

  for i in tickerSymbols:
    i = tickerSymbols.index(i)
    i = str(i)
    currentPercentage = L.cget("text")
    newPercentage = i + "%" + " complete"
    L.configure(text=newPercentage)
    top.update_idletasks()
    
   
B = tk.Button(top, text ="Change text", command = addText)
L = tk.Label(top,text='0%')
   
B.pack()
L.pack()
top.geometry("900x300") 
top.mainloop()

