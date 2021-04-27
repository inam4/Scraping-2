from tkinter import *
main_win=Tk()
button=[Button for _ in range(10)]
print(button)
for i in range(10):
   button[i]=Button(main_win,text=i,bg="gray14",fg="thistle4",width=7,height=4)

b=0
for i in range(3):
   for j in range(3):
      button[b].grid(row=i,column=1)
      b=b+1

main_win.mainloop()