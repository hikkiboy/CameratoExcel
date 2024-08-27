import cv2
import xlwings as xw
import imutils
import cython


cam = cv2.VideoCapture(0)

wb = xw.Book('cam.xlsx')
sht1 = wb.sheets['Sheet']

@cython.boundscheck(False)
def run():
    for row in range(1,36):
        for col in range(1,37):
                ret,frame = cam.read()
                frame = imutils.resize(frame, width=48)
                cv2.imshow('frame', frame)    
    
                colors = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                SendColor(colors=colors,row=row,col=col)
                print('Current Row: ', row, "Current Column: ", col)
                
                # print('RED: ', R, 'GREEN: ', G, 'BLUE: ', B)
    run()

    
def SendColor(colors,row,col):
                R = int(colors[row][col][0])
                G = int(colors[row][col][1])
                B = int(colors[row][col][2])
                sht1.range(row,col).color = (R,G,B)
run()

cam.release()
cv2.destroyAllWindows()