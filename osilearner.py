import xlrd
import os

from Tkinter import *
import tkSimpleDialog
import tkMessageBox

tide_file_location="C:/Users/Unnas/Documents/1UnnasHussain/hudsonTides.xlsx"
tide_workbook = xlrd.open_workbook(tide_file_location)
tide_sheet=tide_workbook.sheet_by_index(0)

wave_file_location="C:/Users/Unnas/Documents/1UnnasHussain/hudsonWaves.xlsx"
wave_workbook = xlrd.open_workbook(wave_file_location)
wave_sheet=wave_workbook.sheet_by_index(0)

#--------------------------------------------------------------------------
wheight_col=3

wperiod_mean=5
wperiod_wind=7
wperiod_peak=8

wdirection_mean=4
wdirection_wind=6
wdirection_peak=8
#--------------------------------------------------------------------------

def get_wave_data(date, col):
    for r in range(0,wave_sheet.nrows):
        if wave_sheet.cell_value(r,0)==date:
            return wave_sheet.cell_value(r,col)
            
    
def get_low_tide(date):
    summ=0
    count=0
    row=0
    for r in range(0,tide_sheet.nrows-1):
        if tide_sheet.cell_value(r,0)==date:
            if tide_sheet.cell_value(r,5)=="Low Tide":
                summ=summ+tide_sheet.cell_value(r,3)
                count=count+1
    if(count==0):
        return 'Date Invalid for current TIDE data set'
    else:
        return summ/count

def get_high_tide(date):
    summ=0
    count=0
    row=0
    for r in range(0,tide_sheet.nrows-1):
        if tide_sheet.cell_value(r,0)==date:
            if tide_sheet.cell_value(r,5)=="High Tide":
                summ=summ+tide_sheet.cell_value(r,3)
                count=count+1
    if(count==0):
        return 'Date Invalid for current TIDE data set'
    else:
        return summ/count
    
def find_best_pos(date):
    tide_val=(get_low_tide(date)+get_high_tide(date))/2
    if(tide_val>get_wave_data(date,wheight_col)):
        return 'Vertical Orientation'
    elif(get_wave_data(date,wheight_col)>tide_val):
        return 'Horizontal Position facing: ' + get_wave_data(date,wdirection_wind)
    elif(get_wave_data(date,wheight_col)==tide_val):
        return 'Vertical Orientation is Suggested'

#------------------------------------------------------------

#wheight_col=3

#wperiod_mean=5
#wperiod_wind=7
#wperiod_peak=8

#wdirection_mean=4
#wdirection_wind=6
#wdirection_peak=8

root=Tk()
x=Label(root,text="M.O.E.C.D. Learner")
tkMessageBox.showinfo("M.O.E.C.D. Learner","Hi There!")

while (1==1):
    
    
    date=tkSimpleDialog.askstring("M.O.E.C.D. is Learning...","Enter Date [Day 00 Month]:")
    
    
    output=("Low Tide: {} m\nHigh Tide: {} m\nWave Height: {} m\nWave Direction: {}".format(get_low_tide(date),                                                                                        get_high_tide(date),
                                                                                                get_wave_data(date,wheight_col),
                                                                                                get_wave_data(date,wdirection_peak)))
    
    tkMessageBox.showinfo("Ocean Motion",output)
    
    tkMessageBox.showinfo('BEST POSITION for M.O.E.C.D.', find_best_pos(date))

root.quit()
