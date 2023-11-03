import xlsxwriter
import os
from datetime import datetime
from meteostat import Monthly, Stations

def get_average_temp():

    ##################################################################
    ############## declair variable section ##########################
    ################################################################## 
      
    states =['AL','AK','AS','AZ','AR','CA','CO','CT','DE','DC','FL','GA','GU','HI','ID','IL','IN','IA','KS','KY','LA','ME','MD','MA','MI','MN','MS','MO','MT','NE','NV','NH','NJ','NM','NY','NC','ND','MP','OH','OK','OR','PA','PR','RI','SC','SD','TN','TX','UM','UT','VT','VI','VA','WA','WV','WI','WY',]
    start = datetime(2000, 1, 1)
    end = datetime(2022, 12, 31)
    destination_path = os.getcwd()+'\\State_Climates.xlsx'

    ##################################################################
    ############## spreadsheet formatting section ####################
    ##################################################################

    workbook = xlsxwriter.Workbook(destination_path)
    worksheet = workbook.add_worksheet('High Temps')
    worksheet.set_column('A:A', 5)
    worksheet.set_column('B:B', 16)
    worksheet.set_column('C:C', 16)
    worksheet.set_column('D:D', 30)
    worksheet.set_column('E:E', 36)
    worksheet.set_column('F:F', 42)
    worksheet.write('A1', "State")
    worksheet.write('B1', "Average high in F")
    worksheet.write('C1', "Average low in F")
    worksheet.write('D1', "Average max wind speed in MPH")
    worksheet.write('E1', "Average max precipitation in inches")
    worksheet.write('F1', "Average daily sunlight in minutes if available")

    ##################################################################
    ################## gather data section ###########################
    ##################################################################

    row_count = 2
    for state in states: 
        try:
            stations = Stations()
            all_stations = stations.region('US', state).count()
            stations = stations.region('US', state).fetch(all_stations, sample=True)
            data = Monthly(stations, start, end)
            data = data.normalize().aggregate(freq="1Y").fetch()

    ##################################################################
    ################### convert data section #########################
    ##################################################################

            av_high = round((data["tmax"].max()*9/5)+32)        #convert to Farenheight
            av_low = round((data["tmin"].min()*9/5)+32)         #convert to Farenheight
            av_windspeed = round((data["tmax"].max()/1.609344)) #convert to MPH
            av_precipitation = round((data["prcp"].max()/25.4)) #convert to inches
            av_sun = round(data["tsun"].max())                  #already in Minutes

    ##################################################################
    ##############  write to spreadsheet section  ####################
    ##################################################################
                                                       
            print("row", row_count, state, all_stations, "stations found\n")
            worksheet.write('A'+str(row_count), state)
            worksheet.write('B'+str(row_count), str(av_high))
            worksheet.write('C'+str(row_count), str(av_low))
            worksheet.write('D'+str(row_count), str(av_windspeed))
            worksheet.write('E'+str(row_count), str(av_precipitation))
            worksheet.write('F'+str(row_count), str(av_sun))                      
            row_count += 1

        except Exception as e:
            print(e)
    workbook.close()

get_average_temp()
