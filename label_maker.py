#%%
import pandas as pd
import numpy as np
import datetime as dt

from string import ascii_uppercase

# from funcs.stritfuncs import excel_cols, get_wwtp_df


def get_wwtp_df(wwtp_data_src):

    wwtp_df = pd.read_csv(wwtp_data_src,header=0)
    wwtp_df = wwtp_df[["ARA ID \n(Eawag/EPFL)", "ARANAME","Kanton","ARANR (GIS)",
                       "StadtNAME","Population (Current)"]]
    wwtp_df.rename(columns = {"ARA ID \n(Eawag/EPFL)":"ARA_ID",
                              "Population (Current)":"Population"}, inplace=True)
    wwtp_df.replace(to_replace="/",value="-",regex=True, inplace=True)
    wwtp_df["ARA_ID"] = wwtp_df["ARA_ID"].str.strip("_")
    wwtp_df["StadtNAME"] = wwtp_df["StadtNAME"].str.lower()

    wwtp_df.set_index("ARA_ID", inplace=True)

    return wwtp_df

def excel_cols(df, index=True):
    cols = list(df.columns.values)
    idx = list(df.index.names)

    if (df.index.names[0] == None) & (index==True):
        cols.insert(0,"index")
    elif index==False:
        pass
    else:
        cols = idx + cols

    i = 0
    j = -1
    excel_col_dict = {}
    for col in cols:
        if (i < 26) & (j < 0):
            excel_col_dict[col]= ascii_uppercase[i]
        elif i == 26:
            i = 0
            j += 1
            excel_col_dict[col]= ascii_uppercase[j]+ascii_uppercase[i]
        else:
            excel_col_dict[col]= ascii_uppercase[j]+ascii_uppercase[i]
        i += 1

    return excel_col_dict

#%% Prints 3 weeks of stickers at a time (15 days)
if __name__ == "__main__":
    # Starting date of our expanded 15 WWTP monitoring
    start_monitor_program = dt.date(year=2023,month=7,day=10)
    # Increase by one everytime you want to print stickers (describes the amount of 3 week intervals to add to the start_monitoring_program)
    factor = 0
    # Determination of the first date of the stickers
    start_date = dt.date(year=2023,month=7,day=10) + (dt.timedelta(days=(21 *factor)))
    # Determining if it is an odd or even week
    alternating_determinant = (start_date - start_monitor_program)
    # Establishing the date range of the stickers
    base = pd.date_range(start_date, start_date + dt.timedelta(days=22))

    if alternating_determinant.days % 2 == 0:
        dr_purple = base[[0,2,4,5,6,8,10,11,12,13,14,16,18,19,20]]
        dr_orange = base[[1,3,5,6,7,9,11,12,13,14,15,17,19,20,21]]
        dr_green = base[[2,4,6,7,8,10,12,13,14,15,16,18,20,21,22]]
        
    if alternating_determinant.days % 2 != 0: 
        dr_purple =  base[[1,3,4,5,6,7,9,11,12,13,15,17,18,19,20]]
        dr_orange = base[[2,4,5,6,7,8,10,12,13,14,16,18,19,20,21]]
        dr_green =  base[[3,5,6,7,8,9,11,13,14,15,17,19,20,21,22]]

    wwtp_df = get_wwtp_df("/Users/charlesgan/Library/Mobile Documents/com~apple~CloudDocs/Eawag Covid Work/Labels/wwtp_info.csv")

#############################################! PURPLE GROUP ###################################################   
    for wwtp in (['10','12','17','34','35']):

        wwtp_long = wwtp + '_' + wwtp_df['ARANAME'].loc[wwtp].replace(" ","_")

        dates_top = ['\n'.join([wwtp,str(dt.datetime.date(val)),"v4"]) for val in dr_purple]
        dates_long = ['\n'.join([wwtp_long,str(dt.datetime.date(val))]) for val in dr_purple]

        for i in range(1,len(dates_top)*2,2):
            
            dates_top.insert(i, pd.NA)

            dates_long.insert(i, pd.NA)

        cols = ['Eluate [Protocol v4]','top1','qPCR 1:10 Dil ☐ (5µL)','top2','Sequencing 15 µL','top3','dPCR 1:X Dil ☐ (X=__)','top4']

        d = pd.DataFrame()
        for i,col in enumerate(cols):

            if col.startswith('top'):
                s = pd.Series(dates_top, name=col)
            else:
                s = pd.Series(dates_long, name=col)
            s_blank = pd.Series(name=i)
            
            if i % 2 == 0:

                s = s.apply(lambda x: x +f'\n{col}')

            d = pd.concat([d,s,s_blank,], axis=1)

            

        output_xlsx = f'/Users/charlesgan/Library/Mobile Documents/com~apple~CloudDocs/Eawag Covid Work/Labels/LabelTemplate_{wwtp}_{start_date}.xlsx' #Change this
        with pd.ExcelWriter(output_xlsx, engine='xlsxwriter', date_format='YYYY-MM-DD') as writer:

            d.to_excel(writer, sheet_name='LabelTemplate', index=False, header=False)

            workbook  = writer.book
            worksheet = writer.sheets['LabelTemplate']

            worksheet.set_margins(left=0.47,right=0,top=0.25,bottom=0)
            worksheet.set_paper(1)
            worksheet.set_page_view()

            cols = list(d.columns.values)
            excel_col_dict = excel_cols(d, index=False)

            blank_cols = {}
            for k in range(0,8):
                blank_cols[k] = excel_col_dict.pop(k)
            last_col = list(excel_col_dict.values())[-1]

            color_list =['black','black','green','green','red','red','blue','blue']

            for i,val in enumerate(excel_col_dict):
                col = excel_col_dict[val]
                
                if 'top' in val:
                    label_format = workbook.add_format({
                                        'bold': True,
                                        'text_wrap': True,
                                        'font_name':"Eurostile",
                                        'font_size': 4,
                                        'valign': 'top',
                                        'align': 'center',
                                        'border': 0})
                    width_pix = 31
                    width = 4.4
                else:
                    label_format = workbook.add_format({
                                        'bold': True,
                                        'text_wrap': True,
                                        'font_name':"Eurostile",
                                        'font_size': 7,
                                        'valign': 'vcenter',
                                        'align': 'center',
                                        'border': 0})
                    width_pix = 102
                    width = 16.3

                
                label_format.set_font_color(color_list[i])

                # worksheet.set_column_pixels(f'{col}:{col}',width_pix, label_format)
                worksheet.set_column(f'{col}:{col}',width, label_format)

            for col in blank_cols.values():

                worksheet.set_column_pixels(f'{col}:{col}',10)

            for row in range(0,35,1):

                if row % 2 == 0:
                    height = 39
                else:
                    height= 9
                worksheet.set_row(row, height)
#############################################! ORANGE GROUP ###################################################   
    for wwtp in (['05','15','18','19','25']):

        wwtp_long = wwtp + '_' + wwtp_df['ARANAME'].loc[wwtp].replace(" ","_")

        dates_top = ['\n'.join([wwtp,str(dt.datetime.date(val)),"v4"]) for val in dr_orange]
        dates_long = ['\n'.join([wwtp_long,str(dt.datetime.date(val))]) for val in dr_orange]

        for i in range(1,len(dates_top)*2,2):
            
            dates_top.insert(i, pd.NA)

            dates_long.insert(i, pd.NA)

        cols = ['Eluate [Protocol v4]','top1','qPCR 1:10 Dil ☐ (5µL)','top2','Sequencing 15 µL','top3','dPCR 1:X Dil ☐ (X=__)','top4']

        d = pd.DataFrame()
        for i,col in enumerate(cols):

            if col.startswith('top'):
                s = pd.Series(dates_top, name=col)
            else:
                s = pd.Series(dates_long, name=col)
            s_blank = pd.Series(name=i)
            
            if i % 2 == 0:

                s = s.apply(lambda x: x +f'\n{col}')

            d = pd.concat([d,s,s_blank,], axis=1)

            

        output_xlsx = f'/Users/charlesgan/Library/Mobile Documents/com~apple~CloudDocs/Eawag Covid Work/Labels/LabelTemplate_{wwtp}_{start_date}.xlsx' #Change this
        with pd.ExcelWriter(output_xlsx, engine='xlsxwriter', date_format='YYYY-MM-DD') as writer:

            d.to_excel(writer, sheet_name='LabelTemplate', index=False, header=False)

            workbook  = writer.book
            worksheet = writer.sheets['LabelTemplate']

            worksheet.set_margins(left=0.47,right=0,top=0.25,bottom=0)
            worksheet.set_paper(1)
            worksheet.set_page_view()

            cols = list(d.columns.values)
            excel_col_dict = excel_cols(d, index=False)

            blank_cols = {}
            for k in range(0,8):
                blank_cols[k] = excel_col_dict.pop(k)
            last_col = list(excel_col_dict.values())[-1]

            color_list =['black','black','green','green','red','red','blue','blue']

            for i,val in enumerate(excel_col_dict):
                col = excel_col_dict[val]
                
                if 'top' in val:
                    label_format = workbook.add_format({
                                        'bold': True,
                                        'text_wrap': True,
                                        'font_name':"Eurostile",
                                        'font_size': 4,
                                        'valign': 'top',
                                        'align': 'center',
                                        'border': 0})
                    width_pix = 31
                    width = 4.4
                else:
                    label_format = workbook.add_format({
                                        'bold': True,
                                        'text_wrap': True,
                                        'font_name':"Eurostile",
                                        'font_size': 7,
                                        'valign': 'vcenter',
                                        'align': 'center',
                                        'border': 0})
                    width_pix = 102
                    width = 16.3

                
                label_format.set_font_color(color_list[i])

                # worksheet.set_column_pixels(f'{col}:{col}',width_pix, label_format)
                worksheet.set_column(f'{col}:{col}',width, label_format)

            for col in blank_cols.values():

                worksheet.set_column_pixels(f'{col}:{col}',10)

            for row in range(0,35,1):

                if row % 2 == 0:
                    height = 39
                else:
                    height= 9
                worksheet.set_row(row, height)
#############################################! GREEN GROUP ###################################################                
    for wwtp in (['16','32','33','36']):

        wwtp_long = wwtp + '_' + wwtp_df['ARANAME'].loc[wwtp].replace(" ","_")

        dates_top = ['\n'.join([wwtp,str(dt.datetime.date(val)),"v4"]) for val in dr_green]
        dates_long = ['\n'.join([wwtp_long,str(dt.datetime.date(val))]) for val in dr_green]

        for i in range(1,len(dates_top)*2,2):
            
            dates_top.insert(i, pd.NA)

            dates_long.insert(i, pd.NA)

        cols = ['Eluate [Protocol v4]','top1','qPCR 1:10 Dil ☐ (5µL)','top2','Sequencing 15 µL','top3','dPCR 1:X Dil ☐ (X=__)','top4']

        d = pd.DataFrame()
        for i,col in enumerate(cols):

            if col.startswith('top'):
                s = pd.Series(dates_top, name=col)
            else:
                s = pd.Series(dates_long, name=col)
            s_blank = pd.Series(name=i)
            
            if i % 2 == 0:

                s = s.apply(lambda x: x +f'\n{col}')

            d = pd.concat([d,s,s_blank,], axis=1)

            

        output_xlsx = f'/Users/charlesgan/Library/Mobile Documents/com~apple~CloudDocs/Eawag Covid Work/Labels/LabelTemplate_{wwtp}_{start_date}.xlsx' #Change this
        with pd.ExcelWriter(output_xlsx, engine='xlsxwriter', date_format='YYYY-MM-DD') as writer:

            d.to_excel(writer, sheet_name='LabelTemplate', index=False, header=False)

            workbook  = writer.book
            worksheet = writer.sheets['LabelTemplate']

            worksheet.set_margins(left=0.47,right=0,top=0.25,bottom=0)
            worksheet.set_paper(1)
            worksheet.set_page_view()

            cols = list(d.columns.values)
            excel_col_dict = excel_cols(d, index=False)

            blank_cols = {}
            for k in range(0,8):
                blank_cols[k] = excel_col_dict.pop(k)
            last_col = list(excel_col_dict.values())[-1]

            color_list =['black','black','green','green','red','red','blue','blue']

            for i,val in enumerate(excel_col_dict):
                col = excel_col_dict[val]
                
                if 'top' in val:
                    label_format = workbook.add_format({
                                        'bold': True,
                                        'text_wrap': True,
                                        'font_name':"Eurostile",
                                        'font_size': 4,
                                        'valign': 'top',
                                        'align': 'center',
                                        'border': 0})
                    width_pix = 31
                    width = 4.4
                else:
                    label_format = workbook.add_format({
                                        'bold': True,
                                        'text_wrap': True,
                                        'font_name':"Eurostile",
                                        'font_size': 7,
                                        'valign': 'vcenter',
                                        'align': 'center',
                                        'border': 0})
                    width_pix = 102
                    width = 16.3

                
                label_format.set_font_color(color_list[i])

                # worksheet.set_column_pixels(f'{col}:{col}',width_pix, label_format)
                worksheet.set_column(f'{col}:{col}',width, label_format)

            for col in blank_cols.values():

                worksheet.set_column_pixels(f'{col}:{col}',10)

            for row in range(0,35,1):

                if row % 2 == 0:
                    height = 39
                else:
                    height= 9
                worksheet.set_row(row, height)
# %%
