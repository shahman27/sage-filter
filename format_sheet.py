import pandas as pd
from xlsxwriter import Workbook
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
import os
import numpy as np

class sheet_filter:
    def __init__(self, df) :
        df.columns = ['Sys No' , 'Asset ID', 'Description', 'Vendor' , 'Serial Number', 'In Svc Date', 'EST Life', 'Prior Thru' , 'Acquired Value' , 'Prior Accum Depreciation' , 'Depreciation This Run', 'Current YTD Deprecitation', 'Current Accum Depreciation']
        reduced_df = df[:-8]
        df1 = reduced_df.copy()
        df1['In Svc Date'] = pd.to_datetime(df1['In Svc Date'])
        df1 = df1.sort_values(by='In Svc Date')
        df1['Quarter'] = df1['In Svc Date'].dt.to_period('Q-OCT')
        df1['Fiscal Year'] = df1['In Svc Date'].dt.to_period('Q-OCT').dt.year
        df1.sort_values(by='Acquired Value', inplace=True, ascending=False)
        self.df = df1

    def asset_filter(self , df1):
        grouped3 = df1.groupby(['Asset ID'])
        new_df3 = [grouped3.get_group(x) for x in grouped3.groups]
        return new_df3

    def multiple_dfs(self, df_list, sheets, file_name, spaces):
        writer = pd.ExcelWriter(file_name,engine='xlsxwriter')   
        row = 0
        for dataframe in df_list:
            dataframe.to_excel(writer,sheet_name=sheets,startrow=row , startcol=0)
            workbook = writer.book
            worksheet = writer.sheets[sheets]
            money_format = workbook.add_format({'num_format': '$#,##0.00'})
            worksheet.set_column('B:C', 20, money_format)  
            row = row + len(dataframe.index) + spaces + 1
        writer.close()
    
    def df_tab_sheets(self, df_list, sheets, spaces, writer):
        row = 0
        for dataframe in df_list:
            dataframe.to_excel(writer,sheet_name=sheets,startrow=row , startcol=0)
            workbook = writer.book
            worksheet = writer.sheets[sheets]
            money_format = workbook.add_format({'num_format': '$#,##0.00'})
            worksheet.set_column('J:N', 20, money_format)     
            row = row + len(dataframe.index) + spaces + 1

    def df_tabs(self, df_list, sheet_list, file_name) :
        writer = pd.ExcelWriter(file_name,engine='xlsxwriter')   
        for dataframe, sheet in zip(df_list, sheet_list):
            self.df_tab_sheets(dataframe, sheet, 1, writer)
        writer.close()

    def graph(self, df , name, legend_names):
        ax = df.plot.bar(logy=True, figsize=(30,25))
        ax.legend(legend_names, bbox_to_anchor=(1.05, 1), loc='upper left')
        ax.yaxis.set_major_locator(plt.MaxNLocator(10))
        ax.yaxis.set_major_formatter(mtick.ScalarFormatter())
        # ax.ticklabel_format(axis="y", style='plain')
        # ax.locator_params(axis='y', nbins=10)
        # plt.savefig(name, bbox_inches='tight')
        plt.show()

class quarter(sheet_filter):
    def __init__(self, df):
        df.columns = ['Sys No' , 'Asset ID', 'Description', 'Vendor' , 'Serial Number', 'In Svc Date', 'EST Life', 'Prior Thru' , 'Acquired Value' , 'Prior Accum Depreciation' , 'Depreciation This Run', 'Current YTD Deprecitation', 'Current Accum Depreciation']
        reduced_df = df[:-8]
        df1 = reduced_df.copy()
        df1['In Svc Date'] = pd.to_datetime(df1['In Svc Date'])
        df1 = df1.sort_values(by='In Svc Date')
        df1['Quarter'] = df1['In Svc Date'].dt.to_period('Q-OCT')
        df1.sort_values(by='Acquired Value', inplace=True, ascending=False)
        grouped = df1.groupby(['Quarter'])
        new_df = [grouped.get_group(x) for x in grouped.groups]
        self.df = new_df

    def reduce_df(self):
        self.df = self.df[-4:]

    def recent(self):
        self.df = [self.df[-1]]

    def specific(self, quarter):
        found = False
        for x in self.df:
            if (str(x['Quarter'].iloc[0]) == quarter):
                self.df = [x]
                found = True
                break
        return found
    
    def same_quarter(self, quarter):
        quarter_list = []
        for x in self.df:
            if (str(x['Quarter'].iloc[0]).__contains__(quarter)):
                quarter_list.append(x)
        self.df = quarter_list
    
    def full_lists(self):
        df_list = []
        names = []
        
        for df in self.df:
            temp_df = self.asset_filter(df)
            df_list.append(temp_df)
            names.append(str(df['Quarter'].iloc[0]))

        return [df_list, names]
        
    def total_assets(self, out_type):
        sorted_by_assets = []

        for x in self.df:
            df = x.groupby(['Asset ID'])
            new_df = [df.get_group(x) for x in df.groups]
            sorted_by_assets.append(new_df)

        df_totals = []
        quarter_name = []

        for x in sorted_by_assets:
            totals = []
            dep_totals = []
            names = []
            for y in x:
                totals.append(y['Acquired Value'].sum())
                dep_totals.append(y['Current YTD Deprecitation'].sum())
                names.append(y['Asset ID'].iloc[0])
                curr_quarter = str(y['Quarter'].iloc[0])

            df = pd.DataFrame({'Asset ID' : names, 'Total Value' : totals, 'Total Depreciation' : dep_totals})
            df.sort_values(by='Total Value', inplace=True, ascending=False)
            df.set_index('Asset ID', inplace=True)
            quarter_index = df.index
            quarter_index.name = 'Quarter: ' + curr_quarter
            if (out_type == 'assets'):
                quarter_name.append(str(curr_quarter) + ' Assets')
                df.drop('Total Depreciation', axis=1, inplace=True)
            elif (out_type == 'depreciation'):
                quarter_name.append(str(curr_quarter) + ' Depreciation')
                df.drop('Total Value', axis=1, inplace=True)
            else:
                quarter_name.append(str(curr_quarter) + ' Assets')
                quarter_name.append(str(curr_quarter) + ' Depreciation')
            df_totals.append(df)


        # self.multiple_dfs(df_totals, 'Validation', file_name, 1)
        # print(df_totals)

        return [df_totals, quarter_name]
        
        # df = pd.concat(df_totals, axis=1)

        # self.graph(df, "happy" + '.png', quarter_name)

        # df.plot.bar(logy=True, figsize=(20,10))
        # plt.legend(labels=quarter_name , bbox_to_anchor=(1.05, 1), loc='upper left')
        # plt.show()
       
class year(sheet_filter):
    def __init__(self, df):
        df.columns = ['Sys No' , 'Asset ID', 'Description', 'Vendor' , 'Serial Number', 'In Svc Date', 'EST Life', 'Prior Thru' , 'Acquired Value' , 'Prior Accum Depreciation' , 'Depreciation This Run', 'Current YTD Deprecitation', 'Current Accum Depreciation']
        reduced_df = df[:-8]
        df1 = reduced_df.copy()
        df1['In Svc Date'] = pd.to_datetime(df1['In Svc Date'])
        df1 = df1.sort_values(by='In Svc Date')
        df1['Quarter'] = df1['In Svc Date'].dt.to_period('Q-OCT')
        df1['Fiscal Year'] = df1['In Svc Date'].dt.to_period('Q-OCT').dt.year
        df1.sort_values(by='Acquired Value', inplace=True, ascending=False)
        grouped = df1.groupby(['Fiscal Year'])
        new_df = [grouped.get_group(x) for x in grouped.groups]
        self.df = new_df  

    def reduce_df(self):
        self.df = self.df[-2:]

    def recent(self):
        self.df = [self.df[-1]]
    
    def specific(self, year):
        found = False
        for x in self.df:
            if (str(x['Fiscal Year'].iloc[0]) == year):
                self.df = [x]
                found = True
                break

        return found               
    
    def full_lists(self):
        df_list = []
        names = []
        
        for df in self.df:
            temp_df = self.asset_filter(df)
            df_list.append(temp_df)
            names.append(str(df['Fiscal Year'].iloc[0]))

        return [df_list, names]
        # df = pd.concat(df_list[0], axis=1)
        # self.graph(df, "happy" + '.png', names)

    def total_assets(self, out_type):
        sorted_by_assets = []

        for x in self.df:
            df = x.groupby(['Asset ID'])
            new_df = [df.get_group(x) for x in df.groups]
            sorted_by_assets.append(new_df)

        df_totals = []
        years = []

        for x in sorted_by_assets:
            totals = []
            dep_totals = []
            names = []
            for y in x:
                totals.append(y['Acquired Value'].sum())
                names.append(y['Asset ID'].iloc[0])
                dep_totals.append(y['Current YTD Deprecitation'].sum())
                curr_year = str(y['Fiscal Year'].iloc[0])
            
            df = pd.DataFrame({'Asset ID' : names, 'Total Value' : totals, 'Total Depreciation' : dep_totals})
            df.sort_values(by='Total Value', inplace=True, ascending=False)
            df.set_index('Asset ID', inplace=True)
            quarter_index = df.index
            quarter_index.name = 'Fiscal Year: ' + curr_year
            if (out_type == 'assets'):
                years.append(str(curr_year) + ' Assets')
                df.drop('Total Depreciation', axis=1, inplace=True)
            elif (out_type == 'depreciation'):
                years.append(str(curr_year) + ' Depreciation')
                df.drop('Total Value', axis=1, inplace=True)
            else:
                years.append(str(curr_year) + ' Assets')
                years.append(str(curr_year) + ' Depreciation')
            df_totals.append(df)

        # self.multiple_dfs(df_totals, 'Validation', file_name + '.xlsx', 1)

        # print(years)
        return [df_totals, years]
        # df = pd.concat(df_totals, axis=1)
        

        # self.graph(df, "happy" + '.png', years)


df = pd.read_excel('C:/Projects/sage_filter/HTS-sage-report5.xls')        

first = quarter(df)

first.recent()
first.total_assets('both')



# df2 = first.compare_q_assets()

# df3 = [df2, df2, df2, df2]
# name = ['Quarter 1', 'Quarter 2', 'Quarter 3', 'Quarter 4']

# first.tryal(df3, name, 'test1.xlsx')