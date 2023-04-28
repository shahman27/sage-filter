import pandas as pd
from xlsxwriter import Workbook
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
import os
import numpy as np

class sheet_filter:
    def __init__(self, df) :
        """ Constructor for this class.
            
            Args:
                df (dataframe): dataframe to be filtered 

        """
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
        """Return a list of dataframes grouped by asset id

            Args:
                df1 (dataframe): dataframe to be grouped
            
            Returns:
                list: list of dataframes grouped by asset id
        """
        grouped3 = df1.groupby(['Asset ID'])
        new_df3 = [grouped3.get_group(x) for x in grouped3.groups]
        return new_df3

    def multiple_dfs(self, df_list, sheets, file_name, spaces):
        """ Saves multiple dataframes to one excel sheet

            Args:
                df_list (list): list of dataframes to be saved
                sheets (string): name of sheet to be saved to
                file_name (string): name of file to be saved to
                spaces (int): number of rows to skip between dataframes
    
        """
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
        """ Writes multiple dataframes to one excel sheet

            Args:  
                df_list (list): list of dataframes to be saved
                sheets (string): name of sheet to be saved to
                spaces (int): number of rows to skip between dataframes
                writer (writer): writer object to write to excel

        """
        row = 0
        for dataframe in df_list:
            dataframe.to_excel(writer,sheet_name=sheets,startrow=row , startcol=0)
            workbook = writer.book
            worksheet = writer.sheets[sheets]
            money_format = workbook.add_format({'num_format': '$#,##0.00'})
            worksheet.set_column('J:N', 20, money_format)     
            row = row + len(dataframe.index) + spaces + 1

    def df_tabs(self, df_list, sheet_list, file_name) :
        """ Saves multiple dataframes to multiple excel sheets

            Args:
                df_list (list): list of dataframes to be saved
                sheet_list (list): list of sheet names to be saved to
                file_name (string): name of file to be saved to

        """

        writer = pd.ExcelWriter(file_name,engine='xlsxwriter')   
        for dataframe, sheet in zip(df_list, sheet_list):
            self.df_tab_sheets(dataframe, sheet, 1, writer)
        writer.close()

    def graph(self, df , name, legend_names):
        """ Saves a bar graph of the dataframe
            
                Args:
                    df (dataframe): dataframe to be graphed
                    name (string): name of file to be saved to
                    legend_names (list): list of names for the legend
    
        """
        ax = df.plot.bar(logy=True, figsize=(30,25))
        ax.legend(legend_names, bbox_to_anchor=(1.05, 1), loc='upper left')
        ax.yaxis.set_major_locator(plt.MaxNLocator(10))
        ax.yaxis.set_major_formatter(mtick.ScalarFormatter())
        ax.ticklabel_format(axis="y", style='plain')
        ax.locator_params(axis='y', nbins=10)
        plt.savefig(name, bbox_inches='tight')
        # plt.show()

class quarter(sheet_filter):
    def __init__(self, df):
        """ Constructor for this class. 
            Args:
                df (dataframe): dataframe to be filtered
        """
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
        """ Reduces the dataframe to the last 4 quarters """
        self.df = self.df[-4:]

    def recent(self):
        """ Reduces the dataframe to the most recent quarter """
        self.df = [self.df[-1]]

    def specific(self, quarter):
        """ Reduces the dataframe to a specific quarter
            Args:
                quarter (string): quarter to be filtered to
            Returns:
                bool: True if quarter is found, False if not
        """
        found = False
        for x in self.df:
            if (str(x['Quarter'].iloc[0]) == quarter):
                self.df = [x]
                found = True
                break
        return found
    
    def same_quarter(self, quarter):
        """ Reduces the dataframe to all quarters with the same quarter
            Args:
                quarter (string): quarter to be filtered to
        """
    
        quarter_list = []
        for x in self.df:
            if (str(x['Quarter'].iloc[0]).__contains__(quarter)):
                quarter_list.append(x)
        self.df = quarter_list
    
    def full_lists(self):
        """ Filters the dataframe to the full list of assets for each quarter
            Returns:
                list: list of dataframes
                list: list of names of the dataframes
        """
        df_list = []
        names = []
        
        for df in self.df:
            temp_df = self.asset_filter(df)
            df_list.append(temp_df)
            names.append(str(df['Quarter'].iloc[0]))

        return [df_list, names]
        
    def total_assets(self, out_type):
        """ Filters the dataframe to the total assets for each quarter
            Args:
                out_type (string): type of output to be returned
            Returns:
                list: list of dataframes
                list: list of names of the dataframes
        """
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

            df = pd.DataFrame({'Asset ID' : names, 'Total Acquired Value' : totals, 'Total Depreciation' : dep_totals})
            df.sort_values(by='Total Acquired Value', inplace=True, ascending=False)
            df.set_index('Asset ID', inplace=True)
            quarter_index = df.index
            quarter_index.name = 'Quarter: ' + curr_quarter
            if (out_type == 'assets'):
                quarter_name.append(str(curr_quarter) + ' Assets')
                df.drop('Total Depreciation', axis=1, inplace=True)
            elif (out_type == 'depreciation'):
                quarter_name.append(str(curr_quarter) + ' Depreciation')
                df.drop('Total Acquired Value', axis=1, inplace=True)
            else:
                quarter_name.append(str(curr_quarter) + ' Assets')
                quarter_name.append(str(curr_quarter) + ' Depreciation')
            df_totals.append(df)

        return [df_totals, quarter_name]
       
class year(sheet_filter):
    def __init__(self, df):
        """ Constructor for this class.
            Args:
                df (dataframe): dataframe to be filtered
        """
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
        """ Reduces the dataframe to the last 2 years"""
        self.df = self.df[-2:]

    def recent(self):
        """ Reduces the dataframe to the most recent year """
        self.df = [self.df[-1]]
    
    def specific(self, year):
        """ Reduces the dataframe to a specific year
            Args:
                year (string): year to be filtered to
            Returns:
                bool: True if year is found, False if not
        """
        found = False
        for x in self.df:
            if (str(x['Fiscal Year'].iloc[0]) == year):
                self.df = [x]
                found = True
                break

        return found               
    
    def full_lists(self):
        """ Filters the dataframe to the full list of assets for each year
            Returns:
                list: list of dataframes
                list: list of names of the dataframes
        """
        df_list = []
        names = []
        
        for df in self.df:
            temp_df = self.asset_filter(df)
            df_list.append(temp_df)
            names.append(str(df['Fiscal Year'].iloc[0]))

        return [df_list, names]

    def total_assets(self, out_type):
        """ Filters the dataframe to the total assets for each year
            Args:
                out_type (string): type of output to be returned
            Returns:
                list: list of dataframes
                list: list of names of the dataframes
        """
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
            
            df = pd.DataFrame({'Asset ID' : names, 'Total Acquired Value' : totals, 'Total Depreciation' : dep_totals})
            df.sort_values(by='Total Acquired Value', inplace=True, ascending=False)
            df.set_index('Asset ID', inplace=True)
            quarter_index = df.index
            quarter_index.name = 'Fiscal Year: ' + curr_year
            if (out_type == 'assets'):
                years.append(str(curr_year) + ' Assets')
                df.drop('Total Depreciation', axis=1, inplace=True)
            elif (out_type == 'depreciation'):
                years.append(str(curr_year) + ' Depreciation')
                df.drop('Total Acquired Value', axis=1, inplace=True)
            else:
                years.append(str(curr_year) + ' Assets')
                years.append(str(curr_year) + ' Depreciation')
            df_totals.append(df)

        return [df_totals, years]

