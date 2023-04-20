# %% [markdown]
# # **WELCOME TO THE SAGE FILTER**
# ## ----------------------------------------------------------------
# ## This program will filter data from a Sage report and output a new report or graph with filtered data.
# ## ----------------------------------------------------------------
# ## First you will need to enter the file path of the Sage report you want to filter

# %%
import ipywidgets as widgets
from IPython.display import display, HTML
import os
import pandas as pd
import io
import format_sheet as fs

period = 0
frequency = 0
report = 0
output = 0
asset_totals = ''
quarter = 0
specific_quarter = ''


# %%
def download_graph(df, name) :
    global report
    global asset_totals
    if (report == 1) :
        data = df.total_assets(asset_totals)
    elif (report == 2) :
        data = df.full_lists()
        
    df1 = pd.concat(data[0], axis=1)
    df.graph(df1, name, data[1]) 


# %%
def download_excel(df, name) :
    global report
    if (report == 1) :
        data = df.total_assets(asset_totals)
        df.multiple_dfs(data[0], name, name + '.xlsx', 1)
    elif (report == 2) :
        data = df.full_lists()
        df.df_tabs(data[0], data[1], name + '.xlsx')

# %%
def filename(df) :
    lbl1= widgets.Label(value='Enter the name you would like to save the file as:')
    txt1 = widgets.Text()
    button1 = widgets.Button(description = "Submit")

    h1 = widgets.HBox([txt1, button1])

    def on_button_clicked(b):
        global output
        if os.path.exists(txt1.value + '.xlsx'):
            display(HTML("<style>.red_label { color:red }</style>"))
            l = widgets.Label(value="File already exists")
            l.add_class("red_label")
            display(l)
            filename(df)
        elif output == 1 :
            download_excel(df, txt1.value)
            button1.disabled = True
            txt1.disabled = True
        elif output == 2 :
            download_graph(df, txt1.value)
            button1.disabled = True
            txt1.disabled = True
        elif output == 3 :
            download_excel(df, txt1.value)
            download_graph(df, txt1.value)
            button1.disabled = True
            txt1.disabled = True
        
        
    display(lbl1)
    display(h1)
    button1.on_click(on_button_clicked)

# %%
def full_list_output(df) :
    lbl1 = widgets.Label(value='Select the type of output you want to have:')
    button1 = widgets.Button(description = "Excel Sheet")
    button2 = widgets.Button(description = "Graph")
    button3 = widgets.Button(description = "Excel Sheet and Graph")

    h1 = widgets.HBox([button1, button2, button3])

    if (report == 2) :
        button2.disabled = True
        button3.disabled = True

    def disable_buttons() :
        button1.disabled = True
        button2.disabled = True
        button3.disabled = True
    def on_button_clicked1(b):
        global output
        output = 1
        filename(df)
        button1.button_style = "success"
        disable_buttons()
    def on_button_clicked2(b):
        global output
        output = 2
        filename(df)
        button2.button_style = "success"
        disable_buttons()
    def on_button_clicked3(b):
        global output
        output = 3
        filename(df)
        button3.button_style = "success"
        disable_buttons()

    display(lbl1)
    display(h1)
    button1.on_click(on_button_clicked1)
    button2.on_click(on_button_clicked2)
    button3.on_click(on_button_clicked3)

# %%
def asset_totals(df) :
    lbl1 = widgets.Label(value='Select the type of data you want to show:')
    button1 = widgets.Button(description = "Asset Totals")
    button2 = widgets.Button(description = "Depreciation Totals")
    button3 = widgets.Button(description = "Both")

    h1 = widgets.HBox([button1, button2, button3])
    
    def disable_buttons() :
        button1.disabled = True
        button2.disabled = True
        button3.disabled = True
    def on_button_clicked1(b):
        global asset_totals
        asset_totals = 'assets'
        full_list_output(df)
        button1.button_style = "success"
        disable_buttons()
    def on_button_clicked2(b):
        global asset_totals
        asset_totals = 'depreciation'
        full_list_output(df)
        button2.button_style = "success"
        disable_buttons()
    def on_button_clicked3(b):
        global asset_totals
        asset_totals = 'both'
        full_list_output(df)
        button3.button_style = "success"
        disable_buttons()

    display(lbl1)  
    display(h1)
    button1.on_click(on_button_clicked1)
    button2.on_click(on_button_clicked2)
    button3.on_click(on_button_clicked3)

# %%
def report_type(df, initial_df) :
    lbl1 = widgets.Label(value='Select the type of information you would like to show:')
    button1 = widgets.Button(description = "Asset Totals")
    button2 = widgets.Button(description = "Full Lists")
    button3 = widgets.Button(description = "Go Back")

    h1 = widgets.HBox([button1, button2, button3])


    def disable_buttons() :
        button1.disabled = True
        button2.disabled = True
        button3.disabled = True
    def on_button_clicked1(b):
        global report
        report = 1
        asset_totals(df)
        button1.button_style = "success"
        disable_buttons()
    def on_button_clicked2(b):
        global report
        report = 2
        full_list_output(df)
        button2.button_style = "success"
        disable_buttons()
    def on_button_clicked3(b):
        handler(df, initial_df)
        disable_buttons()

    display(lbl1)
    display(h1)
    button1.on_click(on_button_clicked1)
    button2.on_click(on_button_clicked2)
    button3.on_click(on_button_clicked3)

# %%
def specific (df, initial_df) :
    lbl1 = widgets.Label(value='Select the specific reporting period you would like:')
    lbl2 = widgets.Label(value='For quarterly reporting, please enter the year then quarter (e.g. 2022Q2)')
    text = widgets.Text()
    button1 = widgets.Button(description = "Submit")
    button2 = widgets.Button(description = "Go Back")

    h1 = widgets.HBox([button1, button2])

    def on_button_clicked1(b):
        button1.disabled = True
        button2.disabled = True
        valid = df.specific(text.value)
        if valid == True :
            report_type(df, initial_df)
        else :
            display(HTML("<style>.red_label { color:red }</style>"))
            l = widgets.Label(value="Invalid Period")
            l.add_class("red_label")
            display(l)
            specific(df, initial_df)
    def on_button_clicked2(b):  
        handler(df, initial_df)
        button1.disabled = True
        button2.disabled = True


    display(lbl1)
    display(lbl2)
    display(text)
    display(h1)
    button1.on_click(on_button_clicked1)
    button2.on_click(on_button_clicked2)

# %%
def same_quarter(df, initial_df) :
    lbl1 = widgets.Label(value='Select the type of information you would like to show:')
    button1 = widgets.Button(description = "Quarter 1")
    button2 = widgets.Button(description = "Quarter 2")
    button3 = widgets.Button(description = "Quarter 3")
    button4 = widgets.Button(description = "Quarter 4")
    button5 = widgets.Button(description = "Go Back")

    h1 = widgets.HBox([button1, button2, button3, button4, button5])


    def disable_buttons() :
        button1.disabled = True
        button2.disabled = True
        button3.disabled = True
        button4.disabled = True
        button5.disabled = True
    def on_button_clicked1(b):
        global quarter
        quarter = 1
        df.same_quarter('Q1')
        report_type(df, initial_df)
        button1.button_style = "success"
        disable_buttons()
    def on_button_clicked2(b):
        global quarter
        quarter = 2
        df.same_quarter('Q2')
        report_type(df, initial_df)
        button2.button_style = "success"
        disable_buttons()
    def on_button_clicked3(b):
        global quarter
        quarter = 3
        df.same_quarter('Q3')
        report_type(df, initial_df)
        button3.button_style = "success"
        disable_buttons()
    def on_button_clicked4(b):
        global quarter
        quarter = 4
        df.same_quarter('Q4')
        report_type(df, initial_df)
        button4.button_style = "success"
        disable_buttons()
    def on_button_clicked5(b):
        handler(df, initial_df)
        disable_buttons()

    display(lbl1)
    display(h1)
    button1.on_click(on_button_clicked1)
    button2.on_click(on_button_clicked2)
    button3.on_click(on_button_clicked3)
    button4.on_click(on_button_clicked4)
    button5.on_click(on_button_clicked5)

# %%
def handler(period_chosen, inital_df) :
    global period
    lbl1 = widgets.Label(value='Select the number of periods you would like to show:')
    if (period == 1) :        
        button1 = widgets.Button(description = "Current Year")
        button2 = widgets.Button(description = "Past 2 Years")
        button3 = widgets.Button(description = "Specific Year")
        button4 = widgets.Button(description = "Go Back")
        h1 = widgets.HBox([button1, button2, button3, button4])
    elif (period == 2) :
        button1 = widgets.Button(description = "Current Quarter")
        button2 = widgets.Button(description = "Past 4 Quarters")
        button3 = widgets.Button(description = "Specific Quarter")
        button5 = widgets.Button(description = "Same Quarter")
        button4 = widgets.Button(description = "Go Back")
        h1 = widgets.HBox([button1, button2, button3, button5, button4])



    def disable_buttons() :
        button1.disabled = True
        button2.disabled = True
        button3.disabled = True
        button4.disabled = True
        if period == 2 :
            button5.disabled = True
    def on_button_clicked1(b):
        global frequency
        frequency = 1
        period_chosen.recent()
        report_type(period_chosen, inital_df)
        button1.button_style = "success"
        disable_buttons()
    def on_button_clicked2(b):
        global frequency
        frequency = 2
        period_chosen.reduce_df()
        report_type(period_chosen, inital_df)
        button2.button_style = "success"
        disable_buttons()
    def on_button_clicked3(b):
        global frequency
        frequency = 3
        specific(period_chosen, inital_df)
        button3.button_style = "success"
        disable_buttons()
    def on_button_clicked4(b):
        button4.button_style = "success"
        period_input(inital_df)
        disable_buttons()
    def on_button_clicked5(b):
        global frequency
        frequency = 4
        same_quarter(period_chosen, inital_df)
        button5.button_style = "success"
        disable_buttons()
        
    


    display(lbl1)
    display(h1)
    button1.on_click(on_button_clicked1)
    button2.on_click(on_button_clicked2)
    button3.on_click(on_button_clicked3)
    button4.on_click(on_button_clicked4)
    if (period == 2) :
        button5.on_click(on_button_clicked5)

# %%
def period_input(df) :
    lbl1 = widgets.Label(value='Select the period you would like to filter to:')
    button1 = widgets.Button(description = "Fiscal Year")
    button2 = widgets.Button(description = "Quarter")

    h1 = widgets.HBox([button1, button2])

    def disable_buttons() :
        button1.disabled = True
        button2.disabled = True
    def on_button_clicked1(b):
        global period
        period = 1
        handler(fs.year(df), df)
        button1.button_style = "success"
        disable_buttons()
    def on_button_clicked2(b):
        global period 
        period = 2
        handler(fs.quarter(df), df)
        button2.button_style = "success"
        disable_buttons()

    display(lbl1)
    display(h1)
    button1.on_click(on_button_clicked1)
    button2.on_click(on_button_clicked2)


# %%
def path_input() :
    lbl1 = widgets.Label(value='Path to file:')
    display(lbl1)
    path = widgets.Text()
    display(path)

    # uploader = widgets.FileUpload(accept='.xls,.xlsx,.csv')
    # display(uploader)
    # print(uploader.value)
    # uploaded_file = uploader.value

    # pd.read_excel(io.BytesIO(uploaded_file.content))

    button1 = widgets.Button(description='find path')
    display(button1)

    def on_button_clicked(b):
        if os.path.exists(path.value):
            button1.disabled = True
            path.disabled = True
            df = pd.read_excel(path.value)
            period_input(df)
            return 
        else:
            display(HTML("<style>.red_label { color:red }</style>"))
            l = widgets.Label(value="File does not exist. Please try again.")
            l.add_class("red_label")
            display(l)
            return path_input()

    button1.on_click(on_button_clicked)

path_input()








# %%



