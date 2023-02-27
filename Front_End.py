import tkinter as tk
from tkinter import ttk
import tkinter.filedialog as filedialog
import os
import tkinter.messagebox
from tkinter import StringVar

LARGEFONT = ("Verdana", 35)
MEDIUMFONT = ("Arial" , 12, "bold")

class tkinterApp(tk.Tk):
    def __init__(self, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)
        container = tk.Frame(self)
        container.pack(side = 'top', fill ='both', expand  = True)
        container.grid_rowconfigure(0, weight=5)
        container.grid_columnconfigure(0, weight=5)
        self.frames = {}
        for F in (Data_Analysis, Data_Collector):
            frame = F(container, self)
            self.frames[F] = frame
            frame.grid(row = 0, column = 0, sticky = 'nsew')
        
        self.show_frame(Data_Collector)
    
    ## Display the current frame passed as parameter
    def show_frame(self, cont):
        frame = self.frames[cont]
        frame.tkraise()

# First window frame startpage
class Data_Colelctor(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        ## label of frame Layout 2
        label = ttk.Label(self, text = "Data Collector", font = LARGEFONT)
        # Putting the grid in its place by using the grid
        label.grid(row = 0, column = 1, padx= 10, pady = 10)
        button1 = ttk.Button(self, text = "Data Analysis", command= lambda: controller.show_frame(Data_Analysis))

        ## Putting the button in its place by using grid
        button1.grid(row=5, column =0, padx = 10, pady =10)

        ## Try to find the API for Bloomberg in API files
        self.input_path_bloomberg = ttk.Label(self, text = "Input Bloomberg xla:", font= MEDIUMFONT)
        self.input_entry_bloomberg = ttk.Entry(self, text = "", width=60)
        self.browse1 = ttk.Button(self, text = 'Browse', command = self.inputs)

        self.input_path_bloomberg.grid(row=1, column=0, padx=10, pady=10)
        self.input_entry_bloomberg.grid(row=1, column=1, padx=10, pady=10)
        self.browse1.grid(row=1, column=2, padx=5, pady=5)

        ## Button to start the extraction
        self.begin_button = tk.Button(self, text='Begin!', command=self.begin, fg = 'red')
        self.begin_button.grid(row=5, column=5, padx=5, pady=5)

    def inputs(self):
        ## This is for Bloomberg
        self.input_path_bloomberg = tk.filedialog.askopenfilename()
        self.input_entry_bloomberg.delete(1, tk.END) ## Remove current text in entry
        self.input_entry_bloomberg.insert(0, self.input_path_bloomberg) ## Re insert the path 

    def begin(self):
        self.input_directory1 = self.input_entry_bloomberg.get()
        ### Check if user inputs a valid input
        if len(self.input_directionary1) == 0:
            return self.Error()

        self.input_directory1 = self.input_entry_bloomberg.get()
        self.path1 = os.path.dirname(os.path.realpath(self.input_directory1))

        import DownloadData
        ### Input the Data Collecting the file here ###
        DownloadData.write_BB_query_in_excel(self.input_directory1)
        value = tk.messagebox.askokcancel(message = 'Completed. Want to try again?')
        if value == False:
            self.master.destroy()
    
    def Error(self):
        value = tk.messagebox.askretrycancel('Input Error', "Re-Input file paths and Try Again")
        if value == False:
            self.destroy()

## Secondary window frame page1
class Data_Analaysis(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        label = ttk.Label(self, text = 'Data Analaysis', font = LARGEFONT)
        label.grid(row=0, column=2, padx=10, pady=10)
        ## Layout 2
        button1 = ttk.Button(self, text='Data Collector', command= lambda: controller.show_frame(Data_Analaysis))

        ## Putting the button in its place by using grid 
        button1.grid(row=5, column=1, padx=10, pady=10)

        ## This is for checkbox of the Data Analysis report that someone wants to have
        self.label_Choice = ttk.Label(self, text= 'Choose Data Analaysis Tools:', font=MEDIUMFONT)
        self.label_Choice.grid(row=1, column=1, padx=10, pady=10)

        ## Creating the Analysis options
        self.Checkbutton1 = StringVar(self)
        self.Checkbutton2 = StringVar(self)
        self.Checkbutton3 = StringVar(self)

        ### Set the inital checkbox to be off
        self.Checkbutton1.set('OFF')
        self.Checkbutton2.set('OFF')
        self.Checkbutton3.set('OFF')

        ### The button and layout
        self.CheckButton1 = tk.Checkbutton(self, text = 'Common Holding Base on Frequency', onvalue= "common_holding_freq", offvalue= 0, variable = self.Checkbutton1)
        self.Checkbutton2 = tk.Checkbutton(self, text = 'Common Holding Base on Number Of Shares', onvalue = 'common_holding_no_share', offvalue = 0, variable  = self.Checkbutton2)
        self.Checkbutton3 = tk.Checkbutton(self, text = 'Common Holding absolute changes in volume', onvalue= 'common_holding_vol_change', offvalue=0, variable = self.Checkbutton3)

        self.CheckButton1.grid(row=2, column=1, padx=10, pady=10)
        self.Checkbutton2.grid(row=2, column=2, padx=10, pady=10)
        self.Checkbutton3.grid(row=2, column=3, padx=10, pady=10)

        ### Creating the number of top few Names
        self.Top = [5, 10, 15, 20, 25, 30, 'ALL']
        ### Creating the Number of firms option
        self.value_inside = tk.StringVar(self)
        ### Set the inital checkbox to choosen an option
        self.value_inside.set('Choose an option')
        ### Selection Option
        self.label_No = tk.Label(self, text = 'Top Number Of Firms:', font = MEDIUMFONT)
        self.label_No.grid(row=3, column=1, padx=10, pady=10)
        self.Timeline = tk.OptionMenu(self, self.value_inside, *self.Top)
        self.Timeline.grid(row=3, column=2, padx=0, pady=10)

        ### Creating a button to run the analysis
        self.begin_button = tk.Button(self, text = 'Start analyzing!', command = self.begin, fg = 'red')
        self.begin_button.grid(row=5, column=5, padx=10, pady=10)

        ### Adding Option to do Stock all holdings breakdown in excel
        self.label_get_others = ttk.Label(self, text = 'Include analysis on other holdings:', font = MEDIUMFONT)
        self.label_get_others.grid(row=4, column=1, padx=10, pady=10)
        ### Creating Selection for detailed analysis
        self.details_section = tk.StringVar(self) ## To store the current section
        ### Set the inital dropdown option to be off
        self.details.selection.set('Choose an option')
        ### Choices for dropdown
        self.details_option_list = [str(True), str(False)]
        self.details = tk.OptionMenu(self, self.details_selections, *self.details_option_list)
        self.details.grid(row=4, column=2, padx=10, pady=10)
    
    def begin(self):
        ### Check if a valid input was inserted
        check_data_analysis = [self.Checkbutton1.get(), self.Checkbutton2.get(), self.Checkbutton3.get()] ## in string or 0 if it is not chosen
        drop_no_firm = self.value_inside.get()
        drop_other_holding = self.details_section.get()
        if check_data_analysis == ['OFF', 'OFF', 'OFF'] or drop_no_firm == 'Choose an option' or drop_other_holding == 'Choose an option':
            return self.Error()
        import OutputFile_for_analysis
        OutputFile_for_analysis.make_Excel(check_data_analysis, drop_no_firm, drop_other_holding)
        value = tk.messagebox.askokcancel(message = 'Completed. Do you want to try again?')
        if value == False:
            self.master.destroy()
    
    def Error(self):
        value = tk.messagebox,askretrycancel("Input Error", 'Reselect and Try Again')
        if value == False:
            self.master.destroy()

# Driver Code
app = tkinterApp()
app.mainloop()


