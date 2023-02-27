import pandas as pd
import xlsxwriter
from win32com.client import DispatchEx
from openpyxl import load_workbook
import itertools
import File_Paths


def data_processing():
    stock_tickers = pd.read_csv(File_Paths.data_path)['Ticker Number'].tolist()
    df_ls = []
    main_df = pd.read_excel(File_Paths.data_path)

    for i in stock_tickers[1:]:
        new_df = pd.read_excel(File_Paths.data_path, sheet_name = i)
        main_df = pd.concat([main_df, new_df], ignore_index=True)
    
    ### Restructuring data
    main_df = main_df.pivot(index = 'ID', columns = 'id().ORIG_IDS', values = 'id().POSITION')
    main_df = main_df.fillna(0)

    ### Storing the shape
    size = main_df.shape

    ### Data that has been parsed (I placed it in output folder)
    main_df.to_excel(File_Paths.output_path + '\\' + 'clean_data.xlsx', sheet_name = 'Main')


def First_order(shock, base_shock, threshold, data, stock_tickers):
    data = data.set_index('ID')
    top_2_stocks = {}
    first_order_outcome = pd.DataFrame
    ### Finding each stock's market value
    for i in stock_tickers:
        top_2_stocks[i] = []
        new_name = 'Mv' + i
        data[new_name] = data[i] * data['Price']
        data = data.fillna(0)
        ### Shock top 2 names
        before = data[new_name].sum()
        top_2_value = data[new_name].nlargest(2).tolist()
        after = 0
        for ind,row in data[new_name].tolist():
            if row in top_2_value:
                top_2_stocks[i].append((ind,row))
                after = after + row * shock
            else:
                after = after + row
        first_order_outcome = first_order_outcome.append({'Hedge Fund Ticker': i,
                                                          'Before': before,
                                                          'After': after,
                                                          'Change': (before-after)/before},
                                                          ignore_index = True )
    
    ### Making all dataframe to come together
    data = data.reset_index()
    data = pd.concat([data,first_order_outcome], axis = 1)
    ### Printing shocked positions of those that are above the threshold
    first_order_outcome = first_order_outcome.set_index("Hedge Fund Ticker")

    hit_first_order = []

    ### Generating the columns which were hit in the first shock
    for ind, row in first_order_outcome['Change'].iteritems():
        if row > threshold:
            hit_first_order.append(ind)
            data['ShockedMV' + ind] = data['MV' + ind].apply(lambda x: x * shock
                                                             if x in list(sum(top_2_stocks[ind], ()))
                                                             else x * base_shock)
    return (data, hit_first_order)

### Second Order
def Second_order(Shock, Threshold, data, first_order, all_tickers):
    ### Second Order Impact
    for i in first_order:
        ### Creating a dictionary of stock and shock due to first order
        data['First Order Shoc for' + i] = 1 - (data['Mv' + i] - data['ShockedMV' + i]) / data['MV' + i]
        data = data.fillna(0)
        first_shock = dict(zip(data['ID'], data['First Order Shock for' + i]))
        all_tickers.remove(i)
        data = data.set_index('ID')
        Second_order_outcome = pd.DataFrame(columns = ['Hedge Fund Ticker', 'Before', 'After', 'Change'])
        for i in all_tickers:
            before = data[i].sum()
            after = 0
            for key,val in first_shock.items():
                for ind, row in data['MV' + i].iteritems():
                    if key == ind:
                        if val == 0:
                            after += row
                        else:
                            after += val * row
            Second_order_outcome = Second_order_outcome.append({'Hedge Fund Ticker':i,
                                                                'Before' : before, 
                                                                'After': after, 
                                                                'Change': (before-after)/before},
                                                                ignore_index =True)
            for ind, row in Second_order_outcome['Change'].iteritems():
                if row > Threshold:
                    hit_second_order.append(ind)
                    data['ShockedMV2' + ind] = data['Mv'+ind].apply(lambda x: x* Shock
                                                                    if x in list(sum(top_2_stocks[ind], ()))
                                                                    else x * base_shock)
            data.to_excel(File_Paths.output_path + '\\' + '2ndorderoutput' + '\\' + str(i) + '_Second_Order.xlsx')


