import pandas as pd
import File_Paths
import os
from itertools import islice

### Creating the DataFrame
def Data_Collector():
    coll = {}
    ticks = {}
    ticker_path = File_Paths.input_path
    ## Open the Excel
    data = pd.read_csv(ticker_path)
    tickers = data['Ticker Number'].tolist()

    ## Populate the Dictionary
    for i in tickers:
        coll[i] = []
        ticks[i] = []
    
    ## Taking the data of what we require only
    for i in tickers:
        df = pd.read_excel(File_Paths.data_path + '\\' + 'Data.xlsx', sheet_name= i)
        ## Iterate through the data frame Filter through
        for ind, row in df.iterrows():
            if row['id().POSITION'] != 0 and row['id().POSITION'] != 'nan':
                coll[i].append((row['ID'], row['id().POSITION'], row['id().POSITION_CHANGE']))
                ticks[i].append(row['ID'])
    return (coll, ticks)

## Finding the top holders in terms of commonality
def common_holding_freq(top):
    combined_dic = Data_Collector()
    ## Base on Frequency
    freq_tick = {}
    for key, val in combined_dic[0].items():
        freq_tick[key] = [i[0] for i in val]
    
    tick = freq_tick.values()
    tick_list = []
    for i in list(tick):
        tick_list += i
    listed_by_order = ranker1(tick_list)
    if top == 'ALL':
        return list(listed_by_order.items())
    else:
        top_n_entries = list(listed_by_order.items())[:int(top)]
        return top_n_entries

### Finding the top holders in terms of volume
def common_holding_no_share(top):
    combined_dic = Data_Collector()
    ## Base on holding of shares
    vol_tick = {}
    for key, val in combined_dic[0].items():
        for i in val:
            if i[0] not in vol_tick:
                vol_tick[i[0]] = i[1]
            else:
                vol_tick[i[0]] += i[1]  ## returns a ticker with the total volume of all the funds chosen

    ## Sort base on the values
    dic_by_order = dict(sorted(vol_tick.items(), key = lambda item: item[1], reverse =True))
    if top == 'ALL':
        return list(dic_by_order.items())
    else:
        top_n_entries = list(dic_by_order.items())[:int(top)]
        return top_n_entries

## Finding the top changes in terms of volume changes in last filing
def common_holding_vol_change(top):
    combined_dic = Data_Collector()
    ## Base on net change, does not account for sell/buy that the fund takes
    vol_change_tick = {}
    for key, val in combined_dic[0].items():
        for i in val:
            if i[0] not in vol_change_tick:
                vol_change_tick[i[0]] = abs(i[2])
            else:
                vol_change_tick[i[0]] += abs(i[2])
    ## Sort Base on the value of greatest change
    dic_by_order = dict(sorted(vol_change_tick.items(), key = lambda item: item[1], reverse = True))
    if top == 'ALL':
        return list(dic_by_order.items())
    else:
        top_n_entries = list(dic_by_order.items())[:int(top)]
        return top_n_entries

## Sorting mechansim to rank for frequency
def ranker1(lst):
    word_count = {}
    for ticks in lst: 
        if ticks in word_count: 
            word_count[ticks] += 1
        else: 
            word_count[ticks] = 1
    
    popular_ticks = dict(sorted(word_count.items(), key = lambda item: item[1], reverse = True))
    return popular_ticks

## Side work to include the names of the fund that has a ticker
def Extract_ticker_name():
    combined_data = Data_Collector()
    names = list(common_holding_freq('ALL').keys()
    
    new_dic = {}
    side_dic = {}
    df = pd.DataFrame({'ticker':names})
    for i in names:
        new_dic[i] = []
    for k in names:
        for key,val in combined_data.items():
            for i in val:
                if k in i:
                    new_dic[k].append([key, i[1]])
    for key, val in new_dic.items():
        new_dic[key] = sorted(val, key = lambda tup: tup[1])
    for key, val in new_dic.items():
        flat_list = []
        for sublist in reversed(val):
            for item in sublist:
                flat_list.append(item)
            new_dic[key] = flat_list
    
    total = []
    for key, val in new_dic.items():
        total.append([key,] + val)
    main_df = pd.DataFrame()
    for i in total:
        df = pd.DataFrame([i])
        main_df = pd.concat([main_df, df])
    
    main_df.to_excel(r'C:\Users\sng987,,,,')
    print(main_df.head())
    