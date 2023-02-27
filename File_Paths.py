import pathlib

### Current File Path
current = pathlib.Path(__file__).parent.resolve()

### Data File for tickers
current_str = str(pathlib.Path(current))
input_path = current_str + "\InputFile" + "\Tickers.csv"
data_path = current_str + "\Data"
output_path = current_str + "\Output"