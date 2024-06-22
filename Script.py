import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.linear_model import LinearRegression
from sklearn.metrics import mean_absolute_error


file_path = '\\patch\\data\\PharmaDrugSales.xlsx'
data = pd.read_excel(file_path)


data['Time'] = pd.to_datetime(data['Time'], errors='coerce')
data.dropna(subset=['Time'], inplace=True)  


data['Year'] = data['Time'].dt.year
data['Month'] = data['Time'].dt.month
data['Date'] = data['Time'].dt.day
data['Hour'] = data['Time'].dt.hour
data['Day'] = data['Time'].dt.day_name()


for drug in data.columns[6:]:
    
    lag_column = f'Lag_{drug}'
    if lag_column not in data.columns:
        data[lag_column] = data[drug].shift(1)

    
    data.dropna(inplace=True)

    
    X = data[['Hour', 'Day', lag_column]]  
    y = data[drug]  

    
    X = pd.get_dummies(X, columns=['Day'], drop_first=True)

    
    X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)

    
    model = LinearRegression()
    model.fit(X_train, y_train)

    
    y_pred = model.predict(X_test)

    
    mae = mean_absolute_error(y_test, y_pred)
    print(f'\nMean Absolute Error for {drug}: {mae}')

    
    comparison = pd.DataFrame({'Actual': y_test, 'Predicted': y_pred})

    

    output_file_path = f'patch\\data\\prediksi\\{drug}_Predictions.xlsx'
    comparison.to_excel(output_file_path, index=False)
    print(f'Predictions for {drug} have been saved to {output_file_path}')


yearly_sales = data.groupby(['Year'])[data.columns[6:]].sum()


ranked_sales = yearly_sales.rank(axis=1, method='min', ascending=False)
for drug in ranked_sales.columns:
    yearly_sales[f'{drug}_Rank'] = ranked_sales[drug].astype(int)


yearly_sales_file_path = '\\patch\\data\\perbandingan\\YearlyDrugSalesComparison.xlsx'
yearly_sales.to_excel(yearly_sales_file_path)
print(f'Yearly drug sales comparison has been saved to {yearly_sales_file_path}')
