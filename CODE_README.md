# Penjelasan Kode

## Library yang Digunakan
1. `import pandas as pd`
    : Library Pandas digunakan untuk membaca file dan memanipulasi data.

2. `from sklearn.model_selection import train_test_split`
    : Mengimpor fungsi `train_test_split` dari `sklearn` untuk membagi data menjadi set pelatihan dan pengujian.

3. `from sklearn.linear_model import LinearRegression`
    : Mengimpor model regresi linier dari `sklearn` untuk membuat model prediksi.

4. `from sklearn.metrics import mean_absolute_error`
    : Mengimpor `mean_absolute_error` dari `sklearn` untuk menghitung kesalahan prediksi.

## Membaca dan Memanipulasi Data
5. `file_path = '\\path\\data\\PharmaDrugSales.xlsx'`
    : Mendefinisikan jalur file data Excel yang akan dimuat.

6. `data = pd.read_excel(file_path)`
    : Membaca data dari file Excel ke dalam DataFrame pandas.

7. `data['Time'] = pd.to_datetime(data['Time'], errors='coerce')`
    : Mengonversi kolom 'Time' ke format datetime, dan menetapkan nilai yang tidak dapat dikonversi menjadi NaT (Not a Time).

8. `data.dropna(subset=['Time'], inplace=True)`
    : Menghapus baris yang memiliki nilai 'Time' tidak valid (NaT).

## Mengekstrak Informasi Waktu
9. `data['Year'] = data['Time'].dt.year`
    : Mengekstrak tahun dari kolom 'Time'.

10. `data['Month'] = data['Time'].dt.month`
    : Mengekstrak bulan dari kolom 'Time'.

11. `data['Date'] = data['Time'].dt.day`
    : Mengekstrak tanggal dari kolom 'Time'.

12. `data['Hour'] = data['Time'].dt.hour`
    : Mengekstrak jam dari kolom 'Time'.

13. `data['Day'] = data['Time'].dt.day_name()`
    : Mengekstrak nama hari dari kolom 'Time'.

## Rekayasa Fitur dan Pelatihan Model
14. `for drug in data.columns[6:]:`
    : Melakukan iterasi untuk setiap kolom yang merupakan kategori obat, dimulai dari kolom ke-6.

15. `    lag_column = f'Lag_{drug}'`
    : Mendefinisikan nama kolom baru untuk fitur lag.

16. `    if lag_column not in data.columns:`
17. `        data[lag_column] = data[drug].shift(1)`
    : Menambahkan kolom lag yang berisi nilai dari baris sebelumnya untuk kategori obat tersebut.

18. `    data.dropna(inplace=True)`
    : Menghapus baris yang memiliki nilai NaN akibat fitur lag.

19. `    X = data[['Hour', 'Day', lag_column]]`
    : Memilih kolom 'Hour', 'Day', dan kolom lag sebagai fitur (X).

20. `    y = data[drug]`
    : Menetapkan kolom kategori obat sebagai variabel target (y).

21. `    X = pd.get_dummies(X, columns=['Day'], drop_first=True)`
    : Mengonversi kolom 'Day' menjadi beberapa kolom biner menggunakan one-hot encoding, dan menghapus kolom pertama untuk menghindari dummy variable trap.

22. `    X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)`
    : Membagi data menjadi set pelatihan (80%) dan pengujian (20%) dengan random state 42 untuk replikasi yang konsisten.

23. `    model = LinearRegression()`
    : Menginisialisasi model regresi linier.

24. `    model.fit(X_train, y_train)`
    : Melatih model menggunakan data pelatihan.

## Evaluasi dan Penyimpanan Prediksi
25. `    y_pred = model.predict(X_test)`
    : Membuat prediksi pada data pengujian.

26. `    mae = mean_absolute_error(y_test, y_pred)`
    : Menghitung mean absolute error antara nilai sebenarnya dan prediksi.

27. `    print(f'\nMean Absolute Error for {drug}: {mae}')`
    : Mencetak nilai mean absolute error untuk kategori obat.

28. `    comparison = pd.DataFrame({'Actual': y_test, 'Predicted': y_pred})`
    : Membuat DataFrame untuk membandingkan nilai aktual dan prediksi.

29. `    output_file_path = f'\\path\\data\\prediksi\\{drug}_Predictions.xlsx'`
    : Mendefinisikan jalur file output untuk menyimpan prediksi.

30. `    comparison.to_excel(output_file_path, index=False)`
    : Menyimpan DataFrame perbandingan ke file Excel tanpa menyertakan indeks.

31. `    print(f'Predictions for {drug} have been saved to {output_file_path}')`
    : Mencetak pesan bahwa prediksi telah disimpan.

## Analisis Penjualan Tahunan
32. `yearly_sales = data.groupby(['Year'])[data.columns[6:]].sum()`
    : Mengelompokkan data berdasarkan tahun dan menjumlahkan penjualan untuk setiap kategori obat.

33. `ranked_sales = yearly_sales.rank(axis=1, method='min', ascending=False)`
    : Membuat DataFrame yang berisi peringkat penjualan obat untuk setiap tahun, dengan peringkat diurutkan dari terbesar ke terkecil.

34. `for drug in ranked_sales.columns:`
35. `    yearly_sales[f'{drug}_Rank'] = ranked_sales[drug].astype(int)`
    : Menambahkan kolom peringkat untuk setiap obat dalam DataFrame yearly_sales.

## Penyimpanan Data Penjualan Tahunan
36. `yearly_sales_file_path = '\\path\\data\\perbandingan\\YearlyDrugSalesComparison.xlsx'`
    : Mendefinisikan jalur file output untuk menyimpan perbandingan penjualan tahunan.

37. `yearly_sales.to_excel(yearly_sales_file_path)`
    : Menyimpan DataFrame yearly_sales ke file Excel.

38. `print(f'Yearly drug sales comparison has been saved to {yearly_sales_file_path}')`
    : Mencetak pesan bahwa perbandingan penjualan tahunan telah disimpan.
