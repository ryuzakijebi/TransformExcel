### Transformation Logic
- For each unique `Order No`, a new `HEADER` row is generated with the mapped fields.
- Following the `HEADER` row, all associated items are listed as `ITEM` rows with the mapped fields.
- A new `HEADER` row is created for each different `Order No`, and the process repeats.

#
- ### Input
The system reads an Excel file containing the following fields:

- **Posted Date**: (A)
- **Order Date**: (B)
- **DSR Code**: (C)
- **DSR Name**: (D)
- **Sales Route**: (E)
- **Customer Code**: (F)
- **Customer Name**: (G)
- **Address**: (H)
- **District**: (I)
- **Customer**: (J)
- **Turfview Code**: (K)
- **Order No**: (L)
- **Status**: (M)
- **Category**: (N)
- **Brand**: (O)
- **Sub Brand**: (P)
- **Variant**: (Q)
- **Packtype**: (R)
- **Item Code**: (S)
- **Item Name**: (T)
- **UOM**: (U)
- **Primary Price**: (V)
- **Unit Price**: (W)
- **Quantity**: (X)
- **Quantity (Liter)**: (Y)
- **Revenue**: (Z)
- **Amount**: (AA)
- **Discount**: (AB)
- **Return Quantity**: (AC)
- **Return Quantity (Liter)**: (AD)
- **Return Amount**: (AE)
- **Return Discount**: (AF)
- **External Document No.**: (AG)
#
### Output Format
The output Excel file includes three main sections: `HEADER`, `ITEM`, and `EXPENSE`, each with specific columns as outlined below.

#### `HEADER` Row (Row 1)
| Column                | Source Field         |
|-----------------------|----------------------|
| HEADER                | Fixed Text           |
| No Form               | `Order No` (L)       |
| Tgl Pesanan           | `Posted Date` (A)    |
| No Pelanggan          | `Customer Code` (F)  |
| Alamat                | `Address` (H)        |
| Kena PPN              | Fixed Null           |
| Total Termasuk PPN    | Fixed Null           |
| Diskon Pesanan (%)    | Fixed Null           |
| Diskon Pesanan (Rp)   | Fixed Null           |
| Keterangan            | Fixed Null           |
| Nama Cabang           | Fixed Null           |
| Pengiriman            | Fixed Null           |
| Tgl Pengiriman        | Fixed Null           |
| FOB                   | Fixed Null           |
| Syarat Pembayaran     | Fixed Null           |

#### `ITEM` Row (Row 2)
| Column                | Source Field         |
|-----------------------|----------------------|
| ITEM                  | Fixed Text           |
| Kode Barang           | `Item Code` (S)      |
| Nama Barang           | `Item Name` (T)      |
| Kuantitas             | `Quantity` (X)       |
| Satuan                | `UOM` (U)            |
| Harga Satuan          | `Unit Price` (W)     |
| Diskon Barang (%)     | `Discount` (AB)      |
| Diskon Barang (Rp)    | Fixed Null           |
| Catatan Barang        | Fixed Null           |
| Nama Dept Barang      | Fixed Null           |
| No Proyek Barang      | Fixed Null           |
| Nama Gudang           | Fixed Null           |
| ID Salesman           | `DSR Code` (C)       |
| Kustom Karakter 1     | Fixed Null           |
| Kustom Karakter 2     | Fixed Null           |
| Kustom Karakter 3     | Fixed Null           |

#### `EXPENSE` Row (Row 3)
| Column                | Description          |
|-----------------------|----------------------|
| EXPENSE               | Fixed Text           |
| No Biaya              | Fixed Null           |
| Nama Biaya            | Fixed Null           |
| Nilai Biaya           | Fixed Null           |
| Catatan Biaya         | Fixed Null           |
| Nama Dept Biaya       | Fixed Null           |
| No Proyek Biaya       | Fixed Null           |
| Kategori Keuangan 1   | Fixed Null           |
| Kategori Keuangan 2   | Fixed Null           |
| Kategori Keuangan 3   | Fixed Null           |
| Kategori Keuangan 4   | Fixed Null           |
| Kategori Keuangan 5   | Fixed Null           |
| Kategori Keuangan 6   | Fixed Null           |
| Kategori Keuangan 7   | Fixed Null           |
| Kategori Keuangan 8   | Fixed Null           |
| Kategori Keuangan 9   | Fixed Null           |


