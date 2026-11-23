# MATMIS

### Materials Management Information System

MATMIS is a Python-based inventory analysis and reporting application
designed to support inventory visibility and stock evaluation.

The application enables users to analyze inventory values,
identify critical stock levels, perform ABC classification,
and generate downloadable reports.

## Live Application

https://matmis.streamlit.app
## Features

- Inventory value calculation and stock monitoring
- Critical stock identification and shortage cost analysis
- ABC inventory classification based on annual consumption value
- Interactive Pareto chart
- Product-level inventory filtering
- Excel and CSV data upload with validation
- Downloadable Excel analysis reports
- Built-in user guide

## Technologies Used

- Python
- Pandas – Data processing and analysis
- Streamlit – Interactive web application
- Plotly – Data visualization
- OpenPyXL – Excel report generation

## How to Use

1. Open the [MATMIS Live Application](https://matmis.streamlit.app).
2. Click "Örnek veri setiyle deneyin" to explore the application using sample inventory data.
3. Alternatively, download the Excel or CSV template from the "Stok Dosyası Yükle" section.
4. Fill in the required inventory fields:
   - `ad`: Product name
   - `stok`: Current stock quantity
   - `birim_fiyat`: Unit price (TRY)
   - `kritik_stok`: Critical stock threshold
   - `yillik_tuketim`: Annual consumption quantity
5. Upload the completed file.
6. Explore the inventory dashboard, critical stock analysis, and ABC classification.
7. Download the generated Excel report or export individual tables and the Pareto chart.

For detailed instructions, visit the "Kullanım Kılavuzu" section inside the application.

## Development Status

**MATMIS V1.0 – Initial Release**

MATMIS V1.0 provides core inventory analysis, data validation, and reporting functionalities.

The application is under active development. Future versions may include:

- Dynamic safety stock calculations
- Reorder point analysis
- Inventory scenario simulations
- Advanced inventory optimization models

These features are planned improvements and are not available in V1.0.
