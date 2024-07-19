import streamlit as st
import pandas as pd
import requests
import json
import re
import streamlit.components.v1 as components
from io import BytesIO
import numpy as np
from sklearn.linear_model import LinearRegression


# Set page configuration
st.set_page_config(layout="wide")

# CSS for header and footer
header_footer_css = """
    <style>
        .banner {
            position: relative;
            width: 100%;
            height: 150px;
            background-image: url('https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcQZJkPpcjFpNhfvdUcHdSytP1-tePz8v8X34Q&s');
            background-size: cover;
            background-position: center;
            color: transparent;
            text-align: center;
            line-height: 150px;
            transition: color 0.5s ease;
        }
        .banner:hover {
            color: white;
        }
        .logo {
            position: fixed;
            top: 0;
            right: 0;
            margin-right: 10px;
            margin-top: 10px;
            width: 100px;
            height: 150px;
        }
    </style>
    """
st.markdown(header_footer_css, unsafe_allow_html=True)
st.markdown("<div class='banner'></div>", unsafe_allow_html=True)

# Dashboard title and description
st.title("Financial Statements Dashboard")
st.image("https://img.freepik.com/premium-vector/business-statistics-financial-analytics-market-trend-analysis-vector-concept-illustration_92926-2486.jpg", caption="Trend Analysis and Data Insights", use_column_width=True)
st.sidebar.image("https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRViRfutvNG9i9GtCPAC6qiwcK_uIOvKU0QP-zvFl3iMHKpUvAvStpetXH8o2AQ_fA4tBg&usqp=CAU", use_column_width=False)

# Radio buttons for navigation
page = st.sidebar.radio("Select Sheets", ["Balance Sheet", "Profit & Loss", "KPI"])

# Initialize session state variables
if 'uploaded_files' not in st.session_state:
    st.session_state['uploaded_files'] = []

if 'balance_files_processed' not in st.session_state:
    st.session_state['balance_files_processed'] = False
if 'balance_data_frames' not in st.session_state:
    st.session_state['balance_data_frames'] = []
if 'balance_results' not in st.session_state:
    st.session_state['balance_results'] = {}
if 'balance_quarters_range' not in st.session_state:
    st.session_state['balance_quarters_range'] = None
if 'balance_terms' not in st.session_state:
    st.session_state['balance_terms'] = []

if 'profit_files_processed' not in st.session_state:
    st.session_state['profit_files_processed'] = False
if 'profit_data_frames' not in st.session_state:
    st.session_state['profit_data_frames'] = []
if 'profit_results' not in st.session_state:
    st.session_state['profit_results'] = {}
if 'profit_quarters_range' not in st.session_state:
    st.session_state['profit_quarters_range'] = None
if 'profit_terms' not in st.session_state:
    st.session_state['profit_terms'] = []

if 'include_forecast_bs' not in st.session_state:
    st.session_state['include_forecast_bs'] = False
if 'include_forecast_pl' not in st.session_state:
    st.session_state['include_forecast_pl'] = False

# Initialize session state variables
if 'kpi_uploaded_files' not in st.session_state:
    st.session_state['kpi_uploaded_files'] = []
if 'kpi_file_processed' not in st.session_state:
    st.session_state['kpi_file_processed'] = False
if 'kpi_data' not in st.session_state:
    st.session_state['kpi_data'] = None
if 'kpi_results' not in st.session_state:
    st.session_state['kpi_results'] = None
if 'kpi_fy_columns' not in st.session_state:
    st.session_state['kpi_fy_columns'] = None


def read_file(file, sheet_name):
    if file.type == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
        df = pd.read_excel(file, sheet_name=None)
        if sheet_name in df.keys():
            df = pd.read_excel(file, sheet_name=sheet_name)
            return df, f"{file.name}_{sheet_name}"
        else:
            st.error(f"The '{sheet_name}' sheet is not found in the uploaded file.")
            return None, file.name
    else:
        st.error("Unsupported file type. Please upload an Excel file.")
        return None, file.name

def extract_and_sort_quarters(df):
    quarter_pattern = re.compile(r'Q[1-4] \d{2}-\d{2}$')
    quarters = [col for col in df.columns if quarter_pattern.match(col)]
    quarters.sort(key=lambda x: (int(x.split()[1].split('-')[0]), x.split()[0][1]))
    return quarters

def process_dataframe(df, term, file_name):
    normalized_df = df.applymap(lambda x: x.strip().lower() if isinstance(x, str) else x)
    term_lower = term.lower()
    matching_rows = normalized_df[normalized_df.apply(lambda row: term_lower in row.values, axis=1)]
    if not matching_rows.empty:
        matching_rows.insert(0, 'File', file_name)
        return matching_rows
    return pd.DataFrame()

def generate_full_quarter_range(all_quarters, start_q, end_q):
    start_index = all_quarters.index(start_q)
    end_index = all_quarters.index(end_q) + 1  
    return all_quarters[start_index:end_index]

def filter_valid_quarters(quarters):
    valid_quarters = [q for q in quarters if not re.search(r'\.\d+', q)]
    return valid_quarters

def forecast_and_plot(df, term, selected_quarters, col, add_legend=False, show_forecast=True):
    combined_data = {}
    for quarter in selected_quarters:
        for idx, row in df.iterrows():
            if quarter in row and pd.notna(row[quarter]):
                combined_data[quarter] = row[quarter]
                break

    valid_quarters = list(combined_data.keys())
    valid_values = list(combined_data.values())
    quarters_numeric = np.arange(1, len(valid_quarters) + 1)
    
    if len(valid_values) > 1:
        model = LinearRegression()
        model.fit(quarters_numeric.reshape(-1, 1), valid_values)
        
        forecasts = []
        if show_forecast:
            for i in range(1, 4):
                next_quarter = len(valid_quarters) + i
                forecast = model.predict(np.array([[next_quarter]]))[0]
                forecasts.append(forecast)
        
        forecast_values = np.append(valid_values, forecasts)
        valid_quarters.extend(['Forecast1', 'Forecast2', 'Forecast3'] if show_forecast else [])
        
        colors = ['#75E2C9'] * len(valid_values) + (['#FFA500', '#800080', '#008080'] if show_forecast else [])
        
        highchart_html = f"""
        <script src="https://code.highcharts.com/highcharts.js"></script>
        <script src="https://code.highcharts.com/modules/zooming.js"></script>
        <div id="container_{term.replace(' ', '_')}" style="width:100%; height:400px;"></div>
        <script>
            Highcharts.chart('container_{term.replace(' ', '_')}', {{
                chart: {{
                    zoomType: 'x',
                    type: 'area'
                }},
                title: {{
                    text: 'Forecast for {term}'
                }},
                xAxis: {{
                    type: 'datetime',
                    categories: {json.dumps(valid_quarters)}
                }},
                yAxis: {{
                    title: {{
                        text: 'Value'
                    }}
                }},
                plotOptions: {{
                    series: {{
                        colorByPoint: true,
                        marker: {{
                            enabled: true
                        }},
                        fillColor: '#75E2C9'
                    }}
                }},
                series: [{{
                    name: '{term}',
                    data: {json.dumps([{'y': v, 'color': c} for v, c in zip(forecast_values.tolist(), colors)])},
                    color: '#75E2C9',
                    dashStyle: 'ShortDash'
                }}]
            }});
        </script>
        """
        with col:
            if add_legend and show_forecast:
                legend_html = """
                <div style="margin-top: 10px; border: 1px solid #ddd; padding: 10px; display: inline-block;">
                    <span style="color:#FFA500;">●</span> Forecast1
                    <span style="color:#800080; margin-left: 10px;">●</span> Forecast2
                    <span style="color:#008080; margin-left: 10px;">●</span> Forecast3
                </div>
                """
                st.markdown(legend_html, unsafe_allow_html=True)
            components.html(highchart_html, height=500)
    else:
        col.warning(f"Not enough data to forecast for {term}")




def balance_sheet_page():
    uploaded_files = st.sidebar.file_uploader("Choose Excel files for BS & PL", type='xlsx', accept_multiple_files=True, key='balance_files')

    if uploaded_files:
        st.session_state['uploaded_files'] = uploaded_files
        response = requests.post('https://financialstatementsforecast.azurewebsites.net/api/upload_bs?code=TVKPGisBDr9ja98dZqvDCrIuy32rZYSpoPsVIxrnlLH0AzFuNjkkrQ%3D%3D', files=[('files', (file.name, file, file.type)) for file in uploaded_files])
        
    if st.session_state['uploaded_files']:
        with st.sidebar.expander("View Uploaded Files"):
            for i, file in enumerate(st.session_state['uploaded_files']):
                col1, col2 = st.columns([8, 2])
                col1.write(file.name)
                if col2.button("X", key=f"remove_balance_{i}"):
                    st.session_state['uploaded_files'].pop(i)
                    st.rerun()

    if st.session_state['uploaded_files']:
        if st.sidebar.button('Process Files'):
            data_frames_info = [read_file(file, 'BS') for file in st.session_state['uploaded_files']]
            if all(df_info[0] is not None for df_info in data_frames_info):
                st.session_state['balance_data_frames'] = data_frames_info
                st.session_state['balance_files_processed'] = True
                st.rerun()

    if st.session_state.get('balance_files_processed'):
        quarters = set()
        for df_info in st.session_state['balance_data_frames']:
            df, _ = df_info
            quarters.update(extract_and_sort_quarters(df))
        quarters = sorted(quarters, key=lambda x: (int(x.split()[1].split('-')[0]), x.split()[0][1]))

        if quarters and st.session_state['balance_quarters_range'] is None:
            st.session_state['balance_quarters_range'] = (quarters[0], quarters[-1])
        
        if st.session_state['balance_quarters_range'] is not None:
            selected_quarter_range = st.select_slider("Select Quarter Range", options=quarters, value=st.session_state['balance_quarters_range'], key='balance_quarters_range')
            if selected_quarter_range:
                selected_quarters = generate_full_quarter_range(quarters, selected_quarter_range[0], selected_quarter_range[1])
            else:
                selected_quarters = []

            search_terms = ["Total non-current assets", "Total current assets", "Total assets", "Total Equity", "Total non-current liabilities", "Total current liabilities", "Total equity and liabilities"]
            selected_search_terms = st.multiselect("Select Search Terms", options=search_terms, default=search_terms, key='balance_terms')

            if st.button('Show Results', key='balance_show_results'):
                last_extracted_quarter = quarters[-1] if quarters else None
                st.session_state['include_forecast_bs'] = last_extracted_quarter == selected_quarters[-1] if selected_quarters else False
                results, sorted_quarters = aggregate_data_bs(selected_search_terms, st.session_state['balance_data_frames'], selected_quarters)
                st.session_state['balance_results'] = results
                st.session_state['sorted_balance_quarters'] = sorted_quarters

    if st.session_state.get('balance_results'):
        display_balance_results(st.session_state['balance_results'], st.session_state['sorted_balance_quarters'])

def aggregate_data_bs(search_terms, data_frames_info, selected_quarters):
    aggregate_results = {term: pd.DataFrame() for term in search_terms}
    all_quarters = []

    for df_info in data_frames_info:
        df, file_name = df_info
        if df is not None:
            current_quarters = extract_and_sort_quarters(df)
            filtered_quarters = [q for q in current_quarters if q in selected_quarters]
            all_quarters.extend(filtered_quarters)
            for term in search_terms:
                processed_df = process_dataframe(df, term, file_name)
                if not processed_df.empty:
                    available_columns = ['File'] + filtered_quarters
                    processed_df = processed_df.reindex(columns=available_columns, fill_value=None)
                    aggregate_results[term] = pd.concat([aggregate_results[term], processed_df], ignore_index=True)

    all_quarters = sorted(list(set(all_quarters)), key=lambda x: (int(x.split()[1].split('-')[0]), x.split()[0][1]))
    return aggregate_results, all_quarters




def display_balance_results(results, quarters):
    valid_quarters = filter_valid_quarters(quarters)
    terms = list(results.keys())
    num_terms = len(terms)
    
    include_forecast = st.session_state.get('include_forecast_bs', False)

    first_graph = True
    for i in range(0, num_terms, 2):
        cols = st.columns(2)
        
        if i < num_terms:
            with cols[0]:
                if not results[terms[i]].empty:
                    forecast_and_plot(results[terms[i]], terms[i], valid_quarters, cols[0], add_legend=first_graph, show_forecast=include_forecast)
                    first_graph = False
                else:
                    st.write(f"No data found for {terms[i]}")
        
        if i + 1 < num_terms:
            with cols[1]:
                if not results[terms[i + 1]].empty:
                    forecast_and_plot(results[terms[i + 1]], terms[i + 1], valid_quarters, cols[1], show_forecast=include_forecast)
                else:
                    st.write(f"No data found for {terms[i + 1]}")

    st.subheader("Balance Sheet Data Tables")
    for term, df in results.items():
        df = df.reindex(columns=['File'] + valid_quarters)
        st.write(f"Results for {term}")
        st.dataframe(df)

def profit_loss_page():
    uploaded_files = st.sidebar.file_uploader("Choose Excel files for BS & PL", type='xlsx', accept_multiple_files=True, key='profit_files')

    if uploaded_files:
        st.session_state['uploaded_files'] = uploaded_files
        response = requests.post('https://financialstatementsforecast.azurewebsites.net/api/upload_pl?code=97sBzFGS2wx3UlJPaovVxxkrn-JLPcOA_jXXwBYb6UKPAzFusqIkFg%3D%3D', files=[('files', (file.name, file, file.type)) for file in uploaded_files])
    
    if st.session_state['uploaded_files']:
        with st.sidebar.expander("View Uploaded Files"):
            for i, file in enumerate(st.session_state['uploaded_files']):
                col1, col2 = st.columns([8, 2])
                col1.write(file.name)
                if col2.button("X", key=f"remove_profit_{i}"):
                    st.session_state['uploaded_files'].pop(i)
                    st.rerun()
    
    if st.session_state['uploaded_files']:
        if st.sidebar.button('Process Files'):
            data_frames_info = [read_file(file, 'PL') for file in st.session_state['uploaded_files']]
            if all(df_info[0] is not None for df_info in data_frames_info):
                st.session_state['profit_data_frames'] = data_frames_info
                st.session_state['profit_files_processed'] = True
                st.rerun()

    if st.session_state.get('profit_files_processed'):
        quarters = set()
        for df_info in st.session_state['profit_data_frames']:
            df, _ = df_info
            quarters.update(extract_and_sort_quarters(df))
        quarters = sorted(quarters, key=lambda x: (int(x.split()[1].split('-')[0]), x.split()[0][1]))
        
        if quarters and st.session_state['profit_quarters_range'] is None:
            st.session_state['profit_quarters_range'] = (quarters[0], quarters[-1])

        if st.session_state['profit_quarters_range'] is not None:
            selected_quarter_range = st.select_slider("Select Quarter Range", options=quarters, value=st.session_state['profit_quarters_range'], key='profit_quarters_range')
            if selected_quarter_range:
                selected_quarters = generate_full_quarter_range(quarters, selected_quarter_range[0], selected_quarter_range[1])
            else:
                selected_quarters = []

            search_terms = ["Total Income", "Total expenses", "Profit before exceptional item and tax", "Profit for the year"]
            selected_search_terms = st.multiselect("Select Search Terms", options=search_terms, default=search_terms, key='profit_terms')

            if st.button('Show Results', key='profit_show_results'):
                last_extracted_quarter = quarters[-1] if quarters else None
                st.session_state['include_forecast_pl'] = last_extracted_quarter == selected_quarters[-1] if selected_quarters else False
                results, sorted_quarters, sorted_fy = aggregate_data_pl(selected_search_terms, st.session_state['profit_data_frames'], selected_quarters)
                st.session_state['profit_results'] = results
                st.session_state['sorted_profit_quarters'] = sorted_quarters
                st.session_state['sorted_profit_fy'] = sorted_fy

    if st.session_state.get('profit_results'):
        display_profit_results(st.session_state['profit_results'], st.session_state['sorted_profit_quarters'], st.session_state['sorted_profit_fy'])


def aggregate_data_pl(search_terms, data_frames_info, selected_quarters):
    aggregate_results = {term: pd.DataFrame() for term in search_terms}
    all_quarters = []
    all_fy = []

    for df_info in data_frames_info:
        df, file_name = df_info
        if df is not None:
            current_quarters = extract_and_sort_quarters(df)
            filtered_quarters = [q for q in current_quarters if q in selected_quarters]
            all_quarters.extend(filtered_quarters)
            
            fy_pattern = re.compile(r'FY \d{2}-\d{2}$')
            fy_columns = [col for col in df.columns if fy_pattern.match(col)]
            fy_columns = sorted(fy_columns, key=lambda x: int(x.split()[1].split('-')[0]))
            all_fy.extend(fy_columns)
            
            for term in search_terms:
                processed_df = process_dataframe(df, term, file_name)
                if not processed_df.empty:
                    if term == "Profit for the year":
                        available_columns = ['File'] + fy_columns
                    else:
                        available_columns = ['File'] + filtered_quarters
                    processed_df = processed_df.reindex(columns=available_columns, fill_value=None)
                    aggregate_results[term] = pd.concat([aggregate_results[term], processed_df], ignore_index=True)

    all_quarters = sorted(list(set(all_quarters)), key=lambda x: (int(x.split()[1].split('-')[0]), x.split()[0][1]))
    all_fy = sorted(list(set(all_fy)), key=lambda x: int(x.split()[1].split('-')[0]))
    return aggregate_results, all_quarters, all_fy


def display_profit_results(results, quarters, fy_columns):
    valid_quarters = filter_valid_quarters(quarters)
    terms = list(results.keys())
    num_terms = len(terms)

    last_extracted_quarter = valid_quarters[-1] if valid_quarters else None
    include_forecast = st.session_state.get('include_forecast_pl', False)

    first_graph = True
    for i in range(0, num_terms, 2):
        cols = st.columns(2)

        if i < num_terms:
            with cols[0]:
                if not results[terms[i]].empty:
                    show_forecast = include_forecast if terms[i] != "Profit for the year" else True
                    forecast_and_plot(results[terms[i]], terms[i], fy_columns if terms[i] == "Profit for the year" else valid_quarters, cols[0], add_legend=first_graph, show_forecast=show_forecast)
                    first_graph = False
                else:
                    st.write(f"No data found for {terms[i]}")

        if i + 1 < num_terms:
            with cols[1]:
                if not results[terms[i + 1]].empty:
                    show_forecast = include_forecast if terms[i + 1] != "Profit for the year" else True
                    forecast_and_plot(results[terms[i + 1]], terms[i + 1], fy_columns if terms[i + 1] == "Profit for the year" else valid_quarters, cols[1], show_forecast=show_forecast)
                else:
                    st.write(f"No data found for {terms[i + 1]}")

    st.subheader("Profit & Loss Tables")
    for term, df in results.items():
        if term == "Profit for the year":
            df = df.reindex(columns=['File'] + fy_columns)
        else:
            df = df.reindex(columns=['File'] + valid_quarters)
        st.write(f"Results for {term}")
        st.dataframe(df)

def kpi_page():
    uploaded_files = st.sidebar.file_uploader("Choose Excel files for KPI", type='xlsx', accept_multiple_files=True, key='kpi_files')

    if uploaded_files:
        for file in uploaded_files:
            if file not in st.session_state['kpi_uploaded_files']:
                st.session_state['kpi_uploaded_files'].append(file)
                response = requests.post('https://financialstatementsforecast.azurewebsites.net/api/upload_kpi?code=BQchLjkFc7rclRAPocLWiYhzehL3Hzqj1HKTr2VmHMENAzFugVBFvA%3D%3D', files={'file': (file.name, file, file.type)})

    if st.session_state['kpi_uploaded_files']:
        with st.sidebar.expander("View Uploaded Files"):
            for i, file in enumerate(st.session_state['kpi_uploaded_files']):
                col1, col2 = st.columns([8, 2])
                col1.write(file.name)
                if col2.button("X", key=f"remove_kpi_{i}"):
                    st.session_state['kpi_uploaded_files'].pop(i)
                    st.rerun()

    if st.session_state['kpi_uploaded_files']:
        if st.sidebar.button('Process KPI Files'):
            data_frames = []
            for file in st.session_state['kpi_uploaded_files']:
                data = pd.read_excel(file, sheet_name='PL')
                data_frames.append(data)
            st.session_state['kpi_data'] = pd.concat(data_frames, ignore_index=True)
            st.session_state['kpi_file_processed'] = True
            st.rerun()

    if st.session_state.get('kpi_file_processed'):
        data = st.session_state['kpi_data']
        if data is not None and not data.empty:
            required_kpis = [
                "Net Sales", "EBITDA", "PAT", "Net Worth", "Debt", "Debtors",
                "Cash", "EPS", "EBITDA Margin", "Net Profit Margin", "RoE/RONW",
                "RoCE", "EBIT", "Net Assets", "Gross Margin"
            ]

            # Extract fiscal years
            fy_columns = [col for col in data.columns if col.startswith('FY')]
            fy_columns.sort(key=lambda x: int(x.split('FY ')[1].split('-')[0]))

            # User selects the KPIs to display
            selected_search_terms = st.multiselect("Select Search Terms", options=required_kpis, default=required_kpis, key='kpi_terms')

            if st.button('Show Results', key='kpi_show_results'):
                results = aggregate_kpi_data(selected_search_terms, data, fy_columns)
                st.session_state['kpi_results'] = results
                st.session_state['kpi_fy_columns'] = fy_columns
                st.rerun()

            if st.session_state['kpi_results'] is not None and st.session_state['kpi_fy_columns'] is not None:
                display_kpi_results(st.session_state['kpi_results'], st.session_state['kpi_fy_columns'])


def aggregate_kpi_data(search_terms, data, fy_columns):
    aggregate_results = {term: pd.DataFrame() for term in search_terms}

    for term in search_terms:
        term_data = data[data.iloc[:, 0].str.strip().str.lower() == term.lower()]  # Ensure to strip and lowercase any leading/trailing whitespace
        if not term_data.empty:
            # Filter columns to only include FY columns
            term_data = term_data[['Unnamed: 0'] + fy_columns]
            aggregate_results[term] = term_data
    
    return aggregate_results


def forecast_and_plot_kpi(df, term, fy_columns):
    valid_years = fy_columns
    valid_values = [df[fy].values[0] for fy in fy_columns if fy in df.columns and pd.notna(df[fy].values[0])]

    if len(valid_values) > 1:
        model = LinearRegression()
        X = np.arange(1, len(valid_values) + 1).reshape(-1, 1)
        y = np.array(valid_values).reshape(-1, 1)
        model.fit(X, y)

        forecasts = []
        for i in range(1, 4):
            next_year = len(valid_values) + i
            forecast = model.predict(np.array([[next_year]]))[0][0]
            forecasts.append(forecast)

        forecast_values = np.append(valid_values, forecasts)
        valid_years.extend(['Forecast1', 'Forecast2', 'Forecast3'])

        colors = ['#7cb5ec'] * len(valid_values) + ['#e29375', '#e29375', '#e29375']

        bar_chart_html = f"""
        <script src="https://code.highcharts.com/highcharts.js"></script>
        <script src="https://code.highcharts.com/modules/exporting.js"></script>
        <script src="https://code.highcharts.com/modules/export-data.js"></script>
        <script src="https://code.highcharts.com/modules/accessibility.js"></script>

        <div id="container_{term.replace(' ', '_')}" style="width:100%; height:400px;"></div>
        <script>
        Highcharts.chart('container_{term.replace(' ', '_')}', {{
            chart: {{
                type: 'column'
            }},
            title: {{
                text: '{term}'
            }},
            xAxis: {{
                categories: {json.dumps(valid_years)},
                crosshair: true
            }},
            yAxis: {{
                min: 0,
                title: {{
                    text: 'Value'
                }}
            }},
            tooltip: {{
                headerFormat: '<span style="font-size:10px">{{point.key}}</span><table>',
                pointFormat: '<tr><td style="color:{{series.color}};padding:0">{{series.name}}: </td>' +
                    '<td style="padding:0"><b>{{point.y:.1f}}</b></td></tr>',
                footerFormat: '</table>',
                shared: true,
                useHTML: true
            }},
            plotOptions: {{
                column: {{
                    pointPadding: 0.2,
                    borderWidth: 0
                }}
            }},
            series: [{{
                name: '{term}',
                data: {json.dumps(forecast_values.tolist())},
                colorByPoint: true,
                colors: {json.dumps(colors)}
            }}]
        }});
        </script>
        """
        components.html(bar_chart_html, height=500)
    else:
        st.warning(f"Not enough data to forecast for {term}")



def display_kpi_results(results, fy_columns):
    for term, df in results.items():
        # st.subheader(f"Data for {term}")
        # st.write(df)  # Display the extracted data
        
        if not df.empty:
            forecast_and_plot_kpi(df, term, fy_columns)
        else:
            st.warning(f"Not enough data to forecast for {term}")



# Load the selected page
if page == "Balance Sheet":
    balance_sheet_page()
elif page == "Profit & Loss":
    profit_loss_page()
elif page == "KPI":
    kpi_page()