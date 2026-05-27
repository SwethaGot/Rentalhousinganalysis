import os
import pandas as pd

path = 'Moving annual median rent by suburb and town - September quarter 2025.xlsx'


def inspect_workbook(path):
    print('file exists:', os.path.exists(path))
    xl = pd.ExcelFile(path)
    print('sheets:', xl.sheet_names)
    return xl


def parse_sheet_headers(df):
    """Extract the row with dates and the row with Count/Median labels."""
    date_row = df.iloc[1, 2:].astype(str).tolist()
    measure_row = df.iloc[2, 2:].astype(str).tolist()
    return date_row, measure_row


def parse_quarter_label(label):
    label = str(label).strip()
    if not label:
        return None
    dt = pd.to_datetime(label, format='%b %Y', errors='coerce')
    if pd.isna(dt):
        return None
    return pd.Period(dt, freq='Q')


def build_long_df(path, sheet):
    """Read one sheet and convert it from wide to long format."""
    df = pd.read_excel(path, sheet_name=sheet, header=None)
    date_row, measure_row = parse_sheet_headers(df)

    data = df.iloc[3:].reset_index(drop=True)
    data.columns = list(range(data.shape[1]))
    data = data.rename(columns={0: 'Region', 1: 'Suburb'})

    value_columns = [f"{d.strip()}||{m.strip()}" for d, m in zip(date_row, measure_row)]
    full_columns = ['Region', 'Suburb'] + value_columns
    full_columns = full_columns[: data.shape[1]]
    data.columns = full_columns

    melt_cols = [c for c in data.columns if c not in ['Region', 'Suburb']]
    long = data.melt(
        id_vars=['Region', 'Suburb'],
        value_vars=melt_cols,
        var_name='Date_Measure',
        value_name='Value',
    )
    long[['Date', 'Measure']] = long['Date_Measure'].str.split(r'\|\|', n=1, expand=True)
    long = long.drop(columns=['Date_Measure'])

    long['Date'] = long['Date'].astype(str).str.strip()
    long['Measure'] = long['Measure'].astype(str).str.strip()
    long['Value'] = pd.to_numeric(long['Value'].replace({'-': None, '': None}), errors='coerce')

    long['Date_period'] = long['Date'].map(parse_quarter_label)
    long['Date_quarter'] = long['Date_period'].astype(str)
    long['Date_dt'] = long['Date_period'].dt.to_timestamp(how='end')

    return long


def summarize_long_df(long):
    print('records:', len(long))
    print('unique measures:', sorted(long['Measure'].dropna().unique()))
    print('Date quarters:', long['Date_quarter'].min(), 'to', long['Date_quarter'].max())
    print('missing Value fraction:', long['Value'].isna().mean())

    median = long[long['Measure'].str.contains('Median', case=False, na=False)]
    if not median.empty:
        latest_period = median['Date_period'].max()
        print('latest quarter:', latest_period)
        print('latest top 5 medians:')
        print(
            median[median['Date_period'] == latest_period]
            .dropna(subset=['Value'])
            .sort_values('Value', ascending=False)
            .head(5)
        )

    if long['Date_period'].isna().any():
        print('warning: some dates failed to parse')


def main():
    xl = inspect_workbook(path)
    if not xl.sheet_names:
        return

    all_long = []
    for sheet in xl.sheet_names:
        print('\nExploring sheet:', sheet)
        long = build_long_df(path, sheet)
        long['PropertyType'] = sheet
        summarize_long_df(long)
        all_long.append(long)

    all_long = pd.concat(all_long, ignore_index=True)
    sample_path = 'sample_all_sheets_long.csv'
    all_long.head(200).to_csv(sample_path, index=False)
    clean_path = 'cleaned_rent_long.csv'
    all_long.to_csv(clean_path, index=False)
    print('saved sample to', sample_path)
    print('saved cleaned combined file to', clean_path)


if __name__ == '__main__':
    main()
