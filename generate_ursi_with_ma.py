import pandas as pd
import plotly.graph_objects as go
import json
from datetime import datetime

# Load the data
df = pd.read_pickle(r"C:\Users\minhdang\OneDrive - DRAGON CAPITAL\CodeVisual\Tai_Training\df_ohlcv_195stocks.pkl")

# Convert day column to datetime for proper sorting
df['date'] = pd.to_datetime(df['day'].str.replace('_', '-'))

# Sort by stock and date to ensure chronological order
df = df.sort_values(['stock', 'date'])

# Calculate previous close for each stock
df['prev_close'] = df.groupby('stock')['close'].shift(1)

# Determine if stock is advancing, declining, or unchanged
df['is_advancing'] = (df['close'] > df['prev_close']).astype(int)
df['is_declining'] = (df['close'] < df['prev_close']).astype(int)
df['is_unchanged'] = (df['close'] == df['prev_close']).astype(int)

# Remove rows where prev_close is NaN (first day for each stock)
df_clean = df.dropna(subset=['prev_close'])

# Calculate daily URSI
daily_stats = df_clean.groupby('date').agg({
    'is_advancing': 'sum',
    'is_declining': 'sum',
    'is_unchanged': 'sum'
}).reset_index()

# Rename columns
daily_stats.columns = ['date', 'advancing_stocks', 'declining_stocks', 'unchanged_stocks']

# Calculate URSI = (Advancing / (Advancing + Declining)) * 100
# Note: Unchanged stocks are excluded from the calculation
daily_stats['URSI'] = (daily_stats['advancing_stocks'] / 
                       (daily_stats['advancing_stocks'] + daily_stats['declining_stocks'])) * 100

# Calculate total stocks for reference
daily_stats['total_stocks'] = (daily_stats['advancing_stocks'] + 
                               daily_stats['declining_stocks'] + 
                               daily_stats['unchanged_stocks'])

# Sort by date
daily_stats = daily_stats.sort_values('date')

# Calculate some example moving averages for the Excel file
for period in [5, 10, 20, 50]:
    daily_stats[f'MA_{period}'] = daily_stats['URSI'].rolling(window=period).mean()

print(f"URSI Calculation Complete!")
print(f"Date range: {daily_stats['date'].min()} to {daily_stats['date'].max()}")
print(f"Total trading days: {len(daily_stats)}")
print(f"Average URSI: {daily_stats['URSI'].mean():.2f}")
print(f"Current URSI (latest): {daily_stats['URSI'].iloc[-1]:.2f}")
print(f"\nLatest day statistics:")
print(f"  - Advancing stocks: {daily_stats['advancing_stocks'].iloc[-1]}")
print(f"  - Declining stocks: {daily_stats['declining_stocks'].iloc[-1]}")
print(f"  - Unchanged stocks: {daily_stats['unchanged_stocks'].iloc[-1]}")
print(f"  - Total stocks: {daily_stats['total_stocks'].iloc[-1]}")

# Export to Excel with multiple sheets for better organization
excel_file = r"C:\Users\minhdang\OneDrive - DRAGON CAPITAL\CodeVisual\Tai_Training\ursi_analysis.xlsx"
with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
    # Sheet 1: Full URSI data with moving averages
    daily_stats.to_excel(writer, sheet_name='URSI_Daily', index=False)
    
    # Sheet 2: Summary statistics
    summary_data = {
        'Metric': [
            'Start Date',
            'End Date',
            'Total Trading Days',
            'Average URSI',
            'Median URSI',
            'Min URSI',
            'Max URSI',
            'Standard Deviation',
            'Days Above 70 (Overbought)',
            'Days Below 30 (Oversold)',
            'Days in Neutral Zone (30-70)',
            'Current URSI',
            'Current Advancing Stocks',
            'Current Declining Stocks',
            'Current Unchanged Stocks'
        ],
        'Value': [
            daily_stats['date'].min().strftime('%Y-%m-%d'),
            daily_stats['date'].max().strftime('%Y-%m-%d'),
            len(daily_stats),
            f"{daily_stats['URSI'].mean():.2f}",
            f"{daily_stats['URSI'].median():.2f}",
            f"{daily_stats['URSI'].min():.2f}",
            f"{daily_stats['URSI'].max():.2f}",
            f"{daily_stats['URSI'].std():.2f}",
            (daily_stats['URSI'] > 70).sum(),
            (daily_stats['URSI'] < 30).sum(),
            ((daily_stats['URSI'] >= 30) & (daily_stats['URSI'] <= 70)).sum(),
            f"{daily_stats['URSI'].iloc[-1]:.2f}",
            daily_stats['advancing_stocks'].iloc[-1],
            daily_stats['declining_stocks'].iloc[-1],
            daily_stats['unchanged_stocks'].iloc[-1]
        ]
    }
    summary_df = pd.DataFrame(summary_data)
    summary_df.to_excel(writer, sheet_name='Summary', index=False)
    
    # Sheet 3: Monthly averages
    daily_stats['year_month'] = daily_stats['date'].dt.to_period('M')
    monthly_avg = daily_stats.groupby('year_month').agg({
        'URSI': 'mean',
        'advancing_stocks': 'mean',
        'declining_stocks': 'mean',
        'unchanged_stocks': 'mean'
    }).round(2)
    monthly_avg.index = monthly_avg.index.to_timestamp()
    monthly_avg.reset_index(inplace=True)
    monthly_avg.columns = ['Month', 'Avg_URSI', 'Avg_Advancing', 'Avg_Declining', 'Avg_Unchanged']
    monthly_avg.to_excel(writer, sheet_name='Monthly_Averages', index=False)

print(f"\nExcel file exported to: {excel_file}")
print(f"  - Sheet 1: URSI_Daily - Complete daily data with moving averages")
print(f"  - Sheet 2: Summary - Statistical summary")
print(f"  - Sheet 3: Monthly_Averages - Monthly aggregated data")

# Convert dates to strings for JSON serialization
daily_stats['date_str'] = daily_stats['date'].dt.strftime('%Y-%m-%d')

# Create the interactive plot using Plotly
fig = go.Figure()

# Main URSI line
fig.add_trace(go.Scatter(
    x=daily_stats['date'].tolist(),
    y=daily_stats['URSI'].tolist(),
    mode='lines',
    name='URSI',
    line=dict(color='blue', width=2),
    hovertemplate='Date: %{x|%Y-%m-%d}<br>URSI: %{y:.2f}<br>Advancing: %{customdata[0]}<br>Declining: %{customdata[1]}<br>Unchanged: %{customdata[2]}<br>Total: %{customdata[3]}<extra></extra>',
    customdata=daily_stats[['advancing_stocks', 'declining_stocks', 'unchanged_stocks', 'total_stocks']].values
))

# Placeholder for MA line (will be updated dynamically)
fig.add_trace(go.Scatter(
    x=daily_stats['date'],
    y=[None] * len(daily_stats),
    mode='lines',
    name='MA',
    line=dict(color='red', width=2, dash='solid'),
    visible=False
))

# Add horizontal lines at 30 and 70
fig.add_hline(y=70, line_dash="dash", line_color="gray", line_width=1, 
              annotation_text="70 - Overbought", annotation_position="right")
fig.add_hline(y=30, line_dash="dash", line_color="gray", line_width=1,
              annotation_text="30 - Oversold", annotation_position="right")
fig.add_hline(y=50, line_dash="dot", line_color="lightgray", line_width=1,
              annotation_text="50 - Neutral", annotation_position="right")

# Add shaded regions
fig.add_hrect(y0=70, y1=100, fillcolor="green", opacity=0.2, 
              layer="below", line_width=0)
fig.add_hrect(y0=0, y1=30, fillcolor="lightpink", opacity=0.3,
              layer="below", line_width=0)

# Update layout
fig.update_layout(
    title=dict(
        text='URSI (Up/Down Relative Strength Index) - Vietnamese Stock Market',
        font=dict(size=20, color='black'),
        x=0.5,
        xanchor='center'
    ),
    xaxis=dict(
        title='Date',
        gridcolor='lightgray',
        showgrid=True,
        rangeslider=dict(visible=True),
        type='date'
    ),
    yaxis=dict(
        title='URSI (%)',
        gridcolor='lightgray',
        showgrid=True,
        range=[0, 100],
        dtick=10
    ),
    plot_bgcolor='white',
    paper_bgcolor='white',
    hovermode='x unified',
    height=600,
    margin=dict(l=50, r=50, t=120, b=50),
    legend=dict(
        orientation="h",
        yanchor="bottom",
        y=1.02,
        xanchor="right",
        x=1
    )
)

# Add annotation for latest value
fig.add_annotation(
    x=daily_stats['date'].iloc[-1],
    y=daily_stats['URSI'].iloc[-1],
    text=f"Latest: {daily_stats['URSI'].iloc[-1]:.1f}",
    showarrow=True,
    arrowhead=2,
    arrowsize=1,
    arrowwidth=2,
    arrowcolor="blue",
    bgcolor="white",
    bordercolor="blue",
    borderwidth=1
)

# Convert figure to JSON
fig_json = fig.to_json()

# Prepare URSI data for JavaScript
ursi_data = {
    'dates': daily_stats['date_str'].tolist(),
    'values': daily_stats['URSI'].tolist()
}

# Create HTML with interactive MA input
html_content = f"""
<!DOCTYPE html>
<html>
<head>
    <title>URSI Chart with Interactive Moving Average</title>
    <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
    <style>
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 20px;
            background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
            min-height: 100vh;
        }}
        .header {{
            text-align: center;
            margin-bottom: 20px;
        }}
        .header h1 {{
            color: #2c3e50;
            font-size: 28px;
            margin-bottom: 10px;
        }}
        .header p {{
            color: #7f8c8d;
            font-size: 14px;
        }}
        .control-panel {{
            background: linear-gradient(135deg, #ffffff 0%, #f8f9fa 100%);
            padding: 20px;
            border-radius: 12px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            margin-bottom: 20px;
            display: flex;
            align-items: center;
            gap: 20px;
            flex-wrap: wrap;
        }}
        .input-group {{
            display: flex;
            align-items: center;
            gap: 10px;
        }}
        label {{
            font-weight: 600;
            color: #2c3e50;
            font-size: 14px;
        }}
        input[type="number"] {{
            padding: 10px 14px;
            border: 2px solid #e0e0e0;
            border-radius: 8px;
            font-size: 14px;
            width: 100px;
            transition: all 0.3s ease;
        }}
        input[type="number"]:focus {{
            outline: none;
            border-color: #3498db;
            box-shadow: 0 0 0 3px rgba(52, 152, 219, 0.1);
        }}
        button {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 10px 24px;
            border: none;
            border-radius: 8px;
            cursor: pointer;
            font-size: 14px;
            font-weight: 600;
            transition: all 0.3s ease;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }}
        button:hover {{
            transform: translateY(-2px);
            box-shadow: 0 4px 8px rgba(0,0,0,0.2);
        }}
        button:active {{
            transform: translateY(0);
        }}
        .clear-button {{
            background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
        }}
        .status {{
            color: #27ae60;
            font-weight: 500;
            font-size: 14px;
            padding: 8px 16px;
            background-color: #e8f5e9;
            border-radius: 20px;
            display: inline-block;
        }}
        #plotDiv {{
            background-color: white;
            border-radius: 12px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            padding: 10px;
        }}
        .info-panel {{
            background: white;
            padding: 20px;
            border-radius: 12px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            margin-top: 20px;
        }}
        .info-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 15px;
            margin-top: 15px;
        }}
        .info-item {{
            padding: 10px;
            background: #f8f9fa;
            border-radius: 8px;
            border-left: 4px solid #3498db;
        }}
        .info-label {{
            font-size: 12px;
            color: #7f8c8d;
            margin-bottom: 4px;
        }}
        .info-value {{
            font-size: 18px;
            font-weight: 600;
            color: #2c3e50;
        }}
    </style>
</head>
<body>
    <div class="header">
        <h1>📈 URSI Interactive Dashboard</h1>
        <p>Up/Down Relative Strength Index - Vietnamese Stock Market</p>
        <p>Formula: URSI = (Advancing Stocks / (Advancing + Declining Stocks)) × 100</p>
    </div>
    
    <div class="control-panel">
        <div class="input-group">
            <label for="maDays">Moving Average Period:</label>
            <input type="number" id="maDays" min="2" max="200" value="20" placeholder="Days">
            <button onclick="updateMA()">📊 Calculate MA</button>
            <button class="clear-button" onclick="clearMA()">🗑️ Clear MA</button>
        </div>
        <span id="status" class="status">Ready</span>
    </div>
    
    <div id="plotDiv"></div>
    
    <div class="info-panel">
        <h3>Current Market Statistics</h3>
        <div class="info-grid">
            <div class="info-item">
                <div class="info-label">Latest URSI</div>
                <div class="info-value">{daily_stats['URSI'].iloc[-1]:.2f}%</div>
            </div>
            <div class="info-item">
                <div class="info-label">Advancing Stocks</div>
                <div class="info-value">{daily_stats['advancing_stocks'].iloc[-1]}</div>
            </div>
            <div class="info-item">
                <div class="info-label">Declining Stocks</div>
                <div class="info-value">{daily_stats['declining_stocks'].iloc[-1]}</div>
            </div>
            <div class="info-item">
                <div class="info-label">Unchanged Stocks</div>
                <div class="info-value">{daily_stats['unchanged_stocks'].iloc[-1]}</div>
            </div>
            <div class="info-item">
                <div class="info-label">Average URSI (All-time)</div>
                <div class="info-value">{daily_stats['URSI'].mean():.2f}%</div>
            </div>
            <div class="info-item">
                <div class="info-label">Date</div>
                <div class="info-value">{daily_stats['date'].iloc[-1].strftime('%Y-%m-%d')}</div>
            </div>
        </div>
    </div>
    
    <script>
        // Store the original figure and URSI data
        const figureData = {fig_json};
        const ursiData = {json.dumps(ursi_data)};
        
        // Initial plot
        Plotly.newPlot('plotDiv', figureData.data, figureData.layout);
        
        // Function to calculate moving average
        function calculateMA(values, period) {{
            const ma = [];
            for (let i = 0; i < values.length; i++) {{
                if (i < period - 1) {{
                    ma.push(null);
                }} else {{
                    let sum = 0;
                    for (let j = 0; j < period; j++) {{
                        sum += values[i - j];
                    }}
                    ma.push(sum / period);
                }}
            }}
            return ma;
        }}
        
        // Function to update the moving average
        function updateMA() {{
            const maDays = parseInt(document.getElementById('maDays').value);
            
            if (isNaN(maDays) || maDays < 2) {{
                document.getElementById('status').textContent = '⚠️ Please enter a valid number (minimum 2 days)';
                document.getElementById('status').style.backgroundColor = '#ffebee';
                document.getElementById('status').style.color = '#c62828';
                return;
            }}
            
            if (maDays > ursiData.values.length) {{
                document.getElementById('status').textContent = `⚠️ Maximum period is ${{ursiData.values.length}} days`;
                document.getElementById('status').style.backgroundColor = '#ffebee';
                document.getElementById('status').style.color = '#c62828';
                return;
            }}
            
            // Calculate moving average
            const maValues = calculateMA(ursiData.values, maDays);
            
            // Parse dates back to Date objects
            const dates = ursiData.dates.map(d => new Date(d));
            
            // Update the MA trace
            const update = {{
                x: [dates],
                y: [maValues],
                name: [`MA-${{maDays}}`],
                visible: [null, true],
                'hovertemplate': [`Date: %{{x|%Y-%m-%d}}<br>MA-${{maDays}}: %{{y:.2f}}<extra></extra>`]
            }};
            
            Plotly.update('plotDiv', update, {{}}, [1]);
            
            document.getElementById('status').textContent = `✅ Showing ${{maDays}}-day moving average`;
            document.getElementById('status').style.backgroundColor = '#e8f5e9';
            document.getElementById('status').style.color = '#27ae60';
        }}
        
        // Function to clear the moving average
        function clearMA() {{
            const update = {{
                visible: [null, false]
            }};
            Plotly.update('plotDiv', update, {{}}, [1]);
            document.getElementById('status').textContent = '🔄 Moving average cleared';
            document.getElementById('status').style.backgroundColor = '#fff3e0';
            document.getElementById('status').style.color = '#e65100';
        }}
        
        // Allow Enter key to calculate MA
        document.getElementById('maDays').addEventListener('keypress', function(event) {{
            if (event.key === 'Enter') {{
                updateMA();
            }}
        }});
        
        // Calculate initial MA on page load
        window.onload = function() {{
            updateMA();
        }};
    </script>
</body>
</html>
"""

# Save to new HTML file
output_file = r"C:\Users\minhdang\OneDrive - DRAGON CAPITAL\CodeVisual\Tai_Training\ursi_interactive_ma.html"
with open(output_file, 'w', encoding='utf-8') as f:
    f.write(html_content)

print(f"\nNew interactive HTML chart saved to: {output_file}")
print(f"Features:")
print(f"  - Interactive moving average with custom period input")
print(f"  - Enhanced UI with gradient design and statistics panel")
print(f"  - Real-time calculation updates")
print(f"  - Hover tooltips showing all stock counts")
print(f"  - Visual indicators for overbought/oversold zones")