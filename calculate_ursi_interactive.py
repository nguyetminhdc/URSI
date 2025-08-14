import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import json

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

print(f"URSI Calculation Complete!")
print(f"Date range: {daily_stats['date'].min()} to {daily_stats['date'].max()}")
print(f"Average URSI: {daily_stats['URSI'].mean():.2f}")
print(f"Current URSI (latest): {daily_stats['URSI'].iloc[-1]:.2f}")
print(f"Latest day - Advancing: {daily_stats['advancing_stocks'].iloc[-1]}, Declining: {daily_stats['declining_stocks'].iloc[-1]}, Unchanged: {daily_stats['unchanged_stocks'].iloc[-1]}")

# Convert dates to strings for JSON serialization
daily_stats['date_str'] = daily_stats['date'].dt.strftime('%Y-%m-%d')

# Create the interactive plot using Plotly
fig = go.Figure()

# Main URSI line
fig.add_trace(go.Scatter(
    x=daily_stats['date'],
    y=daily_stats['URSI'],
    mode='lines',
    name='URSI',
    line=dict(color='blue', width=2),
    hovertemplate='Date: %{x|%Y-%m-%d}<br>URSI: %{y:.2f}<br>Advancing: %{customdata[0]}<br>Declining: %{customdata[1]}<br>Total: %{customdata[2]}<extra></extra>',
    customdata=daily_stats[['advancing_stocks', 'declining_stocks', 'total_stocks']].values
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
            font-family: Arial, sans-serif;
            margin: 20px;
            background-color: #f5f5f5;
        }}
        .control-panel {{
            background-color: white;
            padding: 15px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            margin-bottom: 20px;
            display: flex;
            align-items: center;
            gap: 15px;
        }}
        .input-group {{
            display: flex;
            align-items: center;
            gap: 10px;
        }}
        label {{
            font-weight: bold;
            color: #333;
        }}
        input[type="number"] {{
            padding: 8px 12px;
            border: 2px solid #ddd;
            border-radius: 4px;
            font-size: 14px;
            width: 80px;
        }}
        input[type="number"]:focus {{
            outline: none;
            border-color: #4CAF50;
        }}
        button {{
            background-color: #4CAF50;
            color: white;
            padding: 8px 20px;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-size: 14px;
            font-weight: bold;
            transition: background-color 0.3s;
        }}
        button:hover {{
            background-color: #45a049;
        }}
        button:active {{
            background-color: #3d8b40;
        }}
        .clear-button {{
            background-color: #f44336;
        }}
        .clear-button:hover {{
            background-color: #da190b;
        }}
        .status {{
            color: #666;
            font-style: italic;
            margin-left: 10px;
        }}
        #plotDiv {{
            background-color: white;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            padding: 10px;
        }}
    </style>
</head>
<body>
    <div class="control-panel">
        <div class="input-group">
            <label for="maDays">Moving Average Period (days):</label>
            <input type="number" id="maDays" min="2" max="200" value="20" placeholder="Enter days">
            <button onclick="updateMA()">Calculate MA</button>
            <button class="clear-button" onclick="clearMA()">Clear MA</button>
        </div>
        <span id="status" class="status"></span>
    </div>
    
    <div id="plotDiv"></div>
    
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
                document.getElementById('status').textContent = 'Please enter a valid number (minimum 2 days)';
                return;
            }}
            
            if (maDays > ursiData.values.length) {{
                document.getElementById('status').textContent = `Maximum period is ${{ursiData.values.length}} days`;
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
            
            document.getElementById('status').textContent = `Showing ${{maDays}}-day moving average`;
        }}
        
        // Function to clear the moving average
        function clearMA() {{
            const update = {{
                visible: [null, false]
            }};
            Plotly.update('plotDiv', update, {{}}, [1]);
            document.getElementById('status').textContent = 'Moving average cleared';
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

# Save to HTML file
output_file = r"C:\Users\minhdang\OneDrive - DRAGON CAPITAL\CodeVisual\Tai_Training\ursi_chart.html"
with open(output_file, 'w', encoding='utf-8') as f:
    f.write(html_content)

print(f"\nInteractive HTML chart saved to: {output_file}")
print(f"\nThe chart now includes:")
print(f"- An input field to enter any number of days for the moving average")
print(f"- A 'Calculate MA' button to update the chart with the new MA")
print(f"- A 'Clear MA' button to hide the moving average line")
print(f"- Dynamic calculation of moving average based on user input")

# Export the calculated URSI data to CSV for reference
csv_file = r"C:\Users\minhdang\OneDrive - DRAGON CAPITAL\CodeVisual\Tai_Training\ursi_data.csv"
daily_stats[['date', 'advancing_stocks', 'declining_stocks', 'total_stocks', 'URSI']].to_csv(csv_file, index=False)
print(f"\nURSI data exported to: {csv_file}")