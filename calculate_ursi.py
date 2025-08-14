import pandas as pd
import plotly.graph_objects as go
from datetime import datetime

# Load the data
df = pd.read_pickle(r"C:\Users\minhdang\OneDrive - DRAGON CAPITAL\CodeVisual\Tai_Training\df_ohlcv_195stocks.pkl")

# Convert day column to datetime for proper sorting
df['date'] = pd.to_datetime(df['day'].str.replace('_', '-'))

# Sort by stock and date to ensure chronological order
df = df.sort_values(['stock', 'date'])

# Calculate previous close for each stock
df['prev_close'] = df.groupby('stock')['close'].shift(1)

# Determine if stock is advancing (1) or declining (0)
# Advancing: today's close > previous close
# Declining: today's close <= previous close
df['is_advancing'] = (df['close'] > df['prev_close']).astype(int)

# Remove rows where prev_close is NaN (first day for each stock)
df_clean = df.dropna(subset=['prev_close'])

# Calculate daily URSI
daily_stats = df_clean.groupby('date').agg({
    'is_advancing': ['sum', 'count']
}).reset_index()

# Flatten column names
daily_stats.columns = ['date', 'advancing_stocks', 'total_stocks']

# Calculate declining stocks
daily_stats['declining_stocks'] = daily_stats['total_stocks'] - daily_stats['advancing_stocks']

# Calculate URSI = (Advancing / (Advancing + Declining)) * 100
# Note: Advancing + Declining = Total stocks with data for that day
daily_stats['URSI'] = (daily_stats['advancing_stocks'] / daily_stats['total_stocks']) * 100

# Sort by date
daily_stats = daily_stats.sort_values('date')

print(f"URSI Calculation Complete!")
print(f"Date range: {daily_stats['date'].min()} to {daily_stats['date'].max()}")
print(f"Average URSI: {daily_stats['URSI'].mean():.2f}")
print(f"Current URSI (latest): {daily_stats['URSI'].iloc[-1]:.2f}")

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

# Add horizontal lines at 30 and 70
fig.add_hline(y=70, line_dash="dash", line_color="gray", line_width=1, 
              annotation_text="70 - Overbought", annotation_position="right")
fig.add_hline(y=30, line_dash="dash", line_color="gray", line_width=1,
              annotation_text="30 - Oversold", annotation_position="right")

# Add shaded regions
# Green shading above 70 (bullish sentiment)
fig.add_hrect(y0=70, y1=100, fillcolor="green", opacity=0.2, 
              layer="below", line_width=0)

# Pink shading below 30 (bearish sentiment)
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
        rangeslider=dict(visible=True),  # Add range slider for zooming
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
    margin=dict(l=50, r=50, t=80, b=50),
    legend=dict(
        orientation="h",
        yanchor="bottom",
        y=1.02,
        xanchor="right",
        x=1
    )
)

# Add annotations for market conditions
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

# Save to HTML file
output_file = r"C:\Users\minhdang\OneDrive - DRAGON CAPITAL\CodeVisual\Tai_Training\ursi_chart.html"
fig.write_html(output_file)

print(f"\nHTML chart saved to: {output_file}")
print(f"\nSummary Statistics:")
print(f"- Minimum URSI: {daily_stats['URSI'].min():.2f}")
print(f"- Maximum URSI: {daily_stats['URSI'].max():.2f}")
print(f"- Days above 70 (Bullish): {(daily_stats['URSI'] > 70).sum()} days")
print(f"- Days below 30 (Bearish): {(daily_stats['URSI'] < 30).sum()} days")
print(f"- Days in neutral zone (30-70): {((daily_stats['URSI'] >= 30) & (daily_stats['URSI'] <= 70)).sum()} days")

# Export the calculated URSI data to CSV for reference
csv_file = r"C:\Users\minhdang\OneDrive - DRAGON CAPITAL\CodeVisual\Tai_Training\ursi_data.csv"
daily_stats.to_csv(csv_file, index=False)
print(f"\nURSI data exported to: {csv_file}")