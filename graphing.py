import pandas as pd
import os
import openpyxl
from openpyxl import load_workbook
from tqdm import tqdm
import datetime
import pygal
from pygal.style import Style

path = os.path.dirname(os.path.abspath(__file__))
done_path = os.path.join(path, "final")
svgs = os.path.join(path, "svgs")
# data = os.path.join(done_path, "lawn.xlsx")
cities_to_graph = ["Average", "Karachi", "Lahore", "Islamabad", "Peshawar", "Quetta"]
files = []
# r=root, d=directories, f = files
for r, d, f in os.walk(done_path):
    for file in f:
        if '.xlsx' in file:
            files.append(os.path.join(r, file))

for datafile in tqdm(files):
    filename = datafile.split("/")[-1].split(".xlsx")[0]
    df = pd.read_excel(datafile, sheet_name=0) # can also index sheet by name or fetch all sheets
    df = df.sort_values(by='Date')
    # print(df)
    # df['Date'] = pd.to_datetime(df.Date)
    header = [col for col in df.columns][1:]
    print(filename)
    dates = df["Date"].to_list()
    dates_str = [val.strftime("%b/%Y") for val in dates]
    line_chart = pygal.Line()
    # line_chart.interpolate = 'cubic'
    # line_chart.show_dots = False
    # line_chart.show_legend=False
    custom_style = Style(
        background="#f2f5f8",
        plot_background="rgba(255, 255, 255, 0.66)",
        font_family= "googlefont:Raleway",
        label_font_family= "googlefont:Raleway",
        major_label_font_family= "googlefont:Raleway",
        value_font_family= "googlefont:Raleway",
        value_label_font_family= "googlefont:Raleway",
        tooltip_font_family= "googlefont:Raleway",
        title_font_family= "googlefont:Raleway",
        legend_font_family= "googlefont:Raleway",
        no_data_font_family= "googlefont:Raleway",
        label_font_size = 3,
        major_label_font_size = 8)
    line_chart.style = custom_style
    line_chart.x_label_rotation=45
    line_chart.x_labels_major_every=12
    line_chart.y_labels_major_every=1
    line_chart.dots_size = 2
    line_chart.stroke_style = {'width': 2}
    line_chart.margin_bottom=50
    line_chart.legend_at_right=True
    line_chart.legend_at_top_columns=6
    line_chart.width=800
    line_chart.height=450
    line_chart.title = "Prices from {} to {}".format(dates_str [0], dates_str[-1])
    line_chart.x_labels = dates_str
    for col in header:
        if col in cities_to_graph:
            line_chart.add(col, df[col].to_list())

    line_chart.render_to_file(os.path.join(svgs, filename+".svg"))