from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.chart import XL_LABEL_POSITION
from pptx.enum.chart import XL_TICK_MARK
from constructor import Lab, Building
import matplotlib.pyplot as plt
import io

def demonstration():
    ichan = Building("Carl Ichan", "Ichan", 15.4277)
    months = {
        "May": 5.39672,
        "Jun": 4.39672,
    }
    lab003 = Lab("Lab 114", 3.46, 5.525213, months, -0.064247, "12,256", "586,292", "0.607", "3", ichan)
    ichan.add_lab(lab003)
    return ichan


prs = Presentation('input.pptx')
test = prs.slides[0]
ichan_demo = demonstration()
# for shape in test.shapes:
#     print(test.shapes.index(shape), shape.shape_type, end=' ')
#     if (shape.shape_type == MSO_SHAPE_TYPE.TEXT_BOX):
#         print(shape.text)
#     else:
#         print()


# Creates lab name
test.shapes[1].text_frame.text = ichan_demo.labs[0].name
test.shapes[1].text_frame.fit_text(font_family='title', font_file='Poppins-Regular.ttf', max_size=33, bold=True)
test.shapes[1].text_frame.paragraphs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
# Writes weekly energy report
subtitle = test.shapes.add_textbox(Inches(0.44), Inches(1), Inches(4), Inches(4))
subtitle.text_frame.text = "Weekly Energy Report"
subtitle.text_frame.fit_text(font_family='subtitle', font_file='Poppins-Regular.ttf', max_size=14, bold=True)

# Writes week dates
dates = test.shapes.add_textbox(Inches(0.44), Inches(1.31), Inches(4), Inches(4))
dates.text_frame.text = input("Date (in format: May 22–May 28, 2023): ")
dates.text_frame.fit_text(font_family='dates', font_file='Poppins-Regular.ttf', max_size=11, bold=False)

# Fixes font issue
test.shapes[2].text_frame.text = "Your increase in energy usage this week equates to:*"
test.shapes[2].text_frame.fit_text(font_family='dates', font_file='Poppins-Regular.ttf', max_size=15, bold=True)
test.shapes[2].text_frame.paragraphs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

alerts = test.shapes.add_textbox(Inches(6.01), Inches(1), Inches(4), Inches(4))
alerts.text_frame.text = ichan_demo.labs[0].alerts
alerts.text_frame.fit_text(font_family='dates', font_file='Poppins-Regular.ttf', max_size=28, bold=True)

miles = test.shapes.add_textbox(Inches(0.84), Inches(3.17), Inches(7), Inches(7))
miles.text_frame.text = ichan_demo.labs[0].miles
miles.text_frame.fit_text(font_family='dates', font_file='Poppins-Regular.ttf', max_size=20, bold=True)
miles.text_frame.paragraphs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

phones = test.shapes.add_textbox(Inches(2.88), Inches(3.17), Inches(7), Inches(7))
phones.text_frame.text = ichan_demo.labs[0].phones
phones.text_frame.fit_text(font_family='dates', font_file='Poppins-Regular.ttf', max_size=20, bold=True)
phones.text_frame.paragraphs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

homes = test.shapes.add_textbox(Inches(5), Inches(3.17), Inches(7), Inches(7))
homes.text_frame.text = ichan_demo.labs[0].homes
homes.text_frame.fit_text(font_family='dates', font_file='Poppins-Regular.ttf', max_size=20, bold=True)
homes.text_frame.paragraphs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

best_lab = test.shapes.add_textbox(Inches(5.63), Inches(7.4), Inches(3), Inches(3))
best_lab.text_frame.text = input("Lab of the week: ")
best_lab.text_frame.fit_text(font_family='lab', font_file='Poppins-Regular.ttf', max_size=20, bold=True)
best_lab.text_frame.paragraphs[0].font.color.rgb = RGBColor(0xF0, 0x7C, 0x34)

data = {f"{ichan_demo.shorthand}\nAverage": ichan_demo.average, "Baseline\nAverage": ichan_demo.labs[0].baseline_avg, "Your\nLab": ichan_demo.labs[0].week_avg}
labs = list(data.keys())
values = list(data.values())
plt.figure(figsize=(15, 5))
bar_colors = ["orange", "grey", "pink"]
plt.barh(labs, values, color=bar_colors)
plt.margins(.3)
for spine in plt.gca().spines.values():
    spine.set_visible(False)
plt.xticks([])
for i in range(3):
    plt.text(values[i], i, f"{values[i]: .2f} MTCO2", size=30)
plt.yticks(size=20, fontweight='bold')
plt.title("ENERGY USAGE THIS WEEK", size=35, fontweight='bold')
image_stream = io.BytesIO()
plt.savefig(image_stream)
test.shapes.add_picture(image_stream, Inches(0.3), Inches(4.6), Inches(6), Inches(2))

data = ichan_demo.labs[0].months_avg
names = list(data.keys())
values = list(data.values())
plt.figure(figsize=(10, 5))
plt.plot(names, values, color="orange")
baseline = {}
for month in names:
    baseline[month] = ichan_demo.labs[0].baseline_avg
names = list(baseline.keys())
values = list(baseline.values())
plt.plot(names, values, color="grey")
for spine in plt.gca().spines.values():
    spine.set_visible(False)
plt.ylim([0, 10])
plt.legend(['Current', "Baseline"], fontsize='x-large')
plt.ylabel("MTons CO2", size=20)
plt.xticks(size=20)
plt.yticks(size=15)
plt.title("CURRENT VS. BASELINE ENERGY USAGE", size=27, fontweight='bold', pad=20)
image_stream = io.BytesIO()
plt.savefig(image_stream)
test.shapes.add_picture(image_stream, Inches(0.3), Inches(6.8), Inches(5.13), Inches(2.58))

prs.save("output.pptx")
