from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_SHAPE_TYPE


class Lab:
    def __init__(self, name, baseline_avg, week_avg, month_avg, energy_saved, miles, phones, homes, alerts, building):
        self.name = name
        self.baseline_avg = baseline_avg
        self.week_avg = week_avg
        self.month_avg = month_avg
        self.energy_saved = energy_saved
        self.miles = miles
        self.phones = phones
        self.homes = homes
        self.alerts = alerts
        self.building = building


class Building:
    def __init__(self, name, shorthand, average):
        self.name = name
        self.shorthand = shorthand
        self.average = average
        self.labs = []

    def add_lab(self, lab):
        self.labs.append(lab)


def demonstration():
    ichan = Building("Carl Ichan", "Ichan", 15.4277)
    lab003 = Lab("Lab 114", 5.46, 5.525213, 5.39672, -0.064247, "12,256", "586,292", "0.607", "3", ichan)
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
dates.text_frame.text = "May 22–May 28, 2023"
dates.text_frame.fit_text(font_family='dates', font_file='Poppins-Regular.ttf', max_size=11, bold=False)

# Fixes font issue
test.shapes[4].text_frame.text = "Your increase in energy usage this week equates to:*"
test.shapes[4].text_frame.fit_text(font_family='dates', font_file='Poppins-Regular.ttf', max_size=15, bold=True)
test.shapes[4].text_frame.paragraphs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

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

chart_data = CategoryChartData()
# This will be different for lab 210
chart_data.categories = [f"{ichan_demo.shorthand} Average", "Baseline Average", "Your Lab"]
chart_data.add_series('Series 1', ( ichan_demo.labs[0].week_avg, ichan_demo.labs[0].baseline_avg, ichan_demo.average))
x, y, cx, cy = Inches(0.7), Inches(5.15), Inches(5), Inches(1.5)
test.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, x, y, cx, cy, chart_data)

prs.save("output.pptx")
