from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR
from pptx.util import Inches, Pt


class Lab:
    def __init__(self, name, week_avg, month_avg, energy_saved, miles, phones, homes, alerts, building):
        self.name = name
        self.week_avg = week_avg
        self.month_avg = month_avg
        self.energy_saved = energy_saved
        self.miles = miles
        self.phones = phones
        self.homes = homes
        self.alerts = alerts
        self.building = building


class Building:
    def __init__(self, name, shorthand):
        self.name = name
        self.shorthand = shorthand
        self.labs = []

    def add_lab(self, lab):
        self.labs.append(lab)


def demonstration():
    ichan = Building("Carl Ichan", "Ichan")
    lab003 = Lab("Lab 114", 5.525213, 5.39672, -0.064247, "12,256", "586,292", "0.607", "5", ichan)
    ichan.add_lab(lab003)
    return ichan


prs = Presentation('input.pptx')
test = prs.slides[0]
ichan_demo = demonstration()

# Creates lab name
test.shapes[2].text_frame.text = ichan_demo.labs[0].name
test.shapes[2].text_frame.fit_text(font_family='title', font_file='Poppins-Regular.ttf', max_size=28, bold=True)
test.shapes[2].text_frame.paragraphs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
# Writes weekly energy report
subtitle = test.shapes.add_textbox(Inches(0.44), Inches(1), Inches(4), Inches(4))
subtitle.text_frame.text = "Weekly Energy Report"
subtitle.text_frame.fit_text(font_family='subtitle', font_file='Poppins-Regular.ttf', max_size=14, bold=True)

# Writes week dates
dates = test.shapes.add_textbox(Inches(0.44), Inches(1.31), Inches(4), Inches(4))
dates.text_frame.text = "May 22–May 28, 2023"
dates.text_frame.fit_text(font_family='dates', font_file='Poppins-Regular.ttf', max_size=11, bold=False)

# Fixes font issue
test.shapes[7].text_frame.text = "Your increase in energy usage this week equates to:*"
test.shapes[7].text_frame.fit_text(font_family='dates', font_file='Poppins-Regular.ttf', max_size=15, bold=True)
test.shapes[7].text_frame.paragraphs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

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

prs.save("output.pptx")
