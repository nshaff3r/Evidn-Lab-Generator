import collections
import collections.abc
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt
from constructor import Lab, Building
from setup import add_labs
from sys import exit
from demo import demonstration
import matplotlib.pyplot as plt
import io


prs = Presentation('template.pptx')
if len(prs.slides) != int(input("How many labs do you have? ")):
    exit("Your template presentation contains a different number of slides than labs.\n"
         "Please fix the template presentation to have the same number of slides as labs.")
building = add_labs()
# building = demonstration()
if building is None:
    exit()

if len(prs.slides) != len(building.labs):
    exit("Your template presentation contains a different number of slides than labs.\n"
         "Please fix the template presentation to have the same number of slides as labs.")

timeperiod = input("Date (ex. May 22–May 28, 2023): ")
labofweek = input("Lab of the week (ex. Lab 007): ")
for i, slide in enumerate(prs.slides):
    # Creates lab name
    slide.shapes[1].text_frame.text = building.labs[i].name
    slide.shapes[1].text_frame.fit_text(font_family='title', font_file='Poppins-Regular.ttf', max_size=33, bold=True)
    slide.shapes[1].text_frame.paragraphs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    # Writes weekly energy report
    subtitle = slide.shapes.add_textbox(Inches(0.44), Inches(1), Inches(2), Inches(1))
    subtitle.text_frame.text = "Weekly Energy Report"
    subtitle.text_frame.fit_text(font_family='subtitle', font_file='Poppins-Regular.ttf', max_size=14, bold=True)

    # Writes week dates
    dates = slide.shapes.add_textbox(Inches(0.44), Inches(1.31), Inches(2), Inches(0.5))
    dates.text_frame.text = timeperiod
    dates.text_frame.fit_text(font_family='dates', font_file='Poppins-Regular.ttf', max_size=11, bold=False)

    # Fixes font issue
    slide.shapes[2].text_frame.text = "Your increase in energy usage this week equates to:*"
    slide.shapes[2].text_frame.fit_text(font_family='dates', font_file='Poppins-Regular.ttf', max_size=15, bold=True)
    slide.shapes[2].text_frame.paragraphs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    alerts = slide.shapes.add_textbox(Inches(6.01), Inches(1), Inches(2), Inches(1))
    alerts.text_frame.text = str(building.labs[i].alerts)
    alerts.text_frame.fit_text(font_family='dates', font_file='Poppins-Regular.ttf', max_size=28, bold=True)

    miles = slide.shapes.add_textbox(Inches(0.84), Inches(3.17), Inches(2), Inches(1))
    miles.text_frame.text = str(round(building.labs[i].miles))
    miles.text_frame.fit_text(font_family='dates', font_file='Poppins-Regular.ttf', max_size=20, bold=True)
    miles.text_frame.paragraphs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    phones = slide.shapes.add_textbox(Inches(2.88), Inches(3.17), Inches(2), Inches(1))
    phones.text_frame.text = str(round(building.labs[i].phones))
    phones.text_frame.fit_text(font_family='dates', font_file='Poppins-Regular.ttf', max_size=20, bold=True)
    phones.text_frame.paragraphs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    homes = slide.shapes.add_textbox(Inches(5.25), Inches(3.17), Inches(2), Inches(1))
    homes.text_frame.text = str(round(building.labs[i].homes))
    homes.text_frame.fit_text(font_family='dates', font_file='Poppins-Regular.ttf', max_size=20, bold=True)
    homes.text_frame.paragraphs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    best_lab = slide.shapes.add_textbox(Inches(5.7), Inches(7.4), Inches(2), Inches(1))
    best_lab.text_frame.text = labofweek
    best_lab.text_frame.fit_text(font_family='lab', font_file='Poppins-Regular.ttf', max_size=20, bold=True)
    best_lab.text_frame.paragraphs[0].font.color.rgb = RGBColor(0xF0, 0x7C, 0x34)

    data = {f"{building.shorthand}\nAverage": building.average, "Baseline\nAverage": building.labs[i].baseline_avg, "Your\nLab": building.labs[i].week_avg}
    labs = list(data.keys())
    values = list(data.values())
    plt.figure(figsize=(15, 5))
    bar_colors = ["orange", "grey", "pink"]
    plt.barh(labs, values, color=bar_colors)
    plt.margins(.3)
    for spine in plt.gca().spines.values():
        spine.set_visible(False)
    plt.xticks([])
    for j in range(3):
        plt.text(values[j], j, f"{values[j]: .2f} MTCO2", size=30)
    plt.yticks(size=20, fontweight='bold')
    plt.title("ENERGY USAGE THIS WEEK", size=35, fontweight='bold')
    image_stream = io.BytesIO()
    plt.savefig(image_stream)
    slide.shapes.add_picture(image_stream, Inches(0.3), Inches(4.6), Inches(6), Inches(2))
    data = building.labs[i].months_avg
    names = list(data.keys())
    values = list(data.values())
    plt.figure(figsize=(10, 5))
    plt.plot(names, values, color="orange")
    baseline = {}
    for month in names:
        baseline[month] = building.labs[i].baseline_avg
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
    slide.shapes.add_picture(image_stream, Inches(0.3), Inches(6.8), Inches(5.13), Inches(2.58))

print("Finished! Generated output.pptx.")
prs.save("output.pptx")
