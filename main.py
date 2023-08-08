import collections
import collections.abc
from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt
from constructor import Lab, Building
from setup import add_labs
from sys import exit
import matplotlib.pyplot as plt
import io
import calendar


prs = Presentation('template.pptx')
if len(prs.slides) != int(input("How many labs do you have? ")):
    exit("Your template presentation contains a different number of slides than labs.\n"
         "Please fix the template presentation to have the same number of slides as labs.")
building = add_labs()
# building = demonstration()
if building is None:
    exit()

# building = Building("Icahn", "Icahn", 15.23, 0)
# months = {
#         5: 21,
#         6: 22,
#         7: 21.5,
#         8: 0,
#         9: 0,
#         10: 0,
#         11: 0,
#         12: 0,
#     }
# lab003 = Lab("L003", 29, 28, months, 100, 3433, 162882, 0.169, 0, building)
# building.add_lab(lab003)
if len(prs.slides) != len(building.labs):
    exit("Your template presentation contains a different number of slides than labs.\n"
         "Please fix the template presentation to have the same number of slides as labs.")

timeperiod = input("Enter how you want the date to appear (ex. May 20–May 26, 2023): ")
labofweek = input("Lab of the week (ex. Lab 007): ")
for i, slide in enumerate(prs.slides):
    # Creates lab name
    building.labs[i].name = f"{building.labs[i].name[0]}ab {building.labs[i].name[1:]}"
    slide.shapes[1].text_frame.text = building.labs[i].name
    slide.shapes[1].text_frame.fit_text(font_family='title', font_file='Poppins-Regular.ttf', max_size=33, bold=True)
    slide.shapes[1].text_frame.paragraphs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    # Writes weekly energy report
    subtitle = slide.shapes.add_textbox(Inches(0.44), Inches(1), Inches(3), Inches(1))
    subtitle.text_frame.text = "Weekly Energy Report"
    subtitle.text_frame.fit_text(font_family='subtitle', font_file='Poppins-Regular.ttf', max_size=14, bold=True)

    # Writes week dates
    dates = slide.shapes.add_textbox(Inches(0.44), Inches(1.31), Inches(2), Inches(0.5))
    dates.text_frame.text = timeperiod
    dates.text_frame.fit_text(font_family='dates', font_file='Poppins-Regular.ttf', max_size=11, bold=False)

    # Fixes font issue
    change = ""
    if (building.labs[i].energy_saved > 0):
        change = "Your energy reduction this week equates to:*"
        slide.shapes.add_picture("decrease.png", Inches(6.1), Inches(2.5), Inches(1.072), Inches(0.8))
    else:
        change = "Your increase in energy usage this week equates to:*"
        slide.shapes.add_picture("increase.png", Inches(6.1), Inches(2.5), Inches(1.072), Inches(0.8))

    slide.shapes[2].text_frame.text = change
    slide.shapes[2].text_frame.fit_text(font_family='dates', font_file='Poppins-Regular.ttf', max_size=15, bold=True)
    slide.shapes[2].text_frame.paragraphs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    alerts = slide.shapes.add_textbox(Inches(6.01), Inches(1), Inches(0.5), Inches(0.6))
    alerts.text_frame.text = str(building.labs[i].alerts)
    alerts.text_frame.fit_text(font_family='dates', font_file='Poppins-Regular.ttf', max_size=28, bold=True)

    miles = slide.shapes.add_textbox(Inches(0.44), Inches(3.2), Inches(1.74), Inches(0.5))
    miles.text_frame.text = "{:,.1f}".format(building.labs[i].miles)
    miles.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    miles.text_frame.fit_text(font_family='dates', font_file='Poppins-Regular.ttf', max_size=20, bold=True)
    miles.text_frame.paragraphs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    phones = slide.shapes.add_textbox(Inches(2.39), Inches(3.2), Inches(2), Inches(0.5))
    phones.text_frame.text = "{:,.0f}".format(building.labs[i].phones)
    phones.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    phones.text_frame.fit_text(font_family='dates', font_file='Poppins-Regular.ttf', max_size=20, bold=True)
    phones.text_frame.paragraphs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    homes = slide.shapes.add_textbox(Inches(4.79), Inches(3.2), Inches(1.29), Inches(0.5))
    homes.text_frame.text = "{:,.3f}".format(building.labs[i].homes)
    homes.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    homes.text_frame.fit_text(font_family='dates', font_file='Poppins-Regular.ttf', max_size=20, bold=True)
    homes.text_frame.paragraphs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    best_lab = slide.shapes.add_textbox(Inches(5.53), Inches(7.42), Inches(1.5), Inches(0.5))
    best_lab.text_frame.text = labofweek
    best_lab.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
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

    # Months graph
    plt.figure(figsize=(10, 5))
    data = building.labs[i].months_avg
    names = list(data.keys())
    baseline = {}
    for month in names:
        baseline[month] = building.labs[i].baseline_avg
    baselineValues = list(baseline.values())
    # Data started in May
    names = list(map(lambda x:  x + 4 if x <= 8 else x - 8, names))
    names = list(map(lambda x: calendar.month_abbr[x], names))
    plt.plot(names, baselineValues, color="grey")

    # Changing data
    values = list(data.values())
    values = list(map(lambda x: None if x == 0 else x, values))
    values = values[4:] + [None, None, None, None]
    plt.plot(names, values, color="orange")
    for spine in plt.gca().spines.values():
        spine.set_visible(False)
    plt.legend(['Current', "Baseline"], fontsize='x-large')
    plt.ylabel("MTons CO2", size=20)
    plt.xticks(size=20)
    plt.yticks(size=15)
    # Hide every other month
    ax = plt.gca()
    ax.set_ylim(0, 2 * building.labs[i].baseline_avg)
    temp = ax.xaxis.get_ticklabels()
    temp = list(set(temp) - set(temp[::2]))
    for label in temp:
        label.set_visible(False)
    plt.title("CURRENT VS. BASELINE ENERGY USAGE", size=27, fontweight='bold', pad=20)
    image_stream = io.BytesIO()
    plt.savefig(image_stream)
    slide.shapes.add_picture(image_stream, Inches(0.3), Inches(6.8), Inches(5.13), Inches(2.58))

print("Finished! Generated output.pptx.")
prs.save("output.pptx")