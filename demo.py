from constructor import Lab, Building

def demonstration():
    ichan = Building("Carl Ichan", "Ichan", 15.4277)
    months = {
        "May": 5.39672,
        "Jun": 4.39672,
    }
    months2 = {
        "May": 5.39672,
        "Jun": 6.392,
    }
    months3 = {
        "May": 10.2672,
        "Jun": 1.23672,
    }
    lab003 = Lab("Lab 003", 3.46, 5.525213, months, -0.064247, "15,256", "586,292", "0.607", "3", ichan)
    lab004 = Lab("Lab 004", 6.46, 4.525, months2, -0.064247, "12,256", "586,292", "0.507", "3", ichan)
    lab114 = Lab("Lab 114", 2.44, 8.21, months3, -0.322247, "10,256", "466,292", "1.607", "2", ichan)
    ichan.add_lab(lab003)
    ichan.add_lab(lab004)
    ichan.add_lab(lab114)
    return ichan
