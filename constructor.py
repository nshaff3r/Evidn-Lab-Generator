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