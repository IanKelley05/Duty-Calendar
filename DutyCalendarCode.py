import pandas as pd

def checkDutyStatus(row, weekend):
    status = row % 5
    if status == 4:
        return "Primary"
    elif weekend:
        return "Secondary"
    elif status == 0:
        return "Secondary"
    elif status == 1:
        return "Desk"
    elif status == 2:
        return "Off"

def checkWhoWorkingWith(row, weekend, dayColum):
    if row % 5 == 2: # If we are primary
        primary = dayColum[row]
        secondary1 = dayColum[row + 1]
        secondary2 = dayColum[row + 2]
        secondary3 = dayColum[row + 3]
    elif row % 5 == 3: # If we are 1st under primary
        primary = dayColum[row - 1]
        secondary1 = dayColum[row]
        secondary2 = dayColum[row + 1]
        secondary3 = dayColum[row + 2]
    elif row % 5 == 4: # If we are 2nd under primary
        primary = dayColum[row - 2]
        secondary1 = dayColum[row - 1]
        secondary2 = dayColum[row]
        secondary3 = dayColum[row + 1]
    else: # If we are 3rd under primary
        primary = dayColum[row - 3]
        secondary1 = dayColum[row - 2]
        secondary2 = dayColum[row - 1]
        secondary3 = dayColum[row]
    if weekend == 1:
        return {
                "Primary": primary,
                "1st Secondary": secondary1,
                "2nd Secondary": secondary2,
                "3rd Secondary": secondary3,
        }
    else:
        return {
            "Primary": primary,
            "Secondary": secondary1,
            "Desk": secondary2,
            "Off": secondary3
        }

def checkWeekendDuty(row, day, satDayDuty, sunDayDuty):
    saturday1 = "N/A"
    saturday2 = "N/A"
    sunday1 = "N/A"
    sunday2 = "N/A"
    if row % 5 == 2: # If we are primary
        saturday1 = satDayDuty[row]
        saturday2 = satDayDuty[row + 1]
        sunday1 = sunDayDuty[row]
        sunday2 = sunDayDuty[row + 1]
    elif row % 5 == 3:
        saturday1 = satDayDuty[row - 1]
        saturday2 = satDayDuty[row]
        sunday1 = sunDayDuty[row - 1]
        sunday2 = sunDayDuty[row]
    elif row % 5 == 4:
        saturday1 = satDayDuty[row - 2]
        saturday2 = satDayDuty[row - 1]
        sunday1 = sunDayDuty[row - 2]
        sunday2 = sunDayDuty[row - 1]
    else:
        saturday1 = satDayDuty[row - 3]
        saturday2 = satDayDuty[row - 2]
        sunday1 = sunDayDuty[row - 3]
        sunday2 = sunDayDuty[row - 2]

    return{
        "Saturday 1":saturday1,
        "Saturday 2":saturday2,
        "Sunday 1":sunday1,
        "Sunday 2":sunday2
    }

def main():
    sheetNames = ["JANUARY", "FEBRUARY", "MARCH", "APRIL", "MAY"]
    days = ["SUNDAY", "MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "SATURDAY"]
    name = "Ian"

    for sheet in sheetNames:
        df = pd.read_excel("Final MCUT Duty 24-25.xlsx", sheet_name=sheet)
        for day in range(7):
            weekend = 0
            if day >= 4:
                weekend = 1
                satDayDuty = df.iloc[:, 7]
                sunDayDuty = df.iloc[:, 8]
            dayColum = df.iloc[:, day]
            
            for idx, value in enumerate(dayColum):
                dutyStatus = checkDutyStatus(idx + 2, weekend)
                if value == name:
                    coworkers = checkWhoWorkingWith(idx,weekend, dayColum)
                    print(f"Found '{name}' in sheet '{sheet}' on day {days[day]} at row {idx + 2} and is {dutyStatus}")
                    print(f"Coworkers: {coworkers}")
                    if weekend:
                        weekendDuty = checkWeekendDuty(idx, day, satDayDuty, sunDayDuty)
                        print(f"Weekend Duty Workers: {weekendDuty}")

if __name__ == "__main__":
    main()
