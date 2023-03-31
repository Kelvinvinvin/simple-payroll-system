# Lai Zhi Ming
# TP072714

from contextlib import suppress
import xlsxwriter
from datetime import date

d = {}
months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November",
          "December"]
today_date = date.today()


def addEmployee():
    size = int(input("Enter the number of employee/s: "))

    for i in range(size):
        dict_name = input("Enter the employee id: ")
        if dict_name in d.keys():
            print("--------------*---------------")
            print("Please enter new employee ID: ")
        else:
            d[dict_name] = {}
            Name = input("Enter name: ")
            staff_id = dict_name
            department = input("Enter department: ")
            d[dict_name]["Name"] = Name
            d[dict_name]["ID"] = staff_id
            d[dict_name]["Department"] = department
            try:
                basic_salary = float(input("Enter basic salary: "))
                allowance = float(input("Enter allowance: "))
                bonus = float(input("Enter bonus: "))
                overtime = float(input("Enter overtime: "))
                d[dict_name]["Basic_salary"] = basic_salary
                d[dict_name]["Allowance"] = allowance
                d[dict_name]["Bonus"] = bonus
                d[dict_name]["Overtime"] = overtime
            except ValueError:
                print("Please enter integer value....")
                del d[dict_name]

    print(d)


def generateSalary(name):
    if name in d.keys():
        print(months)
        final_total_salary = 0
        print("* the first character of the month input must be in upper case")
        month_salary = input("Which month is the salary belong: ")
        staff_basic_salary = d[name]["Basic_salary"]
        staff_allowance = d[name]["Allowance"]
        staff_bonus = d[name]["Bonus"]
        staff_overtime = d[name]["Overtime"]
        total_without_epf = float(staff_basic_salary) + float(staff_allowance) + float(staff_bonus) + float(
            staff_overtime)
        epf = total_without_epf * 0.11
        total_salary = total_without_epf - epf
        if month_salary in months:

            if total_salary < 2000:
                commission = total_salary * 0.05
                final_total_salary = total_salary + commission
                print("Employee id = ", d[name]["ID"])
                print("Employee name = ", d[name]["Name"])
                print("Basic Salary = RM %.2f" % staff_basic_salary)
                print("Allowance = RM %.2f" % staff_allowance)
                print("Bonus = RM %.2f" % staff_bonus)
                print("Overtime = RM %.2f" % staff_bonus)
                print("Epf = RM %.2f" % epf)
                print("Commission: RM %.2f" % commission)
                print(final_total_salary)
            elif total_salary > 3000:
                tax_deduction = total_salary * 0.06
                final_total_salary = total_salary - tax_deduction
                print("Employee id = ", d[name]["ID"])
                print("Employee name = ", d[name]["Name"])
                print("Basic Salary = RM %.2f" % staff_basic_salary)
                print("Allowance = RM  %.2f" % staff_allowance)
                print("Bonus = RM  %.2f" % staff_bonus)
                print("Overtime = RM  %.2f" % staff_bonus)
                print("Epf = RM  %.2f" % epf)
                print("Tax deduction = RM  %.2f" % tax_deduction)
                print("Salary = RM %.2f" % final_total_salary)
            else:
                final_total_salary = total_salary
                print("Employee id = ", d[name]["ID"])
                print("Employee name = ", d[name]["Name"])
                print("Basic Salary = RM %.2f" % staff_basic_salary)
                print("Allowance = RM %.2f" % staff_allowance)
                print("Bonus = RM %.2f" % staff_bonus)
                print("Overtime = RM %.2f" % staff_bonus)
                print("Epf = RM %.2f" % epf)
                print("Salary = %.2f", final_total_salary)
        else:
            print("Please enter the month correctly.")

        d[name][month_salary] = final_total_salary
        commission = total_salary * 0.05
        tax_deduction = total_salary * 0.06
        filename = d[name]["Name"] + "_" + month_salary + "" + str(today_date.year) + ".xlsx"
        payslip_filename = xlsxwriter.Workbook(filename)
        payslip_file = payslip_filename.add_worksheet()

        payslip_file.write("A1", "Asia Pacific University Payslip")
        payslip_file.write("A3", "Employee Name : ")
        payslip_file.write("A4", "Employee ID : ")
        payslip_file.write("A5", "Department : ")
        payslip_file.write("A6", "Month/Year : ")
        payslip_file.write("A7", "Basic Salary : RM")
        payslip_file.write("A8", "Allowance : RM")
        payslip_file.write("A9", "Bonus : RM")
        payslip_file.write("A10", "Overtime : RM")
        payslip_file.write("A11", "EPF : RM")
        if total_salary < 2000:
            payslip_file.write("A12", "Commission : RM")
        elif total_salary > 3000:
            payslip_file.write("A12", "Tax Deduction : RM")

        payslip_file.write("A15", "Total Salary : RM")

        month_year = month_salary + " / " + str(today_date.year)

        payslip_file.write("C3", d[name]["Name"])
        payslip_file.write("C4", d[name]["ID"])
        payslip_file.write("C5", d[name]["Department"])
        payslip_file.write("C6", month_year)
        payslip_file.write("C7", d[name]["Basic_salary"])
        payslip_file.write("C8", d[name]["Allowance"])
        payslip_file.write("C9", d[name]["Bonus"])
        payslip_file.write("C10", d[name]["Overtime"])
        payslip_file.write("C11", '{:.2f}'.format(epf))

        if total_salary < 2000:
            payslip_file.write("C12", '{:.2f}'.format(commission))
        elif total_salary > 3000:
            payslip_file.write("C12", '{:.2f}'.format(tax_deduction))

        payslip_file.write("C15", '{:.2f}'.format(final_total_salary))

        payslip_filename.close()

    else:
        print("Please enter a valid employee ID....")


def get_keys(dictionary):
    result = []
    for key, value in dictionary.items():
        if type(value) is dict:
            new_keys = get_keys(value)
            for inner_key in new_keys:
                result.append(f'{inner_key}')
        else:
            result.append(key)
    return result


def updateEmployee():
    if d != {}:
        print(d.keys())
        update_key_input = input("Enter which employee profile to update: ")
        if update_key_input in d.keys():
            print(get_keys(d))
            update_value_input = input("Enter which field you wish to update: ")
            if update_value_input in get_keys(d[update_key_input]):
                if type(d[update_key_input][update_value_input]) == float:
                    update_field_value_input = int(input("Enter the value you wish to update: "))
                    d[update_key_input][update_value_input] = update_field_value_input
                else:
                    update_field_value_input = input("Enter the value you wish to update: ")
                    d[update_key_input][update_value_input] = update_field_value_input

            else:
                print("No such profile field in the list.")
        else:
            print("No such employee profile in the list.")
    else:
        print("There is no any employee profile in the list. Please add first...")


def deleteEmployee():
    if d != {}:
        delete_input = input("Enter which employee ID of dictionary to delete: ")
        if delete_input in d.keys():
            with suppress(KeyError):
                del d[delete_input]
        else:
            print("Name not found in the list.")
    else:
        print("There is no any employee profile in the list. Please add first...")


def searchPaySlip(staff_id):
    if staff_id in d.keys():
        staff_basic_salary = d[staff_id]["Basic_salary"]
        staff_allowance = d[staff_id]["Allowance"]
        staff_bonus = d[staff_id]["Bonus"]
        staff_overtime = d[staff_id]["Overtime"]
        total_without_epf = float(staff_basic_salary) + float(staff_allowance) + float(staff_bonus) + float(
            staff_overtime)
        epf = total_without_epf * 0.11
        print(months)
        payslip_month = input("Which month do you want to search: ")
        if payslip_month in get_keys(d):
            print("There is the record for your payslip in", payslip_month, "which is RM %.2f" % d[staff_id][payslip_month])
            print("Employee id = ", d[staff_id]["ID"])
            print("Employee name = ", d[staff_id]["Name"])
            print("Basic Salary = RM", d[staff_id]["Basic_salary"])
            print("Allowance = RM", d[staff_id]["Allowance"])
            print("Bonus = RM", d[staff_id]["Bonus"])
            print("Overtime = RM", d[staff_id]["Overtime"])
            print("Epf = RM", '{:.2f}'.format(epf))
        elif ValueError:
            print("No required payslip record found")
    else:
        print("Please enter a valid employee id...")


def viewPayslip(staff_id):
    history = []
    staff_basic_salary = d[staff_id]["Basic_salary"]
    staff_allowance = d[staff_id]["Allowance"]
    staff_bonus = d[staff_id]["Bonus"]
    staff_overtime = d[staff_id]["Overtime"]
    total_without_epf = float(staff_basic_salary) + float(staff_allowance) + float(staff_bonus) + float(
        staff_overtime)
    epf = total_without_epf * 0.11
    for item1 in months:
        for item2 in get_keys(d[staff_id]):
            if item1 == item2:
                history.append(item2)
            else:
                pass

    for i in history:
        print("Month =", i, ".Salary =RM %.2f" % d[staff_id][i])
        print("Employee id = ", d[staff_id]["ID"])
        print("Employee name = ", d[staff_id]["Name"])
        print("Basic Salary = RM", d[staff_id]["Basic_salary"])
        print("Allowance = RM", d[staff_id]["Allowance"])
        print("Bonus = RM", d[staff_id]["Bonus"])
        print("Overtime = RM", d[staff_id]["Overtime"])
        print("Epf = RM", '{:.2f}'.format(epf))


def main():
    while True:
        try:
            print("\n")
            print("**********APU PAYROLL**********")
            print("1. Employee Profile")
            print("2. Salary Generator")
            print("3. Pay Slip")
            user_input = int(input("Enter a number to access (-1 to quit): "))
            if user_input == 1:
                while True:
                    try:
                        print("\n")
                        print("**********APU PAYROLL**********")
                        print("1. Add Employee")
                        print("2. Remove Employee")
                        print("3. Update Employee Profile")
                        print("4. View Current List")
                        staff_input = int(input("Enter a number to modify employee profile (-1 to quit): "))
                        if staff_input == 1:
                            addEmployee()
                        elif staff_input == 2:
                            deleteEmployee()
                        elif staff_input == 3:
                            updateEmployee()
                        elif staff_input == 4:
                            print(d)
                        elif staff_input == -1:
                            break
                        else:
                            print("Enter value in range 1 to 3")

                    except ValueError:
                        print("Please enter integer value")

            elif user_input == 2:
                if d == {}:
                    print("Please enter an employee profile")
                else:
                    employee_id = input("Enter the employee ID: ")
                    generateSalary(employee_id)

            elif user_input == 3:
                while True:
                    try:
                        print("\n")
                        print("**********APU PAYROLL**********")
                        print("1. Search payslip")
                        print("2. View payslip")
                        print("* the first character of the month input must be in upper case")
                        payslip_input = int(input("Enter a number to deal with payslip (-1 to quit): "))
                        if payslip_input == 1:
                            payslip_employee_id = input("Enter the employee ID: ")
                            searchPaySlip(payslip_employee_id)
                        elif payslip_input == 2:
                            payslip_employee_id = input("Enter the employee ID: ")
                            viewPayslip(payslip_employee_id)
                        elif payslip_input == -1:
                            break
                        else:
                            print("Enter value in range 1 to 2")

                    except ValueError:
                        print("Please enter integer value")

            elif user_input == -1:
                print("Thanks for using.")
                exit()

            else:
                print("Please enter the number in range 1 to 3")

        except ValueError:
            print("Please input an integer")


if __name__ == '__main__':
    main()
