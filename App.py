# Name: Noah Feng Xiang Al-Soodinay, AdminNo: 234241K, TutorialGrp: IT2653-01
import datetime
import logging
import re
from tabulate import tabulate
from colorama import Fore, Style, init
import csv
import os
import matplotlib.pyplot as plt

# Initialize colorama to auto-reset colors after each print statement
init(autoreset=True)

# Configure logging settings
logging.basicConfig(filename='App_Logs.log', level=logging.INFO,
                    format='%(asctime)s - %(message)s', filemode='w')


# Define the Employee class to represent employee records
class Employee:
    def __init__(self, name, emp_id, department, job_title, annual_salary, employment_status, dob=None, email=None,
                 phone=None):
        self.name = name  # Employee's name
        self.emp_id = emp_id  # Unique employee ID
        self.department = department  # Department the employee belongs to
        self.job_title = job_title  # Job title of the employee
        self.annual_salary = annual_salary  # Annual salary of the employee
        self.employment_status = employment_status  # Employment status (True for active, False for inactive)
        self.dob = dob  # Date of birth of the employee (optional)
        self.email = email  # Email address of the employee (optional)
        self.phone = phone  # Phone number of the employee (optional)

    def to_list(self):
        # Convert the employee's information into a list for easy display
        dob_str = self.dob.strftime("%Y-%m-%d") if self.dob else "-"
        return [
            self.name,
            self.emp_id,
            self.department,
            self.job_title,
            f"${self.annual_salary:,.2f}",
            "Active" if self.employment_status else "Inactive",
            dob_str,
            self.email if self.email else "-",
            self.phone if self.phone else "-"
        ]


# Define the Customer class to represent customer records
class Customer:
    def __init__(self, customer_id, name, email, tier, points):
        self.customer_id = customer_id
        self.name = name
        self.email = email
        self.tier = tier
        self.points = points


# Define the CustomerRequest class to represent customer requests
class CustomerRequest:
    def __init__(self, customer_id, request_details):
        self.customer_id = customer_id
        self.request_details = request_details

    def display(self, customer):
        return [
            ["Customer ID", customer.customer_id],
            ["Name", customer.name],
            ["Email", customer.email],
            ["Tier", customer.tier],
            ["Points", customer.points],
            ["Request", self.request_details]
        ]


# Queue class for managing customer requests
class Queue:
    def __init__(self):
        self.queue = []

    def is_empty(self):
        return len(self.queue) == 0

    def enqueue(self, item, tier):
        # Insert the new request based on the priority of the tier
        if self.is_empty():
            self.queue.append(item)
        else:
            inserted = False
            for i in range(len(self.queue)):
                existing_tier = self.find_customer_by_id(self.queue[i].customer_id, customers).tier
                if existing_tier > tier:
                    self.queue.insert(i, item)
                    inserted = True
                    break
            if not inserted:
                self.queue.append(item)

    def dequeue(self):
        if self.is_empty():
            return None
        return self.queue.pop(0)

    def size(self):
        return len(self.queue)

    def display(self, customers):
        return [request.display(self.find_customer_by_id(request.customer_id, customers)) for request in self.queue]

    def find_customer_by_id(self, customer_id, customers):
        for customer in customers:
            if customer.customer_id.upper() == customer_id.upper():
                return customer
        return None


# Initialize the customer request queue
customer_requests = Queue()

# Predefined list of employees for demonstration purposes
employees = [
    Employee("Abdul Rahman", 10001, "Human Resources", "HR Manager", 82000.00, True, datetime.datetime(1990, 10, 22),
             "abdul123@gmail.com", "91234456"),
    Employee("Bobby Chan", 10002, "Engineering", "Software Engineer", 74000.50, True, None, "bobby12@yahoo.com",
             "86678724"),
    Employee("Charlie Cheng", 10003, "Human Resources", "HR Manager", 74000.50, True, None, None, None),
    Employee("David Dan", 10004, "IT", "Software Engineer", 88888.80, False, datetime.datetime(1988, 3, 20),
             "daviddan111@yahoo.com", None)
]

# Predefined list of customers for demonstration purposes
customers = [
    Customer("S222", "John Tan", "jtan@yahoo.com", "C", 2000),
    Customer("S333", "Alice Lee", "alee@gmail.com", "B", 1500),
    Customer("S444", "Michael Wong", "mwong@hotmail.com", "A", 3000),
    Customer("S555", "Tom Wee", "Twee@hotmail.com", "A", 6000)
]


# Function to display all employees in a tabular format
def display_all_employees():
    if not employees:
        # Print a warning if no employee records are found
        print(Fore.YELLOW + "No employee records found.")
    else:
        # Define table headers with colors
        headers = [
            Fore.LIGHTBLUE_EX + "Name" + Style.RESET_ALL,
            Fore.LIGHTGREEN_EX + "Employee ID" + Style.RESET_ALL,
            Fore.LIGHTCYAN_EX + "Department" + Style.RESET_ALL,
            Fore.LIGHTMAGENTA_EX + "Job Title" + Style.RESET_ALL,
            Fore.LIGHTYELLOW_EX + "Annual Salary" + Style.RESET_ALL,
            Fore.LIGHTWHITE_EX + "Employment Status" + Style.RESET_ALL,
            Fore.LIGHTBLUE_EX + "Date of Birth" + Style.RESET_ALL,
            Fore.LIGHTGREEN_EX + "Email" + Style.RESET_ALL,
            Fore.LIGHTCYAN_EX + "Phone" + Style.RESET_ALL
        ]
        # Convert employee objects to lists for tabular display
        table = [emp.to_list() for emp in employees]
        # Print the table using tabulate
        print(tabulate(table, headers, tablefmt="grid"))


# Function to display an employee by their ID
def display_employee_by_id():
    try:
        # Prompt the user to enter an employee ID
        emp_id = int(input("Enter employee ID to search: "))
        # Find the employee with the matching ID
        found_employee = next((emp for emp in employees if emp.emp_id == emp_id), None)
        if found_employee:
            # Define table headers with colors
            headers = [
                Fore.BLUE + "Name" + Style.RESET_ALL,
                Fore.GREEN + "Employee ID" + Style.RESET_ALL,
                Fore.YELLOW + "Department" + Style.RESET_ALL,
                Fore.CYAN + "Job Title" + Style.RESET_ALL,
                Fore.MAGENTA + "Annual Salary" + Style.RESET_ALL,
                Fore.RED + "Employment Status" + Style.RESET_ALL,
                Fore.LIGHTWHITE_EX + "Date of Birth" + Style.RESET_ALL,
                Fore.LIGHTBLUE_EX + "Email" + Style.RESET_ALL,
                Fore.LIGHTYELLOW_EX + "Phone" + Style.RESET_ALL
            ]
            # Convert the found employee object to a list for display
            table = [found_employee.to_list()]
            # Print the table using tabulate
            print(tabulate(table, headers, tablefmt="grid"))
        else:
            # Print a warning if the employee ID is not found
            print(Fore.RED + "Employee ID not found.")
    except ValueError:
        # Print an error if the input is not a valid integer
        print(Fore.RED + "Invalid input. Employee ID must be an integer.")


# Unique validation functions
def is_unique_employee_id(emp_id):
    return not any(emp.emp_id == emp_id for emp in employees)


def is_unique_email(email):
    return not any(emp.email == email for emp in employees if emp.email)


def is_unique_phone(phone):
    return not any(emp.phone == phone for emp in employees if emp.phone)


# Function to add a new employee
def add_new_employee():
    while True:
        try:
            name = input("Enter employee name: ").strip().title()
            if not name:
                raise ValueError("Name cannot be empty.")
            break
        except ValueError as e:
            print(Fore.RED + f"Invalid input: {e}")
            logging.error(f"Error adding employee - Name: {e}")

    while True:
        try:
            emp_id = int(input("Enter employee ID: "))
            if not is_unique_employee_id(emp_id):
                raise ValueError("Employee ID already exists. Please enter a unique employee ID.")
            break
        except ValueError as e:
            if "invalid literal for int()" in str(e):
                print(Fore.RED + "Invalid input. Employee ID must be an integer.")
                logging.error("Error adding employee - Employee ID: not an integer.")
            else:
                print(Fore.RED + f"Invalid input: {e}")
                logging.error(f"Error adding employee - Employee ID: {e}")

    while True:
        try:
            department = input("Enter department: ").strip().title()
            if not department:
                raise ValueError("Department cannot be empty.")
            break
        except ValueError as e:
            print(Fore.RED + f"Invalid input: {e}")
            logging.error(f"Error adding employee - Department: {e}")

    while True:
        try:
            job_title = input("Enter job title: ").strip().title()
            if not job_title:
                raise ValueError("Job title cannot be empty.")
            break
        except ValueError as e:
            print(Fore.RED + f"Invalid input: {e}")
            logging.error(f"Error adding employee - Job Title: {e}")

    while True:
        try:
            annual_salary = float(input("Enter annual salary: "))
            break
        except ValueError:
            print(Fore.RED + "Invalid input. Please enter a numeric value for annual salary.")
            logging.error("Error adding employee - Annual Salary: not a float.")

    while True:
        try:
            employment_status = input("Enter employment status (T for Active, F for Inactive): ").lower()
            if employment_status not in ['t', 'f']:
                raise ValueError("Employment status must be 'T' for Active or 'F' for Inactive.")
            employment_status = employment_status == 't'
            break
        except ValueError as e:
            print(Fore.RED + f"Invalid input: {e}")
            logging.error(f"Error adding employee - Employment Status: {e}")

    while True:
        try:
            dob_input = input("Enter date of birth (YYYY-MM-DD), optional (Press Enter to skip): ")
            if dob_input:
                dob = datetime.datetime.strptime(dob_input, "%Y-%m-%d")
                if dob > datetime.datetime.now():
                    raise ValueError("Date of birth cannot be in the future.")
            else:
                dob = None
            break
        except ValueError as e:
            print(Fore.RED + f"Invalid input: {e}")
            logging.error(f"Error adding employee - Date of Birth: {e}")

    while True:
        try:
            email = input("Enter email, optional (Press Enter to skip): ").lower()
            if email:
                if not re.match(r"[^@]+@[^@]+\.[^@]+", email):
                    raise ValueError("Invalid email format.")
                if not is_unique_email(email):
                    raise ValueError("Email already exists.")
            break
        except ValueError as e:
            print(Fore.RED + f"Invalid input: {e}")
            logging.error(f"Error adding employee - Email: {e}")

    while True:
        try:
            phone = input("Enter phone number, optional (Press Enter to skip): ")
            if phone:
                if not re.match(r"^\d{8}$", phone):
                    raise ValueError("Invalid phone number format. It must be exactly 8 digits.")
                if not is_unique_phone(phone):
                    raise ValueError("Phone number already exists.")
            break
        except ValueError as e:
            print(Fore.RED + f"Invalid input: {e}")
            logging.error(f"Error adding employee - Phone: {e}")

    employees.append(Employee(name, emp_id, department, job_title, annual_salary, employment_status, dob, email, phone))
    print(Fore.GREEN + "Employee added successfully.")
    logging.info(f"Added new employee: {name}, Employee ID: {emp_id}")


# Unique validation functions for customers
def is_unique_customer_id(customer_id):
    return not any(cust.customer_id.upper() == customer_id.upper() for cust in customers)

def is_unique_customer_email(email):
    return not any(cust.email.lower() == email.lower() for cust in customers if cust.email)


# Function to add a new customer
def add_new_customer():
    while True:
        try:
            customer_id = input("Enter customer ID: ").strip().upper()
            if not customer_id:
                raise ValueError("Customer ID cannot be empty.")
            if not is_unique_customer_id(customer_id):
                raise ValueError("Customer ID already exists. Please enter a unique customer ID.")
            break
        except ValueError as e:
            print(Fore.RED + f"Invalid input: {e}")
            logging.error(f"Error adding customer - Customer ID: {e}")

    while True:
        try:
            name = input("Enter customer name: ").strip().title()
            if not name:
                raise ValueError("Name cannot be empty.")
            break
        except ValueError as e:
            print(Fore.RED + f"Invalid input: {e}")
            logging.error(f"Error adding customer - Name: {e}")

    while True:
        try:
            email = input("Enter email, optional (Press Enter to skip): ").lower()
            if email:
                if not re.match(r"[^@]+@[^@]+\.[^@]+", email):
                    raise ValueError("Invalid email format.")
                if not is_unique_customer_email(email):
                    raise ValueError("Email already exists.")
            break
        except ValueError as e:
            print(Fore.RED + f"Invalid input: {e}")
            logging.error(f"Error adding customer - Email: {e}")

    while True:
        try:
            tier = input("Enter customer tier (A, B, C): ").strip().upper()
            if tier not in ['A', 'B', 'C']:
                raise ValueError("Tier must be 'A', 'B', or 'C'.")
            break
        except ValueError as e:
            print(Fore.RED + f"Invalid input: {e}")
            logging.error(f"Error adding customer - Tier: {e}")

    while True:
        try:
            points = int(input("Enter points: "))
            break
        except ValueError:
            print(Fore.RED + "Invalid input. Points must be an integer.")
            logging.error("Error adding customer - Points: not an integer.")

    customers.append(Customer(customer_id, name, email, tier, points))
    print(Fore.GREEN + "Customer added successfully.")
    logging.info(f"Added new customer: {name}, Customer ID: {customer_id}")


# Function to sort employees by salary using selection sort
def selection_sort_salary():
    while True:
        sort_order = input("Sort by salary in ascending or descending order (A/D): ").upper()
        if sort_order in ['A', 'D']:
            ascending = True if sort_order == 'A' else False
            break
        else:
            print(Fore.RED + "Invalid input. Please enter 'A' for ascending or 'D' for descending.")

    n = len(employees)
    for i in range(n):
        extreme_index = i
        for j in range(i + 1, n):
            if ascending:
                if employees[j].annual_salary < employees[extreme_index].annual_salary:
                    extreme_index = j
            else:
                if employees[j].annual_salary > employees[extreme_index].annual_salary:
                    extreme_index = j
        employees[i], employees[extreme_index] = employees[extreme_index], employees[i]

    sort_direction = "Ascending" if ascending else "Descending"
    logging.info(f"Sorted employees by salary in {sort_direction} order using Selection Sort.")
    print(Fore.GREEN + f"Employees sorted by salary in {sort_direction} order.")
    display_all_employees()

# Quick Sort function for sorting by job title
def quick_sort(arr, low, high, key, ascending=True):
    if low < high:
        pi = partition(arr, low, high, key, ascending)
        quick_sort(arr, low, pi - 1, key, ascending)
        quick_sort(arr, pi + 1, high, key, ascending)


def partition(arr, low, high, key, ascending=True):
    i = low - 1
    pivot = getattr(arr[high], key)
    for j in range(low, high):
        if (getattr(arr[j], key) < pivot and ascending) or (getattr(arr[j], key) > pivot and not ascending):
            i = i + 1
            arr[i], arr[j] = arr[j], arr[i]
    arr[i + 1], arr[high] = arr[high], arr[i + 1]
    return i + 1


def sort_by_job_title():
    while True:
        sort_order = input("Sort by job title in ascending or descending order (A/D): ").upper()
        if sort_order in ['A', 'D']:
            ascending = True if sort_order == 'A' else False
            break
        else:
            print(Fore.RED + "Invalid input. Please enter 'A' for ascending or 'D' for descending.")

    quick_sort(employees, 0, len(employees) - 1, 'job_title', ascending)
    display_all_employees()
    sort_direction = "Ascending" if ascending else "Descending"
    logging.info(f"Sorted employees by job title in {sort_direction} order using Quick Sort.")
    print(Fore.GREEN + f"Employees sorted by job title in {sort_direction} order.")


# Merge Sort function for sorting by department and employee ID
def merge_sort(arr, key1, key2, ascending=True):
    if len(arr) > 1:
        mid = len(arr) // 2
        left_half = arr[:mid]
        right_half = arr[mid:]

        merge_sort(left_half, key1, key2, ascending)
        merge_sort(right_half, key1, key2, ascending)

        i = j = k = 0

        while i < len(left_half) and j < len(right_half):
            if (getattr(left_half[i], key1), getattr(left_half[i], key2)) < (getattr(right_half[j], key1), getattr(right_half[j], key2)) and ascending:
                arr[k] = left_half[i]
                i += 1
            elif not ascending and (getattr(left_half[i], key1), getattr(left_half[i], key2)) > (getattr(right_half[j], key1), getattr(right_half[j], key2)):
                arr[k] = left_half[i]
                i += 1
            else:
                arr[k] = right_half[j]
                j += 1
            k += 1

        while i < len(left_half):
            arr[k] = left_half[i]
            i += 1
            k += 1

        while j < len(right_half):
            arr[k] = right_half[j]
            j += 1
            k += 1


def sort_by_department_and_id():
    while True:
        sort_order = input("Sort by department and employee ID in ascending or descending order (A/D): ").upper()
        if sort_order in ['A', 'D']:
            ascending = True if sort_order == 'A' else False
            break
        else:
            print(Fore.RED + "Invalid input. Please enter 'A' for ascending or 'D' for descending.")

    merge_sort(employees, 'department', 'emp_id', ascending)
    display_all_employees()
    sort_direction = "Ascending" if ascending else "Descending"
    logging.info(f"Sorted employees by department and employee ID in {sort_direction} order using Merge Sort.")
    print(Fore.GREEN + f"Employees sorted by department and employee ID in {sort_direction} order.")


# Bubble Sort function for sorting by department
def bubble_sort_department():
    bubble_sort('department')
    display_all_employees()


# Bubble Sort function with attribute choice
def bubble_sort(attribute):
    while True:
        sort_order = input(f"Sort by {attribute} in ascending or descending order (A/D): ").upper()
        if sort_order in ['A', 'D']:
            ascending = True if sort_order == 'A' else False
            break
        else:
            print(Fore.RED + "Invalid input. Please enter 'A' for ascending or 'D' for descending.")

    n = len(employees)
    for i in range(n - 1):
        for j in range(0, n - i - 1):
            if ascending:
                if getattr(employees[j], attribute) > getattr(employees[j + 1], attribute):
                    employees[j], employees[j + 1] = employees[j + 1], employees[j]
            else:
                if getattr(employees[j], attribute) < getattr(employees[j + 1], attribute):
                    employees[j], employees[j + 1] = employees[j + 1], employees[j]
    sort_direction = "Ascending" if ascending else "Descending"
    logging.info(f"Sorted employees by {attribute} in {sort_direction} order using Bubble Sort.")
    print(Fore.GREEN + f"Employees sorted by {attribute} in {sort_direction} order.")


# Function to check access control based on user role and feature
def check_access(role, feature):
    access_control = {
        'Admin': ['display_all_employees', 'add_new_employee', 'bubble_sort_department', 'selection_sort_salary',
                  'display_employee_by_id', 'sort_by_job_title', 'search_employee_by_name', 'sort_by_department_and_id',
                  'manage_customer_requests', 'export_employee_data', 'import_employee_data', 'generate_department_distribution_chart',
                  'filter_employees', 'view_search_history', 'add_new_customer'],
        'User': ['display_all_employees', 'display_employee_by_id', 'search_employee_by_name']
    }
    return feature in access_control.get(role, [])



# Function to search employee by name
def search_employee_by_name():
    search_name = input("Enter the employee name to search: ").strip().title()
    found_employees = [emp for emp in employees if search_name in emp.name]
    if found_employees:
        headers = [
            Fore.BLUE + "Name" + Style.RESET_ALL,
            Fore.GREEN + "Employee ID" + Style.RESET_ALL,
            Fore.YELLOW + "Department" + Style.RESET_ALL,
            Fore.CYAN + "Job Title" + Style.RESET_ALL,
            Fore.MAGENTA + "Annual Salary" + Style.RESET_ALL,
            Fore.RED + "Employment Status" + Style.RESET_ALL,
            Fore.LIGHTWHITE_EX + "Date of Birth" + Style.RESET_ALL,
            Fore.LIGHTBLUE_EX + "Email" + Style.RESET_ALL,
            Fore.LIGHTYELLOW_EX + "Phone" + Style.RESET_ALL
        ]
        table = [emp.to_list() for emp in found_employees]
        print(tabulate(table, headers, tablefmt="grid"))
    else:
        print(Fore.RED + "No employees found with the given name.")


# Function to merge duplicate customer requests
def merge_duplicate_requests(customer_id, request_details):
    for request in customer_requests.queue:
        if request.customer_id.upper() == customer_id.upper():
            request.request_details += f" | {request_details}"
            print(Fore.GREEN + "Duplicate request detected. Merged with existing request.")
            logging.info(f"Merged duplicate request for Customer ID: {customer_id}")
            return True
    return False


# Manage Customer Request functions
def add_customer_request():
    customer_id = input("Enter customer ID: ").strip().upper()
    customer = next((cust for cust in customers if cust.customer_id.upper() == customer_id), None)

    if not customer:
        print(Fore.RED + "Invalid Customer ID.")
        return

    request_details = input("Enter request details: ").strip()

    if not merge_duplicate_requests(customer_id, request_details):
        customer_requests.enqueue(CustomerRequest(customer_id, request_details), customer.tier)
        print(Fore.GREEN + "Customer's request added successfully")
        logging.info(f"Added new customer request: Customer ID: {customer_id}")


# Function to display customer requests with improved layout
def view_customer_requests():
    if customer_requests.is_empty():
        print(Fore.YELLOW + "No customer requests in the queue.")
    else:
        requests = customer_requests.display(customers)
        for request in requests:
            print(Fore.CYAN + "Customer Request Details:")
            print(Fore.CYAN + "-" * 40)
            print(tabulate(request[:-1], headers=[Fore.LIGHTBLUE_EX + "Field" + Style.RESET_ALL, Fore.LIGHTGREEN_EX + "Value" + Style.RESET_ALL], tablefmt="fancy_grid"))
            print(Fore.LIGHTRED_EX + "Request: " + Fore.LIGHTYELLOW_EX + request[-1][1])
            print(Fore.CYAN + "-" * 40)
            print("\n")
        print(Fore.GREEN + f"Total customer requests: {customer_requests.size()}")


# Function to process the next customer request
def process_next_request():
    if customer_requests.is_empty():
        print(Fore.YELLOW + "No customer requests to process.")
    else:
        request = customer_requests.dequeue()
        customer = customer_requests.find_customer_by_id(request.customer_id, customers)
        print(Fore.CYAN + "Customer Request Details:")
        print(Fore.CYAN + "-" * 40)
        print(tabulate(request.display(customer)[:-1], headers=[Fore.LIGHTBLUE_EX + "Field" + Style.RESET_ALL,
                                                                Fore.LIGHTGREEN_EX + "Value" + Style.RESET_ALL],
                       tablefmt="fancy_grid"))
        print(Fore.LIGHTRED_EX + "Request: " + Fore.LIGHTYELLOW_EX + request.display(customer)[-1][1])
        print(Fore.CYAN + "-" * 40)
        print(Fore.GREEN + f"Remaining requests: {customer_requests.size()}")
        logging.info(f"Processed customer request: Customer ID: {request.customer_id}")

        # Add the processed request to the history
        processed_requests_history.append(request)


# Function to display all customers in a tabular format
def display_all_customers():
    if not customers:
        print(Fore.YELLOW + "No customer records found.")
    else:
        headers = [
            Fore.LIGHTBLUE_EX + "Customer ID" + Style.RESET_ALL,
            Fore.LIGHTGREEN_EX + "Name" + Style.RESET_ALL,
            Fore.LIGHTCYAN_EX + "Email" + Style.RESET_ALL,
            Fore.LIGHTMAGENTA_EX + "Tier" + Style.RESET_ALL,
            Fore.LIGHTYELLOW_EX + "Points" + Style.RESET_ALL
        ]
        table = [[cust.customer_id, cust.name, cust.email, cust.tier, cust.points] for cust in customers]
        print(tabulate(table, headers, tablefmt="grid"))


def manage_customer_requests():
    while True:
        print(Fore.CYAN + "\nManage Customer Requests")
        print(Fore.LIGHTGREEN_EX + "1. Add customer request")
        print(Fore.LIGHTYELLOW_EX + "2. View customer requests")
        print(Fore.LIGHTMAGENTA_EX + "3. Process next customer request")
        print(Fore.LIGHTCYAN_EX + "4. View all customers")
        print(Fore.LIGHTWHITE_EX + "5. View processed requests history")
        print(Fore.LIGHTRED_EX + "0. Exit to main menu")

        choice = input(Fore.CYAN + "Enter your choice: ")

        if choice == '1':
            add_customer_request()
        elif choice == '2':
            view_customer_requests()
        elif choice == '3':
            process_next_request()
        elif choice == '4':
            display_all_customers()
        elif choice == '5':
            view_processed_requests_history()
        elif choice == '0':
            return
        else:
            print(Fore.RED + "Invalid choice. Please enter a number between 0 and 5.")
            logging.error("Invalid choice entered in manage customer requests menu.")


# Function to export employee data to a CSV file
def export_employee_data():
    filename = input("Enter the filename to export to (without extension): ").strip()
    if not filename:
        print(Fore.RED + "Filename cannot be empty.")
        return

    filename += ".csv"
    headers = ["Name", "Employee ID", "Department", "Job Title", "Annual Salary", "Employment Status", "Date of Birth",
               "Email", "Phone"]

    try:
        with open(filename, mode='w', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(headers)
            for emp in employees:
                writer.writerow(emp.to_list())
        print(Fore.GREEN + f"Employee data exported successfully to {filename}")
        logging.info(f"Employee data exported to {filename}")
    except Exception as e:
        print(Fore.RED + f"Error exporting employee data: {e}")
        logging.error(f"Error exporting employee data: {e}")


# Function to import employee data from a CSV file
def import_employee_data():
    filename = input("Enter the filename to import from (with .csv extension): ").strip()
    if not filename:
        print(Fore.RED + "Filename cannot be empty.")
        return

    if not filename.endswith(".csv"):
        print(Fore.RED + "Invalid file type. Please provide a .csv file.")
        return

    if not os.path.isfile(filename):
        print(Fore.RED + f"File {filename} does not exist.")
        return

    required_fields = ["Name", "Employee ID", "Department", "Job Title", "Annual Salary", "Employment Status",
                       "Date of Birth", "Email", "Phone"]

    def parse_date(date_str):
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
            try:
                return datetime.datetime.strptime(date_str, fmt)
            except ValueError:
                continue
        raise ValueError(f"time data '{date_str}' does not match any accepted format")

    try:
        with open(filename, mode='r') as file:
            reader = csv.DictReader(file)
            missing_fields = [field for field in required_fields if field not in reader.fieldnames]
            if missing_fields:
                print(Fore.RED + f"Missing fields in the CSV file: {', '.join(missing_fields)}")
                return

            for row in reader:
                try:
                    name = row["Name"].strip().title()
                    if not name:
                        raise ValueError("Name cannot be empty.")

                    emp_id = int(row["Employee ID"])
                    if not is_unique_employee_id(emp_id):
                        raise ValueError(f"Duplicate Employee ID found: {emp_id}")

                    department = row["Department"].strip().title()
                    if not department:
                        raise ValueError("Department cannot be empty.")

                    job_title = row["Job Title"].strip().title()
                    if not job_title:
                        raise ValueError("Job title cannot be empty.")

                    annual_salary = float(row["Annual Salary"].replace("$", "").replace(",", ""))

                    employment_status = row["Employment Status"].strip().lower()
                    if employment_status not in ['active', 'inactive']:
                        raise ValueError("Employment status must be 'Active' or 'Inactive'.")
                    employment_status = employment_status == 'active'

                    dob_input = row["Date of Birth"].strip()
                    dob = None
                    if dob_input and dob_input != "-":
                        dob = parse_date(dob_input)
                        if dob > datetime.datetime.now():
                            raise ValueError("Date of birth cannot be in the future.")

                    email = row["Email"].strip().lower()
                    if email and email != "-":
                        if not re.match(r"[^@]+@[^@]+\.[^@]+", email):
                            raise ValueError("Invalid email format.")
                        if not is_unique_email(email):
                            raise ValueError(f"Duplicate Email found: {email}")

                    phone = row["Phone"].strip()
                    if phone and phone != "-":
                        if not re.match(r"^\d{8}$", phone):
                            raise ValueError("Invalid phone number format. It must be exactly 8 digits.")
                        if not is_unique_phone(phone):
                            raise ValueError(f"Duplicate Phone found: {phone}")

                    employees.append(Employee(
                        name, emp_id, department, job_title, annual_salary, employment_status, dob, email if email != "-" else None, phone if phone != "-" else None
                    ))
                except ValueError as e:
                    print(Fore.RED + f"Error processing row {row}: {e}")
                    logging.error(f"Error processing row {row}: {e}")

        print(Fore.GREEN + "Employee data imported successfully.")
        logging.info(f"Employee data imported from {filename}")
    except Exception as e:
        print(Fore.RED + f"Error importing employee data: {e}")
        logging.error(f"Error importing employee data: {e}")


def generate_department_distribution_chart():
    try:
        if not employees:
            print(Fore.YELLOW + "No employee records found. Cannot generate chart.")
            return

        # Calculate the distribution of employees across departments
        department_counts = {}
        for emp in employees:
            if emp.department in department_counts:
                department_counts[emp.department] += 1
            else:
                department_counts[emp.department] = 1

        # Generate the bar chart
        departments = list(department_counts.keys())
        counts = list(department_counts.values())

        plt.figure(figsize=(10, 6))
        plt.bar(departments, counts, color='skyblue')
        plt.xlabel('Department')
        plt.ylabel('Number of Employees')
        plt.title('Distribution of Employees Across Departments')
        plt.xticks(rotation=45)
        plt.tight_layout()

        # Save the chart as an image file
        chart_filename = 'department_distribution_chart.png'
        plt.savefig(chart_filename)
        plt.show()

        print(Fore.GREEN + f"Department distribution chart generated and saved as {chart_filename}.")
        logging.info(f"Department distribution chart generated and saved as {chart_filename}.")
    except Exception as e:
        print(Fore.RED + f"Error generating department distribution chart: {e}")
        logging.error(f"Error generating department distribution chart: {e}")


# Function to filter employees
def filter_employees():
    criteria = {}

    department = input("Enter department to filter by (leave empty to skip): ").strip().title()
    if department:
        criteria['department'] = department

    job_title = input("Enter job title to filter by (leave empty to skip): ").strip().title()
    if job_title:
        criteria['job_title'] = job_title

    employment_status = input("Enter employment status to filter by (Active/Inactive, leave empty to skip): ").strip().title()
    if employment_status in ['Active', 'Inactive']:
        criteria['employment_status'] = True if employment_status == 'Active' else False

    filtered_employees = employees
    for key, value in criteria.items():
        if key == 'employment_status':
            filtered_employees = [emp for emp in filtered_employees if getattr(emp, key) == value]
        else:
            filtered_employees = [emp for emp in filtered_employees if getattr(emp, key).lower().startswith(value.lower())]

    if not filtered_employees:
        print(Fore.YELLOW + "No employees match the given criteria.")
    else:
        headers = [
            Fore.LIGHTBLUE_EX + "Name" + Style.RESET_ALL,
            Fore.LIGHTGREEN_EX + "Employee ID" + Style.RESET_ALL,
            Fore.LIGHTCYAN_EX + "Department" + Style.RESET_ALL,
            Fore.LIGHTMAGENTA_EX + "Job Title" + Style.RESET_ALL,
            Fore.LIGHTYELLOW_EX + "Annual Salary" + Style.RESET_ALL,
            Fore.LIGHTWHITE_EX + "Employment Status" + Style.RESET_ALL,
            Fore.LIGHTBLUE_EX + "Date of Birth" + Style.RESET_ALL,
            Fore.LIGHTGREEN_EX + "Email" + Style.RESET_ALL,
            Fore.LIGHTCYAN_EX + "Phone" + Style.RESET_ALL
        ]
        table = [emp.to_list() for emp in filtered_employees]
        print(tabulate(table, headers, tablefmt="grid"))

    return criteria


# Initialize search history list
search_history = []

# Function to view and re-run previous searches
def view_search_history():
    if not search_history:
        print(Fore.YELLOW + "No search history available.")
        return

    print(Fore.CYAN + "\nSearch History")
    for index, criteria in enumerate(search_history):
        display_criteria = criteria.copy()
        if 'employment_status' in display_criteria:
            display_criteria['employment_status'] = 'Active' if display_criteria['employment_status'] else 'Inactive'
        print(Fore.CYAN + f"{index + 1}. {display_criteria}")

    try:
        choice = int(input(Fore.CYAN + "Enter the search number to re-run (0 to cancel): "))
        if choice == 0:
            return
        elif 1 <= choice <= len(search_history):
            criteria = search_history[choice - 1]
            filtered_employees = employees
            for key, value in criteria.items():
                if key == 'employment_status':
                    filtered_employees = [emp for emp in filtered_employees if getattr(emp, key) == value]
                else:
                    filtered_employees = [emp for emp in filtered_employees if getattr(emp, key).lower().startswith(value.lower())]

            if not filtered_employees:
                print(Fore.YELLOW + "No employees match the given criteria.")
            else:
                headers = [
                    Fore.LIGHTBLUE_EX + "Name" + Style.RESET_ALL,
                    Fore.LIGHTGREEN_EX + "Employee ID" + Style.RESET_ALL,
                    Fore.LIGHTCYAN_EX + "Department" + Style.RESET_ALL,
                    Fore.LIGHTMAGENTA_EX + "Job Title" + Style.RESET_ALL,
                    Fore.LIGHTYELLOW_EX + "Annual Salary" + Style.RESET_ALL,
                    Fore.LIGHTWHITE_EX + "Employment Status" + Style.RESET_ALL,
                    Fore.LIGHTBLUE_EX + "Date of Birth" + Style.RESET_ALL,
                    Fore.LIGHTGREEN_EX + "Email" + Style.RESET_ALL,
                    Fore.LIGHTCYAN_EX + "Phone" + Style.RESET_ALL
                ]
                table = [emp.to_list() for emp in filtered_employees]
                print(tabulate(table, headers, tablefmt="grid"))
        else:
            print(Fore.RED + "Invalid choice. Please enter a valid number.")
    except ValueError:
        print(Fore.RED + "Invalid input. Please enter a number.")


processed_requests_history = []
def process_next_request():
    if customer_requests.is_empty():
        print(Fore.YELLOW + "No customer requests to process.")
    else:
        request = customer_requests.dequeue()
        customer = customer_requests.find_customer_by_id(request.customer_id, customers)
        print(Fore.CYAN + "Customer Request Details:")
        print(Fore.CYAN + "-" * 40)
        print(tabulate(request.display(customer)[:-1], headers=[Fore.LIGHTBLUE_EX + "Field" + Style.RESET_ALL, Fore.LIGHTGREEN_EX + "Value" + Style.RESET_ALL], tablefmt="fancy_grid"))
        print(Fore.LIGHTRED_EX + "Request: " + Fore.LIGHTYELLOW_EX + request.display(customer)[-1][1])
        print(Fore.CYAN + "-" * 40)
        print(Fore.GREEN + f"Remaining requests: {customer_requests.size()}")
        logging.info(f"Processed customer request: Customer ID: {request.customer_id}")

        # Add the processed request to the history
        processed_requests_history.append(request)


def view_processed_requests_history():
    if not processed_requests_history:
        print(Fore.YELLOW + "No processed requests history available.")
    else:
        print(Fore.CYAN + "\nProcessed Requests History")
        headers = [
            Fore.LIGHTBLUE_EX + "Customer ID" + Style.RESET_ALL,
            Fore.LIGHTGREEN_EX + "Name" + Style.RESET_ALL,
            Fore.LIGHTCYAN_EX + "Email" + Style.RESET_ALL,
            Fore.LIGHTMAGENTA_EX + "Tier" + Style.RESET_ALL,
            Fore.LIGHTYELLOW_EX + "Points" + Style.RESET_ALL,
            Fore.LIGHTWHITE_EX + "Request Details" + Style.RESET_ALL
        ]
        table = []
        for request in processed_requests_history:
            customer = customer_requests.find_customer_by_id(request.customer_id, customers)
            table.append([
                customer.customer_id,
                customer.name,
                customer.email,
                customer.tier,
                customer.points,
                request.request_details
            ])
        print(tabulate(table, headers, tablefmt="grid"))

        while True:
            export_choice = input(Fore.CYAN + "Do you want to export the processed requests history? (Y/N): ").strip().upper()
            if export_choice == 'Y':
                export_processed_requests_history()
                break
            elif export_choice == 'N':
                break
            else:
                print(Fore.RED + "Invalid input. Please enter 'Y' for Yes or 'N' for No.")


def export_processed_requests_history():
    filename = input("Enter the filename to export the history to (without extension): ").strip()
    if not filename:
        print(Fore.RED + "Filename cannot be empty.")
        return

    filename += ".csv"
    headers = ["Customer ID", "Name", "Email", "Tier", "Points", "Request Details"]

    try:
        with open(filename, mode='w', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(headers)
            for request in processed_requests_history:
                customer = customer_requests.find_customer_by_id(request.customer_id, customers)
                writer.writerow([
                    customer.customer_id,
                    customer.name,
                    customer.email,
                    customer.tier,
                    customer.points,
                    request.request_details
                ])
        print(Fore.GREEN + f"Processed requests history exported successfully to {filename}")
        logging.info(f"Processed requests history exported to {filename}")
    except Exception as e:
        print(Fore.RED + f"Error exporting processed requests history: {e}")
        logging.error(f"Error exporting processed requests history: {e}")


def main_menu():
    try:
        while True:
            role = input(Fore.CYAN + "Enter your role (Admin/User or type 'q' to exit): ").capitalize()
            if role == 'Q':
                print(Fore.YELLOW + "Exiting program. Goodbye!")
                logging.info("Program exited by user.")
                return
            if role == 'Admin':
                password = input(Fore.CYAN + "Enter password for Admin: ")
                if password != 'admin':
                    print(Fore.RED + "Invalid password. Access denied.")
                    logging.error("Access denied - Invalid admin password.")
                    continue
            elif role == 'User':
                pass
            else:
                print(Fore.RED + "Invalid role. Please enter 'Admin', 'User', or 'q'.")
                logging.error("Invalid role entered.")
                continue

            while True:
                print(Fore.CYAN + "\nEmployee Management System")
                print(Fore.LIGHTBLUE_EX + "1. Display all employee records")
                print(Fore.LIGHTGREEN_EX + "2. Display all customer records")
                if check_access(role, 'add_new_employee'):
                    print(Fore.LIGHTYELLOW_EX + "3. Add new employee record")
                if check_access(role, 'add_new_customer'):
                    print(Fore.LIGHTMAGENTA_EX + "4. Add new customer record")
                if check_access(role, 'bubble_sort_department'):
                    print(Fore.LIGHTCYAN_EX + "5. Sort employees by department (Bubble Sort)")
                if check_access(role, 'selection_sort_salary'):
                    print(Fore.LIGHTRED_EX + "6. Sort employees by salary (Selection Sort)")
                if check_access(role, 'sort_by_job_title'):
                    print(Fore.LIGHTWHITE_EX + "7. Sort employees by job title (Quick Sort)")
                if check_access(role, 'sort_by_department_and_id'):
                    print(Fore.LIGHTBLUE_EX + "8. Sort employees by department and employee ID (Merge Sort)")
                print(Fore.LIGHTGREEN_EX + "9. Display employee by ID")
                if check_access(role, 'search_employee_by_name'):
                    print(Fore.LIGHTYELLOW_EX + "10. Search employee by name")
                if check_access(role, 'filter_employees'):
                    print(Fore.LIGHTMAGENTA_EX + "11. Filter employees by criteria")
                if check_access(role, 'view_search_history'):
                    print(Fore.LIGHTCYAN_EX + "12. View search history")
                if check_access(role, 'export_employee_data'):
                    print(Fore.LIGHTRED_EX + "13. Export employee data to CSV")
                if check_access(role, 'import_employee_data'):
                    print(Fore.LIGHTWHITE_EX + "14. Import employee data from CSV")
                if check_access(role, 'generate_department_distribution_chart'):
                    print(Fore.LIGHTBLUE_EX + "15. Generate department distribution chart")
                if check_access(role, 'manage_customer_requests'):
                    print(Fore.LIGHTGREEN_EX + "16. Manage customer requests")
                print(Fore.LIGHTRED_EX + "0. Exit")

                choice = input(Fore.CYAN + "Enter your choice: ")

                if choice == '1':
                    display_all_employees()
                elif choice == '2':
                    display_all_customers()
                elif choice == '3' and check_access(role, 'add_new_employee'):
                    add_new_employee()
                elif choice == '4' and check_access(role, 'add_new_customer'):
                    add_new_customer()
                elif choice == '5' and check_access(role, 'bubble_sort_department'):
                    bubble_sort_department()
                elif choice == '6' and check_access(role, 'selection_sort_salary'):
                    selection_sort_salary()
                elif choice == '7' and check_access(role, 'sort_by_job_title'):
                    sort_by_job_title()
                elif choice == '8' and check_access(role, 'sort_by_department_and_id'):
                    sort_by_department_and_id()
                elif choice == '9':
                    display_employee_by_id()
                elif choice == '10' and check_access(role, 'search_employee_by_name'):
                    search_employee_by_name()
                elif choice == '11' and check_access(role, 'filter_employees'):
                    criteria = filter_employees()
                    search_history.append(criteria)
                elif choice == '12' and check_access(role, 'view_search_history'):
                    view_search_history()
                elif choice == '13' and check_access(role, 'export_employee_data'):
                    export_employee_data()
                elif choice == '14' and check_access(role, 'import_employee_data'):
                    import_employee_data()
                elif choice == '15' and check_access(role, 'generate_department_distribution_chart'):
                    generate_department_distribution_chart()
                elif choice == '16' and check_access(role, 'manage_customer_requests'):
                    manage_customer_requests()
                elif choice == '0':
                    print(Fore.YELLOW + "Exiting program. Goodbye!")
                    logging.info("Program exited by user.")
                    return
                else:
                    if choice not in ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13', '14', '15', '16', '0']:
                        print(Fore.RED + "Invalid choice. Please enter a number between 0 and 16.")
                        logging.error("Invalid choice entered.")
                    else:
                        print(Fore.RED + "Access denied.")
                        logging.error("Access denied.")
    except KeyboardInterrupt:
        print(Fore.YELLOW + "\nProgram interrupted. Exiting...")
        logging.info("Program interrupted by user.")


if __name__ == "__main__":
    main_menu()
