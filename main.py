import sqlite3 as sql
import pandas as pd
import openpyxl

conn = sql.connect('database.db')
cur = conn.cursor()

cur.execute('''CREATE TABLE IF NOT EXISTS Person(
Id INTEGER PRIMARY KEY AUTOINCREMENT,
Name TEXT NOT NULL,
Age INTEGER NOT NULL,
Occupation TEXT NOT NULL,
Salary INTEGER NOT NULL)''')

def add_data():
  print("ADD DATA OF PERSON\n")
  try:
    name = input('Enter person name: ')
    age = int(input('Enter person age: '))
    occupation = input('Enter person occupation: ')
    salary = int(input('Enter person salary: '))

    cur.execute('''INSERT INTO Person (Name, Age, Occupation, Salary) VALUES (?, ?, ?, ?)''',
                (name, age, occupation, salary))
    conn.commit()
    print("\n")
    print("DATA ADDED SUCCESSFULLY!")
    print("\n")
  except ValueError as e:
    print(f'Invalid input for age or salary: {e}')
  except sqlite3.Error as e:
    print(f'Database error: {e}')
  except Exception as e:
    print(f'An error occurred: {e}')

def view_data():
  try:
    data = cur.execute('''SELECT * FROM Person''').fetchall()
    print("\n")
    if data:
      for person in data:
        print(person)
    else:
      print("No data found in the table.")
    print("\n")
  except sqlite3.Error as e:
    print(f'Database error: {e}')
  except Exception as e:
    print(f'An error occurred: {e}')

def update_data():
  print("\n")
  print("1. Update by serial number")
  print("2. Update by name")
  choice = input('Enter your choice: ')

  try:
    if choice == '1':
      serial_number = int(input("Enter serial number: "))
      person = cur.execute("""SELECT * FROM Person WHERE Id=?""", (serial_number,)).fetchone()
      if person:
        print(person, "\n")
        new_name = input('Enter new name: ')
        new_age = int(input('Enter new age: '))
        new_occ = input('Enter new occupation: ')
        new_salary = int(input('Enter new salary: '))

        cur.execute("""UPDATE Person SET Name=?, Age=?, Occupation=?, Salary=? WHERE Id=?""",
                    (new_name, new_age, new_occ, new_salary, serial_number))
        conn.commit()
        print("Updated successfully!\n")
      else:
        print(f'Person with serial number {serial_number} not found.\n')
    elif choice == '2':
      name = input('Enter name: ')
      person = cur.execute("""SELECT * FROM Person WHERE Name=?""", (name,)).fetchone()
      if person:
        print(person, "\n")
        new_name = input('Enter new name: ')
        new_age = int(input('Enter new age: '))
        new_occ = input('Enter new occupation: ')
        new_salary = int(input('Enter new salary: '))

        cur.execute("""UPDATE Person SET Name=?, Age=?, Occupation=?, Salary=? WHERE Name=?""",
                    (new_name, new_age, new_occ, new_salary, name))
        conn.commit()
        print("Update successfully!")
      else:
        print(f'Person with name {name} not found.\n')
    else:
      print('Invalid choice.\n')
  except ValueError as e:
    print(f'Invalid input for age or salary: {e}')
  except sqlite3.Error as e:
    print(f'Database error: {e}')
  except Exception as e:
    print(f'An error occurred: {e}')

def delete_data():
  print("1. Delete by id")
  print("2. Delete by name")
  choice = input('Enter your choice: ')

  try:
    if choice == '1':
      id = int(input('Enter id to delete: '))
      person = cur.execute("""SELECT * FROM Person WHERE Id=?""", (id,)).fetchone()
      if person:
        print(person, "\n")
        print("1. Confirm delete")
        print("2. Cancel")
        confirm = input('Enter your choice: ')
        if confirm == '1':
          cur.execute("""DELETE FROM Person WHERE Id=?""", (id,))
          conn.commit()
          print('Deleted successfully!\n')
        else:
          print("Cancelled !!!")
      else:
        print(f'Person with id {id} not found.\n')
    elif choice == '2':
      name = input('Enter name to delete: ')
      person = cur.execute("""SELECT * FROM Person WHERE Name=?""", (name,)).fetchone()
      if person:
        print(person, "\n")
        print("1. Confirm delete")
        print("2. Cancel")
        confirm = input('Enter your choice: ')
        if confirm == '1':
          cur.execute("""DELETE FROM Person WHERE Name=?""", (name,))
          conn.commit()
          print('Deleted successfully!\n')
        else:
          print('Cancelled !!!')
      else:
        print(f'Person with name {name} not found.\n')
    else:
      print('Invalid choice.\n')
  except ValueError as e:
    print(f'Invalid input for id: {e}')
  except sqlite3.Error as e:
    print(f'Database error: {e}')
  except Exception as e:
    print(f'An error occurred: {e}')

def delete_all_data():
  print('1. Do you really want to delete all data permanently.')
  print('2. Are nahi nahi...')
  choice = input('Enter your choice: ')

  try:
    if choice == '1':
      cur.execute("""DELETE FROM Person""")
      conn.commit()
      print("All Data Deleted Successfully!")
    elif choice == '2':
      print("Cancelled !!!")
    else:
      print('Invalid choice.\n')
  except sqlite3.Error as e:
    print(f'Database error: {e}')
  except Exception as e:
    print(f'An error occurred: {e}')

def xl_data():
  try:
    all_data = cur.execute("""SELECT * FROM Person""").fetchall()
    df = pd.DataFrame(all_data)
    df.columns = ["Id", "Name", "Age", "Occupation", "Salary"]
    df.to_excel('person_data.xlsx', index=False)
    print("Successfully created [person_data.xlsx]")
    print(df)
  except sqlite3.Error as e:
    print(f'Database error: {e}')
  except Exception as e:
    print(f'An error occurred: {e}')

while True:
  print("×××××××××××××")
  print("Data Manager")
  print("1. ADD DATA")
  print("2. VIEW DATA")
  print("3. UPDATE DATA")
  print("4. DELETE DATA")
  print("5. DELETE ALL DATA")
  print("6. CONVERT TO EXCEL SHEET")
  print("7. EXIT")
  print("\n")
  choice = input('Enter your choice: ')

  try:
    if choice == '1':
      add_data()
    elif choice == '2':
      view_data()
    elif choice == '3':
      update_data()
    elif choice == '4':
      delete_data()
    elif choice == '5':
      delete_all_data()
    elif choice == '6':
      xl_data()
    elif choice == '7':
      break
    else:
      print("Invalid choice.\n")
  except ValueError as e:
    print(f'Invalid input: {e}')
  except Exception as e:
    print(f'An error occurred: {e}')

conn.close()
