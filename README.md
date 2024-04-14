# StrickerPT SPECs App

## Overview
StrickerPT SPECs App is a Python application that interfaces with a SQL Server database to manage printing specifications for products. It uses Tkinter for the GUI and `pyodbc` for database interactions. The application allows users to fetch, display, and insert customization data related to printing specifications based on the product reference.

## Requirements
- Python 3.x
- `pyodbc` module
- Tkinter module (usually included with Python)
- SQL Server (Database settings must be configured as per the user's environment)

## Setup
1. **Database Connection**: Ensure SQL Server is running and accessible. You will need to configure the database connection strings in the script to match your database server and credentials.
   
   Replace the placeholders in the connection strings:
   ```python
   'DRIVER={SQL Server};SERVER=your_server_address;DATABASE=your_database_name;UID=your_username;PWD=your_password'

## Execution

When the application starts, use the GUI to interact with the application:

Enter a Product Reference: Start by entering a valid product reference to fetch existing data.
Select Options: Based on fetched data, you will be able to select options like print location, whether full color is allowed, print technique, and size.
Insert Data: After selecting all necessary options and associating them with a subproduct ID, use the "Confirm Insert" button to insert the data into the database
