import mysql.connector
from mysql.connector import Error

class MySQLDatabase:
    def __init__(self, host, database, user, password):
        self.host = host
        self.database = database
        self.user = user
        self.password = password
        self.connection = None
        self.cursor = None

    def connect(self):
        """Establish a connection to the MySQL database."""
        try:
            self.connection = mysql.connector.connect(
                host=self.host,
                database=self.database,
                user=self.user,
                password=self.password
            )
            if self.connection.is_connected():
                self.cursor = self.connection.cursor(dictionary=True)  # Return results as dictionaries
                print("Connected to MySQL database")
                return True
        except Error as e:
            print(f"Error connecting to MySQL: {e}")
            return False

    def disconnect(self):
        """Close the database connection."""
        if self.connection and self.connection.is_connected():
            self.cursor.close()
            self.connection.close()
            print("MySQL connection closed")

    def execute_query(self, query, params=None):
        """Execute a SQL query (SELECT, INSERT, UPDATE, DELETE)."""
        try:
            self.cursor.execute(query, params or ())
            self.connection.commit()  # Commit changes for INSERT/UPDATE/DELETE
            print("Query executed successfully")
            return True
        except Error as e:
            self.connection.rollback()  # Rollback in case of error
            print(f"Error executing query: {e}")
            return False

    def fetch_all(self, query, params=None):
        """Fetch all rows from a SELECT query."""
        try:
            self.cursor.execute(query, params or ())
            return self.cursor.fetchall()
        except Error as e:
            print(f"Error fetching data: {e}")
            return None

    def fetch_one(self, query, params=None):
        """Fetch a single row from a SELECT query."""
        try:
            self.cursor.execute(query, params or ())
            return self.cursor.fetchone()
        except Error as e:
            print(f"Error fetching data: {e}")
            return None

    def insert_record(self, table, data):
        """Insert a record into a table (data is a dictionary: {'column': value})."""
        columns = ', '.join(data.keys())
        placeholders = ', '.join(['%s'] * len(data))
        query = f"INSERT INTO {table} ({columns}) VALUES ({placeholders})"
        self.execute_query(query, tuple(data.values()))
        return self.cursor.lastrowid  # Return the last inserted ID

    def update_record(self, table, data, condition):
        """Update a record in a table (data = {'column': new_value}, condition = 'id=1')."""
        set_clause = ', '.join([f"{key}=%s" for key in data.keys()])
        query = f"UPDATE {table} SET {set_clause} WHERE {condition}"
        return self.execute_query(query, tuple(data.values()))