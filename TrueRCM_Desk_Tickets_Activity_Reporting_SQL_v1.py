import os
import logging
import time
from datetime import datetime, timedelta
import mysql.connector
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import psutil
import win32com.client
import humanize  # type: ignore # For human-readable file sizes

# Configure logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

# Chrome options
chrome_options = Options()
chrome_options.add_argument("--disable-usb-discovery")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--log-level=3")
chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])
chrome_options.add_argument("--headless")

# Helper Functions
def initialize_browser():
    """Initialize the Chrome WebDriver."""
    try:
        driver_path = r"C:\Automation\chromedriver.exe"
        service = Service(driver_path)
        driver = webdriver.Chrome(service=service, options=chrome_options)
        driver.maximize_window()
        logging.info("WebDriver initialized successfully.")
        return driver
    except Exception as e:
        logging.error(f"Error initializing WebDriver: {e}")
        return None


def get_db_connection():
    """Establish database connection."""
    try:
        connection = mysql.connector.connect(
            host="localhost",
            user="root",
            password="SubhanAllah@1DB",
            database="TicketActivityDB"
        )
        logging.info("Database connection established.")
        return connection
    except mysql.connector.Error as err:
        logging.error(f"Database connection error: {err}")
        return None


def get_config_details_from_db(connection):
    """Fetch configuration details from ConfigSetup table."""
    try:
        cursor = connection.cursor(dictionary=True)
        cursor.execute("SELECT * FROM ConfigSetup LIMIT 1;")
        config = cursor.fetchone()
        cursor.close()
        return config
    except Exception as e:
        logging.error(f"Error fetching configuration: {e}")
        return None


def fetch_valid_names(connection):
    """Fetch valid names from the TeamWiseSummary table."""
    try:
        query = "SELECT Name FROM TeamWiseSummary;"
        cursor = connection.cursor()
        cursor.execute(query)
        results = cursor.fetchall()
        valid_names = [row[0] for row in results]  # Extract names from the result
        cursor.close()
        logging.info(f"Fetched {len(valid_names)} valid names from TeamWiseSummary.")
        return valid_names
    except Exception as e:
        logging.error(f"Error fetching valid names: {e}")
        return []


def trigger_load_more(driver, max_attempts=10, pause_time=10):
    """Simulate pressing PageDown repeatedly to load more content."""
    logging.info("Simulating PageDown key presses to load more content.")
    body = driver.find_element(By.TAG_NAME, "body")  # Ensure the page is focused
    click_element_time = driver.find_element(By.XPATH, '//*[@id="app"]/div/div[1]/div/div/div/div/div[2]/div/h4')
    time.sleep(pause_time)
    click_element_time.click()
    for attempt in range(max_attempts):
        body.send_keys(Keys.PAGE_DOWN)
        time.sleep(pause_time)
        elements = driver.find_elements(By.CSS_SELECTOR, ".user-info__user-name")
        current_count = len(elements)
        logging.info(f"Total elements loaded: {current_count} after {attempt + 1} scrolls.")


def extract_activity_data(driver):
    """Extract activity data from the webpage."""
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CLASS_NAME, "user-info__user-name"))
        )
        data = []
        names = driver.find_elements(By.CSS_SELECTOR, ".user-info__user-name")
        activity_types = driver.find_elements(By.CSS_SELECTOR, ".user-info__event-name")
        times = driver.find_elements(By.CSS_SELECTOR, ".activity-group__list-item-content--date span")
        ticket_links = driver.find_elements(By.CSS_SELECTOR, "a.details__ticket-subject")

        for i in range(len(names)):
            name = names[i].get_attribute("title")
            activity_type = activity_types[i].get_attribute("title")
            aria_label = times[i].get_attribute("aria-label")
            date, time_val = aria_label.split(", ")
            datetime_stamp = datetime.strptime(f"{date} {time_val}", "%B %dth %Y %H:%M:%S").strftime("%Y-%m-%d %H:%M:%S")
            ticket_url = ticket_links[i].get_attribute("href")
            data.append({
                "Name": name,
                "ActivityType": activity_type,
                "Date": datetime.strptime(date, "%B %dth %Y").strftime("%Y-%m-%d"),
                "Time": time_val,
                "DateTimeStamp": datetime_stamp,
                "TicketUrl": ticket_url,
                "TimeSinceLast Activity": str(datetime.now() - datetime.strptime(f"{date} {time_val}", "%B %dth %Y %H:%M:%S")).split(".")[0]
            })
        logging.info(f"Extracted {len(data)} rows of data.")
        return data
    except Exception as e:
        logging.error(f"Error extracting data: {e}")
        return []


def populate_email_template(template, variables):
    """Replace placeholders in the email template with actual values."""
    for key, value in variables.items():
        template = template.replace(f"v{key}", str(value))
    return template


def send_summary_email(subject, body, recipient):
    """Send a summary email using Outlook."""
    try:
        if not any("OUTLOOK.EXE" in p.name() for p in psutil.process_iter()):
            os.system("start outlook")
            time.sleep(10)

        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To = recipient
        mail.Subject = subject
        mail.HTMLBody = body
        mail.Send()
        logging.info("Email sent successfully.")
    except Exception as e:
        logging.error(f"Error sending email: {e}")

def save_to_db(connection, data):
    """Save extracted activities to MySQL database."""
    try:
        # query = """
        #     INSERT INTO ExtractedActivities 
        #     (Name, ActivityType, Date, Time, DateTimeStamp, TicketUrl, TimeSinceLastActivity)
        #     VALUES (%s, %s, %s, %s, %s, %s, %s)
        # """
        query = """
            INSERT INTO ExtractedActivities (Name, ActivityType, Date, Time, DateTimeStamp, TicketUrl, TimeSinceLastActivity)
            VALUES (%s, %s, %s, %s, %s, %s, %s)
            ON DUPLICATE KEY UPDATE
            TimeSinceLastActivity = VALUES(TimeSinceLastActivity);
        """

        cursor = connection.cursor()

        for row in data:
            try:
                # Sanitize data to ensure compatibility with the database
                sanitized_row = sanitize_data(row)
                
                #logging.info(f"Inserting row: {sanitized_row}")
                cursor.execute(query, (
                    sanitized_row["Name"],
                    sanitized_row["ActivityType"],
                    sanitized_row["Date"],
                    sanitized_row["Time"],
                    sanitized_row["DateTimeStamp"],
                    sanitized_row["TicketUrl"],
                    sanitized_row["TimeSinceLast Activity"]
                ))
            except Exception as e:
                # Log any error with the specific row
                logging.error(f"Error inserting row: {row} - {e}")

        connection.commit()
        cursor.close()
        logging.info("All valid data saved successfully to the database.")
    except Exception as e:
        logging.error(f"Error saving data to database: {e}")
    finally:
        if cursor:
            cursor.close()
# Check for NULL Values
# The table schema specifies NOT NULL for several columns. Ensure no None or empty strings are being passed. Update the save_to_db function to handle this:

def sanitize_data(row):
    """Ensure data is properly sanitized and conforms to the database schema."""
    def validate_date(date_str):
        try:
            return datetime.strptime(date_str, "%Y-%m-%d").strftime("%Y-%m-%d")
        except ValueError:
            logging.warning(f"Invalid date format: {date_str}. Using default '1900-01-01'.")
            return "1900-01-01"

    def validate_datetime(datetime_str):
        try:
            return datetime.strptime(datetime_str, "%Y-%m-%d %H:%M:%S").strftime("%Y-%m-%d %H:%M:%S")
        except ValueError:
            logging.warning(f"Invalid datetime format: {datetime_str}. Using default '1900-01-01 00:00:00'.")
            return "1900-01-01 00:00:00"

    return {
        "Name": row.get("Name", "Unknown")[:255],  # Limit VARCHAR size
        "ActivityType": row.get("ActivityType", "Unknown")[:255],  # Limit VARCHAR size
        "Date": validate_date(row.get("Date", "1900-01-01")),
        "Time": row.get("Time", "00:00:00"),  # Ensure TIME format
        "DateTimeStamp": validate_datetime(row.get("DateTimeStamp", "1900-01-01 00:00:00")),
        "TicketUrl": row.get("TicketUrl", "")[:65535],  # TEXT limit
        "TimeSinceLast Activity": row.get("TimeSinceLast Activity", "Unknown")[:255]  # Limit VARCHAR size
    }




def filter_by_team(connection, valid_names, max_retries=3):
    """Remove rows from ExtractedActivities that don't match valid names."""
    query = "DELETE FROM ExtractedActivities WHERE Name NOT IN (%s);" % ','.join(['%s'] * len(valid_names))
    retries = 0

    while retries < max_retries:
        try:
            cursor = connection.cursor()
            cursor.execute(query, valid_names)
            connection.commit()
            cursor.close()
            logging.info("Filtered activities by team.")
            return  # Exit after successful execution
        except mysql.connector.Error as e:
            if e.errno == 1205:  # Lock wait timeout error code
                retries += 1
                logging.warning(f"Lock wait timeout exceeded. Retrying... ({retries}/{max_retries})")
                time.sleep(5)  # Wait before retrying
            else:
                logging.error(f"Error filtering activities by team: {e}")
                break
        finally:
            if cursor:
                cursor.close()

    logging.error("Failed to filter activities after multiple retries.")



def update_date_summary(connection, max_retries=3):
    """Update DateWiseSummary table with unique dates."""
    query = """
        INSERT INTO DateWiseSummary (Date)
        SELECT DISTINCT Date FROM ExtractedActivities
        ON DUPLICATE KEY UPDATE Date=VALUES(Date);
    """
    retries = 0

    while retries < max_retries:
        try:
            cursor = connection.cursor()
            cursor.execute(query)
            connection.commit()
            cursor.close()
            logging.info("Date-wise summary updated.")
            return  # Exit after successful execution
        except mysql.connector.Error as e:
            if e.errno == 1205:  # Lock wait timeout error code
                retries += 1
                logging.warning(f"Lock wait timeout exceeded. Retrying... ({retries}/{max_retries})")
                time.sleep(5)  # Wait before retrying
            else:
                logging.error(f"Error updating date-wise summary: {e}")
                break
        finally:
            if cursor:
                cursor.close()

    logging.error("Failed to update date-wise summary after multiple retries.")
#Code to Get Row Counts
def get_table_row_count(connection, table_name):
    """Get the total row count from a specific database table."""
    try:
        query = f"SELECT COUNT(*) FROM {table_name};"
        cursor = connection.cursor()
        cursor.execute(query)
        row_count = cursor.fetchone()[0]
        cursor.close()
        return row_count
    except mysql.connector.Error as e:
        logging.error(f"Error fetching row count for table {table_name}: {e}")
        return 0


def update_activity_summary_counts(connection):
    """Update the ActivitySummary table with counts from ExtractedActivities."""
    try:
        query = """
        UPDATE activitySummary AS a
        LEFT JOIN (
            SELECT activitytype, COUNT(*) AS total
            FROM extractedactivities
            GROUP BY activitytype
        ) AS e
        ON e.activitytype LIKE CONCAT('%', a.activitytype, '%')
        SET a.count = COALESCE(e.total, 0);
        """
        cursor = connection.cursor()
        cursor.execute(query)
        connection.commit()
        cursor.close()
        logging.info("ActivitySummary table counts updated successfully.")
    except Exception as e:
        logging.error(f"Error updating ActivitySummary table: {e}")

def update_total_count(connection, table_name):
    """
    Updates the 'Total Count' column in the specified table by summing up relevant columns.
    Args:
        connection: MySQL database connection object.
        table_name: Name of the table to update ('teamwisesummary' or 'datewisesummary').
    """
    try:
        query = f"""
        UPDATE {table_name}
        SET `Total Count` = 
            `Ticket Received` + `Forwarded ticket` + `Created Ticket` + `Viewed ticket` +
            `Assigned to` + `Changed status` + `Added a note` + `Added a tag` +
            `Followed ticket` + `Moved from` + `Wrote a reply` + `Merged to` +
            `Unassigned ticket` + `Customer` + `Unfollowed ticket` +
            `Deleted message` + `Edited a note` + `Changed priority`;
        """
        cursor = connection.cursor()
        cursor.execute(query)
        connection.commit()
        cursor.close()
        logging.info(f"Total Count column updated successfully in {table_name}.")
    except Exception as e:
        logging.error(f"Error updating Total Count in {table_name}: {e}")
#Function to Update TeamWiseSummary Columns
def update_teamwise_summary(connection):
    """
    Updates columns in teamwisesummary with the total count of records from extractedactivities
    based on activitytype and Name matching.
    Args:
        connection: MySQL database connection object.
    """
    try:
        # List of columns to update
        columns = [
            "Ticket Received", "Forwarded ticket", "Created Ticket", "Viewed ticket", "Assigned to",
            "Changed status", "Added a note", "Added a tag", "Followed ticket", "Moved from",
            "Wrote a reply", "Merged to", "Unassigned ticket", "Customer", "Unfollowed ticket",
            "Deleted message", "Edited a note", "Changed priority"
        ]

        # Generate dynamic SQL for each column
        updates = ", ".join([
            f"`{column}` = (SELECT COUNT(*) FROM extractedactivities AS e WHERE e.activitytype LIKE '%{column}%' AND e.Name = t.Name)"
            for column in columns
        ])

        # Complete SQL query
        query = f"""
        UPDATE teamwisesummary AS t
        SET {updates};
        """

        # Execute the query
        cursor = connection.cursor()
        cursor.execute(query)
        connection.commit()
        cursor.close()

        logging.info("TeamWiseSummary table updated successfully.")
    except Exception as e:
        logging.error(f"Error updating TeamWiseSummary table: {e}")

#Function to Update DateWiseSummary Columns
def update_datewise_summary(connection):
    """
    Updates columns in datewisesummary with the total count of records from extractedactivities
    based on activitytype and Date matching.
    Args:
        connection: MySQL database connection object.
    """
    try:
        # List of columns to update
        columns = [
            "Ticket Received", "Forwarded ticket", "Created Ticket", "Viewed ticket", "Assigned to",
            "Changed status", "Added a note", "Added a tag", "Followed ticket", "Moved from",
            "Wrote a reply", "Merged to", "Unassigned ticket", "Customer", "Unfollowed ticket",
            "Deleted message", "Edited a note", "Changed priority"
        ]

        # Generate dynamic SQL for each column
        updates = ", ".join([
            f"`{column}` = (SELECT COUNT(*) FROM extractedactivities AS e WHERE e.activitytype LIKE '%{column}%' AND e.Date = d.Date)"
            for column in columns
        ])

        # Complete SQL query
        query = f"""
        UPDATE datewisesummary AS d
        SET {updates};
        """

        # Execute the query
        cursor = connection.cursor()
        cursor.execute(query)
        connection.commit()
        cursor.close()

        logging.info("DateWiseSummary table updated successfully.")
    except Exception as e:
        logging.error(f"Error updating DateWiseSummary table: {e}")


def main():
    """Main execution flow."""
    script_start_time = datetime.now()

    connection = get_db_connection()
    if not connection:
        return

    config_details = get_config_details_from_db(connection)
    if not config_details:
        connection.close()
        return
# Generate the dynamic subject
    current_time = datetime.now().strftime("%d-%b-%Y %H:%M:%S")  # Format: "26-Nov-2024 19:32:26"
    email_subject = f"Activity Report Summary as of {current_time}"  # Generate subject dynamically
    
    email_recipient = config_details["email_recipient"]
    #email_subject = config_details["email_subject"]
    email_body = config_details["email_body"]
    base_url = config_details["base_url"]
    login_email = config_details["login_email"]
    login_password = config_details["login_password"]

    driver = initialize_browser()
    if not driver:
        connection.close()
        return

    try:
        driver.get(base_url)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "loginemail"))).send_keys(login_email)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "loginpassword"))).send_keys(login_password)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "loginpassword"))).send_keys(Keys.RETURN)

        WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CLASS_NAME, "user-info__user-name")))

        trigger_load_more(driver, max_attempts=1000, pause_time=1)
        table_data = extract_activity_data(driver)

        if table_data:
            #logging.info(f"Extracted data: {table_data}")
            logging.info(f"Ticket data is extracted successfully!")
            # Get existing row count before insertion
            existing_row_count = get_table_row_count(connection, "ExtractedActivities")
            logging.info(f"Existing row count in ExtractedActivities: {existing_row_count}")
            # Save the extracted data to the database
            save_to_db(connection, table_data)
            # Get latest row count after insertion
            latest_row_count = get_table_row_count(connection, "ExtractedActivities")
            logging.info(f"Latest row count in ExtractedActivities: {latest_row_count}")

            valid_names = fetch_valid_names(connection)
            filter_by_team(connection, valid_names)
            # Get latest row count after insertion
            latest_row_count = get_table_row_count(connection, "ExtractedActivities")
            logging.info(f"Latest row count in ExtractedActivities: {latest_row_count}")
                # Calculate the number of rows inserted
            rows_inserted = latest_row_count - existing_row_count
            logging.info(f"Rows inserted: {rows_inserted}")
            update_date_summary(connection)
            # Get the directory of the current script or executable
            script_dir = os.path.dirname(os.path.abspath(__file__))  # Get script directory
            #file_size = humanize.naturalsize(os.path.getsize(file_path))

            variables = {
                "ScriptName": __file__,
                "ScriptDir": script_dir,
                "ExecStartTime": script_start_time.strftime("%Y-%m-%d %H:%M:%S"),
                "ExecEndTime": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "TotalExecutionTime": str(datetime.now() - script_start_time),
                "ExistingCount": existing_row_count,
                "LatestCount": latest_row_count,
                "TotalCount": rows_inserted
            }
                # Update activity summary counts
            update_activity_summary_counts(connection)
            
            # Update TeamWiseSummary table
            update_teamwise_summary(connection)

            # Update DateWiseSummary table
            update_datewise_summary(connection)
            # Update Total Count columns
            update_total_count(connection, "teamwisesummary")
            update_total_count(connection, "datewisesummary")
            
            populated_email_body = populate_email_template(email_body, variables)
            send_summary_email(email_subject, populated_email_body, email_recipient)
        else:
            logging.warning("No data extracted.")

    finally:
        if driver:
            driver.quit()
        if connection:
            connection.close()


if __name__ == "__main__":
    main()
