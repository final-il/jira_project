import requests
import csv

JIRA_URL = "https://jira-final-il.atlassian.net/"  # or https://<your-jira-domain>/jira if self-managed
JIRA_EMAIL = "maorb@final.co.il"  # Or username for on-prem Jira
JIRA_API_TOKEN = ""     # Or password for on-prem Jira
PARENT_ISSUE_KEY = "ITDVPS-731"              # The parent story key
PROJECT_KEY = "ITDVPS"                       # The project key where the issues live

# CSV_FILE should contain your table with columns: SITE_NAME, CPU, MEM, STORAGE
CSV_FILE = "sites.csv"

# Prepare HTTP headers for Jira
headers = {
    "Accept": "application/json",
    "Content-Type": "application/json"
}

# Use HTTP Basic authentication
auth = (JIRA_EMAIL, JIRA_API_TOKEN)

# Build the sub-task creation function
def create_subtask(site_name, cpu, mem, storage):
    """
    Creates a Jira sub-task under the specified parent story, using
    the site configuration data.
    """
    url = f"{JIRA_URL}/rest/api/2/issue"
    summary_text = f"Deployment for {site_name} (CPU: {cpu}, MEM: {mem}, STORAGE: {storage}GB)"

    payload = {
        "fields": {
            "project": {"key": PROJECT_KEY},
            # Specify the parent story
            "parent": {"key": PARENT_ISSUE_KEY},
            # Summary for the sub-task
            "summary": summary_text,
            # Issue type must be "Sub-task" or whatever your sub-task name is
            "issuetype": {"name": "Sub-task"},
            # You can also add any custom fields here if needed, for example:
            # "customfield_12345": site_name  # if you have a custom field
        }
    }

    # Send POST request to create sub-task
    response = requests.post(url, headers=headers, auth=auth, json=payload)
    if response.status_code == 201:
        print(f"Sub-task created for {site_name}")
    else:
        print(f"Failed to create sub-task for {site_name}. "
              f"Status Code: {response.status_code}, Response: {response.text}")


def main():
    # Read the CSV file
    with open(CSV_FILE, mode="r", newline="", encoding="utf-8") as file:
        reader = csv.DictReader(file, fieldnames=["SITE_NAME", "CPU", "MEM", "STORAGE"])
        next(reader, None)  # Skip header row if your CSV has a header
        
        for row in reader:
            site_name = row["SITE_NAME"]
            cpu = row["CPU"]
            mem = row["MEM"]
            storage = row["STORAGE"]

            # Create a sub-task for each line
            create_subtask(site_name, cpu, mem, storage)


if __name__ == "__main__":
    main()
