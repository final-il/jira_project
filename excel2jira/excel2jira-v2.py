import pandas as pd
from jira import JIRA
from dotenv import load_dotenv
import os
from datetime import datetime
import logging
from difflib import SequenceMatcher
import sys
import argparse  # Add this import

# Logging setup
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',
    force=True  # Ensure our logging configuration takes precedence
)

logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

# Add a console handler if one doesn't exist
if not logger.handlers:
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.DEBUG)
    formatter = logging.Formatter('%(asctime)s [%(levelname)s] %(message)s')
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

logger.info("Starting script execution...")
logger.info("Current working directory: %s", os.getcwd())
logger.info("Python version: %s", sys.version)

# Load credentials
load_dotenv()
logger.info("Environment variables loaded")

JIRA_URL = os.getenv("JIRA_URL")
JIRA_EMAIL = os.getenv("JIRA_EMAIL")
JIRA_API_TOKEN = os.getenv("JIRA_API_TOKEN")

# Validate environment variables
required_env_vars = {
    "JIRA_URL": JIRA_URL,
    "JIRA_EMAIL": JIRA_EMAIL,
    "JIRA_API_TOKEN": JIRA_API_TOKEN
}

missing_vars = [var for var, value in required_env_vars.items() if not value]
if missing_vars:
    logger.critical(f"Missing required environment variables: {', '.join(missing_vars)}")
    logger.critical("Please ensure these variables are set in your .env file")
    exit(1)

logger.info(f"Using Jira URL: {JIRA_URL}")
logger.info(f"Using Jira email: {JIRA_EMAIL}")

DEFAULT_PROJECT_KEY = 'ITDVPS'
QBV_PROJECT_KEY = 'CQ'

# Custom fields
DOD_FIELD_ID = 'customfield_10269'
REQUESTED_BY_GROUP_FIELD_ID = 'customfield_10265'
DOD_NEW_FIELD_ID = 'customfield_10115'
DEST_TEAM_FIELD_ID = 'customfield_10114'
YEAR_FIELD_ID = 'customfield_10268'
QUARTER_FIELD_ID = 'customfield_10257'
IN_QUARTER_PLAN_FIELD_ID = 'customfield_10239'
QBV_GROUP_FIELD_ID = 'customfield_10256'  # Required field for QBV with CSI value

# Connect to Jira
jira = JIRA(server=JIRA_URL, basic_auth=(JIRA_EMAIL, JIRA_API_TOKEN))
logger.info("Connected to Jira successfully.")

def similar(a, b):
    """Calculate string similarity ratio"""
    return SequenceMatcher(None, a.lower(), b.lower()).ratio()

def find_jira_user(search_value):
    """Find a Jira user by their display name, email, or username with fuzzy matching"""
    logger.debug(f"Attempting to find user: {search_value}")
    try:
        # Try email pattern first if no @ in search value
        if '@' not in search_value:
            try:
                # Try different email patterns
                email_patterns = [
                    f"{search_value}@final.co.il",
                    f"{search_value.lower()}@final.co.il",
                    f"{search_value.replace(' ', '')}@final.co.il",
                    f"{search_value.replace(' ', '.').lower()}@final.co.il"
                ]
                
                for email_pattern in email_patterns:
                    logger.debug(f"Trying email pattern: {email_pattern}")
                    params = {
                        'query': email_pattern,
                        'maxResults': 1
                    }
                    users = jira._get_json('user/search', params=params)
                    
                    if users and len(users) > 0:
                        user = users[0]
                        logger.info(f"Found user with email pattern: {user.get('displayName')} "
                                  f"(Email: {user.get('emailAddress', 'N/A')})")
                        return user.get('accountId')
            except Exception as e:
                logger.debug(f"Email pattern search failed: {e}")
        
        # Try direct user search with fuzzy matching
        try:
            logger.debug(f"Attempting direct user search for: {search_value}")
            params = {
                'query': search_value,
                'maxResults': 50
            }
            users = jira._get_json('user/search', params=params)
            logger.debug(f"Found {len(users) if users else 0} potential matches")
            
            if users:
                # Filter and sort users by relevance
                filtered_users = []
                for user in users:
                    # Calculate similarity scores
                    name_similarity = similar(search_value, user.get('name', ''))
                    display_similarity = similar(search_value, user.get('displayName', ''))
                    email_similarity = similar(search_value, user.get('emailAddress', ''))
                    
                    # Get the highest similarity score
                    max_similarity = max(name_similarity, display_similarity, email_similarity)
                    
                    # Use a lower threshold for fuzzy matching
                    if max_similarity > 0.6:
                        filtered_users.append((user, max_similarity))
                
                # Sort by similarity score
                filtered_users.sort(key=lambda x: x[1], reverse=True)
                
                if filtered_users:
                    best_match = filtered_users[0][0]
                    logger.info(f"Found best matching user: {best_match.get('displayName')} "
                              f"(Name: {best_match.get('name', 'N/A')}, "
                              f"Email: {best_match.get('emailAddress', 'N/A')}, "
                              f"Similarity: {filtered_users[0][1]:.2f})")
                    return best_match.get('accountId')
        except Exception as e:
            logger.error(f"Error in direct search: {e}")
            
        logger.warning(f"No user found matching '{search_value}'")
        return None
    except Exception as e:
        logger.error(f"Failed to find user '{search_value}': {e}")
        return None

def get_allowed_value_id(field_id, search_value, project_key=DEFAULT_PROJECT_KEY, issuetype='Epic'):
    fields = jira.createmeta(projectKeys=project_key, issuetypeNames=issuetype, expand='projects.issuetypes.fields')
    allowed_list = []  # Initialize the list outside the loop
    
    for project in fields['projects']:
        for issuetype in project['issuetypes']:
            allowed_values = issuetype['fields'].get(field_id, {}).get('allowedValues', [])
            allowed_list = [val['value'] for val in allowed_values]

            for value in allowed_values:
                if value['value'].lower() == search_value.lower():
                    return {'id': value['id']}
    
    # Only show allowed values if the search value wasn't found
    raise ValueError(f"Value '{search_value}' not found for field '{field_id}'. Allowed values: {allowed_list}")

# Add this before the script execution
def parse_arguments():
    """Parse command line arguments"""
    parser = argparse.ArgumentParser(description='Create Jira issues from Excel file')
    parser.add_argument('-p', '--project', type=str, default='ITDVPS',
                        help='Project key (default: ITDVPS)')
    parser.add_argument('-q', '--quarter', type=str, choices=['Q1', 'Q2', 'Q3', 'Q4'],
                        help='Quarter (Q1, Q2, Q3, Q4). If not provided, will use next quarter')
    return parser.parse_args()

# Test issue creation with minimal fields
try:
    # Parse command line arguments
    args = parse_arguments()
    
    # Set project key from command line argument
    DEFAULT_PROJECT_KEY = args.project
    logger.info(f"Using project key: {DEFAULT_PROJECT_KEY}")
    
    # First verify authentication
    try:
        # Try to get user info to verify authentication
        myself = jira.myself()
        logger.info(f"Successfully authenticated as user: {myself['displayName']}")
    except Exception as e:
        logger.critical(f"Authentication failed: {e}")
        logger.critical("Please check your JIRA_EMAIL and JIRA_API_TOKEN environment variables")
        exit(1)

    # Try to get project info to verify permissions
    try:
        project = jira.project(DEFAULT_PROJECT_KEY)
        logger.info(f"Successfully accessed project: {project.key} - {project.name}")
    except Exception as e:
        logger.critical(f"Failed to access project {DEFAULT_PROJECT_KEY}: {e}")
        logger.critical("Please verify you have permission to access this project")
        exit(1)

    # Get the required field values
    try:
        requested_by_group_id = get_allowed_value_id(REQUESTED_BY_GROUP_FIELD_ID, 'Dev')
        year_value_id = get_allowed_value_id(YEAR_FIELD_ID, '2025')
        
        # Determine quarter based on command line argument or current date
        if args.quarter:
            quarter_value = args.quarter
            logger.info(f"Using specified quarter: {quarter_value}")
        else:
            current_quarter = ((datetime.today().month - 1) // 3 + 1) % 4 + 1
            quarter_value = f'Q{current_quarter}'
            logger.info(f"Using next quarter: {quarter_value}")
        
        quarter_value_id = get_allowed_value_id(QUARTER_FIELD_ID, quarter_value)
        in_quarter_plan_yes_id = get_allowed_value_id(IN_QUARTER_PLAN_FIELD_ID, 'Yes')
        
        # Get CSI value for QBV group field
        qbv_group_value_id = get_allowed_value_id(QBV_GROUP_FIELD_ID, 'CSI', project_key=QBV_PROJECT_KEY, issuetype='QBV')
        
        logger.info("Successfully got all required field values")
    except Exception as e:
        logger.critical(f"Failed to get required field values: {e}")
        exit(1)

    # Process Excel file
    excel_file = r'C:\Users\Maorb\Documents\Devops-plan-25Q2.xlsx'
    try:
        if not os.path.exists(excel_file):
            logger.critical(f"Excel file not found at: {excel_file}")
            logger.critical("Please ensure the file exists at the specified location")
            exit(1)
            
        df = pd.read_excel(excel_file, engine='openpyxl')
        logger.info(f"Excel file '{excel_file}' loaded successfully.")
        logger.info(f"Found {len(df)} rows to process")
    except Exception as e:
        logger.critical(f"Failed loading Excel file: {e}")
        logger.critical(f"Please ensure the file '{excel_file}' exists and is accessible")
        exit(1)

    # Validate columns
    expected_columns = ['Project name', 'Final DoD', 'Q2 DoD', 'Project Manager', 'Issue type', 'Description', 'lead', 'parent']
    missing_columns = [col for col in expected_columns if col not in df.columns]
    if missing_columns:
        logger.critical(f"The following columns are missing in the Excel file: {', '.join(missing_columns)}")
        logger.critical("Please ensure your Excel file has all required columns")
        exit(1)

    # Add these variables at the top of the script after the imports
    failed_tasks = []
    failed_users = []
    failed_fields = []
    skipped_issues = []  # New list to track skipped issues
    created_issues = []  # New list to track created issues

    def check_issue_exists(summary, project_key):
        """Check if an issue with the given summary already exists in the project"""
        try:
            jql = f'project = {project_key} AND summary ~ "{summary}"'
            existing_issues = jira.search_issues(jql, maxResults=1)
            return len(existing_issues) > 0
        except Exception as e:
            logger.error(f"Error checking for existing issue: {e}")
            return False

    # Process rows
    for idx, row in df.iterrows():
        project_name = str(row['Project name']).replace('"', "'")
        
        # Get Project Manager's account ID for proper mention
        project_manager_account_id = find_jira_user(row['Project Manager'])
        project_manager_mention = f"@[~{project_manager_account_id}]" if project_manager_account_id else f"@{row['Project Manager']}"
        
        # Format the description with proper Jira mentions
        description = f"{row['Description']} \n\n**Final DoD:** {row['Final DoD']}\n**Project Manager:** {project_manager_mention}"
        dod_content = str(row['Q2 DoD']) if pd.notna(row['Q2 DoD']) else ""  # Handle NaN values
        lead = row['lead'] if pd.notna(row['lead']) else None
        issue_type = row['Issue type'] if pd.notna(row['Issue type']) else None
        parent_link = row['parent'] if pd.notna(row['parent']) else None

        epic_fields = {
            'project': {'key': DEFAULT_PROJECT_KEY},
            'summary': project_name,
            'description': description,
            DOD_FIELD_ID: dod_content,
            REQUESTED_BY_GROUP_FIELD_ID: requested_by_group_id,
            DOD_NEW_FIELD_ID: dod_content,
            DEST_TEAM_FIELD_ID: {'value': 'IT_DevOps Team'},
            YEAR_FIELD_ID: year_value_id,
            QUARTER_FIELD_ID: quarter_value_id,
            IN_QUARTER_PLAN_FIELD_ID: in_quarter_plan_yes_id,
            'labels': ['excel2jira']
        }

        # Handle user assignment
        if lead:
            user_account_id = find_jira_user(lead)
            if user_account_id:
                epic_fields['assignee'] = {'accountId': user_account_id}
            else:
                logger.warning(f"Could not find Jira user for lead '{lead}' at row {idx + 2}")
                failed_users.append(f"Row {idx + 2}: {project_name} - Could not find Jira user for lead '{lead}'")

        try:
            if issue_type and issue_type.lower() == 'epic':
                # Check if epic already exists
                if check_issue_exists(project_name, DEFAULT_PROJECT_KEY):
                    logger.info(f"Epic '{project_name}' already exists, skipping creation")
                    skipped_issues.append(f"Epic: {project_name}")
                    continue
                
                epic_fields['issuetype'] = {'name': 'Epic'}
                if parent_link:
                    epic_fields['parent'] = {'key': parent_link}
                else:
                    epic_fields['project'] = {'key': DEFAULT_PROJECT_KEY}
                issue = jira.create_issue(fields=epic_fields)
                logger.info(f"Epic created: {issue.key}")
                created_issues.append(f"Epic: {project_name} ({issue.key})")

            elif issue_type and issue_type.lower() == 'qbv':
                # Check if QBV already exists
                if check_issue_exists(project_name, QBV_PROJECT_KEY):
                    logger.info(f"QBV '{project_name}' already exists, skipping creation")
                    skipped_issues.append(f"QBV: {project_name}")
                    continue
                
                # Create QBV-specific fields
                qbv_fields = {
                    'project': {'key': QBV_PROJECT_KEY},
                    'summary': project_name,
                    'description': description,
                    'issuetype': {'id': '10222'},  # QBV issue type ID
                    DEST_TEAM_FIELD_ID: {'value': 'IT_DevOps Team'},
                    YEAR_FIELD_ID: year_value_id,
                    QUARTER_FIELD_ID: quarter_value_id,
                    QBV_GROUP_FIELD_ID: qbv_group_value_id,  # CSI group
                    DOD_NEW_FIELD_ID: dod_content,  # Only use DOD NEW field
                    'labels': ['excel2jira']
                }
                
                # Handle user assignment for QBV
                if lead:
                    user_account_id = find_jira_user(lead)
                    if user_account_id:
                        qbv_fields['assignee'] = {'accountId': user_account_id}
                    else:
                        logger.warning(f"Could not find Jira user for lead '{lead}' at row {idx + 2}")
                        failed_users.append(f"Row {idx + 2}: {project_name} - Could not find Jira user for lead '{lead}'")
                
                issue = jira.create_issue(fields=qbv_fields)
                logger.info(f"QBV issue created: {issue.key}")
                created_issues.append(f"QBV: {project_name} ({issue.key})")

            elif issue_type and issue_type.lower() == 'project':
                existing_issues = jira.search_issues(f'project="{DEFAULT_PROJECT_KEY}" AND summary~"{project_name}"')
                if existing_issues:
                    issue = existing_issues[0]
                    # Check existing fields and preserve them
                    current_fields = issue.fields
                    fields_to_update = {}
                    for field, value in epic_fields.items():
                        if hasattr(current_fields, field):
                            current_value = getattr(current_fields, field)
                            if current_value and str(current_value).strip():
                                logger.warning(f"Field {field} already contains data in issue {issue.key}: {current_value}")
                            else:
                                fields_to_update[field] = value
                        else:
                            fields_to_update[field] = value
                    
                    if fields_to_update:
                        issue.update(fields=fields_to_update)
                        logger.info(f"Project '{project_name}' updated: {issue.key}")
                        created_issues.append(f"Project Update: {project_name} ({issue.key})")
                    else:
                        logger.info(f"No fields to update for project '{project_name}'")
                        skipped_issues.append(f"Project Update: {project_name} (no fields to update)")
                else:
                    logger.warning(f"Project '{project_name}' not found, cannot update.")
                    failed_tasks.append(f"Row {idx + 2}: {project_name} - Project not found, cannot update.")

            elif issue_type and issue_type.lower() == 'on-going':
                # Check if ongoing issue already exists
                if check_issue_exists(project_name, DEFAULT_PROJECT_KEY):
                    logger.info(f"On-going issue '{project_name}' already exists, skipping creation")
                    skipped_issues.append(f"On-going: {project_name}")
                    continue
                
                # Create a new dictionary for ongoing issues instead of copying from epic_fields
                story_fields = {
                    'project': {'key': DEFAULT_PROJECT_KEY},
                    'summary': project_name,
                    'description': description,
                    'issuetype': {'name': 'Story'},
                    'parent': {'key': 'ITDVPS-976'},
                    DOD_NEW_FIELD_ID: dod_content,  # Only use DOD NEW field
                    DEST_TEAM_FIELD_ID: {'value': 'IT_DevOps Team'},
                    'labels': ['excel2jira']
                }
                
                # Handle user assignment
                if lead:
                    user_account_id = find_jira_user(lead)
                    if user_account_id:
                        story_fields['assignee'] = {'accountId': user_account_id}
                    else:
                        logger.warning(f"Could not find Jira user for lead '{lead}' at row {idx + 2}")
                        failed_users.append(f"Row {idx + 2}: {project_name} - Could not find Jira user for lead '{lead}'")
                
                issue = jira.create_issue(fields=story_fields)
                logger.info(f"On-Going story created under ITDVPS-976: {issue.key}")
                created_issues.append(f"On-going: {project_name} ({issue.key})")

            else:
                logger.warning(f"Unknown issue type '{issue_type}' at row {idx + 2}")
                failed_tasks.append(f"Row {idx + 2}: {project_name} - Unknown issue type '{issue_type}'")

        except Exception as e:
            logger.error(f"Failed at row {idx + 2} ('{project_name}'): {e}")
            failed_tasks.append(f"Row {idx + 2}: {project_name} - {str(e)}")

    # Add this before the final "Script execution completed" message
    logger.info("\n=== Execution Summary ===")
    if created_issues:
        logger.info("\nCreated/Updated Issues:")
        for issue in created_issues:
            logger.info(f"- {issue}")
    if skipped_issues:
        logger.info("\nSkipped Issues (already exist):")
        for issue in skipped_issues:
            logger.info(f"- {issue}")
    if failed_tasks:
        logger.warning("\nFailed Tasks:")
        for task in failed_tasks:
            logger.warning(f"- {task}")
    if failed_users:
        logger.warning("\nFailed User Assignments:")
        for user in failed_users:
            logger.warning(f"- {user}")
    if failed_fields:
        logger.warning("\nFailed Field Updates:")
        for field in failed_fields:
            logger.warning(f"- {field}")

    if not any([failed_tasks, failed_users, failed_fields]):
        logger.info("All tasks completed successfully!")

except Exception as e:
    logger.error(f"Script failed: {e}")
    exit(1)

logger.info("Script execution completed.")
