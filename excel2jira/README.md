# Excel to Jira Issue Creator

A Python script that creates Jira issues from an Excel file. This tool supports creating different types of issues (Epics, QBV, Projects, and On-going stories) with customizable fields.

## Features

- Create multiple Jira issue types from an Excel file
- Support for Epics, QBV issues, Projects, and On-going stories
- Customizable project key and quarter
- Option to update existing issues instead of skipping them
- Comprehensive summary of created, updated, and skipped issues
- User assignment with fuzzy matching
- Detailed logging

## Requirements

- Python 3.6+
- Jira API access
- Required Python packages:
  - pandas
  - python-jira
  - python-dotenv
  - openpyxl

## Installation

1. Clone this repository:
   ```
   git clone https://github.com/yourusername/excel2jira.git
   cd excel2jira
   ```

2. Install the required packages:
   ```
   pip install -r requirements.txt
   ```

3. Create a `.env` file in the project root with your Jira credentials:
   ```
   JIRA_URL=https://your-jira-instance.atlassian.net/
   JIRA_EMAIL=your-email@example.com
   JIRA_API_TOKEN=your-api-token
   ```

## Excel File Format

The Excel file should contain the following columns:

- `Project name`: The name of the project/issue
- `Final DoD`: The final Definition of Done
- `Q2 DoD`: The Q2 Definition of Done
- `Project Manager`: The name of the project manager
- `Issue type`: The type of issue (Epic, QBV, Project, or On-going)
- `Description`: The description of the issue
- `lead`: The lead/assignee of the issue
- `parent`: The parent issue key (for Epics)

## Usage

Run the script with the following command:

```
python excel2jira-v2.py [options]
```

### Options

- `-d, --data PATH`: Path to the Excel file containing issue data (required)
- `-p, --project PROJECT`: Project key (default: ITDVPS)
- `-q, --quarter {Q1,Q2,Q3,Q4}`: Quarter (Q1, Q2, Q3, Q4). If not provided, will use next quarter
- `-u, --update`: Update existing issues instead of skipping them

### Examples

Create issues from a specific Excel file in the default project (ITDVPS) with the next quarter:
```
python excel2jira-v2.py -d path/to/your/excel/file.xlsx
```

Create issues from a specific Excel file in a specific project with a specific quarter:
```
python excel2jira-v2.py -d path/to/your/excel/file.xlsx -p CUSTOM -q Q2
```

Update existing issues from a specific Excel file instead of skipping them:
```
python excel2jira-v2.py -d path/to/your/excel/file.xlsx -u
```

## Issue Types and Fields

### Epics
- Project: ITDVPS (or custom)
- Issue Type: Epic
- Fields: Summary, Description, DOD, Requested By Group, Destination Team, Year, Quarter, In Quarter Plan, Labels

### QBV Issues
- Project: CQ
- Issue Type: QBV
- Fields: Summary, Description, Destination Team, Year, Quarter, Group (CSI), DOD NEW, Labels

### Projects
- Project: ITDVPS (or custom)
- Issue Type: Project
- Fields: Summary, Description, DOD, Requested By Group, Destination Team, Year, Quarter, In Quarter Plan, Labels

### On-going Stories
- Project: ITDVPS (or custom)
- Issue Type: Story
- Parent: ITDVPS-976
- Fields: Summary, Description, DOD NEW, Destination Team, Labels

## Output

The script provides a comprehensive summary at the end of the run:

- Created Issues: Lists all issues that were created
- Updated Issues: Lists all issues that were updated
- Skipped Issues: Lists all issues that were skipped
- Failed Tasks: Lists any tasks that failed
- Failed User Assignments: Lists any user assignments that failed
- Failed Field Updates: Lists any field updates that failed

## Troubleshooting

- **Authentication Issues**: Ensure your Jira credentials in the `.env` file are correct.
- **Permission Issues**: Make sure you have the necessary permissions to create issues in the specified project.
- **Field Value Issues**: If you encounter issues with field values, check the allowed values for each field in your Jira instance.

## License

This project is licensed under the MIT License - see the LICENSE file for details. 