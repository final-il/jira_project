#!/usr/bin/env python3
"""
Jira Project Bootstrapping Script for Azure AI Integration Project

This script reads the Jira API token from a file (../../values.env), 
authenticates to Jira Cloud, creates a project (if it doesn’t exist), 
and then populates it with epics and detailed stories based on the project plan.

Requirements:
- Python 3.x
- jira module (install via: pip install jira)
- A file named 'values.env' located two directories up with the structure:
      JIRA_API_TOKEN = "ATATT3xFfGF04hG7JMG5s5XXXXXXX"
- Your account must have Jira administrator permissions to create projects.
"""

import os
import sys
from jira import JIRA

def read_jira_api_token():
    """
    Reads the JIRA_API_TOKEN from a file located two directories up in a file named values.env.
    Expected file content example:
        JIRA_API_TOKEN = "ATATT3xFfGF04hG7JMG5s5XXXXXXX"
    """
    file_path = os.path.join(os.path.dirname(__file__), "..", "..", "values.env")
    token = None
    try:
        with open(file_path, "r") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#"):
                    continue
                if line.startswith("JIRA_API_TOKEN"):
                    parts = line.split("=", 1)
                    if len(parts) == 2:
                        token = parts[1].strip().strip('"')
                        break
    except Exception as e:
        print(f"Error reading API token from {file_path}: {e}")
        sys.exit(1)
    if not token:
        print(f"JIRA_API_TOKEN not found in {file_path}")
        sys.exit(1)
    return token

# Jira connection details
JIRA_URL = "https://jira-final-il.atlassian.net/"  # or https://<your-jira-domain>/jira if self-managed
JIRA_EMAIL = "maorb@final.co.il"  # Or username for on-prem Jira
JIRA_API_TOKEN = read_jira_api_token()   # Read token from file

# Set up Jira connection
options = {'server': JIRA_URL}
jira = JIRA(options, basic_auth=(JIRA_EMAIL, JIRA_API_TOKEN))

def check_authentication():
    """
    Verifies that the current credentials can access the Jira API v3 endpoint.
    """
    url = JIRA_URL.rstrip("/") + "/rest/api/3/myself"
    response = jira._session.get(url)
    if response.status_code != 200:
        print("Authentication failed. Please verify your credentials in values.env.")
        print(f"Status Code: {response.status_code}, Response: {response.text}")
        sys.exit(1)
    return response.json()

# Check authentication before proceeding
me = check_authentication()
print(f"Authenticated as: {me.get('displayName', 'Unknown User')}")

# Define project details
project_key = 'AZOAI'
project_name = 'Azure OpenAI'
project_description = ("Project for integrating Azure AI with on-prem connectivity, "
                       "Terraform automation, and Jenkins integration.")

def create_project():
    """
    Creates a Jira project using the REST API directly.
    For Jira Cloud, it requires the leadAccountId from the authenticated user.
    """
    try:
        lead_account_id = me['accountId']
    except Exception as e:
        print("Failed to retrieve your account details: ", e)
        sys.exit(1)

    payload = {
        "key": project_key,
        "name": project_name,
        "projectTypeKey": "software",
        "projectTemplateKey": "com.pyxis.greenhopper.jira:gh-simplified-agility-scrum",
        "description": project_description,
        "leadAccountId": lead_account_id,
        "assigneeType": "PROJECT_LEAD"
    }
    url = JIRA_URL.rstrip("/") + "/rest/api/3/project"
    response = jira._session.post(url, json=payload)
    if response.status_code != 201:
        print("Failed to create project: " + response.text)
        sys.exit(1)
    else:
        print("Project created successfully.")
        return response.json()

# Check if the project exists; if not, create it.
try:
    project = jira.project(project_key)
    print(f"Project '{project_key}' found: {project.name}")
except Exception as e:
    print(f"Project '{project_key}' not found. Creating it...")
    project = create_project()

# Define epics with their detailed descriptions
epics = [
    {
        "summary": "Network and Private Link Setup",
        "description": ("Provision the Azure Virtual Network, configure private endpoints, establish on-prem connectivity, "
                        "and deploy firewall and NSGs.")
    },
    {
        "summary": "AI Services Deployment",
        "description": ("Deploy Azure OpenAI and Cognitive Services, integrate with API Management, and enable secure connectivity "
                        "with on-prem systems.")
    },
    {
        "summary": "Security and Compliance",
        "description": ("Implement Azure AD with SSO and RBAC, deploy comprehensive monitoring/logging, and conduct security audits "
                        "and compliance checks.")
    },
    {
        "summary": "Terraform Deployment Automation",
        "description": ("Develop, test, and secure Terraform scripts for provisioning the environment, and integrate CI/CD for continuous improvement.")
    },
    {
        "summary": "Jenkins Automation Integration",
        "description": ("Set up a Jenkins server, integrate it with Azure for automation, develop CI/CD pipelines, and configure monitoring/alerts.")
    }
]

created_epics = {}
for epic in epics:
    issue_dict = {
        'project': {'key': project_key},
        'summary': epic["summary"],
        'description': epic["description"],
        'issuetype': {'name': "Epic"}
    }
    new_issue = jira.create_issue(fields=issue_dict)
    created_epics[epic["summary"]] = new_issue.key
    print(f"Created epic: {new_issue.key} - {epic['summary']}")

# Use your provided field id for epic linking
EPIC_LINK_FIELD = 'customfield_10011'

# Define detailed tasks under each epic
detailed_tasks = [
    # Epic: Network and Private Link Setup
    {"epic": "Network and Private Link Setup", "summary": "Define IP Address Space and Subnets", 
     "description": "Determine IP ranges, create subnets for AI services, management, and connectivity."},
    {"epic": "Network and Private Link Setup", "summary": "Create Azure Virtual Network", 
     "description": "Provision the VNet with defined subnets and configure routing tables."},
    {"epic": "Network and Private Link Setup", "summary": "Deploy Private Endpoints", 
     "description": "Deploy private endpoints for target Azure AI services and map to Private DNS zones."},
    {"epic": "Network and Private Link Setup", "summary": "Test DNS Resolution", 
     "description": "Verify that private DNS zones resolve to the correct private IPs."},
    {"epic": "Network and Private Link Setup", "summary": "Evaluate On-Prem Connectivity Options", 
     "description": "Assess ExpressRoute versus VPN Gateway based on requirements."},
    {"epic": "Network and Private Link Setup", "summary": "Configure ExpressRoute/VPN Gateway", 
     "description": ("If using ExpressRoute, provision the circuit and configure BGP; "
                     "if VPN, deploy gateway and configure IPsec/IKE settings.")},
    {"epic": "Network and Private Link Setup", "summary": "Deploy Azure Firewall & NSGs", 
     "description": "Create firewall rules and NSG policies, and configure logging and monitoring."},
    
    # Epic: AI Services Deployment
    {"epic": "AI Services Deployment", "summary": "Provision Azure OpenAI Service", 
     "description": "Deploy the Azure OpenAI resource with a private endpoint for secure access."},
    {"epic": "AI Services Deployment", "summary": "Integrate with API Management", 
     "description": "Configure Azure API Management to control access to the OpenAI service."},
    {"epic": "AI Services Deployment", "summary": "Deploy Cognitive Services", 
     "description": "Provision Cognitive Services with necessary API keys and security configurations."},
    {"epic": "AI Services Deployment", "summary": "Configure Internal DNS", 
     "description": "Set up DNS resolution to ensure on-prem systems can access private endpoints."},
    {"epic": "AI Services Deployment", "summary": "Conduct End-to-End Connectivity Tests", 
     "description": "Validate connectivity from internal applications to deployed AI services."},
    
    # Epic: Security and Compliance
    {"epic": "Security and Compliance", "summary": "Implement Azure AD and SSO", 
     "description": "Configure Azure AD tenant, enable SSO, and ensure secure authentication."},
    {"epic": "Security and Compliance", "summary": "Define RBAC Policies", 
     "description": "Implement role-based access control for all Azure resources."},
    {"epic": "Security and Compliance", "summary": "Set Up Azure Monitor", 
     "description": "Deploy Azure Monitor and Log Analytics for comprehensive resource monitoring."},
    {"epic": "Security and Compliance", "summary": "Integrate Microsoft Defender for Cloud", 
     "description": "Configure threat detection and security alerts using Defender for Cloud."},
    {"epic": "Security and Compliance", "summary": "Schedule Vulnerability Scans", 
     "description": "Set up regular security audits and compliance reviews."},
    
    # Epic: Terraform Deployment Automation
    {"epic": "Terraform Deployment Automation", "summary": "Develop Modular Terraform Scripts", 
     "description": "Write Terraform scripts to provision VNet, private endpoints, and AI services; modularize for reuse."},
    {"epic": "Terraform Deployment Automation", "summary": "Version Control and Testing", 
     "description": "Commit code to Git and run terraform plan/apply in a staging environment."},
    {"epic": "Terraform Deployment Automation", "summary": "Configure Remote State Backend", 
     "description": "Set up Azure Storage for secure, remote Terraform state management."},
    {"epic": "Terraform Deployment Automation", "summary": "Integrate CI/CD for Terraform", 
     "description": "Set up automated linting, validation, and testing pipelines for Terraform code."},
    
    # Epic: Jenkins Automation Integration
    {"epic": "Jenkins Automation Integration", "summary": "Provision Jenkins Server", 
     "description": "Deploy and configure the Jenkins server with essential plugins and backup policies."},
    {"epic": "Jenkins Automation Integration", "summary": "Install Azure CLI and Plugins", 
     "description": "Ensure Jenkins has the Azure CLI and necessary plugins installed."},
    {"epic": "Jenkins Automation Integration", "summary": "Configure Service Connections", 
     "description": "Set up authentication and credentials for Jenkins to interact with Azure and Terraform."},
    {"epic": "Jenkins Automation Integration", "summary": "Develop CI/CD Pipeline Scripts", 
     "description": "Create Jenkins pipelines for automating Terraform deployments with testing and security scans."},
    {"epic": "Jenkins Automation Integration", "summary": "Set Up Monitoring and Notifications", 
     "description": "Configure job monitoring and integrate alerts with tools like Slack or Teams."},
]

for task in detailed_tasks:
    # Create the issue without the epic link field
    issue_fields = {
        'project': {'key': project_key},
        'summary': task["summary"],
        'description': task["description"],
        'issuetype': {'name': "Story"}
    }
    new_issue = jira.create_issue(fields=issue_fields)
    print(f"Created story: {new_issue.key} - {task['summary']}")
    # Now update the issue to set the epic link using your custom field id
    epic_key = created_epics.get(task["epic"])
    if epic_key:
        try:
            new_issue.update(fields={EPIC_LINK_FIELD: epic_key})
            print(f"Updated {new_issue.key} with epic link {epic_key} for epic '{task['epic']}'")
        except Exception as ex:
            print(f"Failed to update epic link for {new_issue.key}: {ex}")
