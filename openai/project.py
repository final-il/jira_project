#!/usr/bin/env python3
"""
Jira Project Bootstrapping Script for Azure AI Integration Project

This script creates a Jira project (if permissions allow) and populates it with epics and detailed stories
based on the project plan for integrating Azure AI with on-prem connectivity, Terraform automation, and Jenkins integration.

Requirements:
- Python 3.x
- jira module (install via: pip install jira)
- Update the JIRA_URL, JIRA_EMAIL, and JIRA_API_TOKEN variables as needed.
"""

from jira import JIRA

# Jira connection details
JIRA_URL = "https://jira-final-il.atlassian.net/"  # or https://<your-jira-domain>/jira if self-managed
JIRA_EMAIL = "maorb@final.co.il"  # Or username for on-prem Jira
JIRA_API_TOKEN =    # Your API token

# Set up Jira connection
options = {'server': JIRA_URL}
jira = JIRA(options, basic_auth=(JIRA_EMAIL, JIRA_API_TOKEN))

# Define project details (adjust as needed)
project_key = 'AZAI'
project_name = 'Azure AI Integration Project'
project_type = 'software'
project_template_key = 'com.pyxis.greenhopper.jira:gh-simplified-agility-scrum'  # Adjust if needed

# Function to create project (requires admin privileges)
def create_project():
    try:
        project = jira.create_project(key=project_key, name=project_name,
                                      projectTypeKey=project_type, templateKey=project_template_key)
        print(f"Project '{project_name}' created successfully.")
        return project
    except Exception as e:
        print(f"Project creation failed (possibly due to permissions): {e}")
        print(f"Assuming project '{project_key}' already exists.")
        return None

# Uncomment the following line to create the project if you have permissions
# create_project()

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

# Note: The custom field for linking to an epic may differ; for many Jira Cloud instances it is 'customfield_10008'.
EPIC_LINK_FIELD = 'customfield_10008'

for task in detailed_tasks:
    issue_dict = {
        'project': {'key': project_key},
        'summary': task["summary"],
        'description': task["description"],
        'issuetype': {'name': "Story"},
        EPIC_LINK_FIELD: created_epics[task["epic"]]
    }
    new_issue = jira.create_issue(fields=issue_dict)
    print(f"Created story: {new_issue.key} - {task['summary']} under epic '{task['epic']}'")
