AllocationIQ - AI-Powered Project Allocation System
Overview
AllocationIQ is a modern, AI-enhanced SharePoint Framework (SPFx) solution designed to streamline project allocation and resource management within Microsoft 365 environments. Built on the Heft build system with SPFx 1.22.2, React, and PnPjs, it leverages Azure AI Foundry agents for intelligent project planning and staff recommendations.

Technology Stack
Component	Technology
Framework	SharePoint Framework (SPFx) 1.22.2
Build System	Heft (Rush Stack)
UI Library	React 17 + FluentUI
Data Access	PnPjs (SP & Graph)
Styling	SCSS + PrimeReact Components
AI Integration	Azure AI Foundry Agents
Target Node.js	22.22.0
Key Features
AI-Powered Project Intake: Upload SOWs or project briefs for automatic field extraction via Azure AI agents
Intelligent Staff Allocation: AI-driven team composition recommendations based on skills and availability
Project Plan Generation: Automated timeline and task planning using PlanningAgent
Approval Workflow: Manager approval tracking for project submissions
Employee Directory: Centralized employee data management
Real-time Dashboard: Project metrics, budget tracking, and team visualization
Prerequisites
Node.js 22.22.0 (required)
Microsoft 365 tenant with SharePoint Online
Azure AI Foundry subscription (for AI features)
SharePoint Lists: Project, ProjectParticipation, TaskDetails, Employees
Getting Started
Installation
# Clone the repository
git clone <repository-url>
cd AllocationIQ

# Install Heft globally (if not already)
npm install -g @rushstack/heft

# Install dependencies
npm install
Development
# Start local development server
heft start
The solution will be available at https://localhost:4321/workbench

Build for Production
# Build and package the solution
heft build
Output: solution/project-allocation.sppkg

Deployment
Upload the .sppkg file to the SharePoint App Catalog
Add the web part to a SharePoint page
Configure required SharePoint lists
Project Structure
AllocationIQ/
├── src/
│   ├── webparts/
│   │   └── projectAllocation/
│   │       ├── components/
│   │       │   ├── AI/              # AI agent integrations
│   │       │   ├── Approvals/       # Approval workflow
│   │       │   ├── Dashboard/       # Main dashboard
│   │       │   ├── Employee/        # Employee directory
│   │       │   ├── Form/            # Project intake form
│   │       │   ├── NavBar/          # Navigation
│   │       │   ├── ProjectDetails/  # Project details view
│   │       │   └── UserContext/     # User context provider
│   │       ├── loc/                # Localization
│   │       └── assets/              # Static assets
│   └── service/
│       └── initservice.ts           # PnPjs initialization
├── config/
│   ├── config.json                  # SPFx configuration
│   ├── package-solution.json        # Package metadata
│   └── serve.json                   # Local serve settings
├── package.json                     # Dependencies
└── README.md                        # This file
Configuration
SharePoint Lists
The solution requires the following SharePoint lists:

List Name	Purpose
Project	Main project registry
ProjectParticipation	Team member assignments
TaskDetails	Task and milestone tracking
Employees	Employee directory
List Columns
Project List: - ProjectCode, ProjectName, Client, StartDate, EndDate - Budget, Priority, RequiredSkills, Description, isManagerApproved

TaskDetails List: - Task, Sno, ProjectCode, Assignee, lane, Effort - Status, StartDate, EndDate

Architecture
┌─────────────────────────────────────────────────────────────┐
│                    AllocationIQ Web Part                     │
├─────────────────────────────────────────────────────────────┤
│  UI Layer (React + FluentUI + PrimeReact)                  │
├─────────────────────────────────────────────────────────────┤
│  Component Layer                                            │
│  ├── Dashboard    │ Approvals  │ Form    │ Employee       │
│  └── ProjectDetails │ AI Agents │ NavBar                      │
├─────────────────────────────────────────────────────────────┤
│  Service Layer (PnPjs)                                       │
│  └── SharePoint Online / Microsoft Graph                    │
├─────────────────────────────────────────────────────────────┤
│  External Integrations                                      │
│  └── Azure AI Foundry (ProjectAnalysis, TeamAllocation,    │
│      PlanningAgent)                                          │
└─────────────────────────────────────────────────────────────┘
AI Agents
The solution integrates three Azure AI Foundry agents:

ProjectAnalysisAgent (v3): Analyzes uploaded documents to extract project specifications
TeamAllocation (v24): Recommends suitable team members based on required skills
PlanningAgent (v8): Generates project timelines and task breakdowns
Version History
Version	Date	Changes
1.0.0	June 2026	Initial release with AI integration
References
SharePoint Framework Documentation
Heft Build System
PnPjs Documentation
Microsoft 365 Patterns and Practices
License
THIS CODE IS PROVIDED AS IS WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.

Built with Heft + SPFx 1.22.2 | AllocationIQ