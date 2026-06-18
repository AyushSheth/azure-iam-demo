# MyApp — Microsoft Entra ID Authentication

A web application demonstrating enterprise identity and access management using Microsoft Entra ID (formerly Azure AD).

## Features
- Microsoft Entra ID authentication via MSAL.js
- Role-based access control using Azure Security Groups
- Secure dashboard displaying user identity and role
- Token-based session management
- Protected routes — unauthenticated users redirected to login

## Tech Stack
- HTML, CSS, JavaScript
- MSAL.js (Microsoft Authentication Library)
- Microsoft Entra ID
- Azure Static Web Apps

## Setup
1. Clone the repo
2. Copy `js/config.template.js` to `js/config.js`
3. Fill in your Azure credentials in `js/config.js`
4. Open with Live Server in VS Code

## Azure Configuration
- App Registration with SPA redirect URI
- Security Groups for role assignment
- Group claims configured in token
- API permissions: openid, profile, email, User.Read

## Architecture
Users authenticate via Microsoft Entra ID. On successful login the ID token is parsed to extract group membership which maps to application roles. Unauthenticated users are automatically redirected to the login page.