# MyApp — Microsoft Entra ID Authentication

A web application demonstrating enterprise identity and access management using Microsoft Entra ID (formerly Azure AD).

## How it works

1. The Global Admin sets up the structure in Microsoft Entra ID — creating users, security groups, and assigning roles
2. A user visits the website and clicks Sign in
3. They authenticate with their Microsoft account through Entra ID — the app never sees their password
4. Entra ID verifies the user's identity
5. Entra ID issues a signed token containing the user's name, email, and group membership, then redirects back to the app
6. The dashboard page checks for a valid token before showing any content
7. The app reads the token and displays content based on the user's role
8. After a period of inactivity, the user receives an idle timeout warning and is automatically signed out if no action is taken

## Features

- Microsoft Entra ID authentication via MSAL.js
- Role-based access control using Azure Security Groups
- Role-specific dashboard content (Sales Admin, Sales Associate, IT Admin, IT Helpdesk)
- Secure dashboard displaying user identity and role
- Token-based session management
- Idle timeout with warning popup and automatic sign-out
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

Users authenticate via Microsoft Entra ID. On successful login, the ID token is parsed to extract group membership, which maps to application roles and determines the dashboard content shown. Unauthenticated users are automatically redirected to the login page. Idle sessions are automatically terminated after a period of inactivity.