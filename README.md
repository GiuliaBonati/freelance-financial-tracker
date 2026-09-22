# Freelance Financial Tracker

A Google Sheets-based financial management system designed for **Italian freelancers with a Partita IVA and INPS Gestione Separata**.

The tracker combines income and expense management, invoice monitoring, tax and contribution reserves, budgeting, forecasting, and automated workflows in a single customizable system.

> **Note:** The tracker was designed around the Italian freelance tax system, but its parameters, categories, and workflows can be adapted to different freelance and tax setups.

## Dashboard

![Financial Tracker Dashboard](screenshots/dashboard.png)

## Key Features

- Income and expense tracking
- Invoice management and payment monitoring
- Tax and INPS contribution reserve calculations
- Fixed and variable cost management
- Monthly budgeting
- Annual financial overview
- Financial forecasting
- Account and reserve monitoring
- Configurable categories, subcategories, accounts, payment methods, services, and channels
- Automated workflows powered by Google Apps Script

## Automation

The tracker includes a custom Google Apps Script automation layer designed to reduce repetitive manual work and maintain consistency across the workbook.

Current automations include:

- **Invoice synchronization:** paid invoices are automatically synchronized with the Income sheet.
- **Fixed cost logging:** recurring costs can be logged to Expenses through a checkbox-based workflow.
- **Duplicate protection:** automated checks help prevent duplicate expense records.
- **Dependent dropdowns:** available subcategories dynamically adapt to the selected expense category.
- **Account updates:** balance changes automatically update the corresponding `Last Updated` date.
- **Weekly finance workflow:** checklist completion dates are recorded automatically and the weekly checklist can be reset for the next accounting cycle.

The source code is available in [`scripts/financial-tracker.gs`](scripts/financial-tracker.gs).

## Invoice Management

![Invoice Management](screenshots/invoices.png)

The invoice workflow provides an overview of issued invoices, payment status, tax reserves, and estimated net income.

Paid invoices can be synchronized automatically with the Income sheet through Google Apps Script.

## Forecasting

![Financial Forecast](screenshots/forecast.png)

The forecasting section supports forward-looking financial planning based on the data and assumptions configured in the tracker.

## Technology

- **Google Sheets** — spreadsheet interface, calculations, dashboards, and financial models
- **Google Apps Script / JavaScript** — workflow automation and data synchronization
- **Git & GitHub** — source control, documentation, and project versioning
- **Markdown** — project documentation

## Demo

The repository includes an exported `.xlsx` version of the tracker populated with fictional demonstration data.

> All financial information shown in the public demo and screenshots is fictional and provided exclusively for demonstration purposes.

The tracker is primarily designed for **Google Sheets**. Some Google Sheets-specific functionality and Google Apps Script automations may not be available in the exported Excel version.

## Setup

See the [Setup Guide](docs/setup.md) for initial configuration instructions.

## Disclaimer

This project is intended as a financial organization, planning, and portfolio tool.

Tax rates, contribution rates, reserve calculations, and other financial parameters should be reviewed and customized according to the user's individual tax situation.

It does not constitute tax, accounting, or financial advice.

## Project Status

**Public Demo — Active Development**

The current repository documents the first public version of the Freelance Financial Tracker. Additional documentation, workflow improvements, and automation features may be introduced in future releases.

## Repository Structure

```text
freelance-financial-tracker/
├── README.md
├── freelance-financial-tracker-demo.xlsx
├── docs/
│   └── setup.md
├── scripts/
│   └── financial-tracker.gs
└── screenshots/
    ├── dashboard.png
    ├── invoices.png
    └── forecast.png

