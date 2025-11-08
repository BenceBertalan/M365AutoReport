# M365AutoReport

Automated Microsoft 365 usage reporting tool that generates Excel reports for SharePoint sites and Exchange mailboxes.

## Features

- **SharePoint Site Usage**: Reports site name, URL, size in GB, members, and owners
- **Exchange Mailbox Usage**: Reports email address, name, usage, quota, and remaining quota percentage
- **Excel Export**: Generates formatted Excel reports with separate worksheets
- **Service Principal Authentication**: Secure authentication using Azure AD App Registration
- **Containerized**: Fully containerized with Docker for easy deployment

## Prerequisites

1. **Azure AD App Registration** (Service Principal)
   - Register an application in Azure AD
   - Create a client secret
   - Grant the following Microsoft Graph API permissions (Application permissions):
     - `Sites.Read.All` - For SharePoint site data
     - `User.Read.All` - For user and mailbox data
     - `Mail.Read` - For mailbox information (optional)
     - `Reports.Read.All` - For usage reports (optional)
   - Grant admin consent for the permissions

2. **Python 3.11+** (if running locally without Docker)

3. **Docker** (if running in container)

## Setup

### 1. Clone the Repository

```bash
git clone https://github.com/BenceBertalan/M365AutoReport.git
cd M365AutoReport
```

### 2. Configure Environment Variables

Copy the example environment file and fill in your credentials:

```bash
cp .env.example .env
```

Edit `.env` with your Azure AD application details:

```env
TENANT_ID=your-tenant-id-here
CLIENT_ID=your-client-id-here
CLIENT_SECRET=your-client-secret-here
OUTPUT_FILE=M365_Usage_Report.xlsx
```

## Usage

### Running with Docker (Recommended)

1. **Build the Docker image:**

```bash
docker build -t m365-auto-report .
```

2. **Run the container:**

```bash
docker run --rm \
  --env-file .env \
  -v $(pwd):/output \
  m365-auto-report
```

The report will be saved to `M365_Usage_Report.xlsx` in the current directory.

### Running Locally with Python

1. **Install dependencies:**

```bash
pip install -r requirements.txt
```

2. **Run the script:**

```bash
python m365_report.py
```

## Output

The application generates an Excel file (`M365_Usage_Report.xlsx` by default) with two worksheets:

### SharePoint Sites Worksheet
- Site Name
- Site URL
- Size (GB)
- Members (semicolon-separated)
- Owners (semicolon-separated)

### Exchange Mailboxes Worksheet
- Email Address
- Name
- Usage (GB)
- Quota (GB)
- Remaining (%)

## Permissions Required

The Service Principal needs the following Microsoft Graph API permissions:

| Permission | Type | Description |
|------------|------|-------------|
| Sites.Read.All | Application | Read SharePoint site collections |
| User.Read.All | Application | Read user profiles and mailbox info |
| Reports.Read.All | Application | Read usage reports (optional) |

## Troubleshooting

### Authentication Errors

- Verify your `TENANT_ID`, `CLIENT_ID`, and `CLIENT_SECRET` are correct
- Ensure the Service Principal has been granted admin consent for the required permissions
- Check that the client secret has not expired

### Missing Data

- Some data may not be available depending on your Microsoft 365 license and configuration
- Ensure the Service Principal has sufficient permissions
- Some sites or mailboxes may have restricted access

### Docker Issues

- Ensure Docker is installed and running
- Check that the `.env` file is in the same directory as the Dockerfile
- Verify volume mounting is working correctly

## Security Notes

- Never commit the `.env` file containing real credentials to version control
- Rotate client secrets regularly
- Follow the principle of least privilege when granting API permissions
- Store credentials securely in production environments (e.g., Azure Key Vault)

## License

This project is provided as-is for usage reporting purposes.

## Contributing

Contributions are welcome! Please feel free to submit issues or pull requests.
