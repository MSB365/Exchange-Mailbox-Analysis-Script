# Exchange Mailbox Analysis Script

A comprehensive PowerShell script for analyzing Exchange Server mailboxes using the Exchange Web Services (EWS) API. This tool provides detailed insights into mailbox usage, permissions, and activity patterns for both user and shared mailboxes in on-premise Exchange environments.

## 🚀 Features

### Mailbox Analysis
- **Message Statistics**: Total, read, and unread message counts
- **Activity Tracking**: Last received and sent message timestamps
- **Recent Activity**: Details of the last 5 read messages
- **Permission Auditing**: Complete overview of mailbox access rights

### Permission Analysis
- **FullAccess Permissions**: Who has full access to each mailbox
- **SendAs Permissions**: Users authorized to send as the mailbox
- **SendOnBehalf Permissions**: Users authorized to send on behalf of the mailbox

### Reporting Capabilities
- **HTML Reports**: Beautiful, responsive reports with visual cards for each mailbox
- **CSV Reports**: Structured data export for further analysis and processing
- **Error Handling**: Comprehensive error reporting and logging

### User Experience
- **GUI File Selection**: Windows file dialog for easy CSV file selection
- **Progress Tracking**: Real-time progress updates in the Exchange Management Shell
- **Batch Processing**: Process multiple mailboxes from a single CSV file

## 📋 Prerequisites

### Software Requirements
1. **Exchange Server** (On-Premise)
   - Exchange 2013, 2016, 2019, or newer
   - Exchange Management Shell installed and configured

2. **Exchange Web Services Managed API 2.2**
   - Download from [Microsoft Download Center](https://www.microsoft.com/en-us/download/details.aspx?id=42951)
   - Install on the machine where the script will run

3. **PowerShell**
   - PowerShell 5.1 or newer
   - Must be run from Exchange Management Shell

### Permissions Required
- **Exchange Organization Management** or **View-Only Organization Management**
- **Mailbox Import Export** role (for EWS access)
- **ApplicationImpersonation** rights for accessing other users' mailboxes

## 📁 CSV File Format

The script accepts CSV files with mailbox identifiers. Use any of these column headers:

| Column Header | Description | Example |
|---------------|-------------|---------|
| `EmailAddress` | Primary SMTP address | john.doe@company.com |
| `UserPrincipalName` | User Principal Name | john.doe@company.com |
| `SamAccountName` | SAM Account Name | jdoe |
| `Identity` | Exchange identity | John Doe |
| `Mailbox` | Mailbox identifier | john.doe |

### Sample CSV Content
```csv
EmailAddress
john.doe@company.com
shared.mailbox@company.com
jane.smith@company.com
```

## 🚀 Usage

### Basic Usage
1. Open **Exchange Management Shell** as Administrator
2. Navigate to the script directory
3. Run the script:
   ```powershell
   .\\Exchange-MailboxAnalysis.ps1
   ```
4. Select your CSV file when the file dialog opens
5. Monitor progress in the console
6. Find generated reports in the same directory as your CSV file

### Output Files
The script generates timestamped files:
- `MailboxAnalysis_YYYYMMDD_HHMMSS.html` - Formatted HTML report
- `MailboxAnalysis_YYYYMMDD_HHMMSS.csv` - Raw data export

## 📊 Report Contents

### HTML Report Features
- **Executive Summary**: Overview of analysis results
- **Individual Mailbox Cards**: Detailed information for each mailbox
- **Visual Statistics**: Message counts, dates, and activity metrics
- **Permission Matrix**: Clear display of access rights
- **Recent Messages**: Last 5 read messages with details
- **Error Reporting**: Clear indication of any processing issues

### CSV Report Columns
| Column | Description |
|--------|-------------|
| EmailAddress | Primary email address |
| TotalMessages | Total message count |
| ReadMessages | Number of read messages |
| UnreadMessages | Number of unread messages |
| LastReceivedDate | Timestamp of last received message |
| LastSentDate | Timestamp of last sent message |
| FullAccessPermissions | Users with FullAccess rights |
| SendAsPermissions | Users with SendAs rights |
| SendOnBehalfPermissions | Users with SendOnBehalf rights |
| LastReadMessagesCount | Count of recent read messages |
| Error | Any error messages encountered |

## 🔧 Use Cases

### IT Administration
- **Mailbox Auditing**: Regular audits of mailbox permissions and access
- **Compliance Reporting**: Generate reports for compliance and security reviews
- **Migration Planning**: Assess mailbox usage before migrations
- **Cleanup Operations**: Identify inactive or underutilized mailboxes

### Security Analysis
- **Permission Reviews**: Identify excessive or inappropriate mailbox permissions
- **Access Monitoring**: Track who has access to sensitive mailboxes
- **Shared Mailbox Management**: Monitor shared mailbox usage and permissions

### Operational Insights
- **Usage Analytics**: Understand mailbox usage patterns
- **Activity Monitoring**: Track mailbox activity and engagement
- **Resource Planning**: Plan storage and resource allocation

## ⚠️ Troubleshooting

### Common Issues

**EWS Managed API Not Found**
```
Error: EWS Managed API not found
Solution: Install Exchange Web Services Managed API 2.2 from Microsoft
```

**Permission Denied**
```
Error: Access denied to mailbox
Solution: Ensure account has ApplicationImpersonation rights
```

**Exchange Management Shell Not Loaded**
```
Error: Exchange cmdlets not available
Solution: Run script from Exchange Management Shell, not regular PowerShell
```

### Performance Considerations
- **Large Mailboxes**: Processing may take longer for mailboxes with many messages
- **Network Latency**: EWS calls depend on network connectivity to Exchange server
- **Throttling**: Exchange may throttle EWS requests for large batch operations


### Development Guidelines
1. Follow PowerShell best practices
2. Include error handling for new features
3. Update documentation for any new functionality
4. Test with different Exchange versions when possible


## 🏷️ Version History

### v1.0.0
- Initial release
- Basic mailbox analysis functionality
- HTML and CSV report generation
- Permission auditing
- EWS integration

---

**Note**: This script is designed for on-premise Exchange environments. For Exchange Online, consider using Microsoft Graph API instead of EWS.
```

