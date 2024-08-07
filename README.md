Certainly! Here's a draft of a README for a collection of PowerShell scripts designed to manage Active Directory (AD), Exchange, and various tools:

---

# PowerShell Scripts for AD, Exchange Management, and Tools

## Overview

This repository contains a collection of PowerShell scripts to facilitate the management of Active Directory (AD), Microsoft Exchange, and other administrative tools. These scripts are designed to simplify routine administrative tasks and improve efficiency for system administrators.

## Table of Contents

- [Active Directory Scripts](#active-directory-scripts)
  - [User Management](#user-management)
  - [Group Management](#group-management)
  - [Password Management](#password-management)
- [Exchange Management Scripts](#exchange-management-scripts)
  - [Mailbox Management](#mailbox-management)
  - [Distribution Group Management](#distribution-group-management)
  - [Exchange Reports](#exchange-reports)
- [Tools](#tools)
  - [System Information](#system-information)
  - [Network Utilities](#network-utilities)
  - [Miscellaneous](#miscellaneous)
- [Getting Started](#getting-started)
  - [Prerequisites](#prerequisites)
  - [Installation](#installation)
  - [Usage](#usage)
- [Contributing](#contributing)
- [License](#license)

## Active Directory Scripts

### User Management

- **Add-ADUser.ps1**: Script to add a new user to Active Directory.
- **Remove-ADUser.ps1**: Script to remove an existing user from Active Directory.
- **Get-ADUserInfo.ps1**: Script to retrieve information about a specific AD user.

### Group Management

- **Add-ADGroup.ps1**: Script to create a new AD group.
- **Add-UserToGroup.ps1**: Script to add a user to an AD group.
- **Remove-UserFromGroup.ps1**: Script to remove a user from an AD group.

### Password Management

- **Reset-ADUserPassword.ps1**: Script to reset a user's password in AD.
- **Set-ADUserPasswordExpiration.ps1**: Script to set password expiration policies for AD users.

## Exchange Management Scripts

### Mailbox Management

- **New-Mailbox.ps1**: Script to create a new mailbox in Exchange.
- **Remove-Mailbox.ps1**: Script to remove an existing mailbox.
- **Get-MailboxInfo.ps1**: Script to retrieve information about a specific mailbox.

### Distribution Group Management

- **New-DistributionGroup.ps1**: Script to create a new distribution group.
- **Add-MemberToDistributionGroup.ps1**: Script to add a member to a distribution group.
- **Remove-MemberFromDistributionGroup.ps1**: Script to remove a member from a distribution group.

### Exchange Reports

- **Get-MailboxUsageReport.ps1**: Script to generate a mailbox usage report.
- **Get-DistributionGroupReport.ps1**: Script to generate a report on distribution groups.

## Tools

### System Information

- **Get-SystemInfo.ps1**: Script to retrieve detailed system information.
- **Get-InstalledSoftware.ps1**: Script to list all installed software on a system.

### Network Utilities

- **Test-NetworkConnection.ps1**: Script to test network connectivity.
- **Get-NetworkConfiguration.ps1**: Script to retrieve network configuration details.

### Miscellaneous

- **Backup-AD.ps1**: Script to backup Active Directory.
- **Restore-AD.ps1**: Script to restore Active Directory from a backup.

## Getting Started

### Prerequisites

- Windows PowerShell 5.1 or later / PowerShell Core 7.0 or later.
- Appropriate permissions to execute AD and Exchange management commands.
- Exchange Management Shell (for Exchange scripts).

### Installation

Clone the repository to your local machine:

```sh
git clone https://github.com/qingfengliu0/powershell-scripts.git
cd powershell-scripts
```

### Usage

1. Open PowerShell with administrative privileges.
2. Navigate to the directory containing the scripts.
3. Execute the desired script:

```sh
.\Add-ADUser.ps1 -Username "jdoe" -Password "P@ssw0rd" -OU "Users"
```

Refer to individual script comments and parameters for detailed usage instructions.

## Contributing

Contributions are welcome! Please submit a pull request or open an issue to discuss improvements or additional scripts.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

Feel free to customize this template according to your specific needs and the scripts you have in your repository.
