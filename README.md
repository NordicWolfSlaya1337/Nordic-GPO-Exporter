# Nordic's GPO Exporter

![PowerShell](https://img.shields.io/badge/PowerShell-5.0%2B-5391FE?logo=powershell&logoColor=white)
![Version](https://img.shields.io/badge/Version-1.0-blue)
![License](https://img.shields.io/badge/License-Proprietary-red)

> PowerShell tool to export all Group Policy Objects from an Active Directory domain into a single organized ZIP archive.

---

## Features

| Feature | Description |
|---------|-------------|
| **XML Reports** | Exports every GPO as an individual XML report |
| **Human-Readable Names** | XML files are named by GPO display name, not GUID |
| **OU Link Mapping** | CSV report detailing which OUs each GPO is linked to |
| **Status Tracking** | CSV includes whether each GPO is enabled or disabled |
| **Full Coverage** | Unlinked GPOs are still captured in the CSV for complete visibility |
| **Single Archive** | Everything bundled into one ZIP: `{domain}_GPOs_DD-MM-YYYY.zip` |

---

## Requirements

| Requirement | Details |
|-------------|---------|
| **PowerShell** | Version 5.0 or later |
| **RSAT** | Group Policy Management Tools (`GroupPolicy` module) |
| **Machine** | Domain-joined Windows workstation or server |
| **Permissions** | Read access to Group Policy Objects in Active Directory |

---

## Usage

Run from the script's directory on a domain-joined machine:

```powershell
.\Export-GPOs.ps1
```

The ZIP archive is created in the same folder as the script.

---

## Output

### ZIP Archive

**Filename:** `{domain}_GPOs_DD-MM-YYYY.zip`

| File | Description |
|------|-------------|
| `*.xml` | One XML report per GPO, named by display name |
| `GPO_OU_Links.csv` | GPO-to-OU link mapping with enable/disable status |

### CSV Columns

| Column | Description |
|--------|-------------|
| `GPOName` | Display name of the Group Policy Object |
| `GPOStatus` | Whether the GPO is enabled or disabled (see below) |
| `LinkedOU` | Distinguished path of the linked OU (empty if unlinked) |
| `LinkEnabled` | Whether the link itself is enabled |

<details>
<summary><b>GPO Status Values</b></summary>

<br>

| Value | Meaning |
|-------|---------|
| `AllSettingsEnabled` | Both computer and user settings are active |
| `AllSettingsDisabled` | GPO is fully disabled |
| `UserSettingsDisabled` | Only computer settings are active |
| `ComputerSettingsDisabled` | Only user settings are active |

</details>

---

## How It Works

```
  Domain Controller
        |
        v
  Get-GPO -All              Enumerate every GPO in the domain
        |
        v
  Get-GPOReport -XML        Export each GPO as an XML report
        |
        v
  Parse XML Links           Extract OU link paths and status
        |
        v
  GPO_OU_Links.csv          Build link mapping spreadsheet
        |
        v
  Compress-Archive           Bundle everything into a single ZIP
        |
        v
  {domain}_GPOs_DD-MM-YYYY.zip
```

---

## Author

**NordicWolfSlaya1337**

- GitHub: [@NordicWolfSlaya1337](https://github.com/NordicWolfSlaya1337)

---

## License

This software is proprietary and provided under a custom restrictive license. See the [LICENSE](LICENSE) file for full terms.

- Non-commercial / non-profit use only
- No modification or derivative works permitted
- No redistribution or reuse in other projects
- All rights reserved by NordicWolfSlaya1337
- Violators are subject to legal action
