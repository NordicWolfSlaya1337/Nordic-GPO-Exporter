# Nordic's GPO Exporter

> PowerShell tool to export all Group Policy Objects from an Active Directory domain into a single organized zip archive.

---

## Features

- **XML Reports** — Exports every GPO as an individual XML report
- **Human-Readable Names** — XML files are named by GPO display name, not GUID
- **OU Link Mapping** — CSV report detailing which OUs each GPO is linked to
- **Status Tracking** — CSV includes whether each GPO is enabled or disabled
- **Full Coverage** — Unlinked GPOs are still captured in the CSV for complete visibility
- **Single Archive** — Everything bundled into one zip: `{domainname}_GPOs_DD-MM-YYYY.zip`

---

## Requirements

| Requirement | Details |
|---|---|
| **PowerShell** | Version 5.0 or later |
| **RSAT** | Group Policy Management Tools (`GroupPolicy` module) |
| **Machine** | Domain-joined Windows workstation or server |
| **Permissions** | Read access to Group Policy Objects in Active Directory |

---

## Usage

```powershell
.\Export-GPOs.ps1
```

Run from the script's directory. The zip archive is created in the same folder.

---

## Output

### Zip Archive

**Filename:** `{domainname}_GPOs_DD-MM-YYYY.zip`

**Contents:**

| File | Description |
|---|---|
| `*.xml` | One XML report per GPO, named by display name |
| `GPO_OU_Links.csv` | GPO-to-OU link mapping with status |

### CSV Columns

| Column | Description |
|---|---|
| `GPOName` | Display name of the Group Policy Object |
| `GPOStatus` | Whether the GPO is enabled or disabled (see below) |
| `LinkedOU` | Distinguished path of the linked OU (empty if unlinked) |
| `LinkEnabled` | Whether the link itself is enabled |

### GPO Status Values

| Value | Meaning |
|---|---|
| `AllSettingsEnabled` | Both computer and user settings are active |
| `AllSettingsDisabled` | GPO is fully disabled |
| `UserSettingsDisabled` | Only computer settings are active |
| `ComputerSettingsDisabled` | Only user settings are active |

---

## Version

**v1.0** — Initial release
