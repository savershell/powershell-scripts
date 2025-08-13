# PowerShell Scripts Collection

A collection of **PowerShell scripts** created and maintained by [Saverio Sacchetti](https://github.com/savershell).  
These scripts are designed for automation, system administration, and Azure Active Directory management.

---

## 📂 Included Scripts

### `ADD_ADconnect.ps1` – Azure AD Guest User Inviter
A PowerShell 7 script with a Windows Forms GUI to bulk-invite guest users into Azure Active Directory using the Microsoft Graph API.

**Features:**
- Bulk import guest users from CSV
- Duplicate detection
- Invitation sending via Microsoft Graph API
- Simple GUI interface
- Summary report of actions taken

**Usage:**
```powershell
./ADD_ADconnect.ps1
