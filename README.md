
```markdown
# Jump Client Fix Script 🚀

This PowerShell script automates the process of connecting to a remote host, downloading and installing the BeyondTrust Jump Client, and performing cleanup tasks. 🛠️

## Disclaimer 🤖 
AI-Assisted Development This script was created with significant assistance from AI technologies. As the creator is not a professional PowerShell or software developer, the code reflects a collaborative approach between human intent and AI-generated solution.

## Features ✨

- Prompts for hostname input if not provided as an argument
- Ensures required directories exist
- Downloads and extracts PSExec if not already present
- Loads environment variables from a `.env` file
- Connects to the remote host using PSExec
- Executes remote commands for uninstallation and installation of Jump Client
- Cleans up local files after execution

## Usage 📋

1. **Run the script:**
   ```powershell
   .\JumpClientFix.ps1 <hostname>
   ```
   If no hostname is provided, the script will prompt for it.

2. **Environment Variables:**
   - Ensure you have a `.env` file in the script directory with the following variables:
     ```
     DOWNLOAD_URL=<URL to the Jump Client installer>
     KEY_SECRET=<Installation key>
     ```

## License 📜

This project is licensed under the GNU General Public License v3.0. See the [LICENSE](LICENSE) file for details.



## GPL-3.0 License 📝

```markdown
GNU GENERAL PUBLIC LICENSE
Version 3, 29 June 2007

```

END OF TERMS AND CONDITIONS
