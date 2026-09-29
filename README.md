# SASPy Studio

**SASPy Studio** is a lightweight desktop application for running local SAS programs against **SAS OnDemand for Academics** using Python and SASPy.

It provides a simple development workflow for SAS programmers who want to work with local `.sas` files, reusable SAS macros, Git/version control, and VS Code while using a remote SAS session for execution.

## Features

- Run local SAS programs through SASPy
- Connect to SAS OnDemand for Academics
- Start a fresh SAS session for each run
- Execute a local `autoexec.sas` before the selected program
- Automatically load reusable SAS macros from the `macros` folder
- Display initialization and execution progress
- Save SAS logs locally
- Save returned HTML results locally
- Check SAS logs for errors, warnings, and selected important notes
- Download generated files from the remote SAS WORK directory
- Support common output formats including CSV, TXT, RTF, PDF, XLS/XLSX, XML, JSON, ZIP, PNG/JPG/SVG, HTML, LST, and XPT
- Keep SAS OnDemand credentials outside source control
- Work naturally with Git and VS Code

## Project Structure

```text
SASPyStudio/
├── SASPyStudio.py
├── sas_config.json
├── _authinfo                 # Private - do not commit
├── _authinfo_example
├── autoexec.sas
├── requirements.txt
├── README.md
├── .gitignore
├── macros/
├── TestPrograms/
├── logs/
├── results/
└── output/
```

The `logs`, `results`, and `output` directories contain runtime-generated files and are excluded from source control.

## Requirements

- Python 3
- SASPy
- Java compatible with the SASPy IOM connection
- A SAS OnDemand for Academics account

Install the Python dependencies with:

```powershell
pip install -r requirements.txt
```

## Configuration

Connection settings are stored in `sas_config.json`.

Example:

```json
{
    "config_name": "oda",
    "region": "AP1",
    "java_path": "C:\\Program Files\\Eclipse Adoptium\\jdk-17.0.20.8-hotspot\\bin\\java.exe",
    "iom_host": "odaws01-apse1.oda.sas.com",
    "iom_port": 8591,
    "authkey": "oda",
    "authinfo": "_authinfo",
    "encoding": "utf-8"
}
```

The example above reflects an Asia Pacific SAS OnDemand configuration. Update the connection settings as required for your own environment.

## Authentication

SAS credentials must **not** be stored in the Python source code or committed to Git.

Create a local file named:

```text
_authinfo
```

beside `SASPyStudio.py`.

Use `_authinfo_example` as the template:

```text
oda user YOUR_SAS_ONDEMAND_USER_ID password YOUR_SAS_ONDEMAND_PASSWORD
```

Replace the placeholders with your own SAS OnDemand credentials.

The real `_authinfo` file is intentionally excluded through `.gitignore`. Never commit or share it.

## Running SASPy Studio

From PowerShell:

```powershell
python SASPyStudio.py
```

You can also open the project directory in VS Code and run `SASPyStudio.py`.

In SASPy Studio:

1. Select a `.sas` program using **Browse**.
2. Choose whether to run `autoexec.sas`.
3. Choose whether to load the `macros` folder.
4. Choose whether to download generated WORK outputs.
5. Select **Run SAS**.

## Execution Model

For each run, SASPy Studio:

1. Starts a fresh SAS OnDemand session.
2. Submits `autoexec.sas`, if enabled.
3. Submits each `macros/*.sas` file separately, if enabled.
4. Executes the selected SAS program.
5. Checks the combined SAS log.
6. Downloads supported files generated in SAS WORK, if enabled.
7. Saves local log/results.
8. Ends the SAS session.

Submitting initialization files individually keeps the execution model simple while providing visible initialization progress.

## Log Checking

SASPy Studio checks the combined SAS log for:

- `ERROR:`
- `WARNING:`
- Selected important `NOTE:` messages

Important notes currently include conditions such as:

- Uninitialized variables
- Invalid data
- Missing values generated
- Numeric-to-character conversions
- Character-to-numeric conversions
- Division by zero
- Mathematical operation issues
- Insufficient `W.D` format width
- MERGE statements with repeated BY values

A summary is displayed in the application, for example:

```text
Errors: 0   Warnings: 0   Important Notes: 0
```

The complete SAS log is still retained for review.

## Output Files

When **Download WORK outputs** is selected, supported files created in the remote SAS WORK directory are downloaded to:

```text
output/<run_timestamp>/
```

Supported extensions currently include:

```text
.csv .txt .rtf .pdf .xlsx .xls .xml .json .zip
.png .jpg .jpeg .svg .html .htm .lst .xpt
```

SAS datasets (`.sas7bdat`) are not automatically downloaded by this feature.

## Logs and Results

SAS logs are stored under:

```text
logs/
```

HTML results returned by SASPy are stored under:

```text
results/
```

Downloaded WORK output files are stored under:

```text
output/
```

These runtime files are excluded from Git.

## SAS Macros

Reusable SAS macros can be placed in:

```text
macros/
```

When **Load macros folder** is selected, each `.sas` file in this directory is loaded before the selected program runs.

Keeping macros in individual files also makes them easier to maintain and version with Git.

## Test Programs

Example programs are provided under:

```text
TestPrograms/
```

They can be used to verify the SASPy connection, program execution, log collection, and output-download workflow.

## Security Check Before Committing

Verify that the real authentication file is ignored:

```powershell
git check-ignore -v _authinfo
```

Verify that it is not tracked:

```powershell
git ls-files | Select-String "authinfo"
```

Only `_authinfo_example` should be tracked.

Never add passwords, API keys, access tokens, or other credentials to the repository.

## Development Status

SASPy Studio is an independently developed utility intended to simplify SAS programming workflows using SASPy and SAS OnDemand.

Additional features and improvements may be added as the project evolves.

## Version

**1.0.0 — September 2026**

## License

No license has currently been assigned to this project.
