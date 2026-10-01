# SASPy Studio

**SASPy Studio** is a lightweight desktop application for running local
SAS programs against **SAS OnDemand for Academics** using Python and
SASPy.

It provides a study-oriented development workflow for SAS programmers
who want to maintain SAS programs locally, use Git/version control and
VS Code, organize study code separately from study data, and use a
remote SAS OnDemand session for execution.

## Features

-   Run local SAS programs through SASPy and SAS OnDemand for Academics
-   Start a fresh SAS session for each run
-   Separate the SASPy Studio application, study code repository, study
    data, and reusable global macros
-   Select a **Study Location**, **Programs Repository**, and SAS
    **Program**
-   Automatically derive the SAS programming root from
    `<Programs Repository>/sas`
-   Automatically execute `<Programs Repository>/sas/autoexec.sas`
-   Automatically load study macros from
    `<Programs Repository>/sas/macros`
-   Optionally load reusable global macros from a configurable external
    repository
-   Support a study `program_plan.xlsx`
-   Synchronize required local study inputs to a remote SAS workspace
-   Assign and work with RAW, SDTM, and ADAM libraries
-   Display initialization and execution progress in the SAS Console
-   Check SAS logs for errors, warnings, and selected important notes
-   Download changed permanent SAS datasets back to the local study
    location
-   Download supported output files generated during execution
-   Include a Dataset Viewer for reviewing SAS datasets
-   Keep SAS OnDemand credentials outside source control
-   Work naturally with Git and VS Code

## Recommended Project Architecture

SASPy Studio is designed so that the application itself does not contain
study-specific programs or reusable macro libraries.

Example:

``` text
D:\Work\
├── MyGithub\
│   ├── SASPyStudio\
│   │   ├── SASPyStudio.py
│   │   ├── sas_config.json
│   │   ├── _authinfo
│   │   ├── _authinfo_example
│   │   ├── requirements.txt
│   │   ├── README.md
│   │   └── LICENSE
│   │
│   ├── SASGlobalMacros\
│   │   └── <reusable SAS macro files>
│   │
│   └── TRN001-Code\
│       └── sas\
│           ├── autoexec.sas
│           ├── program_plan.xlsx
│           ├── setup\
│           ├── macros\
│           ├── sdtm\
│           ├── adam\
│           └── tlf\
│
└── Projects\
    └── TRN001\
        └── <study data, datasets, logs, results, and outputs>
```

This separates:

-   **SASPyStudio** --- application and execution engine
-   **SASGlobalMacros** --- reusable SAS macros
-   **Study code repository** --- study-specific SAS programs, autoexec,
    macros, and program plan
-   **Study Location** --- study data and generated runtime content

## Requirements

-   Python 3
-   SASPy
-   pandas
-   openpyxl
-   Java compatible with the SASPy IOM connection
-   A SAS OnDemand for Academics account

Install the Python dependencies with:

``` powershell
pip install -r requirements.txt
```

Current `requirements.txt`:

``` text
saspy>=5.100
pandas>=2.0
openpyxl>=3.1
```

## Configuration

Connection and application defaults are stored in `sas_config.json`.

Example:

``` json
{
    "config_name": "oda",
    "region": "AP1",
    "java_path": "C:\\Program Files\\Eclipse Adoptium\\jdk-17.0.20.8-hotspot\\bin\\java.exe",
    "iom_host": "odaws01-apse1.oda.sas.com",
    "iom_port": 8591,
    "authkey": "oda",
    "authinfo": "_authinfo",
    "encoding": "utf-8",
    "global_macros": "D:\\Work\\MyGithub\\systems-nexus\\SASGlobalMacros"
}
```

Update machine-specific paths and SAS OnDemand connection information as
required for your environment.

`global_macros` defines the default reusable global macro repository
displayed by SASPy Studio. The Global Macros location is optional and
can also be changed from the application.

## Authentication

SAS credentials must **not** be stored in the Python source code or
committed to Git.

Create a local file named:

``` text
_authinfo
```

beside `SASPyStudio.py`.

Use `_authinfo_example` as the template:

``` text
oda user YOUR_SAS_ONDEMAND_USER_ID password YOUR_SAS_ONDEMAND_PASSWORD
```

Replace the placeholders with your own SAS OnDemand credentials.

The real `_authinfo` file is excluded through `.gitignore`. Never commit
or share it.

## Running SASPy Studio

From PowerShell:

``` powershell
python SASPyStudio.py
```

You can also open the SASPyStudio repository in VS Code and run
`SASPyStudio.py`.

### Main Inputs

In SASPy Studio, select:

1.  **Study Location** --- local study data/work area.
2.  **Programs Repository** --- Git repository containing the study SAS
    code.
3.  **Program** --- SAS program to execute.
4.  **Global Macros (Optional)** --- reusable macro repository, if
    required.
5.  **Use Program Plan** --- use the study `program_plan.xlsx` workflow
    when applicable.

The Programs Repository is expected to contain:

``` text
<Programs Repository>\
└── sas\
    ├── autoexec.sas
    ├── program_plan.xlsx
    ├── macros\
    ├── setup\
    ├── sdtm\
    ├── adam\
    └── tlf\
```

`autoexec.sas` is part of the study repository and is executed
automatically. It is not stored in the SASPyStudio application
repository.

## Execution Model

For a normal run, SASPy Studio:

1.  Connects to SAS OnDemand for Academics.
2.  Resolves the remote SAS WORK location.
3.  Creates a temporary remote `SASPyStudio_Study` workspace.
4.  Synchronizes required local study inputs to the remote workspace.
5.  Executes the study `autoexec.sas`.
6.  Assigns the study RAW, SDTM, and ADAM libraries.
7.  Loads configured global macros when available.
8.  Loads study macros from the study repository.
9.  Executes the selected SAS program or program-plan workflow.
10. Identifies changed permanent SAS datasets.
11. Downloads changed datasets and generated outputs to the local study
    area.
12. Completes synchronization and ends the SAS session.

A fresh SAS session is used for each normal run.

## Study Autoexec

The study autoexec is expected at:

``` text
<Programs Repository>\sas\autoexec.sas
```

SASPy Studio uses the study repository selected in the GUI to locate
this file automatically.

This keeps environment setup and study-specific library/configuration
logic with the study code rather than with the SASPyStudio application.

## SAS Macros

### Global Macros

Reusable macros can be maintained in a separate Git repository, for
example:

``` text
D:\Work\MyGithub\systems-nexus\SASGlobalMacros
```

The default location can be specified using `global_macros` in
`sas_config.json`.

Global macros are optional.

### Study Macros

Study-specific macros are maintained under:

``` text
<Programs Repository>\sas\macros
```

SASPy Studio discovers this location from the Programs Repository. A
separate study-macro path does not need to be maintained in the
SASPyStudio repository.

## Program Plan

A study program plan can be maintained at:

``` text
<Programs Repository>\sas\program_plan.xlsx
```

When the **Use Program Plan** option is selected, SASPy Studio uses the
study program plan to drive the applicable execution workflow.

Keeping the plan with the study code allows it to be maintained and
version controlled with the corresponding SAS programs.

## Study Synchronization

SAS OnDemand executes programs in a remote environment, while the source
code and study files can remain on the local Windows computer.

SASPy Studio bridges these environments by creating a temporary remote
study workspace and synchronizing the required files.

Windows paths used by local study programs are adapted for remote
execution where required.

## Permanent SAS Dataset Synchronization

SASPy Studio tracks permanent datasets in the RAW, SDTM, and ADAM
libraries.

Before program execution, it records the existing dataset state. After
execution, it compares dataset metadata such as modification time,
observation count, and variable count.

Only datasets identified as changed need to be downloaded back to the
local study location. This reduces unnecessary data transfer compared
with downloading every permanent SAS dataset after each run.

## Log Checking

SASPy Studio checks the SAS log for:

-   `ERROR:`
-   `WARNING:`
-   Selected important `NOTE:` messages

Important notes can include conditions such as:

-   Uninitialized variables
-   Invalid data
-   Missing values generated
-   Numeric-to-character conversions
-   Character-to-numeric conversions
-   Division by zero
-   Mathematical operation issues
-   Insufficient `W.D` format width
-   MERGE statements with repeated BY values

The complete SAS execution log remains available for review.

## Output Files

SASPy Studio can retrieve generated output files from the remote SAS
environment.

Supported output types include common formats such as:

``` text
.csv .txt .rtf .pdf .xlsx .xls .xml .json .zip
.png .jpg .jpeg .svg .html .htm .lst .xpt
```

Runtime outputs belong with the selected study rather than inside the
SASPyStudio application repository.

## Dataset Viewer

SASPy Studio includes a Dataset Viewer for reviewing SAS datasets from
the application.

The viewer is intended to make it easier to inspect datasets during
development without changing the separation between application code,
study programs, and study data.

## Git and Version Control

The recommended model is to version the different code components
independently:

``` text
SASPyStudio     -> application source
SASGlobalMacros -> reusable/global SAS macros
TRN001-Code     -> study-specific SAS code
```

Study data and generated runtime files should normally remain outside
these source-code repositories.

This structure allows SAS programmers to use Git and VS Code for
development while SAS OnDemand provides the SAS execution environment.

## Security Check Before Committing

Verify that the real authentication file is ignored:

``` powershell
git check-ignore -v _authinfo
```

Verify that it is not tracked:

``` powershell
git ls-files | Select-String "authinfo"
```

Only `_authinfo_example` should be tracked.

Never add passwords, API keys, access tokens, or other credentials to
the repository.

## Development Status

SASPy Studio is an independently developed utility intended to simplify
SAS programming workflows using Python, SASPy, Git, VS Code, and SAS
OnDemand for Academics.

Version 2.0 introduces the study-oriented repository architecture,
external global macro configuration, automatic study autoexec and macro
discovery, remote study synchronization, permanent dataset
synchronization, program-plan support, and the updated desktop
interface.

Planned enhancements can extend the workflow for specification-driven
programming, data import, transport files, and additional
output-generation capabilities.

## Version

**2.0 --- 1 October 2026**

## License

See the `LICENSE` file included in this repository.
