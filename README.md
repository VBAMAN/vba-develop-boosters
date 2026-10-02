# VBA Develop Boosters - INST

**A small tool for installing VBA modules and updating procedures.**

INST is a small VBA development tool designed to make everyday code management a little easier.

When developing with VBA, we often need to import modules from other workbooks, replace existing procedures, or add new code to an existing module.

INST provides three simple operations to help with these tasks.

## Features

### 1. Module Install

Import a standard VBA module into the current workbook from:

* Clipboard
* An existing `.bas` file
* Another Excel workbook

### 2. Procedure Replace

Replace an existing procedure with new code.

The previous implementation is preserved as commented-out code, allowing you to review the old version if needed.

### 3. Procedure Add

Add a new procedure to an existing standard module.

If a procedure with the same name already exists, INST stops instead of overwriting it.

## Project Structure

```text
vba-develop-boosters/
├── vba-develop-boosters.xlsm
├── README.md
├── LICENSE
├── src/
│   ├── INST_01_Config.bas
│   ├── INST_02_Clipboard.bas
│   ├── INST_03_Parser.bas
│   ├── INST_04_ModuleInstall.bas
│   ├── INST_05_Replace.bas
│   ├── INST_06_AddProcedure.bas
│   └── INST_90_Test.bas
└── samples/
    ├── INST_TestBas.bas
    ├── INST_TestSource.xlsm
    └── INST_TestCode.txt
```

* `vba-develop-boosters.xlsm` — The tool workbook.
* `src/` — Exported VBA source modules.
* `samples/` — Sample files and code for testing the installation and procedure operations.

## Samples

The `samples` folder contains the files used by `INST_90_Test`.

```text
samples/
├── INST_TestBas.bas
├── INST_TestSource.xlsm
└── INST_TestCode.txt
```

For testing, copy these sample files to the same directory as
`vba-develop-boosters.xlsm`.

The test procedures in `INST_90_Test` expect the sample files
to be located there.

## Getting Started

1. Download or clone this repository.
2. Open `vba-develop-boosters.xlsm` in Excel.
3. Enable macros if required by your Excel security settings.
4. Open the `samples` folder.
5. Copy the sample files to the same directory as `vba-develop-boosters.xlsm`.

The sample files are:

* `INST_TestBas.bas`
* `INST_TestSource.xlsm`
* `INST_TestCode.txt`

You can then run the test procedures in `INST_90_Test`.


### Option 1: Use the Workbook

1. Download `vba-develop-boosters.xlsm` from this repository.
2. Open the workbook in Excel.
3. Enable macros if required by your Excel security settings.
4. Run the appropriate INST entry point for the operation you want to perform.

### Option 2: Import the Source Modules

1. Create a new macro-enabled Excel workbook (`.xlsm`).
2. Open the Visual Basic Editor.
3. Import the required `.bas` files from `src/`.
4. Compile the VBA project and resolve any errors before use.

INST uses VBA project access to work with modules and procedures. Depending on your Excel security settings, access to the VBA project object model may need to be enabled.

## Code Header Format

INST uses simple metadata headers to identify the target module and procedure.

### Module Install

```vb
Option Explicit

'@ENGINE: VBA Develop Boosters - INST
'@VERSION: 1.0
'@MODULE: INST_10_Modules
```

### Procedure Replace

Add an `@SUB` header to identify the procedure to replace.

```vb
'@ENGINE: VBA Develop Boosters - INST
'@VERSION: 1.0
'@MODULE: INST_10_Modules
'@SUB: Test
```

### Procedure Add

Specify the target module and the new procedure name.

```vb
'@ENGINE: VBA Develop Boosters - INST
'@VERSION: 1.0
'@MODULE: INST_10_Modules
'@SUB: NewTest
```

See `samples/INST_TestCode.txt` for complete examples of all three operations.

## Scope

INST focuses on standard VBA modules.

UserForms and their associated form resources are outside the scope of this version. Keeping application logic in standard modules makes code easier to reuse, maintain, and transfer between projects.

INST is intended as a small development aid, not a replacement for source control or a full code-diff and merge system.

## Compatibility

Originally developed for the following environment:

* Microsoft Excel 2007
* VBA 6.5
* Windows 7
* 32-bit environment

Compatibility with other Excel and Windows versions has not been fully verified.

## License

This project is released under the MIT License. See [LICENSE](LICENSE) for details.

## Notes

Always keep a backup of your workbook before modifying its VBA project.

INST is intended to simplify routine VBA development tasks. Please test its behavior in a separate workbook before using it on important projects.
