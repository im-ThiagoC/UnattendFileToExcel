# Unattend XML to Excel Converter

This Python script parses a Windows `unattend.xml` file and generates a structured Excel file (`.xlsx`) with three organized sheets:

- **GeneralSettings**: hierarchical settings from each component and configuration pass.
- **SynchronousCommands**: all commands listed under `FirstLogonCommands`, ordered and labeled.
- **Drivers**: all driver paths listed under `DriverPaths`, with their key identifiers and actions.

## 📦 Requirements

- Python 3.7+
- `pandas`
- `openpyxl` or `xlsxwriter`

Install dependencies (if needed):

```bash
pip install pandas xlsxwriter
```

## 🚀 How to Use

1. Place your `main.xml` (the unattend file) in the same directory as the script.
2. Run the script:

```bash
python xml_to_excel.py
```

3. The output file `unattend_configuration.xlsx` will be generated in the same directory.

## 📁 Output

The resulting Excel file includes:

- **GeneralSettings**: e.g. `ImageInstall - OSImage - InstallFrom - MetaData`
- **SynchronousCommands**: e.g. `Order: 10`, `CommandLine: Shutdown.exe /r /t 0`
- **Drivers**: with `Path`, `Action`, and `KeyValue` fields from `PathAndCredentials`

## 🛠 Example Call in Script

You can modify and call the main function like this:

```python
create_excel_from_xml('main.xml', 'unattend_configuration.xlsx')
```

---

This is ideal for documenting or reviewing automated Windows deployment configurations.
