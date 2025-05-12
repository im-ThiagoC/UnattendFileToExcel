import xml.etree.ElementTree as ET
import pandas as pd

# Helper to remove XML namespaces
def strip_ns(tag):
    return tag.split('}')[-1] if '}' in tag else tag

# Parse XML file
def parse_unattend_xml(xml_path):
    tree = ET.parse(xml_path)
    root = tree.getroot()

    general_rows = []
    sync_rows = []
    driver_rows = []

    for settings in root:
        if strip_ns(settings.tag) != 'settings':
            continue
        pass_name = settings.attrib.get('pass', '')

        for comp in settings:
            if strip_ns(comp.tag) != 'component':
                continue
            comp_name = comp.attrib.get('name', '')

            # Synchronous Commands
            for sc in comp.iter():
                if strip_ns(sc.tag) == 'SynchronousCommand':
                    order = ''
                    cmdline = ''
                    description = ''
                    for sub in sc:
                        if strip_ns(sub.tag) == 'Order':
                            order = sub.text.strip() if sub.text else ''
                        elif strip_ns(sub.tag) == 'CommandLine':
                            cmdline = sub.text.strip() if sub.text else ''
                        elif strip_ns(sub.tag) == 'Description':
                            description = sub.text.strip() if sub.text else ''
                    sync_rows.append({
                        'Component': comp_name,
                        'ConfigurationPass': pass_name,
                        'Order': int(order) if order.isdigit() else order,
                        'CommandLine': cmdline,
                        'Description': description
                    })

            # Driver Paths
            for dp in comp.iter():
                if strip_ns(dp.tag) == 'PathAndCredentials':
                    key_value = next((v for k, v in dp.attrib.items() if k.endswith('keyValue')), '')
                    action = next((v for k, v in dp.attrib.items() if k.endswith('action')), '')
                    path_val = ''
                    for child in dp:
                        if strip_ns(child.tag) == 'Path' and child.text:
                            path_val = child.text.strip()
                    driver_rows.append({
                        'Component': comp_name,
                        'ConfigurationPass': pass_name,
                        'KeyValue': key_value,
                        'Action': action,
                        'Path': path_val
                    })

            # General settings (recursive)
            def recurse(node, path):
                children = [c for c in node if isinstance(c.tag, str)]
                if not children:
                    value = node.text.strip() if node.text else ''
                    general_rows.append({
                        'Component': comp_name,
                        'ConfigurationPass': pass_name,
                        'ComponentSetting': path,
                        'Value': value
                    })
                else:
                    for c in children:
                        recurse(c, f"{path} - {strip_ns(c.tag)}")

            for child in comp:
                if strip_ns(child.tag) not in ['FirstLogonCommands', 'DriverPaths']:
                    recurse(child, strip_ns(child.tag))

    return general_rows, sync_rows, driver_rows

# Create Excel file with sheets
def create_excel_from_xml(xml_path, excel_path):
    general, sync, drivers = parse_unattend_xml(xml_path)

    with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
        pd.DataFrame(general).sort_values(['Component', 'ConfigurationPass', 'ComponentSetting']).to_excel(writer, sheet_name='GeneralSettings', index=False)
        pd.DataFrame(sync).sort_values('Order').to_excel(writer, sheet_name='SynchronousCommands', index=False)
        pd.DataFrame(drivers).sort_values(['Component', 'KeyValue']).to_excel(writer, sheet_name='Drivers', index=False)

# Example usage
create_excel_from_xml('main.xml', 'unattend_configuration.xlsx')
