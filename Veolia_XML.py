# This program transfers all inspection data under the Veolia contract (extracted from WinCan VX) into an XML format, which is then transferred to Sydney Water for review
import pandas as pd
import xml.etree.ElementTree as ET
import os
import shutil
from xml.dom import minidom
from datetime import datetime

def handle_ASSETandCWONumbers(value):
    value = str(value)
    parts = value.split(',')
    converted_parts = []
    for part in parts:
        part = part.strip()
        try:
            converted_parts.append(str(int(float(part))))
        except ValueError:
            converted_parts.append(part)
    return ', '.join(converted_parts)

def process_excel_to_xml(file_path, base_folder):
    df = pd.read_excel(file_path, sheet_name=0)
    df = df.fillna('')
    root = ET.Element("mrss", {"xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance"})
    channel = ET.SubElement(root, "channel")
    
    for index, row in df.iterrows():
        if row.iloc[0] == '':
            break

        item = ET.SubElement(channel, "item")
        action = ET.SubElement(item, "action")
        action.text = "add"
        type_element = ET.SubElement(item, "type")
        type_element.text = "1"
        userId = ET.SubElement(item, "userId")
        userId.text = "VEOLIAWATE004"
        name = ET.SubElement(item, "name")
        name.text = f"CCTV report for CWO {handle_ASSETandCWONumbers(str(row.iloc[12]))}"  # Column M
        description = ET.SubElement(item, "description")
        description.text = row.iloc[13]  # Column N
        tags = ET.SubElement(item, "tags")
        tag = ET.SubElement(tags, "tag")
        tag.text = "CCTV"
        categories = ET.SubElement(item, "categories")
        category = ET.SubElement(categories, "category")
        category.text = "MediaSpace>site>galleries>CCTV>VEOLIAWATE-004"
        media = ET.SubElement(item, "media")
        mediaType = ET.SubElement(media, "mediaType")
        mediaType.text = "1"
        contentAssets = ET.SubElement(item, "contentAssets")
        content = ET.SubElement(contentAssets, "content")
        dropFolderFileContentResource = ET.SubElement(content, "dropFolderFileContentResource", filePath=row.iloc[1])  # Column B
        attachments = ET.SubElement(item, "attachments")
        action = ET.SubElement(attachments, "action")
        action.text = "update"
        attachment = ET.SubElement(attachments, "attachment", format="3")
        dropFolderFileContentResource = ET.SubElement(attachment, "dropFolderFileContentResource", filePath=row.iloc[14])  # Column O
        filename = ET.SubElement(attachment, "filename")
        filename.text = f"{handle_ASSETandCWONumbers(str(row.iloc[12]))}.pdf"  # Column M
        title = ET.SubElement(attachment, "title")
        title.text = f"CCTV report for {handle_ASSETandCWONumbers(str(row.iloc[12]))}"  # Column M
        description = ET.SubElement(attachment, "description")
        description.text = f"CCTV report for CWO {handle_ASSETandCWONumbers(str(row.iloc[12]))} PWO {handle_ASSETandCWONumbers(str(row.iloc[11]))}"  # Columns M and L
        customDataItems = ET.SubElement(item, "customDataItems")
        customData = ET.SubElement(customDataItems, "customData", metadataProfileId="187")
        xmlData = ET.SubElement(customData, "xmlData")
        metadata = ET.SubElement(xmlData, "metadata")
        parentWorkOrderNumber = ET.SubElement(metadata, "ParentWorkOrderNumber")
        parentWorkOrderNumber.text = handle_ASSETandCWONumbers(str(row.iloc[11]))  # Column L
        childWorkOrderNumbers = ET.SubElement(metadata, "ChildWorkOrderNumbers")
        childWorkOrderNumbers.text = handle_ASSETandCWONumbers(str(row.iloc[12]))  # Column M
        workOrderDescription = ET.SubElement(metadata, "WorkOrderDescription")
        workOrderDescription.text = str(row.iloc[18])  # Column S
        assetNumbers = ET.SubElement(metadata, "AssetNumbers")
        if row.iloc[17] != '':  # Column R
            assetNumbers.text = ', '.join([handle_ASSETandCWONumbers(val) for val in str(row.iloc[17]).split('_')])  # Column R
        else:
            assetNumbers.text = ', '.join([handle_ASSETandCWONumbers(val) for val in str(row.iloc[10]).split('_')])  # Column K
        taskCode = ET.SubElement(metadata, "TaskCode")
        taskCode.text = str(row.iloc[19])  # Column T
        suburb = ET.SubElement(metadata, "Suburb")
        suburb.text = str(row.iloc[16])  # Column Q
        addressStreet = ET.SubElement(metadata, "AddressStreet")
        addressStreet.text = str(row.iloc[15])  # Column P
        product = ET.SubElement(metadata, "Product")
        product.text = "WasteWater"
        contractor = ET.SubElement(metadata, "Contractor")
        contractor.text = "VEOLIAWATE-004"
        upstreamMH = ET.SubElement(metadata, "UpstreamMH")
        upstreamMH.text = handle_ASSETandCWONumbers(str(row.iloc[2]))  # Column C
        downstreamMH = ET.SubElement(metadata, "DownstreamMH")
        downstreamMH.text = handle_ASSETandCWONumbers(str(row.iloc[3]))  # Column D
        directionOfSurvey = ET.SubElement(metadata, "DirectionOfSurvey")
        directionOfSurvey.text = str(row.iloc[4])  # Column E
        dateOfCompletedInspection = ET.SubElement(metadata, "DateOfCompletedInspection")
        dateOfCompletedInspection.text = str(row.iloc[5])  # Column F
        timeOfCompletedInspection = ET.SubElement(metadata, "TimeOfCompletedInspection")
        timeOfCompletedInspection.text = str(row.iloc[6])  # Column G
        packageName = ET.SubElement(metadata, "PackageName")
        packageName.text = str(row.iloc[7])  # Column H
        cleaned = ET.SubElement(metadata, "Cleaned")
        cleaned.text = str(row.iloc[8])  # Column I
        surveyedLength = ET.SubElement(metadata, "SurveyedLength")
        surveyedLength.text = str(row.iloc[9])  # Column J
    
    tree = ET.ElementTree(root)
    xml_str = ET.tostring(root, encoding='utf-8', method='xml')
    parsed_str = minidom.parseString(xml_str)
    pretty_xml_as_str = parsed_str.toprettyxml(indent="\t", newl="\n")
    pretty_xml_as_str = pretty_xml_as_str.replace('<?xml version="1.0" ?>', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    pretty_xml_as_str = pretty_xml_as_str.rstrip("\n")
    
    earliest_date = min(pd.to_datetime(df.iloc[:, 5], dayfirst=True, errors='coerce').dropna()).strftime('%d-%m-%Y')
    column_l_value = handle_ASSETandCWONumbers(str(df.iloc[0, 11]))
    column_q_value = str(df.iloc[0, 16])
    default_filename = f"{column_l_value}_{column_q_value}_{earliest_date}.xml"
    
    # Create a folder named after the XML file (without .xml)
    output_dir = os.path.join(base_folder, default_filename.replace('.xml', ''))
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    output_file_parent = os.path.join(output_dir, default_filename)
    
    with open(output_file_parent, 'w', encoding='utf-8') as f:
        f.write(pretty_xml_as_str)
    
    # Copy the Excel file to the new folder
    shutil.copy(file_path, output_dir)
    
    # Copy all folders in the same directory as the Excel file to the new folder
    excel_folder = os.path.dirname(file_path)
    for item in os.listdir(excel_folder):
        item_path = os.path.join(excel_folder, item)
        if os.path.isdir(item_path):
            dest_path = os.path.join(output_dir, item)
            shutil.copytree(item_path, dest_path)
    
    print(f"'{default_filename}' generated and saved successfully.\n")

def traverse_and_process():
    upload_pending_folder = "F:\\VEOLIA UPLOADS\\READY"
    uploaded_folder = "F:\\VEOLIA UPLOADS\\UPLOADED\\" + datetime.now().strftime("%B").upper() + ' ' + datetime.now().strftime("%Y")

    for root, dirs, files in os.walk(upload_pending_folder):
        for subdir in dirs:
            misc_path = os.path.join(root, subdir, "misc")
            docu_path = os.path.join(misc_path, "docu")
            if os.path.exists(docu_path):
                for docu_root, docu_dirs, docu_files in os.walk(docu_path):
                    for file in docu_files:
                        if file.endswith(".xlsx") or file.endswith(".xls"):
                            file_path = os.path.join(docu_root, file)
                            base_folder = os.path.abspath(os.path.join(root, subdir))
                            process_excel_to_xml(file_path, base_folder)
                            folder_to_move = base_folder
                            destination_folder = uploaded_folder

                            if not os.path.exists(destination_folder):
                                os.makedirs(destination_folder)

                            shutil.move(folder_to_move, destination_folder)
                            print(f"Moved '{folder_to_move}' to '{destination_folder}'.")
                            break

if __name__ == "__main__":
    traverse_and_process()
