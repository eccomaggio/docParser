# https://stackoverflow.com/questions/29243/how-do-i-create-an-xml-document-in-python


import os
import zipfile
import xml.dom.minidom
from xml.dom.minidom import Node


def write_xml_file(file_name, xml_output):
    out = file_name
    with open(out, "w", encoding="utf-8") as f:
        f.write(xml_output)


def process_xml(xml):
    root = xml.documentElement
    print(root.tagName)
    tmp = shadow_xml.createElement(root.tagName)
    for attr_name, attr_value in root.attributes.items():
        tmp.setAttribute(attr_name, attr_value)
    shadow_root = shadow_xml.appendChild(tmp)
    print(root.childNodes)

    for node in root.childNodes:
        process_node(root, node, shadow_root)


def process_node(parent, current_node, shadow_parent):
    # tmp = ""
    if current_node.nodeType == Node.TEXT_NODE:
        tmp = shadow_xml.createTextNode(current_node.data)
    else:
        tmp = shadow_xml.createElement(current_node.tagName)
        for attr_name, attr_value in current_node.attributes.items():
            tmp.setAttribute(attr_name, attr_value)
    new_element = shadow_parent.appendChild(tmp)

    if current_node.childNodes:
        for node in current_node.childNodes:
            process_node(current_node, node, new_element)


homeDir = "/Users/paulwakelin/Dropbox/coding/python/levelChecker/"
docName = "sample2"
docSuffix = "docx"
docFile = docName + "." + docSuffix

zipped_doc = zipfile.ZipFile(homeDir + docFile)
doc_xml = xml.dom.minidom.parseString(zipped_doc.read("word/document.xml"))
shadow_xml = xml.dom.minidom.Document()


process_xml(doc_xml)

# write_xml_file(homeDir + "tweakDocX.out.xml", doc_xml.toprettyxml(indent="  "))
write_xml_file(homeDir + "out.xml", shadow_xml.toprettyxml(indent="  "))
