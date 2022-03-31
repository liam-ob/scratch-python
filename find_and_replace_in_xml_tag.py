def find_and_replace_in_xml_tag(docx_path, xml_path_keyword, xml_end_tag, tag_insert):
    document = zipfile.ZipFile(docx_path)
    list = document.namelist()
    print(list)
    raw_string = ""
    i = ""
    for i in list:
        if xml_path_keyword in i:
            raw_string = str(document.read(i, pwd=None))
            break
    # xml_end_tag should be something like '</Company>'
    index = raw_string.find(xml_end_tag)
    new_xml_string = raw_string[:index] + tag_insert + raw_string[index:]
    write = document.write(docx_path)