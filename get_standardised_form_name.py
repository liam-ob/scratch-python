def get_standard_form_name(string):
    file_extension = os.path.splitext(string)
    file_extension = file_extension[1]
    result = filter(None, [x.strip() for x in re.findall(r"\b[A-Z]?[a-z\s]+\b", string)])
    mylist = list(result)
    end_string = ""
    for num, i in enumerate(mylist):
        if num != (len(mylist) - 1):
            end_string += i
            if num != (len(mylist) - 2):
                end_string += " "
        else:
            end_string += file_extension
    return end_string