def xlsx_to_dict(path):
    book = pandas.read_excel(path)
    data_frame = pandas.DataFrame(book)
    data_frame = data_frame.dropna()

    # puts data frame into a dictionary and manipulates it to a usable dictionary
    dicti = data_frame.to_dict(orient='tight')
    small_dict = dicti['data']
    final_dict = {}
    for x in small_dict:
        final_dict[x[0]] = x[1]

    return final_dict