converts a pdf to one giant string (does not convert whole pdf)
    def convert_pdf_to_txt_and_find_title(self, possible_titles):
        rsrcmgr = PDFResourceManager()
        retstr = StringIO()
        laparams = LAParams()
        device = TextConverter(rsrcmgr, retstr, laparams=laparams)
        try:
            fp = open(self.pdf_path, 'rb')
        except:
            pass
        interpreter = PDFPageInterpreter(rsrcmgr, device)
        password = ""
        maxpages = 0
        caching = True
        pagenos = set()
    
        for page in PDFPage.get_pages(fp, pagenos,
                                      maxpages=maxpages,
                                      password=password,
                                      caching=caching,
                                      check_extractable=True):
            interpreter.process_page(page)
    
        fp.close()
        device.close()
        text = retstr.getvalue()
        retstr.close()
        for j in possible_titles:
            for k in possible_titles[j]:
                if k in text:
                    self.pdf_title = j
                    return