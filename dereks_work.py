from __future__ import print_function
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdftypes import resolve1
from pdfminer.layout import LTTextBoxHorizontal
from pdfminer.layout import LAParams
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.converter import PDFPageAggregator
from pdfminer.pdfpage import PDFPage
import os
from glob import glob
from xlsxwriter.workbook import Workbook
import numpy as np


def returnFieldDictValue(dictionary, keyname):
    try:
        retVal = False if dictionary[keyname] is None else True
    except KeyError:
        retVal = False
    return retVal


def searchPDF4String(pdf, str):
    retVal = False

    resource_manager = PDFResourceManager()
    laparams = LAParams()
    device = PDFPageAggregator(resource_manager, laparams=laparams)
    interpreter = PDFPageInterpreter(resource_manager, device)

    for page in PDFPage.create_pages(pdf):
        interpreter.process_page(page)

        layout = device.get_result()
        for element in layout:
            if type(element) is LTTextBoxHorizontal:
                if element.get_text().find(str) != -1:
                    retVal = True
                    break

    return retVal


def testPAF(pdf):
    testOutput = False
    PDMsigned = False
    GMSigned = False
    PTLSigned = False
    PSSigned = False

    try:
        fields = resolve1(pdf.catalog['AcroForm'])['Fields']
    except KeyError:
        testOutput = False
    else:
        fieldDict = {}
        for i in fields:
            field = resolve1(i)
            name, value = field.get('T'), field.get('V')
            fieldDict[name] = value

        if 'Contract Review completed by' in fieldDict:
            testOutput = True
            PDMsigned = returnFieldDictValue(fieldDict, 'PDMSign')
            GMSigned = returnFieldDictValue(fieldDict, 'GMSign')
            PTLSigned = returnFieldDictValue(fieldDict, 'PTLSign')
            PSSigned = returnFieldDictValue(fieldDict, 'PSSign')

    return [testOutput, PDMsigned, GMSigned, PTLSigned, PSSigned]


def testRiskAss(pdf):
    testOutput = False
    sig2 = False
    sig3 = False
    sig4 = False

    try:
        fields = resolve1(pdf.catalog['AcroForm'])['Fields']
    except KeyError:
        testOutput = False
    else:
        fieldDict = {}
        for i in fields:
            field = resolve1(i)
            name, value = field.get('T'), field.get('V')
            fieldDict[name] = value

        if searchPDF4String(pdf, u'NEW\u00A0WORK\u00A0RISK\u00A0ASSESSMENT'):
            testOutput = True
            sig2 = returnFieldDictValue(fieldDict, 'Signature2')
            sig3 = returnFieldDictValue(fieldDict, 'Signature3')
            sig4 = returnFieldDictValue(fieldDict, 'Signature4')
        else:
            testOutput = False

    return [testOutput, sig2, sig3, sig4]


def testCSF(pdf):
    testOutput = False
    sub_A = False
    PTL_A = False
    App_A = False
    GM_A = False
    subASign = False
    sub_B = False
    PTL_B = False
    App_B = False
    GM_B = False
    subBSign = False

    try:
        fields = resolve1(pdf.catalog['AcroForm'])['Fields']
    except KeyError:
        testOutput = False
    else:
        fieldDict = {}
        for i in fields:
            field = resolve1(i)
            name, value = field.get('T'), field.get('V')
            fieldDict[name] = value

        if 'CSFNo' in fieldDict:
            testOutput = True
            sub_A = returnFieldDictValue(fieldDict, 'SubmACost')
            PTL_A = returnFieldDictValue(fieldDict, 'PTLSign')
            App_A = returnFieldDictValue(fieldDict, 'AppSign')
            GM_A = returnFieldDictValue(fieldDict, 'GMSign')
            subASign = returnFieldDictValue(fieldDict, 'SubmABy')

            sub_B = returnFieldDictValue(fieldDict, 'SubmBCost')
            PTL_B = returnFieldDictValue(fieldDict, 'PTLSignB')
            App_B = returnFieldDictValue(fieldDict, 'AppSignB')
            GM_B = returnFieldDictValue(fieldDict, 'GMSignB')
            subBSign = returnFieldDictValue(fieldDict, 'SubBBy')

    return [testOutput, sub_A, PTL_A, App_A, GM_A, subASign, sub_B, PTL_B, App_B, GM_B, subBSign]


def testPQEP(pdf):
    testOutput = False
    sig2 = False
    sig3 = False
    sig4 = False
    sig5 = False

    try:
        fields = resolve1(pdf.catalog['AcroForm'])['Fields']
    except KeyError:
        testOutput = False
    else:
        fieldDict = {}
        for i in fields:
            field = resolve1(i)
            name, value = field.get('T'), field.get('V')
            fieldDict[name] = value

        if searchPDF4String(pdf, u"PROJECT QUALITY AND EXECUTION PLAN"):
            testOutput = True
            sig2 = returnFieldDictValue(fieldDict, 'Signature2')
            sig3 = returnFieldDictValue(fieldDict, 'Signature3')
            sig4 = returnFieldDictValue(fieldDict, 'Signature4')
            sig5 = returnFieldDictValue(fieldDict, 'Signature5')

    return [testOutput, sig2, sig3, sig4, sig5]


def testPQP(pdf):
    testOutput = False
    return testOutput


def testPEP(pdf):
    testOutput = False
    return testOutput


def testProjCloseOut(pdf):
    testOutput = False
    PTLSigned = False
    GMSigned = False
    PDMsigned = False
    PSSigned = False

    try:
        fields = resolve1(pdf.catalog['AcroForm'])['Fields']
    except KeyError:
        testOutput = False
    else:
        fieldDict = {}
        for i in fields:
            field = resolve1(i)
            name, value = field.get('T'), field.get('V')
            fieldDict[name] = value

        if 'Feedback Received' in fieldDict:
            testOutput = True
            PTLSigned = returnFieldDictValue(fieldDict, 'PTL Signature')
            GMSigned = returnFieldDictValue(fieldDict, 'GM Signature')
            PDMsigned = returnFieldDictValue(fieldDict, 'PDM Signature')
            PSSigned = returnFieldDictValue(fieldDict, 'PDM Signature')

    return [testOutput, PTLSigned, GMSigned, PDMsigned, PSSigned]


def testClientFeedback(pdf):
    testOutput = False
    sig1 = False
    sig2 = False

    try:
        fields = resolve1(pdf.catalog['AcroForm'])['Fields']
    except KeyError:
        testOutput = False
    else:
        fieldDict = {}
        for i in fields:
            field = resolve1(i)
            name, value = field.get('T'), field.get('V')
            fieldDict[name] = value

        if 'Feedback Provided byName' in fieldDict:
            testOutput = True
            sig1 = returnFieldDictValue(fieldDict, 'Signature1')
            sig2 = returnFieldDictValue(fieldDict, 'Signature2')

    return [testOutput, sig1, sig2]


def runAssessment(locPath, projNumber, book):
    print('Running Project {0}\n'.format(projNumber))

    PRA = []  # Proposal Risk Assessment
    CSF = []  # Commercial Submission Form
    PAF = []  # Project Assignment Form
    PQEP = []  # Combined Project Quality and Execution Plan
    PQP = []  # Project Quality Plan
    PEP = []  # Project Execution Plan
    PCO = []  # Project Close Out Form
    CFB = []  # Client Feedback Form
    ICC = []  # Internal Check Copies (Reports)
    Other = []  # Other files

    unknownCount = 0

    # Proposal stage files
    f000 = [y for x in os.walk(mypath + "000 Project Tender\\020 Proposal Forms\\")
            for y in glob(os.path.join(x[0], '*.pdf'))]
    # Project management files
    f100 = [y for x in os.walk(mypath + "100 Project Control\\") for y in glob(os.path.join(x[0], '*.pdf'))]

    totalF = len(f000) + len(f100)
    idx = 1

    for f in f000:
        print('Assessing File {0} of {1}'.format(idx, totalF), end='\r')
        fp = open(f, 'rb')
        parser = PDFParser(fp)
        doc = PDFDocument(parser)

        # Test if Risk Assessment
        PRAResults = testRiskAss(doc)
        if PRAResults[0]:
            # print 'Risk Assessment Found'
            PRA.append(np.hstack((f, PRAResults[1:])))

        # Test if CSF
        else:
            CSFResults = testCSF(doc)
            if CSFResults[0]:
                # print 'CSF Found'
                CSF.append(np.hstack((f, CSFResults[1:])))
            else:
                unknownCount = unknownCount + 1
        idx = idx + 1

    for f in f100:
        print('Assessing File {0} of {1}'.format(idx, totalF), end='\r')
        fp = open(f, 'rb')
        parser = PDFParser(fp)
        doc = PDFDocument(parser)

        # Test if PAF
        PAFResults = testPAF(doc)
        if PAFResults[0]:
            # print 'PAF Found'
            PAF.append(np.hstack((f, PAFResults[1:])))

        # Test if CSF
        else:
            CSFResults = testCSF(doc)
            if CSFResults[0]:
                # print 'CSF Found'
                CSF.append(np.hstack((f, CSFResults[1:])))

            # Test if PQEP
            else:
                PQEPResults = testPQEP(doc)
                if PQEPResults[0]:
                    # print 'PQEP Found'
                    PQEP.append(np.hstack((f, PQEPResults[1:])))

                # Test if PQP
                else:
                    PQPResults = [testPQP(doc)]
                    if PQPResults[0]:
                        # print 'PQP Found'
                        PQP.append(np.hstack((f, PQPResults[1:])))

                    # Test if PEP
                    else:
                        PEPResults = [testPEP(doc)]
                        if PEPResults[0]:
                            # print 'PEP Found'
                            PQP.append(np.hstack((f, PEPResults[1:])))

                        # Test if Project Close Out
                        else:
                            PCOResults = testProjCloseOut(doc)
                            if PCOResults[0]:
                                # print 'Project Close Out Form Found'
                                PCO.append(np.hstack((f, PCOResults[1:])))

                            # Test if Client Feedback Form
                            else:
                                CFBResults = testClientFeedback(doc)
                                if CFBResults[0]:
                                    # print 'Client Feedback Found'
                                    CFB.append(np.hstack((f, CFBResults[1:])))
                                else:
                                    Other.append(f)
                                    unknownCount = unknownCount + 1
        idx = idx + 1

    print('\n')

    outputResults(projNumber, PRA, CSF, PAF, PQEP, PQP, PEP, PCO, CFB, ICC, Other, book)


def outputIndCase(header, data, offset, book, sheet):
    header_format = book.add_format({'bold': True,
                                     'valign': 'vcenter',
                                     'bg_color': 'gray',
                                     'border': 1})

    value_format = book.add_format({'valign': 'vcenter',
                                    'border': 1})

    for col1, curHeader in enumerate(header):
        sheet.write(offset, col1, curHeader, header_format)
    if len(data):
        for row1, curData in enumerate(data):
            if col1 > 1:
                for col in range(len(curData)):
                    if col == 0:
                        hyplink = '=HYPERLINK("{0}")'.format(curData[col])
                        #hyplink = curData[col]
                        sheet.write_formula(offset + row1 + 1, col, hyplink, value_format)
                    else:
                        sheet.write(offset + row1 + 1, col, curData[col], value_format)
            else:
                hyplink = '=HYPERLINK("{0}")'.format(curData)
                #hyplink = curData
                sheet.write_formula(offset + row1 + 1, 0, hyplink, value_format)

        offset = offset + row1 + 3
    else:
        sheet.write(offset + 1, 0, 'None found!')
        offset = offset + 3

    return offset


def outputResults(pNum, PRA, CSF, PAF, PQEP, PQP, PEP, PCO, CFB, ICC, Other, book):
    print('Outputting Data to File')
    sh = book.add_worksheet(pNum)

    left_format = book.add_format({'align': 'left'})
    center_format = book.add_format({'align': 'center'})

    offset = 0
    offset = outputIndCase(['Risk Assessment Filename',
                            'Signature 2',
                            'Signature 3',
                            'Signature 4'],
                           PRA, offset, book, sh)

    offset = outputIndCase(['Commercial Submissions Form Filename',
                            'Submission A',
                            'PTL Signed',
                            'Approver Signed',
                            'GM Signed',
                            'Submitter Signed',
                            'Submission B',
                            'PTL Signed',
                            'Approver Signed',
                            'GM Signed',
                            'Submitter Signed'],
                           CSF, offset, book, sh)

    offset = outputIndCase(['Project Assignement Form Filename',
                            'PDM Signed',
                            'GM Signed',
                            'PTL Signed',
                            'Sponsor Signed'],
                           PAF, offset, book, sh)

    offset = outputIndCase(['PQEP Filename',
                            'Signature 2',
                            'Signature 3',
                            'Signature 4',
                            'Signature 5'],
                           PQEP, offset, book, sh)

    offset = outputIndCase(['PEP Filename',
                            'Signature 2',
                            'Signature 3',
                            'Signature 4',
                            'Signature 5'],
                           PEP, offset, book, sh)

    offset = outputIndCase(['PQP Filename',
                            'Signature 2',
                            'Signature 3',
                            'Signature 4',
                            'Signature 5'],
                           PQP, offset, book, sh)

    offset = outputIndCase(['Close Out Form Filename',
                            'Signature 2',
                            'Signature 3',
                            'Signature 4',
                            'Signature 5'],
                           PCO, offset, book, sh)

    offset = outputIndCase(['Client Feedback Form Filename',
                            'Client Signature',
                            'Atteris Signature'],
                           CFB, offset, book, sh)

    offset = outputIndCase(['Other Files'],
                           Other, offset, book, sh)

    sh.set_column(0, 0, 100, left_format)
    sh.set_column(1, 10, 16, center_format)

    error_format = book.add_format({'bg_color':   '#FFC7CE',
                                    'font_color': '#9C0006'})

    sh.conditional_format('B1:K' + str(offset), {'type':     'cell',
                                                 'criteria': '==',
                                                 'value':    r'"False"',
                                                 'format':   error_format})


if __name__ == "__main__":
    pNumList = [r"17-059",
                r"18-021",
                r"18-018"]

    book = Workbook('Output.xlsx')
    for ProjectNumber in pNumList:
        for mypath in glob('P:\\*' + ProjectNumber + '*'):
            mypath = mypath + '\\'

        if mypath is None:
            for mypath in glob('R:\\*' + ProjectNumber + '*'):
                mypath = mypath + '\\'

        if mypath is None:
            print("Directory for Project Number {0} can not be found on the P: or R: drive".format(ProjectNumber))
        else:
            runAssessment(mypath, ProjectNumber, book)

    book.close()
