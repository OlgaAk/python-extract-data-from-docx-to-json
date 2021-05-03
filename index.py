import os
from docx import Document
import json
from datetime import datetime


def getDocxs():
    files = os.listdir()
    docxFiles = []
    for file in files:
        if file.endswith(".docx") and "NS" in file and not file.startswith("~"):
            docxFiles.append(file)
    return docxFiles


def extractDataFromDocx(documentName):
    document = Document(documentName)
    searchedLines = []
    isSearchedLine = False
    for para in document.paragraphs:
        for line in para.text.split("\n"):
            if(line == "Russland"):
                isSearchedLine = True
            if(line == "MOEL" or line == "Polen"):
                # isSearchedLine = False
                return searchedLines
            if(isSearchedLine):
                searchedLines.append(line)
    return searchedLines


def saveDataToJson(jsonName, data):
    with open(jsonName, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)


def getJsonName():
    now = datetime.now().strftime("%d.%m.%Y-%H:%M:%S")
    return 'data-' + now + '.json'


def sortByDocNumber(doc):
    return doc["number"]


def getMissingDocuments(docArray):
    firstDocNumber = docArray[0]["number"]
    lastDocNumber = docArray[-1]["number"]+1
    totalDocumentShouldCount = lastDocNumber - firstDocNumber
    missingDocuments = list(range(firstDocNumber, lastDocNumber))
    missingDocumentCount = totalDocumentShouldCount - len(docArray)
    for doc in docArray:
        missingDocuments.remove(doc["number"])
    return missingDocuments, missingDocumentCount, totalDocumentShouldCount


def prepareFinalData(documents, resultArray):
    missingDocuments, missingDocumentCount, totalDocumentShouldCount = getMissingDocuments(
        resultArray)
    return {
        "documentsCount": len(documents),
        "processedDocumentsCount": len(resultArray),
        "totalDocumentShouldCount": totalDocumentShouldCount,
        "missingDocumentCount": missingDocumentCount,
        "missingDocuments": missingDocuments,
        "result": resultArray
    }


def extractDataFromDocxsToJson():
    jsonName = getJsonName()
    resultArray = []
    documents = getDocxs()
    for index, document in enumerate(documents, start=1):
        textArray = extractDataFromDocx(document)
        number = int(document.split(" ")[0][len("FS"):])
        data = {"number": number, document: textArray}
        resultArray.append(data)
    resultArray.sort(key=sortByDocNumber)
    for index, item in enumerate(resultArray):
        item["index"] = index
    saveDataToJson(jsonName, prepareFinalData(documents, resultArray))


extractDataFromDocxsToJson()
