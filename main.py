from createFile import CreateFile, ParsingShop



def runParsing():
    parsing = ParsingShop()

    # Создание списка элементов ссылок
    catalogElem = parsing.pageLoading()

    # Создание списка ссылок на разделы с товарами
    catalogUrl = parsing.createLincs(catalogElem)

    # Создание списка со словарями, включающими данные о товарах
    info = parsing.parsingProducts(catalogUrl)

    # Создание таблицы и занесение в неё данных
    CreateFile().addContent(info)


if __name__ == '__main__':
    runParsing()
