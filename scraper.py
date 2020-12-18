from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium import webdriver
import pandas as pd
import numpy as np
import xlsxwriter
import time
import os


class WindowsInhibitor:
    """Prevent OS sleep/hibernate in windows; code from:
    https://github.com/h3llrais3r/Deluge-PreventSuspendPlus/blob/master/preventsuspendplus/core.py
    API documentation:
    https://msdn.microsoft.com/en-us/library/windows/desktop/aa373208(v=vs.85).aspx"""

    ES_CONTINUOUS = 0x80000000
    ES_SYSTEM_REQUIRED = 0x00000001

    def __init__(self):
        pass

    def inhibit(self):
        import ctypes

        print("Preventing Windows from going to sleep")
        ctypes.windll.kernel32.SetThreadExecutionState(
            WindowsInhibitor.ES_CONTINUOUS | WindowsInhibitor.ES_SYSTEM_REQUIRED
        )

    def uninhibit(self):
        import ctypes

        print("Allowing Windows to go to sleep")
        ctypes.windll.kernel32.SetThreadExecutionState(WindowsInhibitor.ES_CONTINUOUS)


def moveTableSlider(numberOfClicks, sliderButton):
    for i in range(numberOfClicks):
        sliderButton.click()
    return


def selectField(browser, visualContainersPath, mainDivsPath, fieldName):
    visualContainers = browser.find_elements_by_xpath(visualContainersPath)
    for i in range(len(visualContainers)):
        containerText = visualContainers[i].text
        if containerText[:5] == "Campo":
            containerNumber = i + 1
            break
    filterPath = (
        visualContainersPath
        + "["
        + str(containerNumber)
        + "]/transform/div/div[3]/div/visual-modern/div/div/div[2]/div/i"
    )
    fieldFilter = browser.find_element_by_xpath(filterPath)
    fieldFilter.click()
    time.sleep(0.5)
    mainDivs = browser.find_elements_by_xpath(mainDivsPath)
    slicerDropdowns = []
    for i in range(len(mainDivs)):
        if mainDivs[i].get_attribute("class") == "slicer-dropdown-popup visual":
            slicerDropdowns.append(i + 1)
    for i in slicerDropdowns:
        searchBoxPath = mainDivsPath + "[" + str(i) + "]/div[1]/div/div[1]"
        elementFound = browser.find_element_by_xpath(searchBoxPath)
        if elementFound.get_attribute("class") == "searchHeader show":
            searchBox = browser.find_element_by_xpath(searchBoxPath + "/input")
            searchBoxDivNumber = i
            break
    firstBoxPath = (
        mainDivsPath
        + "["
        + str(searchBoxDivNumber)
        + "]/div[1]/div/div[2]/div/div[1]/div/div/div/div/span"
    )
    time.sleep(0.1)
    searchBox.clear()
    time.sleep(0.1)
    searchBox.send_keys(fieldName)
    time.sleep(1)
    browser.find_element_by_xpath(firstBoxPath).click()
    fieldFilter.click()
    time.sleep(1)
    return


def getCompleteTimeSeries(browser, sliderPath):
    source_element = browser.find_element_by_css_selector(sliderPath)
    dest_element = browser.find_element_by_css_selector("span.textRun")
    ActionChains(browser).drag_and_drop(source_element, dest_element).perform()
    return


def getFieldData(browser, fieldName):

    start = time.time()

    menuItems = browser.find_elements_by_xpath(
        '//*[@id="pvExplorationHost"]/div/div/exploration/div/explore-canvas-modern/div/div[2]/div/div[2]/div[2]/visual-container-repeat/visual-container-modern/transform'
    )
    for i in range(len(menuItems)):
        if menuItems[i].text == "Poços":
            containerNumber = i + 1
            browser.find_element_by_xpath(
                '//*[@id="pvExplorationHost"]/div/div/exploration/div/explore-canvas-modern/div/div[2]/div/div[2]/div[2]/visual-container-repeat/visual-container-modern['
                + str(containerNumber)
                + "]/transform/div/div[3]/div/visual-modern/div/div"
            ).click()
            break
    time.sleep(2)

    selectField(
        browser,
        '//*[@id="pvExplorationHost"]/div/div/exploration/div/explore-canvas-modern/div/div[2]/div/div[2]/div[2]/visual-container-repeat/visual-container-modern',
        "/html/body/div",
        fieldName,
    )

    getCompleteTimeSeries(browser, "div.noUi-handle-lower")

    columns = [
        "Date",
        "ANP Well Name",
        "Oil (bbl)",
        "Gas (Mm³)",
        "Barrels of Oil Equivalent",
    ]

    tableScrollButtons = browser.find_elements_by_css_selector(
        "div.scroll-bar-part-arrow"
    )

    values = {}
    for col in columns:
        values[col] = []

    scrollBarVertical = browser.find_elements_by_css_selector(
        "div.scroll-bar-part-bar"
    )[1]
    scrollBarVerticalTop = float(scrollBarVertical.value_of_css_property("top")[:-2])
    prevTop = -1
    breakFlag = False
    emptyTableCellsFlag = False

    while True:

        if scrollBarVerticalTop == prevTop:
            if breakFlag:
                print("Last date collected: " + values[columns[0]][-1])
                break
            else:
                breakFlag = True
        else:
            breakFlag = False

        dataObjs = browser.find_element_by_css_selector("div.bodyCells")
        data = dataObjs.text.replace(",", "")
        data = data.splitlines()

        childrenDivs = browser.find_elements_by_css_selector(
            "div.bodyCells > div > div > div"
        )

        for i in range(len(columns)):
            sweepData = childrenDivs[i].text.replace(",", "").splitlines()
            for cell in sweepData:
                values[columns[i]].append(cell)

        print("Last date collected: " + values[columns[0]][-1], end="\r")

        moveTableSlider(25, tableScrollButtons[3])
        prevTop = scrollBarVerticalTop
        scrollBarVerticalTop = float(
            scrollBarVertical.value_of_css_property("top")[:-2]
        )
        time.sleep(0.1)

    df = pd.DataFrame()
    for col in columns:
        df[col] = values[col]
    df.index.name = "Índice"
    df.replace(r"^\s*$", np.nan, regex=True, inplace=True)
    df.dropna(inplace=True)
    df.iloc[:, 2:5] = df.iloc[:, 2:5].astype(float)
    df.drop_duplicates(keep="first", inplace=True)
    df.reset_index(drop=True, inplace=True)

    browser.find_element_by_css_selector("button.themableBackgroundColor").click()
    time.sleep(3)

    end = time.time()

    if end - start < 60:
        seconds = end - start
        print("Time to obtain data: {0:.0f}s".format(seconds))
    elif end - start < 3600:
        minutes = (end - start) // 60
        seconds = (end - start) % 60
        print("Time to obtain data: {0:.0f}min{1:.0f}s".format(minutes, seconds))
    else:
        hours = (end - start) // 3600
        minutes = ((end - start) % 3600) // 60
        seconds = ((end - start) % 3600) % 60
        print(
            "Time to obtain data: {0:.0f}h{1:.0f}min{2:.0f}s".format(
                hours, minutes, seconds
            )
        )

    return df


def main():
    fileName = "fields.txt"
    scriptDirPath = os.path.dirname(os.path.abspath(__file__))
    with open(scriptDirPath + "/" + fileName, "r", encoding="UTF-8") as f:
        fieldNames = f.read()
        fieldNames = fieldNames.splitlines()
    n = len(fieldNames)
    options = webdriver.ChromeOptions()
    # options.add_argument("--headless")
    options.add_argument("log-level=3")
    browser = webdriver.Chrome(
        executable_path="ChromeDriver/chromedriver.exe", options=options
    )
    osSleep = None
    # in Windows, prevent the OS from sleeping while we run
    if os.name == "nt":
        osSleep = WindowsInhibitor()
        osSleep.inhibit()
    try:
        browser.get(
            "https://www.gov.br/anp/pt-br/centrais-de-conteudo/paineis-dinamicos-da-anp/paineis-dinamicos-de-producao-de-petroleo-e-gas-natural"
        )
        time.sleep(1)
        urlPainel = browser.find_element_by_xpath(
            '//*[@id="parent-fieldname-text"]/p[9]/a'
        ).get_attribute("href")
        browser.get(urlPainel)
        time.sleep(3)
        part = 1
        for name in fieldNames:
            xlsxFile = pd.ExcelWriter(
                scriptDirPath + "/Field Data/{0}.xlsx".format(name), engine="xlsxwriter"
            )
            print("Obtaining data from {0} ({1}/{2})".format(name, part, n))
            part += 1
            df = getFieldData(browser, name)
            df.to_excel(xlsxFile, sheet_name="Data " + name)
            xlsxFile.save()
    except:
        raise
    finally:
        if osSleep:
            osSleep.uninhibit()
        browser.quit()


main()