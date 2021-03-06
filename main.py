from selenium import webdriver
from selenium.webdriver import Chrome
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import logging
from PIL import Image
import io
import argparse

logger = logging.getLogger()


def show_element_screen(element):
    imageStream = io.BytesIO(element.screenshot_as_png)
    img = Image.open(imageStream)
    img.show()


def save_element_screen(element, newfilename):
    img_bytes = element.screenshot_as_png
    image_stream = io.BytesIO(img_bytes)
    img = Image.open(image_stream)
    img.save(newfilename)


def set_datepicker_date(driver, date: str):
    logger.info(f"Selecting date {date} in date picker...")
    date = date.split("-")
    driver.switch_to.frame(driver.find_element_by_tag_name("iframe"))
    dropdowns = WebDriverWait(driver, 10).until(
        EC.visibility_of_all_elements_located((By.XPATH, "//div[@id='dpTitle']//input[@class='yminput']"))
    )

    # Set Month and Year
    dropdowns[1].click()
    dropdowns[1].send_keys(Keys.BACKSPACE)
    dropdowns[1].send_keys(Keys.BACKSPACE)
    dropdowns[1].send_keys(date[0])
    dropdowns[0].click()
    dropdowns[0].click()
    dropdowns[0].send_keys(Keys.BACKSPACE)
    dropdowns[0].send_keys(Keys.BACKSPACE)
    dropdowns[0].send_keys(Keys.BACKSPACE)
    dropdowns[0].send_keys(Keys.BACKSPACE)
    dropdowns[0].send_keys(date[1])
    dropdowns[1].click()

    date_element = driver.find_elements_by_xpath(f"//table[@class='WdayTable']//td[text()='{date[2]}']")
    if int(date[2]) > 15:
        date_element[-1].click()
    else:
        date_element[0].click()

    driver.switch_to.default_content()


def get_single_data(table, results, screen):
    td_elements = table.find_elements_by_xpath("//tr[@id='tr0']/td")
    for td in td_elements:
        if td.get_attribute('id') == "dcq0":
            maturity = td.get_attribute('innerHTML')
        elif td.get_attribute('id') == "syl0":
            yield_pc = td.get_attribute('innerHTML')
    data = (float(maturity), float(yield_pc))
    if data not in results:
        results.append(data)
        save_element_screen(screen, f"screensave/{int(data[0])}.png")


def collect_data_on_graph(webdriver):
    screen = webdriver.find_element_by_xpath("//div[@id='main']")
    actions = ActionChains(webdriver)
    actions.move_to_element(screen).perform()
    table = webdriver.find_element_by_xpath("//div[@id='dataTable']//table[@id='table1']")
    results = []

    graph = webdriver.find_element_by_xpath("//div[@id='container']//*[name()='svg']")
    logger.info(f"Graph size is {graph.size}")
    y = int(graph.size['height'] / 2)

    get_single_data(table, results, screen)
    action = ActionChains(webdriver)
    action.move_to_element_with_offset(graph, 56, y).click().perform()
    get_single_data(table, results, screen)

    while len(results) < 51:
        action = ActionChains(webdriver)
        action.key_down(Keys.ARROW_RIGHT).key_up(Keys.ARROW_RIGHT).perform()
        get_single_data(table, results, screen)

    logger.info(f"{len(results)} data points collected")
    return results


def write_results_to_csv(results, directory, page, date):
    if directory == "":
        filename = f"cn_yield_data {date}.csv"
    else:
        filename = f"{directory}\\cn_yield_data {date}.csv"

    logger.info(f"Writing collected data to csv: {filename}")
    with open(filename, 'w') as f:
        f.write(f"Source,{page}\n")
        f.write(f"As of {date}\n\n")
        f.write("Maturity,Yield\n")
        for data in results:
            f.write(f"{data[0]},{data[1]}\n")


def main(date: str, driver_path: str):
    logger.setLevel(logging.INFO)
    handler = logging.StreamHandler()
    formatter = logging.Formatter("%(levelname)s - %(message)s")
    handler.setFormatter(formatter)
    logger.addHandler(handler)

    logger.info("Starting browser...")
    # Config options for Chrome
    options = webdriver.ChromeOptions()
    options.add_argument('--ignore-certificate-errors')
    options.add_argument("--test-type")

    # Init Chrome Driver and negivate to target page
    driver = Chrome(executable_path=driver_path, options=options)
    page = "http://yield.chinabond.com.cn/cbweb-mn/yield_main?locale=en_US"
    logger.info(f"Loading page {page}")
    driver.get(page)
    logger.info("Page loaded successfully")

    dropdownlist = driver.find_element_by_class_name("chartQuota")
    dropdownlist.click()
    drop_down_options = driver.find_elements_by_xpath(
        "//div[@class='chartOptionsFlowTrend']//input[@name='xycheck']")

    # Select spot rate curve
    logger.info("Seleting Spot rate curve type")
    drop_down_options[0].click()
    drop_down_options[1].click()
    # Close drop down list
    dropdownlist.click()

    # Input dates
    date_img = driver.find_element_by_xpath("//div[@id='bondyId']//img[@id='img_date']")
    date_img.click()
    set_datepicker_date(driver, date)
    date_img.click()

    # Click search
    search_button = driver.find_element_by_xpath(
        "//div[@style='height: 23px;']//button[contains(text(), 'Search')]")
    search_button.click()

    # Collect all yield data
    try:
        results = collect_data_on_graph(driver)
    except:
        logger.warning("Failed, retrying")
        results = collect_data_on_graph(driver)

    driver.close()

    write_results_to_csv(results, "", page, date)


if __name__ == '__main__':
    date = "2020-12-31"  # yyyy-mm-dd
    driver_path = ""  # path to your Chrome web driver
    main(date, driver_path)
