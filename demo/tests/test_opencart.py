import pytest
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import openpyxl

# Initialize the WebDriver
@pytest.fixture
def driver():
    driver = webdriver.Chrome()
    yield driver
    driver.quit()

def create_excel(file_name, headers):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(headers)
    wb.save(file_name)

def add_row_to_excel(file_name, row):
    wb = openpyxl.load_workbook(file_name)
    ws = wb.active
    ws.append(row)
    wb.save(file_name)

def test_opencart(driver):
    # Constants
    base_url = "https://www.opencart.com/index.php?route=cms/demo"
    excel_file = "product_details.xlsx"

    # Test Steps
    print("[INFO] Navigating to OpenCart demo page")
    driver.get(base_url)

    print("[INFO] Clicking on 'Demo' link")
    demo_link = driver.find_element(By.XPATH, '//*[@id="navbar-collapse-header"]/ul/li[2]/a')
    demo_link.click()

    # Step 2: Click on 'View Store Front'
    try:
        print("[INFO] Waiting for 'View Store Front' link to be clickable")
        view_store_front = WebDriverWait(driver, 40).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="cms-demo"]/div[2]/div/div[1]/div/a/span'))
        )
        view_store_front.click()
        print("[INFO] 'View Store Front' link clicked")
    except TimeoutException:
        print("[ERROR] 'View Store Front' link not clickable within the specified time.")
        driver.quit()
        return

    driver.switch_to.window(driver.window_handles[-1])
    print("[INFO] Switched to new tab")

    # Step 3a: Create Excel sheet and extract product details
    create_excel(excel_file, ["Product Name", "MRP", "Discount Price"])
    print("[INFO] Excel workbook created")

    try:
        print("[INFO] Extracting product details from the PLP")
        products = WebDriverWait(driver, 100).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".product-thumb"))
        )
        for product in products:
            name = product.find_element(By.CSS_SELECTOR, "h4 a").text
            price = product.find_element(By.CSS_SELECTOR, ".price").text
            if '\n' in price:
                mrp, discount_price = price.split('\n')
            else:
                mrp = price
                discount_price = "N/A"
            add_row_to_excel(excel_file, [name, mrp, discount_price])
    except TimeoutException:
        print("[ERROR] Products not found within the specified time.")
        driver.quit()
        return

    print("[INFO] Product details extracted and saved to Excel")

    # Step 4: Click on any product to see product details
    product_link = products[0].find_element(By.CSS_SELECTOR, "h4 a")
    product_link.click()

    try:
        print("[INFO] Waiting for product details page to load")
        WebDriverWait(driver, 40).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, "#button-cart"))
        )
        print("[INFO] Product details page loaded")
    except TimeoutException:
        print("[ERROR] Product details page not loaded within the specified time.")
        driver.quit()
        return

    # Step 5: Change the quantity of the product in the PDP
    quantity_input = driver.find_element(By.ID, "input-quantity")
    quantity_input.clear()
    quantity_input.send_keys("2")
    print("[INFO] Changed product quantity to 2")

    # Step 6: Add product to the cart
    add_to_cart_button = driver.find_element(By.ID, "button-cart")
    add_to_cart_button.click()
    print("[INFO] Product added to cart")

    try:
        print("[INFO] Waiting for cart to update")
        WebDriverWait(driver, 40).until(
            EC.text_to_be_present_in_element((By.CSS_SELECTOR, "#cart-total"), "2 item(s)")
        )
        print("[INFO] Cart updated with 2 items")
    except TimeoutException:
        print("[ERROR] Cart not updated within the specified time.")
        driver.quit()
        return

    # Step 7: Navigate to the cart page
    cart_button = driver.find_element(By.CSS_SELECTOR, "#cart > button")
    cart_button.click()
    view_cart_button = driver.find_element(By.XPATH, '//*[@id="cart"]/ul/li[2]/div/p/a[1]/strong')
    view_cart_button.click()

    try:
        print("[INFO] Waiting for cart page to load")
        WebDriverWait(driver, 40).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, ".table-responsive"))
        )
        print("[INFO] Cart page loaded")
    except TimeoutException:
        print("[ERROR] Cart page not loaded within the specified time.")
        driver.quit()
        return

    # Validate product details in the cart page
    cart_product_name = driver.find_element(By.CSS_SELECTOR, ".table-responsive tbody tr td:nth-child(2) a").text
    assert cart_product_name == name, "[ERROR] Product name in cart does not match the product added."

    # Step 8: Change the quantity of the product in the cart
    cart_quantity_input = driver.find_element(By.CSS_SELECTOR, ".table-responsive tbody tr td:nth-child(4) input")
    cart_quantity_input.clear()
    cart_quantity_input.send_keys("1")
    update_button = driver.find_element(By.CSS_SELECTOR, ".table-responsive tbody tr td:nth-child(4) button")
    update_button.click()
    print("[INFO] Changed product quantity to 1 in the cart")

    try:
        print("[INFO] Waiting for cart to update")
        WebDriverWait(driver, 40).until(
            EC.text_to_be_present_in_element((By.CSS_SELECTOR, "#cart-total"), "1 item(s)")
        )
        print("[INFO] Cart updated with 1 item")
    except TimeoutException:
        print("[ERROR] Cart not updated within the specified time.")
        driver.quit()
        return

    # Step 9: Proceed to checkout
    checkout_button = driver.find_element(By.CSS_SELECTOR, "a.btn.btn-primary")
    checkout_button.click()

    try:
        print("[INFO] Waiting for checkout page to load")
        WebDriverWait(driver, 40).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, "#collapse-checkout-option"))
        )
        print("[INFO] Checkout page loaded")
    except TimeoutException:
        print("[ERROR] Checkout page not loaded within the specified time.")
        driver.quit()
        return

    print("[INFO] Order placed successfully")
