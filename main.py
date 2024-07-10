from playwright.sync_api import sync_playwright
import os


# Set the download path
download_path = r"C:\Users\Andy Weigl\Documents\KeHe Downloads"

def run(playwright):
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context(accept_downloads=True)
    page = context.new_page()

    # Navigate to the website
    page.goto("https://connect.kehe.com/#/dashboard")

    # Wait for the page to load
    page.wait_for_load_state("networkidle")

    page.wait_for_selector('input#username', state='visible', timeout=30000)  # Wait up to 60 seconds

    # Find the username input field and type the username
    username_field = page.locator('input#username')
    username_field.fill("garrett.culligan@kodiakcakes.com")

    # Wait for the Next button to be visible
    page.wait_for_selector('button.login-button', state='visible', timeout=30000)

    # click the Next button
    next_button = page.locator('button.login-button')
    next_button.click()

    page.wait_for_selector('input#password', state='visible', timeout=30000)  # Wait up to 60 seconds

    # Find the password input field and type the password
    password_field = page.locator('input#password')
    password_field.fill("Apple#22")

    # Wait for the Next button to be visible
    page.wait_for_selector('button.login-button', state='visible', timeout=30000)

    # Click the Next button
    next_button = page.locator('button.login-button')
    next_button.click()

    # Wait for the page to load
    page.wait_for_load_state("networkidle")

    # Wait for the "Stay on Site" button to be visible
    page.wait_for_selector('button.btn-secondary:has-text("Stay on Site")', state='visible', timeout=90000)

    # Find and click the "Stay on Site" button
    stay_button = page.locator('button.btn-secondary:has-text("Stay on Site")')
    stay_button.click()

    # Wait for the page to load
    page.wait_for_load_state("networkidle")

    # Wait for the K-Solve button to be visible
    page.wait_for_selector('h2.text-right:has-text("K-Solve")', state='visible', timeout=60000)

    # Click the K-Solve button
    k_solve_element = page.locator('h2.text-right:has-text("K-Solve")')
    k_solve_element.click()

    # Wait for the page to load
    page.wait_for_load_state("networkidle")

    # Wait for the search input field to be visible
    page.wait_for_selector('input.search-input', state='visible', timeout=60000)

    # Enter the check number into the input field
    search_field = page.locator('input.search-input')
    search_field.fill("1073551")

    # Wait for the search button to be enabled (it's disabled when the input is empty)
    page.wait_for_selector('button.search-btn:not([disabled])', state='visible', timeout=60000)

    # Click the search button
    search_button = page.locator('button.search-btn')
    search_button.click()

    # Wait for the page to load
    page.wait_for_load_state("networkidle")

    # Function to wait for spinner
    def wait_for_spinner_to_disappear():
        print("Waiting for spinner to appear...")
        page.wait_for_selector('sham-spinner .sham-spinner-blocker:not(.ng-hide)', state='visible', timeout=10000)
        print("Spinner appeared, waiting for it to disappear...")
        page.wait_for_selector('sham-spinner .sham-spinner-blocker:not(.ng-hide)', state='hidden', timeout=60000)
        print("Spinner disappeared, data should be loaded.")

    # Wait for the spinner to appear and then disappear
    wait_for_spinner_to_disappear()

    # Wait for the page grid to load
    print("Waiting for grid to load...")
    page.wait_for_selector('#transactionGrid', state='visible')
    print("Grid loaded.")

    # Get all rows in the grid
    rows = page.query_selector_all('#transactionGrid .k-grid-content table tbody tr')
    print(f"Found {len(rows)} rows in the grid.")

    for index, row in enumerate(rows, start=1):
        print(f"Processing row {index}")

        # Find the Invoice Number from the current row
        invoice_number_element = row.query_selector('td span[ng-bind="dataItem.InvoiceNumber"]')
        if invoice_number_element:
            invoice_number = invoice_number_element.inner_text()
            print(f"Found Invoice Number: {invoice_number}")
        else:
            print(f"No Invoice Number found in row {index}")
            continue  # Skip this row if no invoice number is found

        # Check if the row has a visible download button
        download_button = row.query_selector('span.glyphicon-file:not(.ng-hide)')
        if download_button:
            print(f"Found visible download button in row {index}")

            # Click the download button
            try:
                download_button.click()
                print("Clicked download button")
            except Exception as e:
                print(f"Error clicking download button: {str(e)}")
                continue

            # Wait for the download popup to appear
            try:
                print("Waiting for download modal...")
                page.wait_for_selector('img[title="Download"][cdn-image="download.png"]', state='visible', timeout=5000)
                print("Download modal appeared")
            except Exception as e:
                print(f"Error waiting for download modal: {str(e)}")
                continue

            # Click the download button in the modal
            try:
                with page.expect_download() as download_info:
                    page.click('img[title="Download"][cdn-image="download.png"]')
                download = download_info.value
                # Use the Invoice Number for the filename
                download.save_as(os.path.join(download_path, f"{invoice_number}.pdf"))
                print(f"Downloaded document for Invoice Number: {invoice_number}")
            except Exception as e:
                print(f"Error downloading document: {str(e)}")
                continue

            # checks if model is open and visable, then closes it
            try:
                is_modal_open = page.is_visible('img[title="Download"][cdn-image="download.png"]')
                if is_modal_open:
                    close_button = page.query_selector('button.btn-default:has-text("Cancel")')
                    if close_button:
                        close_button.click()
                        print("Clicked close button")
                    else:
                        print("Close button not found, but modal was open")
                else:
                    print("Modal was already closed")

                # Wait for the modal to be fully closed
                page.wait_for_selector('img[title="Download"][cdn-image="download.png"]', state='hidden', timeout=5000)
                print("Modal closed")
            except Exception as e:
                print(f"Error during modal closing process: {str(e)}")

            # Add a small delay to ensure the page is stable before moving to the next row
            page.wait_for_timeout(500)  # Wait for .5 seconds
        else:
            print(f"No visible download button found in row {index}")

    print("Finished processing all rows.")



    # # This line will keep the browser open until you press Enter
    # input("Press Enter to close the browser...")

    # Close the browser
    browser.close()


with sync_playwright() as playwright:
    run(playwright)