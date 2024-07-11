from playwright.sync_api import sync_playwright
import os
import configparser
import sys
import pandas as pd


########################################################################################################################
# Defines paths for the script to use
########################################################################################################################


def find_valid_path(base_dir, alternate_paths):
    for path in alternate_paths:
        full_path = os.path.join(base_dir, *path)
        if os.path.exists(full_path):
            return full_path, True
    return base_dir, False


def get_folder_path(alternate_paths, folder_components):
    home_dir = os.path.expanduser('~')  # Get the user's home directory
    onedrive_dir = 'Kodiak Cakes'  # Folder name for OneDrive
    base_dir = os.path.join(home_dir, onedrive_dir)

    valid_path, found = find_valid_path(base_dir, alternate_paths)
    if not found:
        show_error(f"Could not find the valid base path starting from: '{base_dir}'")
        return None

    for component in folder_components:
        next_path = os.path.join(valid_path, component)
        if not os.path.exists(next_path):
            show_error(f"Path resolution stopped at: '{valid_path}'\n\nCould not find the next folder: '{component}'")
            return None
        valid_path = next_path

    return valid_path


def show_error(message):
    print("Error:", message)


def get_config_file():
    alternate_paths = [
        ['Kodiak Cakes Team Site - Public'],
        ['Kodiak Cakes Team Site - Accounting', 'Public']
    ]
    folder_components = ['AR', 'KeHe', 'config.ini']
    return get_folder_path(alternate_paths, folder_components)


def get_check_number_excel_file():
    alternate_paths = [
        ['Kodiak Cakes Team Site - Public'],
        ['Kodiak Cakes Team Site - Accounting', 'Public']
    ]
    folder_components = ['AR', 'KeHe', 'Check Numbers.xlsx']
    return get_folder_path(alternate_paths, folder_components)


config_file = get_config_file()
check_number_excel_file = get_check_number_excel_file()

if not config_file or not check_number_excel_file:
    print("Error: One or more required folders could not be found. Script cannot continue")
    sys.exit(1)


config = configparser.ConfigParser()
config.read(config_file)

# Get the username from the config file
kahe_username = config['Credentials']['Username']
kehe_password = config['Credentials']['Password']

# Get the download path from the config file
base_download_path = config['Download Path']['Path']

base_download_path = os.path.normpath(base_download_path)


# Read the Excel file
check_number_df = pd.read_excel(check_number_excel_file)

# Extract the "Check Number" column into a list
check_numbers = check_number_df["Check Number"].tolist()

def run(playwright):
    browser = playwright.chromium.launch(headless=True)
    context = browser.new_context(
        viewport={"width": 1920, "height": 1080},  # Set the desired width and height
        accept_downloads=True
    )
    page = context.new_page()
    
    # Navigate to the website
    page.goto("https://connect.kehe.com/#/dashboard")

    # Wait for the page to load
    page.wait_for_load_state("networkidle")

    page.wait_for_selector('input#username', state='visible', timeout=30000)  # Wait up to 60 seconds

    # Find the username input field and type the username
    username_field = page.locator('input#username')
    username_field.fill(kahe_username)

    # Wait for the Next button to be visible
    page.wait_for_selector('button.login-button', state='visible', timeout=30000)

    # click the Next button
    next_button = page.locator('button.login-button')
    next_button.click()

    page.wait_for_selector('input#password', state='visible', timeout=30000)  # Wait up to 60 seconds

    # Find the password input field and type the password
    password_field = page.locator('input#password')
    password_field.fill(kehe_password)

    # Wait for the Next button to be visible
    page.wait_for_selector('button.login-button', state='visible', timeout=30000)

    # Click the Next button
    next_button = page.locator('button.login-button')
    next_button.click()

    print('Successfully logged in')

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

    print('Waiting for K-Solve to load')
    # Wait for the page to load
    page.wait_for_load_state("networkidle")

    # Wait for the search input field to be visible
    page.wait_for_selector('input.search-input', state='visible', timeout=60000)

    # Function to wait for spinner
    def wait_for_spinner_to_disappear():
        page.wait_for_selector('sham-spinner .sham-spinner-blocker:not(.ng-hide)', state='visible', timeout=10000)
        page.wait_for_selector('sham-spinner .sham-spinner-blocker:not(.ng-hide)', state='hidden', timeout=60000)

# Loop through each check number
    for check_number in check_numbers:
        print(f"Processing check number: {check_number}")

        # Enter the check number into the input field
        search_field = page.locator('input.search-input')
        search_field.fill(str(check_number))

        # Wait for the search button to be enabled
        page.wait_for_selector('button.search-btn:not([disabled])', state='visible', timeout=60000)

        # Click the search button
        search_button = page.locator('button.search-btn')
        search_button.click()

        # Wait for the page to load
        page.wait_for_load_state("networkidle")

        # Wait for the spinner to appear and then disappear
        wait_for_spinner_to_disappear()

        # Create a new folder based on the check number
        download_path = os.path.join(base_download_path, str(check_number))
        os.makedirs(download_path, exist_ok=True)

        # Click the export button
        export_button = page.locator('button.btn.export-button')
        export_button.click()

        # Wait for the download to complete
        with page.expect_download() as download_info:
            download = download_info.value
            download.save_as(os.path.join(download_path, download.suggested_filename))
            print(f"Download CSV file")

        # Wait for the page grid to load
        page.wait_for_selector('#transactionGrid', state='visible')

        # Get all rows in the grid
        rows = page.query_selector_all('#transactionGrid .k-grid-content table tbody tr')
        print(f"Found {len(rows)} rows in the grid.")

        for index, row in enumerate(rows, start=1):
            print(f"Processing row {index}")
            # Find the Invoice Number from the current row
            invoice_number_element = row.query_selector('td span[ng-bind="dataItem.InvoiceNumber"]')
            if invoice_number_element:
                invoice_number = invoice_number_element.inner_text()
            else:
                continue
            # Check if the row has a visible download button
            download_button = row.query_selector('span.glyphicon-file:not(.ng-hide)')
            if download_button:
                # Click the download button
                try:
                    download_button.click()
                except Exception as e:
                    print(f"Error clicking download button: {str(e)}")
                    continue
                # Wait for the download popup to appear
                try:
                    page.wait_for_selector('img[title="Download"][cdn-image="download.png"]', state='visible', timeout=30000)
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
                except Exception as e:
                    print(f"Error downloading document: {str(e)}")
                    continue
                # Check if modal is open and visible, then close it
                try:
                    is_modal_open = page.is_visible('img[title="Download"][cdn-image="download.png"]')
                    if is_modal_open:
                        close_button = page.query_selector('button.btn-default:has-text("Cancel")')
                        if close_button:
                            close_button.click()
                    # Wait for the modal to be fully closed
                    page.wait_for_selector('img[title="Download"][cdn-image="download.png"]', state='hidden', timeout=30000)
                except Exception:
                    pass
                # Add a small delay to ensure the page is stable before moving to the next row
                page.wait_for_timeout(500)  # Wait for .5 seconds

    print("Finished processing all check numbers.")

    # # This line will keep the browser open until you press Enter
    # input("Press Enter to close the browser...")

    # Close the browser
    browser.close()

with sync_playwright() as playwright:
    run(playwright)