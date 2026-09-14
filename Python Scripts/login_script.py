from playwright.sync_api import sync_playwright
from playwright.sync_api import TimeoutError as TimeoutError
from pathlib import Path
import yaml

# Define the path to the config.yaml file
BASE_DIR = Path(__file__).resolve().parent
file_path = BASE_DIR / "yaml files" / "config.yaml"

# Load email and password from config.yaml
with open(file_path, "r") as f:
    config = yaml.safe_load(f)

with sync_playwright() as p:
    try:
        # Launch the browser (Edge) and create a new context
        browser = p.chromium.launch(
            # executable_path=r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe", if you want to specify the path to the Edge browser executable
            channel="msedge",
            headless=False
        )

        # Create a new browser context
        context = browser.new_context()
        page = context.new_page()
        page.goto("https://outlook.office.com")

        # Wait for the login page to load and fill in the email and password fields
        page.wait_for_timeout(1000)
        page.get_by_role("textbox", name="Enter your email, phone, or").fill(config["email"])
        page.keyboard.press("Enter")

        page.wait_for_timeout(1000)
        page.get_by_role("textbox", name="Enter the password for").fill(config["password"])
        page.keyboard.press("Enter")

        # Check if the "New" button is present to confirm successful login
        new_button = page.get_by_role("button", name="New", exact=True)

        try:
            new_button.wait_for(state="visible", timeout=120_000)
        except TimeoutError:
            raise SystemExit("Login not detected within 2 minutes. Session not saved.")

        # Save the session state to a file
        context.storage_state(path="outlook_session.json" )

        browser.close()

    finally:
        print("Session saved to outlook_session.json")
