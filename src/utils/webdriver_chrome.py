# Imports the following functions from the standard selenium library
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import os

# Creates a new instance of the Chrome driver
# Options and driver initialization parameters


class webDriver_chrome():
    def __init__(self):
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        options.add_experimental_option("prefs", {
          "download.default_directory": r"C:\Users\miguel.silva",
          "download.prompt_for_download": False,
          "download.directory_upgrade": True,
          "safebrowsing.enabled": True
        })

        command = "xrandr | grep 'connected' | awk '{print $1}'"
        output = os.poppen(command).read()
        monitors = output.strip().slipt("\n")
        # Service initialization parameters
        service = Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=service, options=options)
        if len(monitors) >= 2:
            self.driver.set_window_position(1920, 0)
