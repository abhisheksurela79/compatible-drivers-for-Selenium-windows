import requests
from win32com.client import Dispatch
import wget
import zipfile
import os
import platform
from winreg import HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, OpenKey, QueryValueEx


def get_browser_version():
    browser, path = None, None
    os_platform = platform.system()

    if os_platform == 'Windows':

        # Find the default browser by interrogating the registry
        try:
            with OpenKey(HKEY_CURRENT_USER,
                         r'SOFTWARE\Microsoft\Windows\Shell\Associations\UrlAssociations\http\UserChoice') as reg_key:
                # Get the user choice
                browser = QueryValueEx(reg_key, 'ProgId')[0]

            with OpenKey(HKEY_CLASSES_ROOT, r'{}\shell\open\command'.format(browser)) as reg_key:
                # Get the application the user's choice refers to in the application registrations
                browser_path_tuple = QueryValueEx(reg_key, "")

                # This is a bit sketchy and assumes that the path will always be in double quotes
                path = browser_path_tuple[0].split('"')[1]

            parser = Dispatch("Scripting.FileSystemObject")
            version = parser.GetFileVersion(path)
            return browser, version

        except Exception as registryError:
            return registryError


def download_compatible_driver():
    browser, latest_driver_zip, download_url, comp = get_browser_version(), None, None, None

    if platform.architecture()[0] == "32bit":
        comp = 32
    elif platform.architecture()[0] == "64bit":
        comp = 64

    if browser[0] == "ChromeHTML":  # Chrome
        download_url = "https://chromedriver.storage.googleapis.com/" + browser[1] + f"/chromedriver_win32.zip"
        # download the zip file using the url built above

    elif browser[0] == "FirefoxURL-308046B0AF4A39CB":  # Firefox
        url = "https://github.com/mozilla/geckodriver/releases/latest"
        latest_version = requests.get(url)
        latest_version = "v" + latest_version.url.split("/v")[-1]

        download_url = f"https://github.com/mozilla/geckodriver/releases/download/" \
                       f"{latest_version}/geckodriver-v0.30.0-win{comp}.zip"

    elif browser[0] == "MSEdgeHTM":  # Microsoft Edge
        download_url = f"https://msedgedriver.azureedge.net/" + browser[1] + f"/edgedriver_win{comp}.zip"

    elif browser[0] == "OperaStable":  # Opera
        url = "https://github.com/operasoftware/operachromiumdriver/releases/latest"
        latest_version = requests.get(url)
        latest_version = "v" + latest_version.url.split("/v")[-1]

        download_url = f"https://github.com/operasoftware/operachromiumdriver/releases/download/{latest_version}/operadriver_win{comp}.zip"

    try:

        latest_driver_zip = wget.download(download_url, 'driver.zip')  # extract the zip file

        with zipfile.ZipFile(latest_driver_zip, 'r') as zip_ref:
            zip_ref.extractall()  # you can specify the destination folder path here

        os.remove(latest_driver_zip)  # delete the zip file downloaded above

    except TypeError:
        return f"{browser[0]} is not supported currently. please install Chrome, Firefox or at least Opera "






