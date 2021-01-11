from win32api import GetFileVersionInfo, LOWORD, HIWORD
from selenium import webdriver
import zipfile
import wget
import sys
import re
import os
import requests
import pickle
from time import sleep

class Session:
    def __init__(self, log=None, profile=True):
        self.log = log
        self.browser = None
        self.headless = False
        self.download_path = os.getcwd()
        self.chromedriver_path = os.path.dirname(__file__)
        self.profile = {"download.prompt_for_download": False,
                        "download.directory_upgrade": True,
                        "download.default_directory": self.download_path}
        self.session_id = None
        self.executor_url = None
        self.chrome_version = self.get_chrome_version()
        self.chromedriver_version = None
        if profile:
          self.profile_folder = f'{os.path.dirname(__file__)}/ChromeProfile'
        else: 
          self.profile_folder = False
        
    def set_download_path(self, path):
        self.download_path = path
        self.profile["download.default_directory"] = self.download_path
        
    def set_chromedriver_path(self, path):
        self.chromedriver_path = path
        
    def get_browser(self):
        self.setup_browser()
        self.session_id = self.browser.session_id
        self.executor_url = self.browser.command_executor._url
        if self.log:
            self.log.info(f'Sessão iniciada. ID: {self.session_id} - EXECUTOR_URL: {self.executor_url}')
        else:
            print(f'Sessão iniciada. ID: {self.session_id} - EXECUTOR_URL: {self.executor_url}')
        return self.browser
        
    def set_profile(self, prof):
        self.profile = prof
        
    def headless(self, h):
        self.headless = h
        
    def remote_session(self, session_id, executor_url):
        # Code by tarunlalwani@GitHub
        
        from selenium.webdriver.remote.webdriver import WebDriver as RemoteWebDriver

        # Save the original function, so we can revert our patch
        org_command_execute = RemoteWebDriver.execute

        def new_command_execute(self, command, params=None):
            if command == "newSession":
                # Mock the response
                return {'success': 0, 'value': None, 'sessionId': session_id}
            else:
                return org_command_execute(self, command, params)

        # Patch the function before creating the driver object
        RemoteWebDriver.execute = new_command_execute

        new_driver = webdriver.Remote(command_executor=executor_url, desired_capabilities={})
        new_driver.session_id = session_id

        # Replace the patched function with original function
        RemoteWebDriver.execute = org_command_execute
        
        self.browser = new_driver
        self.session_id = self.browser.session_id
        self.executor_url = self.browser.command_executor._url
        
        return self.browser
        
    def get_chrome_version(self):
        try:
            info = GetFileVersionInfo(r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe', "\\")
            ms = info['FileVersionMS']
            ls = info['FileVersionLS']
            return HIWORD(ms)
        except:
            return "Unknown Version"
        
    def get_chromedriver_version(self):
        stream = os.popen('{}\\chromedriver.exe --version'.format(self.chromedriver_path))
        output = stream.read()
        chromedriver_version = re.search('\d*\.\d*\.\d*\.\d*', output).group().split('.')[0]
        return chromedriver_version
      
    def download_chromedriver(self):
        latest = requests.get('https://chromedriver.storage.googleapis.com/LATEST_RELEASE_{}'.format(self.chrome_version))
        wget.download('https://chromedriver.storage.googleapis.com/{}/chromedriver_win32.zip'.format(latest.text), '{}\\chromedriver_win32.zip'.format(self.chromedriver_path))

        with zipfile.ZipFile('{}\\chromedriver_win32.zip'.format(self.chromedriver_path), 'r') as zip_ref:
            zip_ref.extractall(self.chromedriver_path)
          
        os.remove(f'{self.chromedriver_path}\\chromedriver_win32.zip')
            
        if os.path.exists(f'{self.chromedriver_path}\\chromedriver.exe'):
            return True
        else: 
            return False
        
    def setup_browser(self):
        updated = True
        browser = None

        while True:
            if not os.path.exists(f'{self.chromedriver_path}\\chromedriver.exe'):
                if self.log:
                    self.log.info('chromedriver.py não encontrado. Fazendo download...')
                else:
                    print('chromedriver.py não encontrado. Fazendo download...')
                downloaded = self.download_chromedriver()
            else: 
                downloaded = True
                
            if downloaded:
                self.chromedriver_version = self.get_chromedriver_version()
                
                if int(self.chromedriver_version) != int(self.chrome_version):
                    updated = False
                    
                    if self.log:
                        self.log.info('Atualizando chromedriver...')
                    else:
                        print('Atualizando chromedriver...')
                        
                    self.download_chromedriver()
                    self.chromedriver_version = self.get_chromedriver_version()
                    self.chrome_version = self.get_chrome_version()

                    updated = int(self.chromedriver_version) == int(self.chrome_version)
            else:
                continue

            if updated:
                self.options = webdriver.ChromeOptions()
                self.options.add_experimental_option("prefs", self.profile)

                if self.headless:
                   self.options.add_argument('--headless')
                
                if self.profile_folder:
                    if not os.path.exists(self.profile_folder):
                        os.makedirs(self.profile_folder)
                        
                    self.options.add_argument(f'user-data-dir={os.path.dirname(__file__)}/ChromeProfile')
                    
                self.options.add_argument('--start-maximized')
                
                self.browser = webdriver.Chrome('{}\\chromedriver.exe'.format(self.chromedriver_path), options=self.options)

                break
            else:
                sleep(3)
            continue
