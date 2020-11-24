import unittest
import time
from appium import webdriver
from selenium.webdriver.common.keys import Keys

class NotepadTests(unittest.TestCase):
    def test_edit(self):
        for i in range(1,5):
            desired_caps = {}
            desired_caps['app'] = r"C:\Program Files\tengyue\XXX.exe"
            self.driver = webdriver.Remote(
                command_executor='http://127.0.0.1:4723',
                desired_capabilities=desired_caps)
            self.driver.implicitly_wait(10)
            try:
                self.driver.find_element_by_name("重新检测").click()
                time.sleep(5)
            except:
                print("第"+str(i)+"执行成功")
                self.driver.quit()

if __name__ == "__main__":
    aa=NotepadTests()
    aa.test_edit()