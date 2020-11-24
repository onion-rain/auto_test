import unittest

from appium import webdriver
from selenium.webdriver.common.keys import Keys


class NotepadTests(unittest.TestCase):

    @classmethod
    def setUpClass(self):
        desired_caps = {}
        desired_caps['app'] = r"C:\Windows\System32\notepad.exe"
        self.driver = webdriver.Remote(
            command_executor='http://127.0.0.1:4723',
            desired_capabilities=desired_caps)

    @classmethod
    def tearDownClass(self):
        self.driver.quit()

    def test_edit(self):
        self.driver.find_element_by_name("文本编辑器").send_keys("polyv")
        self.driver.find_element_by_name("文件(F)").click()
        self.driver.find_element_by_xpath(
            '//MenuItem[starts-with(@Name, "保存(S)")]').click()
        self.driver.find_element_by_xpath(
            '//Pane[starts-with(@ClassName, "Address Band Root")]').find_element_by_xpath(
            '//ProgressBar[starts-with(@ClassName, "msctls_progress32")]').click()
        self.driver.find_element_by_xpath(
            '//Edit[starts-with(@Name, "地址")]').send_keys(
            r"D:\test" + Keys.ENTER)
        self.driver.find_element_by_accessibility_id(
            'FileNameControlHost').send_keys("note_test.txt")
        self.driver.find_element_by_name("保存(S)").click()
        self.driver.find_element_by_name("关闭").click()


if __name__ == "__main__":
    suite = unittest.TestLoader().loadTestsFromTestCase(NotepadTests)
    unittest.TextTestRunner(verbosity=2).run(suite)
