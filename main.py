import os
import shutil
import time
import win32clipboard
import csv

from appium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

data_root = "data\\"
ledock_path = r"D:/Ledock.win32/LeDock.exe"
output_path = "output.csv"
box_path = "data\\box.txt"

def check_dir(dir):
    if not os.path.exists(dir):
        os.makedirs(dir)

def check_path(path):
    if not os.path.exists(path):
        raise Exception('path: "{}"" is not exist'.format(path))

def file_preprocess():
    # 1
    dirs = os.listdir(data_root)
    new_roots = []
    for file_name in dirs:
        if file_name.endswith(".mol2"):
            new_root = data_root+file_name[:-5]+"\\"
            check_dir(new_root)
            # shutil.move(data_root+file_name, new_root+file_name)
            shutil.copy(data_root+file_name, new_root+file_name)
            new_roots.append(new_root)
    # 2
    new_absolute_roots = []
    pdb_name = None
    for file_name in dirs:
        if file_name.endswith(".pdb"):
            pdb_name = file_name
    if pdb_name is None:
        raise Exception('cannot find ".pdb" file')
    for new_root in new_roots:
        shutil.copy(data_root+pdb_name, new_root+pdb_name)
        new_absolute_roots.append(os.getcwd()+"\\"+new_root)
    return pdb_name, new_absolute_roots

def change_dock(path, box_str):
    check_path(path)
    with open(path, "r", encoding="utf-8") as dock:
        dock.seek(0)
        dock_str = dock.read()
    with open(path, "w", encoding="utf-8") as dock:
        dock.write(dock_str.replace("xmin xmax\nymin ymax\nzmin zmax", box_str))

def extract_result(str):
    feature = "REMARK Cluster   1 of Poses:  1 Score: "
    idx_start = str.find("\n")
    idx_end = str.find("\n", idx_start+1)
    return str[idx_start+len(feature):idx_end]

def set_clipboard(str):
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, str)
    win32clipboard.CloseClipboard()

def data_process(absolute_root, pdb_name, box_str, ledock_path):
    pdb_path = absolute_root + pdb_name
    mol2_path = absolute_root + absolute_root.split("\\")[-2] + ".mol2"
    desired_caps = {}
    desired_caps['app'] = ledock_path
    driver = webdriver.Remote(
        command_executor='http://127.0.0.1:4723',
        desired_capabilities=desired_caps)
    
    driver.find_elements_by_name("Choose ...")[0].click()
    time.sleep(1)
    set_clipboard(pdb_path)
    driver.find_elements_by_name("文件名(N):")[2].send_keys(Keys.CONTROL, 'v')
    driver.find_element_by_name("打开(O)").click()
    time.sleep(1)
    driver.find_element_by_name("Add Hydrogen").click()
    time.sleep(1) # 不加此延时则get_dock_str无法读取
    
    change_dock(absolute_root+"dock.in", box_str)

    driver.find_element_by_name("LeDock").click()
    driver.find_elements_by_name("Choose ...")[0].click()
    time.sleep(1)
    set_clipboard(mol2_path)
    driver.find_elements_by_name("文件名(N):")[2].send_keys(Keys.CONTROL, 'v')
    driver.find_element_by_name("打开(O)").click()
    time.sleep(1)
    driver.find_element_by_name("Start docking").click()
    WebDriverWait(driver, 1e6).until(EC.element_to_be_clickable((By.NAME, "Check docking results"))).click()
    out_str = driver.find_elements_by_xpath('//Edit[starts-with(@Name, "")]')[0].text

    driver.quit()
    
    return extract_result(out_str)

def read_box(path):
    check_path(path)
    with open(path, "r", encoding="utf-8") as box:
        box_str = box.read()
    return box_str

def write2csv(path, str):
    with open(path, 'a+', encoding='utf-8', newline='') as f:
        # csv_writer = csv.writer(f)
        # csv_writer.writerow(str)
        f.write(str+',')
    print()

if __name__ == "__main__":
    if os.path.exists(output_path):
        os.remove(output_path)
    pdb_name, new_absolute_roots = file_preprocess()
    box_str = read_box(box_path)
    for new_absolute_root in new_absolute_roots:
        result = data_process(new_absolute_root, pdb_name, box_str, ledock_path)
        write2csv(output_path, result)
    print("done")
