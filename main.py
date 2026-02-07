import time
import os
import json  

# 修复网络报错
os.environ['WDM_SSL_VERIFY'] = '0'
os.environ['WDM_LOCAL'] = '1'

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from webdriver_manager.microsoft import EdgeChromiumDriverManager

# -------------------------- 读取配置文件函数 --------------------------
def load_config():
    """读取config.json配置文件，返回配置字典"""
    try:
        with open("config.json", "r", encoding="utf-8") as f:
            config = json.load(f)
        # 校验配置项是否完整
        required_keys = ["url", "student_number", "password"]
        for key in required_keys:
            if key not in config or config[key] in ["", "山理学子的学号", "山理学子的密码", "请替换为你的教务系统登录URL"]:
                raise ValueError(f"配置文件中 {key} 未正确填写！")
        # 拼接选课页面URL（基于登录URL推导，避免硬编码）
        config["target_url"] = config["url"].replace("xtgl/login_slogin.html", "xsxk/zzxkyzb_cxZzxkYzbIndex.html?gnmkdm=N253512&layout=default")
        return config
    except FileNotFoundError:
        print("未找到 config.json 文件！请复制 config.example.json 并修改为你的信息")
        exit(1)
    except ValueError as e:
        print(f" 配置文件错误：{e}")
        exit(1)
    except json.JSONDecodeError:
        print("config.json 格式错误！请检查JSON语法是否正确")
        exit(1)

# -------------------------- 初始化驱动函数 --------------------------
def init_driver():
    """初始化 Microsoft Edge 驱动（使用本地驱动文件）"""
    print("正在启动 Microsoft Edge (本地驱动模式)...")
    
    options = Options()
    options.add_argument("--start-maximized")
    options.add_argument("--ignore-certificate-errors")
    
    # 彻底禁用自动化标志，防止网站检测
    options.add_experimental_option("detach", True)
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    
    # 使用本地驱动文件路径 
    driver_path = os.path.join(os.getcwd(), "msedgedriver.exe")
    if not os.path.exists(driver_path):
        print(f"找不到驱动文件: {driver_path}")
        print("请检查 msedgedriver.exe 是否在该文件夹内")
        exit(1)
        
    service = Service(driver_path)
    
    driver = webdriver.Edge(service=service, options=options)
    return driver

# -------------------------- 原有函数：登录 + 进入选课 --------------------------
def login_jwglxt(driver, url, username, password):
    """登录教务系统"""
    try:
        print(f"正在访问: {url}")
        driver.get(url)
        wait = WebDriverWait(driver, 10)
        
        user_input = wait.until(EC.presence_of_element_located((By.ID, "yhm")))
        user_input.clear()
        user_input.send_keys(username)
        
        pwd_input = driver.find_element(By.ID, "mm")
        pwd_input.clear()
        pwd_input.send_keys(password)
        
        login_btn = driver.find_element(By.ID, "dl")
        login_btn.click()
        
        WebDriverWait(driver, 10).until(
            EC.any_of(EC.url_contains("index_initMenu"), EC.url_contains("xtgl/index_initMenu.html"))
        )
        print("登录成功！")
        return True
    except Exception as e:
        print(f"登录失败: {str(e)}")
        return False

def enter_course_selection(driver, target_url):
    """进入选课页面（接收动态target_url，不再硬编码）"""
    print("\n开始执行选课流程...")
    
    try:
        driver.get(target_url)
        WebDriverWait(driver, 10).until(EC.url_contains("jwglxt"))
        print("选课页面加载完成")
        time.sleep(2)

        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        if iframes:
            driver.switch_to.frame(iframes[-1])
            print("已切换到选课区域iframe")
        else:
            print("未找到选课区域iframe，可能影响后续操作")

    except Exception as e:
        print(f"进入选课页面出错: {str(e)}")

# -------------------------- 核心修改：主函数 --------------------------
if __name__ == "__main__":
    # 1. 先加载配置文件（核心：移除硬编码敏感信息）
    config = load_config()
    
    driver = None  # 初始化driver变量，避免后续异常时找不到
    try:
        # 2. 初始化驱动（自动下载，无需本地msedgedriver.exe）
        driver = init_driver()
        
        # 3. 登录（使用配置文件中的信息）
        login_success = login_jwglxt(
            driver=driver,
            url=config["url"],
            username=config["student_number"],
            password=config["password"]
        )
        
        # 4. 登录成功则进入选课页面
        if login_success:
            enter_course_selection(driver=driver, target_url=config["target_url"])
            input("\n✅ 流程执行完成，按回车键关闭浏览器...")
        
    # 5. 优化异常处理：不再静默pass，打印错误+确保浏览器退出
    except Exception as e:
        print(f"\n程序运行出错: {str(e)}")
    finally:
        # 无论是否出错，最终都关闭浏览器（避免残留进程）
        if driver is not None:
            print("\n正在关闭浏览器...")
            driver.quit()
            print("浏览器已关闭")