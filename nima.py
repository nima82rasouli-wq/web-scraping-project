import selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import time
import pandas as pd

# مشخصات ورود
USERNAME = "xcxxx"
PASSWORD = "xzxxz8"

# راه‌اندازی مرورگر
service = Service("chromedriver.exe")
options = Options()
options.add_argument("--start-maximized")
driver = webdriver.Chrome(service=service, options=options)
wait = WebDriverWait(driver, 20)

# رفتن به صفحه لاگین
driver.get("https://mmai.atu.ac.ir/contacts?_action=loginForm")

# ورود اطلاعات
wait.until(EC.presence_of_element_located((By.NAME, "username"))).send_keys(USERNAME)
driver.find_element(By.NAME, "_password").send_keys(PASSWORD)
driver.find_element(By.XPATH, "//button[@type='submit']").click()

# ورود به نقش "دبیر علمی"
wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "نقش های کاربر"))).click()
wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "دبیر علمی"))).click()

time.sleep(2)

# کلیک روی "مقالات پذیرفته شده"
wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(),'مقالات پذیرفته شده')]"))).click()

time.sleep(3)

results = []

# فرض می‌کنیم فقط یک صفحه است؛ اگر چند صفحه هست باید pagination هم اضافه کنیم
article_links = driver.find_elements(By.XPATH, "//a[contains(@href, 'javascript:prepareToAccept')]")

for i in range(len(article_links)):
    try:
        # دوباره گرفتن لینک‌ها
        article_links = driver.find_elements(By.XPATH, "//a[contains(@href, 'javascript:prepareToAccept')]")
        link = article_links[i]
        article_code = link.text.strip()

        driver.execute_script("arguments[0].scrollIntoView(true);", link)
        link.click()
        time.sleep(2)

        # کلیک روی تب
        tab = wait.until(EC.element_to_be_clickable((By.ID, "rvr-tab")))
        tab.click()
        time.sleep(2)

        # کلیک روی آیکن توضیحات ویراستار
        icon_button = wait.until(EC.element_to_be_clickable((
            By.XPATH,
            "//legend[contains(text(), 'توضیحات ویراستار ادبی')]/ancestor::fieldset[1]//table//a[.//i[contains(@class, 'fa-file-alt')]]"
        )))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", icon_button)
        time.sleep(1)
        icon_button.click()
        time.sleep(2)

        # گرفتن متن نهایی داور
        try:
            modal_content = wait.until(EC.presence_of_element_located((
                By.XPATH, "//div[contains(@class, 'modal-body')]"
            )))
            final_review_text = modal_content.text.strip()
        except:
            final_review_text = "نظر نهایی داور یافت نشد"

        results.append({
            "کد مقاله": article_code,
            "نظر نهایی داور": final_review_text
        })

        print(f"✅ کد مقاله: {article_code} | نظر نهایی داور: {final_review_text}")

        # برگشت به لیست مقالات
        driver.get("https://mmai.atu.ac.ir/editor?_action=final_accept")
        time.sleep(2)

    except Exception as e:
        results.append({
            "کد مقاله": article_code if 'article_code' in locals() else "نامشخص",
            "نظر نهایی داور": f"خطا: {str(e)}"
        })
        driver.get("https://mmai.atu.ac.ir/editor?_action=final_accept")
        time.sleep(2)

# ذخیره خروجی
df = pd.DataFrame(results)
df.to_excel("reviewer_final_comments.xlsx", index=False)
driver.quit()
