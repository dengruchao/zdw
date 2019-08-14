# coding: utf-8

import requests
from PIL import Image
from selenium import webdriver
import time
import traceback
import re
import xlrd
import xlwt
import json
import chardet
import base64
import os
import pdfplumber
import pickle
import shutil
import copy


import sys
reload(sys)
sys.setdefaultencoding('utf-8')


class ZhongDengWang:

    def __init__(self, parse_pdf=True, auto_captcha=False):
        self.captcha_cache = None
        self.browser = None

        self.data_list = []
        self.name_list = []

        self.pdf_dir = os.path.join(os.getcwd(), 'pdf')
        self.headless = None
        self.parse_pdf = parse_pdf
        self.auto_captcha = auto_captcha

    def browser_init(self, headless=True):
        self.headless = headless
        if os.path.exists(self.pdf_dir):
            shutil.rmtree(self.pdf_dir)
        os.mkdir(self.pdf_dir)

        options = webdriver.ChromeOptions()

        prefs = {
            'profile.default_content_settings.popups': 0,
            'download.default_directory': self.pdf_dir,
            'download.prompt_for_download': False,
            "plugins.always_open_pdf_externally": True,
        }
        options.add_experimental_option('prefs', prefs)

        if headless:
            options.add_argument('--no-sandbox')  # 解决DevToolsActivePort文件不存在的报错
            options.add_argument('window-size=1366x768')  # 指定浏览器分辨率
            options.add_argument('--disable-gpu')  # 谷歌文档提到需要加上这个属性来规避bug
            options.add_argument('--hide-scrollbars')  # 隐藏滚动条, 应对一些特殊页面
            # options.add_argument('blink-settings=imagesEnabled=false')  # 不加载图片, 提升速度
            options.add_argument('--headless')  # 浏览器不提供可视化页面. linux下如果系统不支持可视化不加
            options.add_argument('log-level=3')

        self.browser = webdriver.Chrome(chrome_options=options, executable_path='chromedriver_win32_75/chromedriver.exe')

        self.browser.maximize_window()
        self.browser.minimize_window()

    def enable_download_headless(self):
        self.browser.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
        params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': self.pdf_dir}}
        command_result = self.browser.execute("send_command", params)

    def _get_captcha(self, captcha_xpath='//*[@id="imgId"]', iframe_loc={'x': 0, 'y': 0}):
        try:
            self.browser.save_screenshot('screenshot.png')
            self.browser.minimize_window()
            captcha_elm = self.browser.find_element_by_xpath(captcha_xpath)
            captcha_loc = captcha_elm.location
            captcha_size = captcha_elm.size
            left = int(captcha_loc['x']) + int(iframe_loc['x'])
            top = int(captcha_loc['y']) + int(iframe_loc['y'])
            right = left + captcha_size['width']
            bottom = top + captcha_size['height']
            captcha_rangle = (left, top, right, bottom)
            screenshot = Image.open('screenshot.png')
            captcha = screenshot.crop(captcha_rangle)
            captcha.save('captcha.png')

            print u'正在识别验证码...',
            captcha_input = self.browser.find_element_by_xpath('//*[@id="validateCode"]')
            if self.auto_captcha:
                url = 'http://op.juhe.cn/vercode/index'
                data2 = {
                    'key': '57e595193b88b3ae027390fb95bdd858',
                    'codeType': '1004',
                    'image': open('captcha.png', 'rb'),
                    'base64Str': base64.b64encode(open('captcha.png', 'rb').read())
                }
                response = requests.post(url, data=data2)
                response_js = json.loads(response.content)
                if response_js['error_code'] != 0:
                    print response_js['reason']
                    self.browser.quit()
                    raw_input(u'请退出程序并重新运行'.encode('gbk'))
                captcha_text = response_js['result']
            else:
                captcha.show()
                captcha_text = raw_input()
            print captcha_text
            captcha_input.send_keys(captcha_text)

            return captcha_text
        except:
            traceback.print_exc()
            self.browser.quit()

    def login(self, user, pswd):
        index_url = 'https://www.zhongdengwang.org.cn/zhongdeng/index.shtml'

        try:
            print u'正在登录...'
            self.browser.get(index_url)
            time.sleep(1)
            iframe = self.browser.find_element_by_xpath('/html/body/div[2]/div[1]/div[2]/iframe')
            iframe_loc = iframe.location

            self.browser.switch_to.frame(iframe)
            time.sleep(1)

            username = self.browser.find_element_by_xpath('//*[@id="userCode"]')
            username.clear()
            username.send_keys(user)
            # password = self.browser.find_element_by_xpath('//*[@id="password"]')
            # password.clear()
            # password.send_keys('testtest03')
            js = "$('#%s').val('%s')" % ('password', pswd)
            self.browser.execute_script(js)

            self._get_captcha(iframe_loc=iframe_loc)

            submit = self.browser.find_element_by_xpath('//*[@id="login_btn"]/img')
            submit.click()
            time.sleep(1)

            if self.browser.current_url != index_url:
                print u'登录成功'
            else:
                print u'账号或密码错误'
                self.browser.quit()
                raw_input(u'请退出程序并重新运行'.encode('gbk'))
        except Exception, e:
            traceback.print_exc()
            self.browser.quit()

    def query_by_name_list(self):
        try:
            for name_idx, name in enumerate(self.name_list):
                zulin_list = []
                if self.captcha_cache is None:
                    main = self.browser.find_element_by_xpath('/html/body/div[2]/table[6]/tbody/tr[1]/td[2]/table/tbody/tr[2]/td/table/tbody/tr[1]/td/a')
                    main.click()
                    byname = self.browser.find_element_by_xpath('/html/body/div[2]/table[1]/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[1]/td/a')
                    byname.click()
                name_input = self.browser.find_element_by_xpath('//*[@id="name"]')
                name_input.send_keys(name)
                if self.captcha_cache is None:
                    captcha_xpath = '/html/body/div[2]/table[1]/tbody/tr/td/table[2]/tbody/tr/td/table[1]/tbody/tr[4]/td[2]/img'
                    captcha_text = self._get_captcha(captcha_xpath)
                    self.captcha_cache = captcha_text
                else:
                    captcha_input = self.browser.find_element_by_xpath('//*[@id="validateCode"]')
                    captcha_input.send_keys(self.captcha_cache)
                print u'\n正在查询：%s...' % name,
                query = self.browser.find_element_by_xpath('//*[@id="query"]')
                query.click()
                # page_source = self.browser.page_source

                try:
                    zulin_reg = self.browser.find_element_by_xpath('//table[@id="summaryLinks"]/tbody/tr[2]/td[2]/a')
                except:
                    print u'验证码错误'
                    self.browser.quit()
                    raw_input(u'请退出程序并重新运行'.encode('gbk'))
                zulin_reg.click()

                # baibiao_list = self.browser.find_elements_by_xpath('//tr[@class="baibiao"]/td[7]/span')
                baibiao_list = self.browser.find_elements_by_xpath('//tr[@class="baibiao"]')
                print u'找到%d条记录' % (len(baibiao_list)-2)
                if len(baibiao_list) > 2:
                    for n, baibiao in enumerate(baibiao_list[1:-1]):
                        print u'正在解析第%d条记录...' % (n+1),
                        reg_date = baibiao.find_element_by_xpath('td[3]').text.split(' ')[0]
                        ter_date = baibiao.find_element_by_xpath('td[4]').text
                        reg_type = baibiao.find_element_by_xpath('td[5]').text
                        reg_name = baibiao.find_element_by_xpath('td[6]').text
                        # print reg_date, reg_name
                        if self.parse_pdf:
                            download = baibiao.find_element_by_xpath('td[7]/span')
                            download.click()
                            self.browser.minimize_window()
                            handles = self.browser.window_handles
                            self.browser.switch_to.window(handles[1+n])
                            self.browser.minimize_window()
                            if self.headless:
                                self.enable_download_headless()
                            pdf_a = self.browser.find_element_by_xpath('//*[@id="tab"]/tbody/tr[2]/td[2]/a')
                            regno = re.search("'(.+?)'", pdf_a.get_attribute('href')).group(1)
                            pdf_path = os.path.join(self.pdf_dir, regno + '_A.pdf')
                            if not os.path.exists(pdf_path):
                                pdf_a.click()
                            while not os.path.exists(pdf_path):
                                time.sleep(0.5)
                            time.sleep(0.5)
                            money, desc = self.read_pdf(pdf_path)
                            self.browser.switch_to.window(handles[0])
                        else:
                            money, desc = ['', '']
                        zulin_list.append([reg_date, ter_date, reg_type, reg_name, money, desc])

                        print reg_date, ter_date, reg_type, reg_name, money
                else:
                    zulin_list.append(['', '', '', '', '', ''])

                # if name_idx == 0:
                #     zulin_list.append(copy.deepcopy(zulin_list[-1]))

                self.data_list.append(zulin_list)

                handles = self.browser.window_handles
                if len(handles) > 1:
                    for handle in handles[1:]:
                        self.browser.switch_to.window(handle)
                        self.browser.close()
                    self.browser.switch_to.window(handles[0])

                self.browser.back()

        except Exception, e:
            traceback.print_exc()
            self.browser.quit()

    def save_data_as_pickle(self):
        with open('data_list.pk', 'wb') as f:
            pickle.dump(self.data_list, f)
            pickle.dump(self.name_list, f)
        # time.sleep(0.5)

    def save_data_as_excel(self, filename):
        if os.path.exists('data_list.pk'):
            with open('data_list.pk', 'rb') as f:
                last_data_list = pickle.load(f)
                last_name_list = pickle.load(f)
            is_first = False
        else:
            last_data_list = []
            last_name_list = []
            is_first = True
        self.save_data_as_pickle()

        wb = xlwt.Workbook(encoding='utf-8')
        ws = wb.add_sheet('Sheet1')
        ws.col(0).width = 10000
        ws.col(1).width = 20000
        ws.col(2).width = 25000

        style = xlwt.XFStyle()  # 初始化样式
        style.alignment.wrap = 1  # 自动换行

        alignment = xlwt.Alignment()
        alignment.vert = xlwt.Alignment.VERT_CENTER
        alignment.wrap = xlwt.Alignment.WRAP_AT_RIGHT
        style.alignment = alignment

        old_style = copy.deepcopy(style)
        font = xlwt.Font()  # 为样式创建字体
        font.bold = True
        style.font = font  # 设定样式
        style.alignment.horz = xlwt.Alignment.HORZ_CENTER
        ws.write(0, 0, u'查询主体名称', style)
        ws.write(0, 1, u'记录', style)
        ws.write(0, 2, u'记录描述', style)
        style = old_style

        style_yellow = copy.deepcopy(style)
        is_first = True
        if not is_first:
            ptn = xlwt.Pattern()
            ptn.pattern = xlwt.Pattern.SOLID_PATTERN
            ptn.pattern_fore_colour = 5
            style_yellow.pattern = ptn

        row = 1
        for i, name in enumerate(self.name_list):
            first_row = row
            col0 = name
            data = self.data_list[i]
            last_has = False
            if name in last_name_list:
                last_data = last_data_list.pop(last_name_list.index(name))
                last_name_list.remove(name)
                last_has = True
            else:
                last_data = []
                last_has = False
            for j, d in enumerate(data):
                if j > 0:
                    row += 1
                col1 = ' '.join(d[0:-1])
                col2 = d[-1]
                if d not in last_data:
                    ws.write(row, 1, col1, style_yellow)
                    ws.write(row, 2, col2, style_yellow)
                else:
                    ws.write(row, 1, col1, style)
                    ws.write(row, 2, col2, style)
                    last_data.remove(d)
            if not last_has:
                ws.write_merge(first_row, row, 0, 0, col0, style_yellow)
            else:
                ws.write_merge(first_row, row, 0, 0, col0, style)
            row += 1

        wb.save(filename)
        print u'\n数据已存储到%s' % filename

    def get_name_list(self):
        wb = xlrd.open_workbook(u'查询目录.xls')
        ws = wb.sheet_by_index(0)
        col_values = ws.col_values(0)
        for v in col_values[1:]:
            v = v.strip()
            if v == '':
                continue
            self.name_list.append(v)

    @staticmethod
    def read_pdf(path):
        pdf = pdfplumber.open(path)

        pages = pdf.pages
        pdf_text = ''
        for page in pages:
            pdf_text += page.extract_text()+'\n'
        match = re.search(u'租金总额\s*(.*?)租赁财产唯一标识码.*?租赁财产描述(.*?)租赁财产信息附件', pdf_text, re.S)
        money = match.group(1).strip()
        desc = match.group(2).strip()

        pdf.close()
        return money, desc

    def browser_quit(self):
        self.browser.quit()


if __name__ == '__main__':
    zdw = ZhongDengWang(parse_pdf=True, auto_captcha=False)

    zdw.browser_init(headless=True)
    user = raw_input(u'请输入账号: '.encode('gbk'))
    pswd = raw_input(u'请输入密码: '.encode('gbk'))
    zdw.login(user, pswd)
    zdw.get_name_list()
    zdw.query_by_name_list()
    zdw.browser_quit()

    # zdw.read_pdf(u'pdf\\04859604000580597100_A.pdf')
    filename = time.strftime('%Y%m%d%H%M%S') + '.xls'
    zdw.save_data_as_excel(filename)

    raw_input()



