import sys, json, requests, time, openpyxl, string, os, ctypes
from urllib3 import disable_warnings
from PyQt5.QtWidgets import QWidget,QTextEdit,QPushButton,QComboBox,QListView,QListWidget,QMessageBox,QApplication
from PyQt5.QtCore import QCoreApplication,Qt,QRect,QMetaObject,QThreadPool,pyqtSlot
from PyQt5.QtGui import QIcon
from concurrent.futures import ThreadPoolExecutor

disable_warnings()


class Ui_MainWindow(QWidget):
    with open('./city_code.json', 'r', encoding='utf-8') as f:
        city_code = json.load(f)['data']

    phone = None
    verify_code = None
    province_index = 0
    city_index = 0
    district_index = 0

    headers = {
        'authorization': '',
        'Cookie': 'SESSION=NjBmMGUxMDYtNDYxNy00ZGRhLWFhNzktNzNjZGFkZjhkNGRi'
    }
    rf_token = ''

    def __init__(self):
        super(Ui_MainWindow, self).__init__()
        self.setObjectName("MainWindow")
        self.resize(1128, 683)
        self.setObjectName("centralwidget")
        self.setWindowIcon(QIcon('sys.ico'))

        self.phoneEdit = QTextEdit(self)
        self.phoneEdit.setGeometry(QRect(20, 20, 221, 31))
        self.phoneEdit.setObjectName("phoneEdit")

        self.sendCodeButton = QPushButton(self)
        self.sendCodeButton.setGeometry(QRect(150, 20, 91, 32))
        self.sendCodeButton.setObjectName("sendCodeButton")
        self.sendCodeButton.clicked.connect(self.send_code)

        self.verifyCodeEdit = QTextEdit(self)
        self.verifyCodeEdit.setGeometry(QRect(270, 20, 101, 31))
        self.verifyCodeEdit.setObjectName("verifyCodeEdit")

        self.loginButton = QPushButton(self)
        self.loginButton.setGeometry(QRect(390, 20, 80, 32))
        self.loginButton.setObjectName("loginButton")
        self.loginButton.clicked.connect(self.login)

        self.provinceSelect = QComboBox(self)
        self.provinceSelect.setGeometry(QRect(20, 60, 110, 31))
        self.provinceSelect.view().setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.provinceSelect.setStyleSheet("QAbstractItemView::item {width: 40px;}")
        self.provinceSelect.setMaxVisibleItems(10)
        self.provinceSelect.setView(QListView())
        self.provinceSelect.setObjectName("provinceSelect")
        self.provinceSelect.addItem('全部')
        for province in self.city_code:
            self.provinceSelect.addItem(province['name'])

        self.citySelect = QComboBox(self)
        self.citySelect.setGeometry(QRect(150, 60, 90, 31))
        self.citySelect.setObjectName("citySelect")
        self.citySelect.addItem('全部')

        self.districtSelect = QComboBox(self)
        self.districtSelect.setGeometry(QRect(270, 60, 100, 31))
        self.districtSelect.setObjectName("districtSelect")
        self.districtSelect.addItem('全部')

        self.provinceSelect.currentIndexChanged[int].connect(self.province_val)
        self.citySelect.currentIndexChanged[int].connect(self.city_val)
        self.districtSelect.currentIndexChanged[int].connect(self.district_val)

        self.startButton = QPushButton(self)
        self.startButton.setGeometry(QRect(390, 60, 80, 32))
        self.startButton.setObjectName("startButton")
        self.startButton.clicked.connect(self.start_spider)

        self.listView = QListWidget(self)
        self.listView.setGeometry(QRect(20, 100, 1101, 551))
        self.listView.setObjectName("listView")

        self.retranslateUi()
        QMetaObject.connectSlotsByName(self)

    def province_val(self, index):
        self.citySelect.clear()
        self.citySelect.addItem('全部')
        self.city_index = 0
        self.districtSelect.clear()
        self.districtSelect.addItem('全部')
        self.district_index = 0

        if index <= 0:
            self.province_index = 0
            return

        self.citySelect.view().setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.citySelect.setStyleSheet("QAbstractItemView::item {width: 40px;}")
        self.citySelect.setMaxVisibleItems(10)
        self.citySelect.setView(QListView())
        self.citySelect.setStyleSheet("QComboBox {combobox-popup: 0;}")
        for city in self.city_code[index-1]['children']:
            self.citySelect.addItem(city['name'])

        self.province_index = index

    def city_val(self, index):
        self.districtSelect.clear()
        self.districtSelect.addItem('全部')
        self.district_index = 0
        if index <= 0:
            self.city_index = 0
            return

        self.districtSelect.view().setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.districtSelect.setStyleSheet("QAbstractItemView::item {width: 40px;}")
        self.districtSelect.setMaxVisibleItems(10)
        self.districtSelect.setView(QListView())
        self.districtSelect.setStyleSheet("QComboBox {combobox-popup: 0;}")
        for district in self.city_code[self.province_index-1]['children'][index-1]['children']:
            self.districtSelect.addItem(district['name'])

        self.city_index = index

    def district_val(self, index):
        if index <= 0:
            self.district_index = 0
            return

        self.district_index = index

    def send_code(self):
        phone = self.phoneEdit.toPlainText().strip()
        if len(phone) != 11:
            QMessageBox.warning(self, "输入异常", "请检查手机号是否正确", QMessageBox.Yes)
        else:
            url = 'https://brm.zhoupudata.com/api/oauth/send-mobile-code'
            headers = {
                'Content-Type': 'application/json',
                'Cookie': 'SESSION=NjBmMGUxMDYtNDYxNy00ZGRhLWFhNzktNzNjZGFkZjhkNGRi',
                'User-Agent': 'okhttp/3.12.12'
            }
            data = {
                "mobile": phone,
                "devicetoken": "00000000-6a74-6210-0000-00002bc63564"
            }
            response = requests.post(url, headers=headers, json=data, verify=False)
            result = response.json()
            if result.get('msg', '') == '成功':
                self.phone = phone
                QMessageBox.information(self, "成功", "发送验证码成功", QMessageBox.Yes)
            elif result.get('message'):
                QMessageBox.warning(self, "异常", result['message'], QMessageBox.Yes)
            else:
                QMessageBox.critical(self, "失败", '请等待60秒后再次发送', QMessageBox.Yes)

    def login(self):
        verify_code = self.verifyCodeEdit.toPlainText().strip()
        if self.phone and verify_code:
            try:
                url = 'https://brm.zhoupudata.com/api/oauth/app-login'
                headers = {
                    'Content-Type': 'application/json',
                    'Cookie': 'SESSION=NjBmMGUxMDYtNDYxNy00ZGRhLWFhNzktNzNjZGFkZjhkNGRi',
                    'User-Agent': 'okhttp/3.12.12'
                }
                data = {
                    "mobile": self.phone,
                    "deviceToken": "00000000-6a74-6210-0000-00002bc63564",
                    "smsCode": verify_code
                }
                response = requests.post(url, headers=headers, json=data, verify=False)

                self.headers['authorization'] = response.json()['data']['accessToken']
                self.rf_token = response.json()['data']['refreshToken']

                QMessageBox.information(self, "成功", "登陆成功", QMessageBox.Yes)
                self.listView.addItem(f'用户: {self.phone} 已成功登陆')
                self.phoneEdit.clear()
                self.verifyCodeEdit.clear()

            except:
                QMessageBox.critical(self, "失败", '验证码输入错误, 请重新输入', QMessageBox.Yes)
        else:
            QMessageBox.warning(self, "异常", '请输入手机号和验证码', QMessageBox.Yes)

    def start_spider(self):
        if not self.headers['authorization'] or not self.rf_token:
            QMessageBox.critical(self, "失败", '未登陆, 请先登录', QMessageBox.Yes)
            return

        area_final = {}
        if self.province_index != 0:
            province_dict = self.city_code[self.province_index-1]
            area_final['province'] = province_dict['name']
            if self.city_index != 0:
                city_dict = province_dict['children'][self.city_index-1]
                area_final['city'] = city_dict['name']
                if self.district_index != 0:
                    area_final['district'] = city_dict['children'][self.district_index-1]['name']

        self.thread = QThreadPool()
        self.thread.start(lambda: self.get_data(area_final))

        self.loginButton.setEnabled(False)
        self.startButton.setEnabled(False)

    def retranslateUi(self):
        _translate = QCoreApplication.translate
        self.setWindowTitle(_translate("MainWindow", "舟谱数据导出工具"))
        self.phoneEdit.setHtml(_translate("MainWindow", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'.AppleSystemUIFont\'; font-size:13pt; font-weight:400; font-style:normal;\">\n"
"<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p></body></html>"))
        self.phoneEdit.setPlaceholderText(_translate("MainWindow", "请输入手机号"))
        self.sendCodeButton.setText(_translate("MainWindow", "发送验证码"))
        self.verifyCodeEdit.setHtml(_translate("MainWindow", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'.AppleSystemUIFont\'; font-size:13pt; font-weight:400; font-style:normal;\">\n"
"<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p></body></html>"))
        self.verifyCodeEdit.setPlaceholderText(_translate("MainWindow", "请输入验证码"))
        self.loginButton.setText(_translate("MainWindow", "登陆"))
        self.startButton.setText(_translate("MainWindow", "开始运行"))

        self.provinceSelect.setStyleSheet("QComboBox {combobox-popup: 0;}")

    @pyqtSlot()
    def get_data(self, area_dict):
        area_str = f'{area_dict.get("province", "")}{area_dict.get("city", "")}{area_dict.get("district", "")}'
        self.listView.addItem(f'开始抓取: {area_str}')

        filename = './template.xlsx'
        workbook = openpyxl.load_workbook(filename)
        worksheet = workbook.active

        url = 'https://brm.zhoupudata.com/api/high-sea/list'
        page, max_page = 1, 1

        while page <= max_page:
            t = ThreadPoolExecutor(max_workers=5)
            params = {
                'current': page,
                'size': 20,
                'orderRule': 1
            }
            params.update(area_dict)
            try:
                response = requests.get(url, params=params, headers=self.headers, verify=False, proxies={'http': None, 'https': None})
                if response.json()['msg'] == "登录已过期":
                    self.refresh_token()
                    continue
            except:
                time.sleep(3)
                continue

            data = response.json()['data']

            business_ids = [item['businessId'] for item in data['records']]

            # result = []
            # for b_id in business_ids:
            #     b_data = self.get_b_data(b_id)
            #     result.append(b_data)

            tasks = [t.submit(self.get_b_data, idx, b_id) for idx, b_id in enumerate(business_ids)]

            result = [task.result() for task in tasks]
            result.sort(key=lambda x: x[0])

            [worksheet.append(row) for item in result for row in item[1]]
            workbook.save(f'./{area_str}.xlsx')

            max_page = data['pages']

            self.listView.clear()
            self.listView.addItem(f'列表页第 {page} 页保存完毕, 共 {max_page} 页')

            page += 1

        self.listView.addItem(f'数据已全部保存完毕, 正在开始合并单元格...')
        type_list = []
        i = 2
        while True:
            r = worksheet.cell(i, 2).value
            if r:
                type_list.append(r)
            else:
                break
            i += 1

        for col in string.ascii_uppercase[:string.ascii_uppercase.index('S') + 1]:
            self.listView.addItem(f'正在合并{col}列中...')
            # 判断合并单元格的始末位置
            s = 0
            e = 0
            flag = type_list[0]
            for i in range(len(type_list)):
                if type_list[i] != flag:
                    flag = type_list[i]
                    e = i - 1
                    if e >= s:
                        worksheet.merge_cells(col + str(s + 2) + f":{col}" + str(e + 2))
                        s = e + 1
                if i == len(type_list) - 1:
                    e = i
                    worksheet.merge_cells(col + str(s + 2) + f":{col}" + str(e + 2))

        workbook.save(f'./{area_str}.xlsx')
        workbook.close()

        self.listView.addItem(f'合并单元格完成, 程序结束. Excel路径: {os.path.dirname(__file__)}/{area_str}.xlsx')

        self.loginButton.setEnabled(True)
        self.startButton.setEnabled(True)

    def get_b_data(self, idx, b_id):
        b_data = self.b_detail(b_id)
        visit_data = self.get_visit(b_id)

        final_b_data = []
        for item in visit_data:
            new_b_data = b_data.copy()
            new_b_data.extend(item)
            final_b_data.append(new_b_data)

        self.listView.addItem(f'{b_data[1]} 已请求完成. 负责人: {b_data[2]}. 公司地址: {b_data[6]}')
        return idx, final_b_data

    def b_detail(self, b_id):
        url = f'https://brm.zhoupudata.com/api/business/queryDetail?businessId={b_id}'
        try:
            response = requests.get(url, headers=self.headers, verify=False, proxies={'http': None, 'https': None})
            if response.json()['msg'] == "登录已过期":
                self.refresh_token()
                return self.b_detail(b_id)
        except:
            time.sleep(3)
            return self.b_detail(b_id)

        data = response.json()['data']
        district = f"{data['province']}{data['city']}{data['district']}"      # 区
        b_name = data['businessName']       # 公司名称
        owner = data['ownerName']     # 负责人
        last_follow_date = data['lastFollowedAt']       # 最后跟进时间
        customer = data['contactName']      # 客户姓名
        customer_phone = data['mobile']     # 客户电话
        # 客户地址
        customer_addr = f"{district}{data['address']}"
        industry = ', '.join([item['tag'] for item in data['industries']]) if data.get('industries') else '无'      # 客户行业
        brand = ', '.join([item['tag'] for item in data['brands']]) if data.get('brands') else '无'     # 代理品牌
        sales_man_num = data['salesmanNum'] if data.get('salesmanNum') else '无'     # 业务员数量
        sale_amount = data['annualSaleAmount'] if data.get('annualSaleAmount') else '无'     # 年销售额
        cur_p_name = data['curSys'] if data.get('curSys') else '无'      # 竞品名称
        expire_time = data['curSysExpireTime'] if data.get('curSysExpireTime') else '无'      # 到期日期
        cur_virtue = data['curSysVirtue'] if data.get('curSysVirtue') else '无'      # 竞品优势
        remark = '无'
        create_time = data['createdAt'] if data.get('createdAt') else '无'       # 创建时间
        creator = data['creatorName'] if data.get('creatorName') else '无'       # 创建人
        last_assign_time = data['businessAssignedAt'] if data.get('businessAssignedAt') else '无'    # 最新分配时间
        assigner = data['assignerName'] if data.get('assignerName') else '无'    # 分配人

        return [
            district, b_name.strip(), owner, last_follow_date, customer,
            customer_phone, customer_addr.strip(), industry, brand,
            sales_man_num, sale_amount, cur_p_name, expire_time,
            cur_virtue, remark, create_time, creator, last_assign_time, assigner
        ]

    def get_visit(self, b_id):
        url = f'https://brm.zhoupudata.com/api/visit/queryByBusiness'

        page, max_page = 1, 1
        final_data = []
        while page <= max_page:
            params = {
                'current': page,
                'size': 20,
                'businessId': b_id
            }
            try:
                response = requests.get(url, params=params, headers=self.headers, verify=False, proxies={'http': None, 'https': None})
                if response.json()['msg'] == "登录已过期":
                    self.refresh_token()
                    continue
            except:
                time.sleep(3)
                continue

            data = response.json()['data']
            for item in data['records']:
                visit_data = item['visits'][0]

                visitor = item['user']['realName']     # 跟进人
                post = item['user']['post']           # 跟进人类型
                visit_time = visit_data['day']    # 跟进日期
                visit_count = visit_data['businessVisitCount']    # 跟进次数
                intent = visit_data['intent']       # 意向
                visit_kind = visit_data['visitKind']      # 拜访类型
                visit_type = visit_data['visitType']      # 拜访方式
                product = visit_data['product']     # 跟进产品
                result = visit_data['visitResult'] if visit_data.get('visitResult') else ''  # 跟进结果
                detail = visit_data['detail'] if visit_data.get('detail') else ''       # 内容详述
                contact_name = visit_data['contactName'] if visit_data.get('contactName') else ''    # 客户对接人

                final_data.append([
                    visitor, post, visit_time, visit_count, intent,
                    visit_kind, visit_type, product, result, detail, contact_name
                ])

            time.sleep(1)
            page += 1
            max_page = data['pages']

        return final_data

    def refresh_token(self):
        url = 'https://brm.zhoupudata.com/api/oauth/refresh-token'
        headers = {
            'Content-Type': 'application/json',
            'Cookie': 'SESSION=NjBmMGUxMDYtNDYxNy00ZGRhLWFhNzktNzNjZGFkZjhkNGRi'
        }
        data = {"refreshToken": self.rf_token}

        response = requests.post(url, headers=headers, json=data, verify=False, proxies={'http': None, 'https': None})
        self.headers['authorization'] = response.json()['data']['accessToken']
        self.rf_token = response.json()['data']['refreshToken']


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ui = Ui_MainWindow()
    appid = 'app.CRM.project.version'
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(appid)
    ui.show()
    sys.exit(app.exec_())
