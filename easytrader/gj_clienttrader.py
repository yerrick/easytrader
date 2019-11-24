# coding:utf8
from __future__ import division

import functools
import io
import os
import re
import tempfile
import time
import sys

import easyutils
import pandas as pd
import pywinauto
import pywinauto.clipboard
from . import exceptions
from . import helpers
from .yh_clienttrader import YHClientTrader
from .log import log


class GJClientTrader(YHClientTrader):
    @property
    def broker_type(self):
        return 'gj'

    def login(self, user, password, exe_path, comm_password=None, **kwargs):
        """
        登陆客户端
        :param user: 账号
        :param password: 明文密码
        :param exe_path: 客户端路径类似 r'C:\中国银河证券双子星3.2\Binarystar.exe', 默认 r'C:\中国银河证券双子星3.2\Binarystar.exe'
        :param comm_password: 通讯密码, 华泰需要，可不设
        :return:
        """
        try:
            self._app = pywinauto.Application().connect(path=self._run_exe_path(exe_path), timeout=1)
            try:
                tmp = self._app.window(title_re='网上股票交易系统.*').wait('visible',timeout=5)
                print('Reconnect',tmp.window_text())
            except Exception as e:
                print('Reconnect fail')
                print(e)
                self._app.kill()
                raise e

        except Exception:
            self._app = pywinauto.Application().start(exe_path)

            # wait login window ready
            while True:
                try:
                    self._app.top_window().Edit1.wait('ready')
                    print('Edit1 ready')
                    break
                except RuntimeError:
                    pass
            self._app.top_window().Edit1.type_keys(user)
            while True:
                try:
                    self._app.top_window().Edit2.wait('ready')
                    print('Edit2 ready')
                    break
                except RuntimeError:
                    pass
            self._app.top_window().Edit2.type_keys(password)
            edit3 = self._app.top_window().window(control_id=0x3eb)
            while True:
                try:
                    code = self._handle_verify_code()
                    print('verify code=',code)
                    edit3.type_keys(
                        code
                    )
                    time.sleep(0.5)
                    self._app.top_window()['确定(Y)'].click()
                    print("Check Login...")
                    # detect login is success or not
                    try:
                        self._app.window(title_re='用户登录.*').wait_not('exists',timeout=5)
                        print('Login window closed')
                        tmp = self._app.window(title_re='网上股票交易系统.*').wait('visible',timeout=5)
                        print('Open',tmp.window_text())
                        print("ReCheck Login...")
                        self._wait(2)
                        self._app.window(title_re='用户登录.*').wait_not('exists',timeout=5)
                        print('Login window closed -- Recheck Pass')                        
                        print("Login Success")
                        break
                    except Exception as e:
                        print("Login Fail")
                        print(e)
                        self._app.top_window()['确定'].click()
                        pass
                except Exception as e:
                    print("Exception,",e)
                    pass
        self._main = self._app.window(title='网上股票交易系统5.0')

    def _handle_verify_code(self):
        control = self._app.top_window().window(control_id=0x5db)
        control.click()
        time.sleep(0.2)
        file_path = tempfile.mktemp()+'.jpg'
        control.capture_as_image().save(file_path)
        time.sleep(0.2)
        vcode = helpers.recognize_verify_code(file_path, 'gj_client')
        return ''.join(re.findall('[a-zA-Z0-9]+', vcode))

    @property
    def balance(self):
        self._switch_left_menus(['查询[F4]', '资金股票'])
        retv={}
        retv['enable_balance'] = self._main.window(control_id=0x3f8).window_text()
        retv['total_balance'] = self._main.window(control_id=0x3f7).window_text()
        return [retv]

    @property
    def position(self):
        self._switch_left_menus(['查询[F4]', '资金股票'])
        return self._get_grid_data_fromfile(self._config.COMMON_GRID_CONTROL_ID)

    def _get_grid_data_fromfile(self, controlid):
        grid = self._main.window(
            control_id=controlid, class_name="CVirtualGridCtrl"
        )

        # ctrl+s 保存 grid 内容为 xls 文件
        grid.type_keys("^s")
        time.sleep(1)

        temp_path = tempfile.mktemp(suffix=".csv", prefix="easytrader_position_")
        self._app.window(title='另存为',control_id=0).Edit.set_edit_text(temp_path)

        # alt+s保存，alt+y替换已存在的文件
        self._app.top_window().type_keys("%{s}%{y}")
        time.sleep(1)
        return self._format_grid_data_fromfile(temp_path)

    def _format_grid_data_fromfile(self,filepath):
        df = pd.read_csv(filepath,
            delimiter="\t",
            dtype=self._config.GRID_DTYPE,
            na_filter=False,
            encoding='GBK'
        )
        return df.to_dict("records")


    def cancel_all_entrusts(self):
        self._refresh()
        self._switch_left_menus(['撤单[F3]'],1)
        total_len = len(self._get_grid_data_fromfile(self._config.COMMON_GRID_CONTROL_ID))
        if total_len==1:            
            print('%d Entrusts to Cancel'%total_len)
            self._main.window(control_id=0x7531).click()
            self._wait(1)
            self._handle_cancel_entrust_pop_dialog()
        elif total_len>1:            
            print('%d Entrusts to Cancel'%total_len)
            self._main.window(control_id=0x7531).click()
            self._wait(1)
            self._handle_cancel_entrust_pop_dialog()
        else:
            print('No Entrusts to Cancel')
        return

