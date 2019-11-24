# coding:utf8
from __future__ import division

import functools
import io
import os
import re
import tempfile
import time
import sys
import numpy as np
import win32api

import easyutils
# import pandas as pd
import pywinauto
import pywinauto.clipboard
from . import exceptions
from . import helpers
from .yh_clienttrader import YHClientTrader
from .log import log
from .clienttrader import ClientTrader


class GJClientTraderV2(ClientTrader):
    @property
    def broker_type(self):
        return 'gj_v2'

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
            self._app = pywinauto.Application(backend='uia').connect(path=exe_path, timeout=1)
            try:
                tmp = self._app.window(title_re='%s.*'%self._config.TITLE).wait('visible',timeout=5)
                print('Reconnect',tmp.window_text())
            except Exception as e:
                print('Reconnect fail')
                print(e)
                self._app.kill()
                raise e

        except Exception:
            self._app = pywinauto.Application(backend='win32').start(exe_path)

            # wait login window ready
            time.sleep(2)
            self._topw = self._app.window(control_id=0,title_re='.*国金.*')
            self._topw.child_window(control_id=int('16E',16)).click()
            self._topw.set_focus()

            while True:
                try:
                    self._topw.Edit1.wait('ready')
                    print('Edit1 ready')
                    break
                except RuntimeError:
                    pass
            self._topw.Edit1.click()
            for k in user:
                win32api.keybd_event(ord(k),0,0,0)

            while True:
                try:
                    self._topw.Edit2.wait('ready')
                    print('Edit2 ready')
                    break
                except RuntimeError:
                    pass
            self._topw.Edit2.click()
            for k in password:
                win32api.keybd_event(ord(k),0,0,0)

            # edit3 = self._topw.Edit4.wrapper_object().rectangle()
            while True:
                try:
                    self._topw = self._app.top_window()
                    code = self._handle_verify_code()
                    print('verify code=',code)
                    self._topw.Edit4.click()
                    for k in code:
                        win32api.keybd_event(ord(k),0,0,0)
                    time.sleep(0.5)
                    self._app.top_window().child_window(control_id=int('1',16)).click()
                    print("Check Login...")

                    # detect login is success or not
                    try:
                        self._app.top_window().child_window(title_re='脱机运行.*').wait_not('exists',timeout=5)
                        print('Login window closed')
                        tmp = self._app.window(title_re='国金.*').wait('visible',timeout=5)
                        print('Open',tmp.window_text())
                        print("ReCheck Login...")
                        self._wait(2)
                        self._app.top_window().child_window(title_re='脱机运行.*').wait_not('exists',timeout=5)
                        print('Login window closed -- Recheck Pass')                        
                        print("Login Success")
                        break
                    except Exception as e:
                        print("Login Fail")
                        print(e)
                        # self._app.top_window()['确定'].click()
                        pass
                except Exception as e:
                    print("Exception,",e)
                    pass
        self._main = self._app.window(title_re='国金太阳至强版.*')

    def _handle_verify_code(self):
        editpos = self._topw.Edit4.wrapper_object().rectangle()
        import PIL
        from PIL import ImageGrab
        import tempfile,re
        from easytrader import helpers
        img = ImageGrab.grab(bbox=(editpos.right+2, editpos.top-5, editpos.right+78, editpos.bottom+5))
        def binary_pixel(x):
            if x>=218 and x<=239:
                return 255
            else:
                return 0

        iimg = img.point(binary_pixel)
        ti = np.array(iimg)
        ti = np.min(ti,axis=2)
        iimggray = PIL.Image.fromarray(ti)
        file_path = tempfile.mktemp()+'.tif'
        iimggray.save(file_path)
        time.sleep(0.2)

        vcode = helpers.recognize_verify_code(file_path, 'gj_client_v2')
        return vcode

    @property
    def balance(self):
        self._switch_left_menus(['查询[F4]', '资金股票'])
        retv={}
        retv['enable_balance'] = self._main.window(control_id=0x3f8).window_text()
        retv['total_balance'] = self._main.window(control_id=0x3f7).window_text()
        return [retv]

    @property
    def position(self):
        self._main.child_window(title='持仓',control_type='Button').click()
        self._get_grid_data(self._config.COMMON_GRID_CONTROL_ID)

    def cancel_all_entrusts(self):
        self._refresh()
        self._switch_left_menus(['撤单[F3]'],1)
        total_len = len(self._get_grid_data(self._config.COMMON_GRID_CONTROL_ID))
        if total_len==1:            
            print('%d Entrusts to Cancel'%total_len)
            self._main.window(control_id=0x7531).click()
            self._wait(1)
            self._handle_cancel_entrust_pop_dialog()
        elif total_len>1:            
            print('%d Entrusts to Cancel'%total_len)
            self._main.window(control_id=0x7531).click()
            self._wait(1)
        else:
            print('No Entrusts to Cancel')
        return

