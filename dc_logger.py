import os
import sys
import csv
import time
import random
import datetime
import ConfigParser
import logging
from Queue import Queue

from serial import serialutil
from PyQt4.QtCore import Qt, QThread, QRect, QTimer, SIGNAL
from PyQt4 import QtGui

from dcload import DCLoad
from conf import DEFAULT_PORT, DEFAULT_BAUD, DEFAULT_TIMEOUT, DEFAULT_TIME,\
    DEFAULT_FILENAME, NEW_FILE_DIALOG_TEXT
from test_data import TEST_CONSTANTS_VALUES, TEST_INPUT_VALUES

log = logging.getLogger('dc_logger')

DEFAULT_DC_LOGGER_QUEUE_SIZE = 10
UNKNOWN_STATE = 0
LOCAL_STATE = 1
REMOTE_STATE = 2
TEST_MODE_STATUS = False#True
TEST_CONSTANT_IDX = ('cc', 'cv', 'cw', 'cr',)

DEFAULT_FILE_LOG_STRUCT = {
    'timestamp': None,
    'voltage': None,
    'power': None,
    'current': None,
    'mode': None,
    'total_seconds': None,
}


class DCLoggerWorker(QThread):
    def __init__(self, queue, parent=None):
        super(DCLoggerWorker, self).__init__(parent)
        self._dc_load_obj = DCLoad()
        self._com_port = None
        self._baud_rate = None
        self._dc_logger_state = UNKNOWN_STATE
        self._is_com_port_open = False        
        self._queue = queue
        self._test_constant_idx = None
        self._test_constant_values = TEST_CONSTANTS_VALUES

    def _emit_msg(self, error_msg):
        self.emit(SIGNAL('error_msg_posted'), error_msg)

    def _get_dispatch_method(self, in_val):
        mapper = {
            'set_remote_control': self._set_remote_control,
            'set_local_control': self._set_local_control,
            'turn_load_on': self._turn_load_on,
            'turn_load_off': self._turn_load_off,
            'get_input_data': self._read_input_values,
            'get_constants_values': self._get_constants_values,
            'set_constants_values': self._set_constants_values,
        }
        return mapper.get(in_val)
    
    def run(self):
        while 1:
            try:
                msg = self._queue.get()
                cmd = msg[0]
                val = msg[1] if len(msg) > 1 else None
                if TEST_MODE_STATUS:
                    self._test_dispatch_request(cmd, val)
                else:
                    self._dispatch_request(cmd, val)
            except:
                log.exception('Got unexpected exception @ request dispatcher')
            finally:
                self._queue.task_done()

    def _test_dispatch_request(self, cmd, val):
        log.info('Got payload - cmd: %s val: %s' % (cmd, val,))
        try:
            if not self._is_com_port_open:
                self._test_open_com_port()
            if self._dc_logger_state != REMOTE_STATE:
                self._set_remote_control()
        except:
            log.exception('Failed to dispatch request')
            return

        dispatch_method = self._get_dispatch_method(cmd)
        if dispatch_method is None:
            log.error('Empty cmd was issued')
            return

        dispatch_method() if val is None else dispatch_method(val)

    def _dispatch_request(self, cmd, val):
        try:
            if not self._is_com_port_open:
                self._open_com_port()
            if self._dc_logger_state != REMOTE_STATE:
                self._set_remote_control()
        except:
            return

        dispatch_method = self._get_dispatch_method(cmd)
        if dispatch_method is None:
            log.exception('Failed to dispatch request')
            log.error('Empty cmd was issued')
            return

        dispatch_method() if val is None else dispatch_method(val)

    def connect(self, com_port, baud_rate):
        self._com_port = com_port
        self._baud_rate = baud_rate

        if TEST_MODE_STATUS:
            self._test_open_com_port()
        else:
            self._open_com_port()

    def disconnect(self):
        if TEST_MODE_STATUS:
            self._test_close_com_port()
        else:
            self._close_com_port()

    def _test_open_com_port(self):
        self._emit_msg('Test com port %s with baud rate: %s is open' % (self._com_port, self._baud_rate))
        self._is_com_port_open = True
        self.emit(SIGNAL('com_port_state_changed'), True)

    def _test_close_com_port(self):
        self._emit_msg('Test com closed')
        self._is_com_port_open = False
        self.emit(SIGNAL('com_port_state_changed'), False)

    def _open_com_port(self):
        try:
            if self._dc_load_obj is None:
                self._dc_load_obj.Initialize(self._com_port, self._baud_rate, timeout=DEFAULT_TIMEOUT)
            else:
                self._dc_load_obj.connect(self._com_port, self._baud_rate)
            self._is_com_port_open = True
        except:
            self._emit_msg('Failed to open COM port %s' % (self._com_port))
            log.error('Failed to open COM: %s, %s' % (self._com_port, self._baud_rate))
            raise
            
        self.emit(SIGNAL('com_port_state_changed'), True)

    def _close_com_port(self):
        try:
            self._dc_load_obj.disconnect()
            self._is_com_port_open = False
        except:
            self._emit_msg('Failed to close COM port')
            log.exception('Failed to close COM')
            raise
            
        self.emit(SIGNAL('com_port_state_changed'), False)

    def _set_remote_control(self):
        if TEST_MODE_STATUS:
            log.info('Invoked _set_remote_control')
        try:
            if self._dc_logger_state != REMOTE_STATE:
                if not TEST_MODE_STATUS:
                    self._dc_load_obj.SetRemoteControl()
            self._dc_logger_state = REMOTE_STATE
            self._emit_msg('Control set to remote')
        except:
            self._emit_msg('Failed to set remote control')
            log.exception('Failed to set remote control')
            raise

    def _set_local_control(self):
        try:
            if self._dc_logger_state != LOCAL_STATE:
                if not TEST_MODE_STATUS:
                    self._dc_load_obj.SetLocalControl()
            self._dc_logger_state = LOCAL_STATE
        except:
            self._emit_msg('Failed to set local control')
            log.exception('Failed to set local control')
            raise
    
        self.emit(SIGNAL('control_state_changed'), 'local')
    
    def _turn_load_on(self):
        try:
            if not TEST_MODE_STATUS:
                self._dc_load_obj.TurnLoadOn()
            self._emit_msg('Load turned on')
        except:
            self._emit_msg('Failed to turn load on')
            log.exception('Failed to turn load on')
            raise
    
        self.emit(SIGNAL('load_state_changed'), 'on')

    def _turn_load_off(self):
        error_msg = 'Failed to turn load off'
        try:
            if not TEST_MODE_STATUS:
                self._dc_load_obj.TurnLoadOff()
            self._emit_msg('Load turned off')
        except:
            self._emit_msg(error_msg)
            log.exception(error_msg)
            raise

        self.emit(SIGNAL('load_state_changed'), 'off')
            
    def _read_input_values(self, data_receiver='display'):
        ret_val = None
        error_msg = 'Failed to obtain input values'
        try:
            rand_int = random.randint(0, 7)
            if TEST_MODE_STATUS:
                res = TEST_INPUT_VALUES[rand_int]
            else:
                res = self._dc_load_obj.GetInputValues()

            if len(res) >= 3:
                ret_val = {
                    'input_voltage': res[0],
                    'input_current': res[1],
                    'input_power': res[2],
                }
        except IndexError:
            # this exceptions can be encountered in testing mode
            pass
        except serialutil.SerialException:
            # do not log serialException's otherwise log file will get very large very fast
            pass
        except:
            self._emit_msg(error_msg)
            log.exception(error_msg)
            raise

        if data_receiver == 'file':
            self.emit(SIGNAL('file_input_data_available'), ret_val)
        else:
            self.emit(SIGNAL('display_input_data_available'), ret_val)


    def _get_constants_values(self):
        error_msg = 'Failed to obtain constants values'
        ret_val = {}
        const_getter_objects = {
            'cc': {
                'val': self._dc_load_obj.GetCCCurrent,
                'is_active': False,
            },
            'cv': {
                'val': self._dc_load_obj.GetCVVoltage,
                'is_active': False,
            },
            'cw': {
                'val': self._dc_load_obj.GetCWPower,
                'is_active': False,
            },
            'cr': {
                'val': self._dc_load_obj.GetCRResistance,
                'is_active': False,
            },
        }

        try:
            if TEST_MODE_STATUS:
                if self._test_constant_idx is None:
                    ret_val = self._test_constant_values[random.randint(0,3)]
                else:
                    ret_val = self._test_constant_values[self._test_constant_idx]
            else:
                for const_mode_name, const_mode_obj in const_getter_objects.items():
                    obj_method = const_mode_obj['val']
                    ret_val[const_mode_name] = {
                        'val': obj_method(),
                    }

                active_mode = self._dc_load_obj.GetMode()
                ret_val[active_mode]['is_active'] = True        
        except:
            self._emit_msg(error_msg)
            log.exception(error_msg)
            raise
 
        log.debug('Emitting: %s' % (str(ret_val)))
        self.emit(SIGNAL('constants_data_available'), ret_val)

    def _set_constants_values(self, in_val):
        const_mode, const_value = in_val
        error_msg = 'Failed to update "%s" value: %s' % (const_mode, const_value,)

        #XXX: move it elsewhere
        const_setter_objects = {
            'cc': self._dc_load_obj.SetCCCurrent,
            'cv': self._dc_load_obj.SetCVVoltage,
            'cw': self._dc_load_obj.SetCWPower,
            'cr': self._dc_load_obj.SetCRResistance,
            }

        const_getter_objects = {
            'cc': self._dc_load_obj.GetCCCurrent,
            'cv': self._dc_load_obj.GetCVVoltage,
            'cw': self._dc_load_obj.GetCWPower,
            'cr': self._dc_load_obj.GetCRResistance,
            }
        
        try:
            if not TEST_MODE_STATUS:
                const_setter_objects[const_mode](float(const_value))
                self._dc_load_obj.SetMode(const_mode)
                verification_mode = self._dc_load_obj.GetMode()
                verification_value = const_getter_objects[const_mode]()
            else:
                self._test_constant_idx = TEST_CONSTANT_IDX.index(const_mode)
                verification_value = 1.1
                verification_mode = 'cc'

            log.debug('Verification: %s - %s, in_val: %s' % (
                str(verification_mode), str(verification_value), str(in_val)
            ))
            if verification_mode != const_mode:
                raise

            if float(verification_value) != float(const_value):
                raise

            self._emit_msg('Mode %s activated, with value: %s' % (const_mode, const_value,))
        except Exception, e:
            self._emit_msg(error_msg)
            log.exception(error_msg)
            raise

        # inform that const mode change was successful
        self.emit(SIGNAL('constant_mode_changed'))

