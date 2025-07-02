import os
import sys
import csv
import time
import errno
import datetime
import ConfigParser
import logging
import logging.handlers
from Queue import Queue

from PyQt4.QtCore import *
from PyQt4 import QtGui
from PyQt4.QtGui import *

from dcload import DCLoad
from dc_logger import DCLoggerWorker, TEST_MODE_STATUS
from conf import DEFAULT_PORT, DEFAULT_BAUD, DEFAULT_TIME, DEFAULT_FILENAME,\
    NEW_FILE_DIALOG_TEXT, EXISTING_FILE_DIALOG_TEXT

app_dir = os.path.normpath(os.path.expandvars('%APPDATA%/dc_logger'))
if not os.path.exists(app_dir):
    os.makedirs(app_dir)
log_file_location = os.path.join(app_dir, 'dc_logger.log')
log = logging.getLogger('dc_logger')
log.setLevel(logging.DEBUG)
handler = logging.handlers.RotatingFileHandler(log_file_location, maxBytes=10*1024*1024, backupCount=5)
formatter = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")
handler.setFormatter(formatter)
log.addHandler(handler)

DEFAULT_DC_LOGGER_QUEUE_SIZE = 10
UNKNOWN_STATE = 0
LOCAL_STATE = 1
REMOTE_STATE = 2
RELEASE_VERSION = 8

DEFAULT_FILE_LOG_STRUCT = ('timestamp', 'voltage', 'power', 'current', 'mode', 'total_seconds',)

CONSTANT_MODES = {
    'cr': {
        'idx': 1,
        'id': 'const_resistance',
        'name': 'Constant resistance',
    },
    'cv': {
        'idx': 2,
        'id': 'const_voltage',
        'name': 'Constant voltage',
    },
    'cc': {
        'idx': 3,
        'id': 'const_current',
        'name': 'Constant current',
    },
    'cw': {
        'idx': 4,
        'id': 'const_power',
        'name': 'Constant power',
    },
}


class Logger(QtGui.QMainWindow):
    def __init__(self, parent=None):
        super(Logger, self).__init__(parent)

        self._queue = Queue(DEFAULT_DC_LOGGER_QUEUE_SIZE)
        self._dc_logger_state = UNKNOWN_STATE
        self._is_com_port_open = False
        self._port = None
        self._baud = None
        self._file_log_interval = 0
        self._display_interval = 0
        self._file_name = None
        self._log_dir = None
        self._full_filename = None
        self._seconds_from_logging = 0
        self._is_logging_active = False        
        self._is_load_on = False
        self._constants_mode = None
        self._worker = DCLoggerWorker(self._queue)
        self._worker.start()
        self._display_logger = None
        self._file_logger = None
        self._file_obj = None
        self._csv_obj = None
        self._constants_cache = {}

        #control objects
        self._start_logging_button = None
        self._stop_logging_button = None
        self._load_on_button = None
        self._load_off_button = None
        self._com_port_edit = None
        self._baud_rate_edit = None
        self._com_connect_button = None 
        self._constants_combobox = None
        self._constants_value_edit = None
        self._constants_update_button = None
        self._file_log_edit = None
        self._file_log_button = None
        self._log_file_interval_edit = None
        self._log_file_interval_button = None
        self._log_display_interval_edit = None
        self._log_display_interval_button = None
        self._constant_fields = None
        self._reading_fields = None
        self._log_memo = None
        self._com_status_label = None
        self._logging_status_label = None
        self._constant_mode_status_label = None

    def main(self):
        self._setup_ui()
        self._build_logger_timers()
        self._reset_logger_objects(False)
        self._connect_actions()
        self._connect_logger_signals()
        self._load_user_preferences()
        self.show()

    def _setup_ui(self):
        self._update_window_title()
        central_widget = QtGui.QWidget(self)
        self.setCentralWidget(central_widget)
        layout_panel = QVBoxLayout(central_widget)

        control_frame = self._build_control_fields(central_widget)
        layout_panel.addWidget(control_frame)
        
        log_settings_frame = self._build_log_settings_objects(central_widget)
        layout_panel.addWidget(log_settings_frame)

        readings_frame = self._build_readings_fields(central_widget)
        layout_panel.addWidget(readings_frame)

        self._log_memo = QtGui.QPlainTextEdit(central_widget)
        self._log_memo.setMinimumHeight(400)
        self._log_memo.setReadOnly(True)
        layout_panel.addWidget(self._log_memo)
        self._build_status_bar(central_widget, layout_panel)

    def closeEvent(self, event):
        if self._is_com_port_open:
            try:
                self._stop_logging()
                self._worker.disconnect()
            except:
                log.exception('Failed to close connection')
        self._save_user_preferences()

    def _build_status_bar(self, central_widget, parent_panel):
        """ create custom status bar - don't like qt's own status bar
    
        """

        status_bar_panel = QHBoxLayout()
        status_bar_panel.setContentsMargins(2, 0, 2, 0)
        status_bar_panel.setSpacing(0)
        parent_panel.addLayout(status_bar_panel)
        
        self._com_status_label = self._build_status_bar_label(central_widget, status_bar_panel, 'COM status: inactive')
        self._logging_status_label = self._build_status_bar_label(central_widget, status_bar_panel, 'Logging inactive')
        self._constant_mode_status_label = self._build_status_bar_label(central_widget, status_bar_panel, 'Current mode: N/A')

    def _build_status_bar_label(self, central_widget, parent_panel, text=''):        
        status_label = QLabel(central_widget)
        status_label.setText(text)
        status_label.setMinimumHeight(22)
        status_label.setFrameShape(QFrame.Box)
        status_label.setFrameShadow(QFrame.Raised)
        bold_font = QtGui.QFont()
        bold_font.setBold(True)
        status_label.setFont(bold_font)
        parent_panel.addWidget(status_label)
        return status_label

    def _build_control_fields(self, central_widget):

        inner_frame = QFrame(central_widget)
        inner_frame.setFrameShape(QFrame.StyledPanel)
        inner_frame.setFrameShadow(QFrame.Sunken)
        inner_layer_panel = QHBoxLayout(inner_frame)

        start_log_button = QPushButton(inner_frame)
        start_log_button.setText('Start logging')
        start_log_button.setMinimumWidth(100)
        inner_layer_panel.addWidget(start_log_button)
        self._start_logging_button = start_log_button

        stop_log_button = QPushButton(inner_frame)
        stop_log_button.setText('Stop logging')
        stop_log_button.setMinimumWidth(100)
        inner_layer_panel.addWidget(stop_log_button)
        self._stop_logging_button = stop_log_button

        load_on_button = QPushButton(inner_frame)
        load_on_button.setMinimumWidth(100)
        load_on_button.setText('Turn load on')
        inner_layer_panel.addWidget(load_on_button)
        self._load_on_button = load_on_button

        load_off_button = QPushButton(inner_frame)
        load_off_button.setMinimumWidth(100)
        load_off_button.setText('Turn load off')
        inner_layer_panel.addWidget(load_off_button)
        self._load_off_button = load_off_button

        inner_layer_panel_hspacer = QtGui.QSpacerItem(0, 0, QtGui.QSizePolicy.Expanding, QtGui.QSizePolicy.Minimum)
        inner_layer_panel.addItem(inner_layer_panel_hspacer)

        com_port_label = QLabel(inner_frame)
        com_port_label.setText('COM port:')
        inner_layer_panel.addWidget(com_port_label)

        com_port_edit = QLineEdit(inner_frame)
        com_port_edit.setMaximumWidth(40)
        com_port_edit.setMaxLength(2)
        com_port_edit.setValidator(QtGui.QIntValidator())
        inner_layer_panel.addWidget(com_port_edit)
        self._com_port_edit = com_port_edit

        baud_rate_label = QLabel(inner_frame)
        baud_rate_label.setText('Baud:')
        inner_layer_panel.addWidget(baud_rate_label)

        baud_rate_edit = QLineEdit(inner_frame)
        baud_rate_edit.setMaximumWidth(75)
        baud_rate_edit.setMaxLength(7)
        baud_rate_edit.setValidator(QtGui.QIntValidator())
        inner_layer_panel.addWidget(baud_rate_edit)
        self._baud_rate_edit = baud_rate_edit

        com_open_button = QPushButton(inner_frame)
        com_open_button.setText('Connect')
        com_open_button.setMinimumWidth(120)
        inner_layer_panel.addWidget(com_open_button)
        self._com_connect_button = com_open_button

        return inner_frame

    def _build_log_settings_objects(self, central_widget):
        inner_frame = QFrame(central_widget)
        inner_frame.setFrameShape(QFrame.StyledPanel)
        inner_frame.setFrameShadow(QFrame.Sunken)

        inner_v_layout = QVBoxLayout(inner_frame)
        file_settings_panel = QHBoxLayout()
        display_settings_panel = QHBoxLayout()

        inner_v_layout.addLayout(file_settings_panel)
        inner_v_layout.addLayout(display_settings_panel)

        log_file_label = QLabel(inner_frame)
        log_file_label.setText('Log to file:')
        file_settings_panel.addWidget(log_file_label)

        log_file_edit = QLineEdit(inner_frame)
        log_file_edit.setMinimumWidth(500)
        log_file_edit.setReadOnly(True)
        file_settings_panel.addWidget(log_file_edit)
        self._file_log_edit = log_file_edit

        log_file_button = QPushButton(inner_frame)
        log_file_button.setText('Browse...')
        file_settings_panel.addWidget(log_file_button)
        self._file_log_button = log_file_button

        file_settings_panel_spacer = QtGui.QSpacerItem(0, 0, QtGui.QSizePolicy.Expanding, QtGui.QSizePolicy.Minimum)
        file_settings_panel.addItem(file_settings_panel_spacer)

        file_res = self._build_log_interval_objects(file_settings_panel, inner_frame, 'Log interval')
        self._log_file_interval_edit = file_res[0]
        self._log_file_interval_button = file_res[1]

        constants_mode_set_label = QLabel(inner_frame)
        constants_mode_set_label.setText('Set mode:')
        display_settings_panel.addWidget(constants_mode_set_label)

        constants_mode_combobox = QtGui.QComboBox(inner_frame)
        constants_mode_combobox.setMinimumWidth(175)
        constants_mode_combobox.setMaximumWidth(175)
        constants_mode_combobox.insertItem(0, '---', 'DUMMY')
        for key, val in CONSTANT_MODES.items():
            constants_mode_combobox.insertItem(val['idx'], val['name'], key)

        display_settings_panel.addWidget(constants_mode_combobox)
        self._constants_combobox = constants_mode_combobox

        constants_value_edit = QLineEdit(inner_frame)
        constants_value_edit.setMinimumWidth(50)
        constants_value_edit.setMaximumWidth(50)
        constants_value_edit.setValidator(QtGui.QRegExpValidator(QRegExp("[\d\.\,]+")))
        constants_value_edit.editingFinished.connect(self._parse_float)
        display_settings_panel.addWidget(constants_value_edit)
        self._constants_value_edit = constants_value_edit

        constants_mode_update_button = QPushButton(inner_frame)
        constants_mode_update_button.setText('Update')
        constants_mode_update_button.setMaximumWidth(75)
        display_settings_panel.addWidget(constants_mode_update_button)
        self._constants_update_button = constants_mode_update_button

        display_settings_panel_spacer = QtGui.QSpacerItem(0, 0, QtGui.QSizePolicy.Expanding, QtGui.QSizePolicy.Minimum)
        display_settings_panel.addItem(display_settings_panel_spacer)
        display_res = self._build_log_interval_objects(display_settings_panel, inner_frame, 'Display interval')
        self._log_display_interval_edit = display_res[0]
        self._log_display_interval_button = display_res[1]

        return inner_frame

    def _build_log_interval_objects(self, parent_panel, parent_frame, log_type_name):
        separate_line = QFrame(parent_frame)
        separate_line.setFrameShape(QFrame.VLine)
        separate_line.setFrameShadow(QFrame.Sunken)
        parent_panel.addWidget(separate_line)

        log_label = QLabel(parent_frame)
        log_label.setText(log_type_name)
        log_label.setMinimumWidth(105)
        log_label.setMaximumWidth(105)
        log_label.setAlignment(Qt.AlignRight|Qt.AlignVCenter)
        parent_panel.addWidget(log_label)

        log_interval_edit = QLineEdit(parent_frame)
        log_interval_edit.setMinimumWidth(55)
        log_interval_edit.setMaximumWidth(55)
        log_interval_edit.setValidator(QtGui.QRegExpValidator(QRegExp("[\d\.\,]+")))
        log_interval_edit.editingFinished.connect(self._parse_float)
        parent_panel.addWidget(log_interval_edit)

        log_seconds_label = QLabel(parent_frame)
        log_seconds_label.setText('sec')

        log_seconds_label.setMinimumWidth(30)
        log_seconds_label.setMaximumWidth(30)
        parent_panel.addWidget(log_seconds_label)

        log_interval_button = QPushButton(parent_frame)
        log_interval_button.setText('Update')
        log_interval_button.setMaximumWidth(75)
        parent_panel.addWidget(log_interval_button)

        return (log_interval_edit, log_interval_button)

    def _parse_float(self):
        obj = self.sender()
        obj_text = obj.text()
        if not obj_text:
            obj.setText('0')
        else:
            obj.setText('%0.2f' % (float(obj_text.replace(',', '.'))))

    def _build_readings_fields(self, central_widget):
        """ Render line edit and label object
            @return QFrame

        """
    
        self._constant_fields = {
            'const_current': {
                'edit_obj': None,
                'label_obj': None,
                'label': 'CC:',
                },
            'const_voltage': {
                'edit_obj': None,
                'label_obj': None,
                'label': 'CV:',
                },
            'const_power': {
                'edit_obj': None,
                'label_obj': None,
                'label': 'CW:',
                },
            'const_resistance': {
                'edit_obj': None,
                'label_obj': None,
                'label': 'CR:',
            },
        }
        
        self._reading_fields = {
            'input_current': {
                'edit_obj': None,
                'label_obj': None,
                'label': 'A',
                },
            'input_voltage': {
                'edit_obj': None,
                'label_obj': None,
                'label': 'V',
                },
            'input_power': {
                'edit_obj': None,
                'label_obj': None,
                'label': 'W',
                },
        }

        inner_frame = QFrame(central_widget)
        inner_frame.setFrameShape(QFrame.StyledPanel)
        inner_frame.setFrameShadow(QFrame.Sunken)
        inner_layer_panel = QHBoxLayout(inner_frame)

        for key, elem in self._constant_fields.items():
            new_label = QLabel(inner_frame)
            new_label.setText(elem['label'])
            new_label.setMinimumWidth(33)
            new_label.setMaximumWidth(33)
            new_label.setAlignment(Qt.AlignRight|Qt.AlignVCenter)
            elem['label_obj'] = new_label
            inner_layer_panel.addWidget(new_label)

            new_edit = QLineEdit(inner_frame)
            new_edit.setText('-')
            new_edit.setReadOnly(True)
            new_edit.setAlignment(Qt.AlignHCenter)
            new_edit.setMaximumWidth(55)
            elem['edit_obj'] = new_edit
            inner_layer_panel.addWidget(new_edit)

        inner_layer_spacer = QtGui.QSpacerItem(22, 0, QSizePolicy.Expanding, QSizePolicy.Minimum)
        inner_layer_panel.addItem(inner_layer_spacer)

        for key, elem in self._reading_fields.items():
            new_edit = QLineEdit(inner_frame)
            new_edit.setText('-')
            new_edit.setReadOnly(True)
            new_edit.setAlignment(Qt.AlignHCenter)
            new_edit.setMaximumWidth(55)
            elem['edit_obj'] = new_edit
            inner_layer_panel.addWidget(new_edit)

            new_label = QLabel(inner_frame)
            new_label.setText(elem['label'])
            new_label.setMaximumWidth(20)
            elem['label_obj'] = new_label
            inner_layer_panel.addWidget(new_label)

        return inner_frame

    def _build_logger_timers(self):
        self._display_logger = QTimer(self)
        self._display_logger.stop()
        self._display_logger.setInterval(self._display_interval * 1000)
        
        self._file_logger = QTimer(self)
        self._file_logger.stop()
        self._file_logger.setInterval(self._file_log_interval * 1000)
    
    def _reset_logger_objects(self, is_com_open=None):
        if is_com_open is not None:
            self._com_port_edit.setDisabled(is_com_open)
            self._baud_rate_edit.setDisabled(is_com_open)
            self._com_connect_button.setText('Disconnect' if is_com_open else 'Connect')
            self._start_logging_button.setEnabled(is_com_open)
            self._stop_logging_button.setEnabled(False)
            self._load_on_button.setEnabled(is_com_open)
            self._load_off_button.setEnabled(False)
            self._constants_combobox.setEnabled(is_com_open)
            self._constants_update_button.setEnabled(is_com_open)

        self._constants_combobox.setCurrentIndex(0)
        self._constants_value_edit.setText('0')

        self._update_constant_mode_status_label(self._constants_mode)

        disabled_bold_font = QtGui.QFont()
        disabled_bold_font.setBold(False)

        for key, elem in self._constant_fields.items():
            elem['edit_obj'].setText('-')
            elem['label_obj'].setFont(disabled_bold_font)

        for key, elem in self._reading_fields.items():
            elem['edit_obj'].setText('-')
            elem['label_obj'].setFont(disabled_bold_font)               

    def _save_user_preferences(self):
        config = ConfigParser.RawConfigParser()
        config.add_section('Main')
        config.set('Main', 'log_dir', self._log_dir)
        config.set('Main', 'file_name', self._file_name)
        config.set('Main', 'port', self._port)
        config.set('Main', 'baud', self._baud)
        config.set('Main', 'file_log_interval', self._file_log_interval)
        config.set('Main', 'display_interval', self._display_interval)
        with open((os.path.join(os.getcwd(), 'preferences.cfg')), 'wb') as config_file:
            config.write(config_file)

    def _load_user_preferences(self):
        default_values = {
            'log_dir': os.getcwd(),
            'file_name': DEFAULT_FILENAME,
            'file_log_interval': DEFAULT_TIME,
            'display_interval': DEFAULT_TIME,
            'port': DEFAULT_PORT,
            'baud': DEFAULT_BAUD,            
        }
        config = ConfigParser.RawConfigParser(default_values)
        config.read(os.path.join(os.getcwd(), 'preferences.cfg'))
        if not config.has_section('Main'):
            config.add_section('Main')
            
        self._log_dir = config.get('Main', 'log_dir')
        self._file_log_interval = config.getfloat('Main', 'file_log_interval')
        self._display_interval = config.getfloat('Main', 'display_interval')
        self._file_name = config.get('Main', 'file_name')
        self._full_filename = os.path.normpath(os.path.join(self._log_dir, self._file_name))
        self._port = config.get('Main', 'port')
        self._baud = config.get('Main', 'baud')
        
        self._file_log_edit.setText(str(self._full_filename))
        self._log_file_interval_edit.setText(str(self._file_log_interval))
        self._log_display_interval_edit.setText(str(self._display_interval))
        self._com_port_edit.setText(str(self._port))
        self._baud_rate_edit.setText(str(self._baud))
        self._file_logger.setInterval(self._file_log_interval * 1000)
        self._display_logger.setInterval(self._display_interval * 1000)

    def _update_window_title(self, appendix=''):
        test_mode_str = 'TEST MODE ACTIVE' if TEST_MODE_STATUS else ''
        self.setWindowTitle('DC Logger :: r%s %s %s' % (RELEASE_VERSION, appendix, test_mode_str,))

    def _update_constant_mode_status_label(self, mode):
        new_const_str = 'Current mode: %s' % (str('N/A' if mode is None else mode).upper())
        self._constant_mode_status_label.setText(new_const_str)

    def _file_has_write_access(self, file_name):
        try:
            fp = open(file_name, 'a+')
            fp.close()
        except IOError as err:
            if err.errno == errno.EACCES:
                return False
            else:
                raise
        return True

    def _check_for_existing_file(self):
        new_file_location = False
        file_exists = os.path.exists(self._full_filename)

        # assume failure
        file_has_write_access = False

        if not file_exists:
            return self._select_new_file_location()

        if file_exists:
            file_has_write_access = self._file_has_write_access(self._full_filename)

        if file_exists and file_has_write_access:
            q = QMessageBox()
            new_file = q.addButton('Create new file', QMessageBox.YesRole)
            append_file = q.addButton('Append', QMessageBox.RejectRole)
            replace_file = q.addButton('Replace', QMessageBox.ActionRole)
            q.setIcon(QMessageBox.Information)

            q.setWindowTitle('Confirm file overwrite')
            q.setText(NEW_FILE_DIALOG_TEXT % (self._file_name))
            q.exec_()

            if q.clickedButton() is append_file:
                return True
            elif q.clickedButton() is replace_file:
                os.remove(self._full_filename)
                return True
            elif q.clickedButton() is new_file:
                return self._select_new_file_location()

        if not file_has_write_access:
            return self._issue_file_write_access_denied()

        return True
           
    def _issue_file_write_access_denied(self):
        q = QMessageBox(QMessageBox.Critical, 'Access denied',
            EXISTING_FILE_DIALOG_TEXT % (str(self._full_filename)), QMessageBox.Ok)
        q.exec_()           
        return False

    def _select_new_file_location(self):
        new_location = None
        try:
            new_location = QtGui.QFileDialog.getSaveFileName (self, 'Save log file to...')
            if not new_location:
                return False
            new_location = os.path.normpath(str(new_location))

            self._log_dir = os.path.dirname(new_location)
            new_filename = os.path.basename(new_location)

            if not os.path.splitext(new_filename)[1]:
                new_filename += '.csv'
            self._file_name = new_filename

            parsed_new_location = os.path.normpath(os.path.join(self._log_dir, self._file_name))
            if not self._file_has_write_access(parsed_new_location):
                return self._issue_file_write_access_denied()

            self._full_filename = parsed_new_location
            self._update_file_location_edit()
        except Exception, e:
            log.exception('New file "%s" selection failed' % (new_location))
            self._update_log(str(e))
            return False
            
        return True

    def _update_file_location_edit(self):
        try:
            self._file_log_edit.setText(self._full_filename)
            self._save_user_preferences()
        except Exception, e:
            log.exception('Failed to update file location edit')
            self._update_log(str(e))
    
    # SIGNALS & SLOTS
    def _connect_actions(self):
            self._start_logging_button.clicked.connect(self._start_logging)
            self._stop_logging_button.clicked.connect(self._stop_logging)
            self._load_on_button.clicked.connect(self._load_on)
            self._load_off_button.clicked.connect(self._load_off)
            self._com_connect_button.clicked.connect(self._com_port_connection_handle)
            self._log_file_interval_button.clicked.connect(self._update_log_file_interval)
            self._log_display_interval_button.clicked.connect(self._update_log_display_interval)
            self._constants_update_button.clicked.connect(self._update_constants_settings)
            self._file_log_button.clicked.connect(self._select_new_file_location)
            self._display_logger.timeout.connect(self._request_display_data)
            self._file_logger.timeout.connect(self._request_file_data)
            self._constants_combobox.currentIndexChanged.connect(self._on_constant_mode_selection)
            
    def _connect_logger_signals(self):
        self.connect(self._worker, SIGNAL('error_msg_posted'), self._update_log)
        self.connect(self._worker, SIGNAL('com_port_state_changed'), self._on_com_port_state_change)
        self.connect(self._worker, SIGNAL('load_state_changed'), self._on_load_state_change)
        self.connect(self._worker, SIGNAL('display_input_data_available'), self._on_get_display_input_data)
        self.connect(self._worker, SIGNAL('file_input_data_available'), self._on_get_file_input_data)
        self.connect(self._worker, SIGNAL('constants_data_available'), self._on_get_constants_data)
        self.connect(self._worker, SIGNAL('constant_mode_changed'), self._on_constant_mode_changed)

    def _com_port_connection_handle(self):
        if self._is_com_port_open:
            try:
                self._stop_logging()
                self._worker.disconnect()
            except:
                log.exception('Failed to close connection')
        else:
            self._port = int(self._com_port_edit.text())
            self._baud = int(self._baud_rate_edit.text())
            self._worker.connect(self._port, self._baud)

    def _on_constant_mode_selection(self, idx):
        self._constants_update_button.setDisabled(idx == 0)
        try:
            selected_mode = str(self._constants_combobox.itemData(idx).toString())
            self._constants_value_edit.setText(str(self._constants_cache[selected_mode]['val']))
        except:
            pass

    def _initialize_dc_logger(self):
        self._update_log('Initializing logger')
        try:
            if not self._check_for_existing_file():
                return

            self._file_obj = open(self._full_filename, 'a+')
            self._csv_obj = csv.DictWriter(self._file_obj, DEFAULT_FILE_LOG_STRUCT)
            if os.path.getsize(self._full_filename) == 0:
                self._csv_obj.writeheader()

        except:
            self._update_log('File check failed')
            return False

        self._queue.put(('get_constants_values',))
        return True

    def _start_logging(self):
        if not self._initialize_dc_logger():
            self._update_log('Initialization failed')
            return

        self._seconds_from_logging = time.time()
        self._toggle_logging()
        if not self._is_load_on:
            self._load_on()

        if self._display_logger.interval() >= 500:
            self._display_logger.start()

            self._display_logger.start()
        if self._file_logger.interval() >= 500:
            self._file_logger.start()

        self._file_log_edit.setEnabled(False)
        self._file_log_button.setEnabled(False)
        self._logging_status_label.setText('Logging: active')
        self._update_log('logging started @ %s' % (str(datetime.datetime.now())))
    
    def _stop_logging(self):
        self._toggle_logging()
        if self._file_obj is not None:
            self._file_obj.close()

        self._file_log_edit.setEnabled(True)
        self._file_log_button.setEnabled(True)
        self._display_logger.stop()
        self._file_logger.stop()

        self._logging_status_label.setText('Logging: inactive')
        self._update_log('logging stopped @ %s' % (str(datetime.datetime.now())))

    def _toggle_logging(self):
        self._start_logging_button.setEnabled(self._is_logging_active)
        self._stop_logging_button.setDisabled(self._is_logging_active)
        self._is_logging_active = not self._is_logging_active

    def _load_on(self):
        self._queue.put(('turn_load_on',))

    def _load_off(self):
        self._queue.put(('turn_load_off',))

    def _update_log_file_interval(self):
        try:
            new_interval_value = float(self._log_file_interval_edit.text())
            self._file_log_interval = new_interval_value

            if new_interval_value < 0.5:
                self._file_logger.stop()
                self._update_log('File log interval is too low, minimum allowed value is 0.5 second.')
            else:
                self._file_logger.setInterval(new_interval_value * 1000)
                if not self._file_logger.isActive() and self._is_logging_active:
                    self._file_logger.start()

            self._save_user_preferences()
        except:
            self._update_log('ERROR: failed to update file log interval')
            log.exception('Failed to update file log interval')

    def _update_log_display_interval(self):
        try:
            new_interval_value = float(self._log_display_interval_edit.text())
            self._display_interval = new_interval_value
            if new_interval_value < 0.5:
                self._display_logger.stop()
                self._update_log('Display log interval is too low, minimum allowed value is 0.5 second.')
            else:
                self._display_logger.setInterval(new_interval_value * 1000)
                    
                if not self._display_logger.isActive() and self._is_logging_active:
                    self._display_logger.start()

            self._save_user_preferences()
        except:
            self._update_log('ERROR: failed to update display interval')
            log.exception('Failed to update display log interval')

    def _update_constants_settings(self):
        if self._constants_combobox.currentIndex() == 0:
            if TEST_MODE_STATUS:
                log.warning('Tried to update dummy constant')
            return

        obj_full_name = self._constants_combobox.currentText()

        try:
            idx = self._constants_combobox.currentIndex()
            obj = str(self._constants_combobox.itemData(idx).toString())
            val = str(self._constants_value_edit.text())
            payload = ('set_constants_values', (obj, val,),)
            self._queue.put(payload)

        except:
            log.exception('Failed to update constants settings')
            self._update_log('Failed to update "%s"' % (obj_full_name,))
  
    def _request_display_data(self):
        self._queue.put(('get_input_data', 'display',))

    def _request_file_data(self):
        self._queue.put(('get_input_data', 'file',))

    def _update_log(self, log_text=''):
        try:
            self._log_memo.appendPlainText(str(log_text))
        except:
            log.exception('Failed to update logger memo')

    def _on_com_port_state_change(self, is_com_open):
        if is_com_open:
            self._reset_logger_objects(True)
            self._com_status_label.setText('COM status: port %s open' % (self._port))
        else:
            self._is_load_on = False
            self._constants_mode = None
            self._reset_logger_objects(False)
            self._com_status_label.setText('COM status: inactive')
        self._is_com_port_open = is_com_open
  
    def _on_load_state_change(self, state):
        load_on = True if state == 'on' else False
        self._is_load_on = load_on
        self._load_on_button.setDisabled(load_on)
        self._load_off_button.setEnabled(load_on)

    def _on_get_display_input_data(self, in_val):
        if not in_val or in_val is None:
            for reading_obj in self._reading_fields.values():
                reading_obj['edit_obj'].setText('-')
            return

        try:
            for key, val in in_val.items():
                self._reading_fields[key]['edit_obj'].setText(str(val))
        except:
            if not TEST_MODE_STATUS:
                log.exception('Failed to update display fields with data %s' % (str(in_val)))

    def _on_get_file_input_data(self, in_val):
        log_data = {}
        if in_val is None or not in_val:
            return

        now = datetime.datetime.now()
        current_date = now.strftime('%Y-%m-%d %H:%M:%S')
       
        try:
            total_seconds = int(time.time() - self._seconds_from_logging)
            voltage = in_val.get('input_voltage')
            power = in_val.get('input_power')
            current = in_val.get('input_current')       

            log_data['timestamp'] = current_date
            log_data['voltage'] = voltage
            log_data['power'] = power
            log_data['current'] = current
            log_data['mode'] = self._constants_mode
            log_data['total_seconds'] = total_seconds

            display_log_text = '%s: mode: %s, voltage: %s V, power: %s W, current: %s A' % \
                (current_date, str(self._constants_mode).upper(), voltage, power, current)

            self._csv_obj.writerow(log_data)
            self._update_log(display_log_text)
        except:
            if not TEST_MODE_STATUS:
                log.exception('Failed to update file log')    

    def _on_get_constants_data(self, in_val):
        bold_font_enabled = QtGui.QFont()
        bold_font_enabled.setBold(True)
        bold_font_disabled = QtGui.QFont()
        bold_font_disabled.setBold(False)
        self._constants_cache = in_val
        log.debug(str(in_val))
        for const_name, const_obj in in_val.items():
            try:
                mode_obj = CONSTANT_MODES[const_name]
                mode_id = mode_obj['id']
                mode_idx = mode_obj['idx']
                const_value = const_obj['val']

                self._constant_fields[mode_id]['label_obj'].setFont(bold_font_disabled)
                self._constant_fields[mode_id]['edit_obj'].setText(str(const_value))

                if const_obj.get('is_active', False):
                    self._constants_mode = const_name
                    self._constants_combobox.setCurrentIndex(mode_idx)
                    self._constant_fields[mode_id]['label_obj'].setFont(bold_font_enabled)
                    self._constants_value_edit.setText(str(const_value))
                    self._update_constant_mode_status_label(const_name)
            except:
                if TEST_MODE_STATUS:
                    log.exception('Got unexpected exception on getting contants data')
                log.exception('Failed to fill new constant data')

    def _on_constant_mode_changed(self):
        self._queue.put(('get_constants_values',))

if __name__ == '__main__':
    app = QtGui.QApplication(sys.argv)
    dc_logger = Logger()
    dc_logger.main()
    sys.exit(app.exec_())
