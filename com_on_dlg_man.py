"""
Common Dialog Box - methods to interact with Windows Common Dialog Boxes
    Open File Dialog Box: This dialog box is used to open a file from the user's computer.
    Save File Dialog Box: This dialog box is used to save a file to the user's computer.
    Print Dialog Box: This dialog box is used to print a document.
    Find Dialog Box: This dialog box is used to find text in a document.
    Replace Dialog Box: This dialog box is used to replace text in a document.

    To Import:
    - Add path to sys.path.append("C:\\...\\Documents\\Dev\\Python\\Projects\\Common")
        And/Or
    - Add to Environment Variable "PYTHONPATH"
"""

import os
import inspect
import win32con
import win32gui
from time import time
from time import sleep
import logging


# class MsoFileDialogType(enum.IntEnum):
#     FileDialogOpen = 1  # Win API
#     FileDialogSaveAs = 2  # Win API
#     # FileDialogFilePicker = 3  # Win API
#     # FileDialogFolderPicker = 4  # MSO


"""Global variables"""
_common_dlg_window_handle = None
_common_dlg_path = ""


def _get_child_windows_by_class_and_caption_path_call_back(hwnd, param):
    """Called by find_window_by_class_and_caption_path()."""
    child_class_name = win32gui.GetClassName(hwnd)
    child_text_name = win32gui.GetWindowText(hwnd)
    child_text_name = None if child_text_name == "" else child_text_name
    if param[0][2] == win32gui.GetParent(hwnd):
        if param[0][0] == child_class_name and param[0][1] == child_text_name:
            param[1].append(hwnd)  # Found


def _get_child_windows_by_class_and_caption_path(hwnd_parent, element_path):
    """Returns a list containing the hwnd of all child window items
    found matching the class name and caption from ordered list passed into element_path"""
    # noinspection SpellCheckingInspection
    """Class names found using uuspy.exe Source:https://uuware.com/st_l.en/st_p2.uw_spy.html"""
    element_path.append(["Final", "Result"])
    element_path[0].append(hwnd_parent)
    while len(element_path) > 1:
        while len(element_path[0][2:]) > 0:
            win32gui.EnumChildWindows(element_path[0][2], _get_child_windows_by_class_and_caption_path_call_back,
                                      element_path)
            del element_path[0][2]
        del element_path[0]
    return element_path[0][2:]


def _error_message(e, details):
    calling_function_name = inspect.getframeinfo(details).function
    logging.error("The exception type is:", type(e), f"An error occurred in function: "
                                                     f"{calling_function_name}", f"Error details: {e}")


def _wait_for_window_open(class_name, caption, sleep_time=0.1, wait_time=120):
    """Waits for a window with the given caption to open and returns its window handle."""
    """Can return incorrect window if multiple are open at the same time."""
    try:
        start_time = time()  # Seconds
        hwnd_save_as = 0
        while not win32gui.IsWindow(hwnd_save_as) and (time() - start_time) < wait_time:
            hwnd_save_as = win32gui.FindWindowEx(0, 0, class_name, caption)
            # print(win32gui.IsWindowVisible(hwnd_save_as))

        if hwnd_save_as == 0:
            raise TimeoutError("A timeout error occurred in: wait_for_window_to_close()")
        else:
            sleep(sleep_time)  # If not here edit_handle isn't set.
            return hwnd_save_as
    except Exception as e:
        _error_message(e, inspect.currentframe())


def _get_text_from_dialog_box(hwnd):
    try:
        """Gets the text from the Save As dialog box."""
        buffer_size = 1024
        buffer = win32gui.PyMakeBuffer(buffer_size)  # win32gui.PyMakeBuffer(text_length + 1)

        # Get the encoded text.
        result = win32gui.SendMessage(hwnd, win32con.WM_GETTEXT, len(buffer), buffer)

        if result > 0:
            decoded_text = buffer.tobytes().decode("utf-16", errors="replace")
            # decoded_text = win32gui.PyGetString(buffer)
            return decoded_text[:result]
    except Exception as e:
        _error_message(e, inspect.currentframe())


def _wait_for_window_close(sleep_time=0.1, wait_time=120):
    """Waits for a window with the given handle to close."""
    """Before closing window verify file path to avoid additional popup windows:"""
    """'Path doesn't' exist & 'File Overwrite'"""
    try:
        start_time = time()  # Seconds
        global _common_dlg_window_handle
        while win32gui.IsWindow(_common_dlg_window_handle) and (time() - start_time) < wait_time:
            sleep(sleep_time)
        if win32gui.IsWindow(_common_dlg_window_handle):
            raise TimeoutError("A timeout error occurred in: wait_for_window_to_close()")
        else:
            return True
    except Exception as e:
        _error_message(e, inspect.currentframe())


def _wait_for_file_exist(sleep_time=0.1, wait_time=120):
    try:
        start_time = time()  # Seconds
        global _common_dlg_path
        while not os.path.isfile(_common_dlg_path) and (time() - start_time) < wait_time:
            sleep(sleep_time)
        if not os.path.isfile(_common_dlg_path):
            raise TimeoutError("A timeout error occurred in: wait_for_file_download()")
        else:
            return True
    except Exception as e:
        _error_message(e, inspect.currentframe())


class OpenCommonDlg:
    __open_file_name_handle = int
    __open_type_handle = int
    __open_open_button_handle = int
    __open_cancel_button_handle = int

    def open_window_interact(self, file_path):
        """File paths are verified earlier in code. Overwrite dialog box if overwrite_files:
        find_and_confirm_yes_button() else: find_and_confirm_no_button()  # Do not overwrite files. Path not
        found dialog box find_and_confirm_OK_button() sleep(sleep_time)
        """
        if os.path.isfile(file_path):
            self.file_path = file_path
            try:  # Open Window
                # win32gui.SetActiveWindow(self.window_handle)
                sleep(.5)
                win32gui.SendMessage(self.open_file_name_handle, win32con.WM_SETTEXT, 0, self.file_path)
                win32gui.SendMessage(self.open_open_button_handle, win32con.BM_CLICK, 0, 0)  # Open
                # win32gui.SendMessage(self.save_as_cancel_button_handle, win32con.BM_CLICK, 0, 0)

                _wait_for_window_close()
                return
            except Exception as e:
                _error_message(e, inspect.currentframe())
        else:
            raise FileNotFoundError()

    def __set_open_window_handles(self, sleep_time=0.1, wait_time=120):
        """Keep trying to set these until successful or time runs out."""
        start_time = time()  # Seconds
        while True and (time() - start_time) < wait_time:
            try:
                if self.__set_open_open_button_handle():
                    if self.__set_open_cancel_button_handle():
                        if self.__set_open_file_name_handle():
                            if self.__set_open_type_handle():
                                return True
            except TimeoutError:
                pass
                # print("The function timed out after {} seconds.".format(wait_time))
                # return None
            sleep(sleep_time)

    def __set_open_file_name_handle(self):
        element_path = [["ComboBoxEx32", None], ["ComboBox", None], ["Edit", None]]
        # noinspection SpellCheckingInspection
        hwnds = _get_child_windows_by_class_and_caption_path(self.window_handle, element_path)
        if len(hwnds) > 1:
            raise IndexError
        else:
            self.open_file_name_handle = hwnds[0]
            return True

    def __set_open_type_handle(self):
        element_path = [["ComboBox", None]]
        # noinspection SpellCheckingInspection
        hwnds = _get_child_windows_by_class_and_caption_path(self.window_handle, element_path)
        if len(hwnds) > 1:
            raise IndexError
        else:
            self.open_type_handle = hwnds[0]
            return True

    def __set_open_open_button_handle(self):
        """Get the handle to the child window with the class name "Button" and caption &Open (top level)"""
        self.open_open_button_handle = win32gui.FindWindowEx(self.window_handle, None, "Button", "&Open")
        return True

    def __set_open_cancel_button_handle(self):
        """Get the handle to the child window with the class name "Button" and caption &Save (top level)"""
        self.open_cancel_button_handle = win32gui.FindWindowEx(self.window_handle, None, "Button", "Cancel")
        return True

    def __init__(self):
        """The constructor for the class."""
        # self.dlg_type = MsoFileDialogType.FileDialogOpen
        self.window_handle = _wait_for_window_open("#32770", "Open")
        self.__set_open_window_handles()

    """__save_as_file_name_handle"""
    @property  # Get
    def file_name_handle(self):
        return self.__open_file_name_handle

    @file_name_handle.setter  # Set
    def file_name_handle(self, file_name_handle):
        self.__open_file_name_handle = file_name_handle

    """__find_save_as_type_handle"""
    @property  # Get
    def type_handle(self):
        return self.__open_type_handle

    @type_handle.setter  # Set
    def type_handle(self, type_handle):
        self.__open_type_handle = type_handle

    """__find_open_button_handle"""
    @property  # Get
    def open_button_handle(self):
        return self.__open_open_button_handle

    @open_button_handle.setter  # Set
    def open_button_handle(self, open_button_handle):
        self.__open_open_button_handle = open_button_handle

    """__find_cancel_button_handle"""
    @property  # Get
    def cancel_button_handle(self):
        return self.__open_cancel_button_handle

    @cancel_button_handle.setter  # Set
    def cancel_button_handle(self, cancel_button_handle):
        self.__open_cancel_button_handle = cancel_button_handle

    """___Global accessors___"""
    """__common_dlg_type"""
    # @property  # Get
    # def dlg_type(self):
    #     global __common_dlg_type
    #     return __common_dlg_type
    #
    # @dlg_type.setter  # Set
    # def dlg_type(self, dlg_type):
    #     global __common_dlg_type
    #     __common_dlg_type = dlg_type

    """_common_dlg_window_handle"""
    @property  # Get
    def window_handle(self):
        # _common_dlg_window_handle
        return _common_dlg_window_handle

    @window_handle.setter  # Set
    def window_handle(self, window_handle):
        global _common_dlg_window_handle
        _common_dlg_window_handle = window_handle

    """__common_dlg_file_path"""
    @property  # Get
    def file_path(self):
        global _common_dlg_path
        return _common_dlg_path

    @file_path.setter  # Set
    def file_path(self, file_path):
        global _common_dlg_path
        _common_dlg_path = file_path


class SaveAsCommonDlg:
    __save_as_file_name_handle = int
    __save_as_type_handle = int
    __save_as_save_button_handle = int
    __save_as_cancel_button_handle = int

    def save_as_window_interact(self, file_path, make_directory=True, file_overwrite=True):
        """Overwrite dialog box if overwrite_files:
        find_and_confirm_yes_button() else: find_and_confirm_no_button()
        # Do not overwrite files. Path not
        found dialog box find_and_confirm_OK_button() sleep(sleep_time)
        """
        # These checks are necessary to avoid additional pop-up dialog boxes.
        if not os.path.exists(os.path.dirname(file_path)):  # Check directory.
            if make_directory:
                os.makedirs(os.path.dirname(file_path), exist_ok=True)  # Proceed if path already exists.
        elif os.path.exists(file_path) and file_overwrite:
            os.remove(file_path)

        if os.path.exists(os.path.dirname(file_path)):  # Does file path exist.
            self.file_path = file_path
            try:  # Save As Window
                win32gui.SendMessage(self.file_name_handle, win32con.WM_SETTEXT, 0, self.file_path)
                # sleep(0.25)
                if os.path.exists(file_path) and not file_overwrite:
                    win32gui.SendMessage(self.cancel_button_handle, win32con.BM_CLICK, 0, 0)
                else:
                    win32gui.SendMessage(self.save_button_handle, win32con.BM_CLICK, 0, 0)
                    # win32gui.SendMessage(self.save_as_cancel_button_handle, win32con.BM_CLICK, 0, 0)
                    logging.info("Downloading: " + self.file_path)
                _wait_for_window_close()
                _wait_for_file_exist(0.1, 300)
            except Exception as e:
                _error_message(e, inspect.currentframe())
                return
        else:
            raise FileNotFoundError()

    def __set_save_as_window_handles(self, sleep_time=0.1, wait_time=120):
        """Keep trying to set these until successful or time runs out."""
        start_time = time()  # Seconds
        while True and (time() - start_time) < wait_time:
            try:
                if self.__set_save_as_button_handle():
                    if self.__set_save_as_cancel_button_handle():  # if self.__set_save_as_cancel_button_handle():
                        if self.__set_save_as_file_name_handle():
                            if self.__set_save_as_type_handle():
                                return True
            except TimeoutError:
                pass
                # print("The function timed out after {} seconds.".format(wait_time))
                # return None
            sleep(sleep_time)

    def __set_save_as_button_handle(self,):
        """Get the handle to the child window with the class name "Button" and caption &Save (top level)"""
        self.save_button_handle = win32gui.FindWindowEx(self.window_handle, None, "Button", "&Save")
        return True

    def __set_save_as_cancel_button_handle(self):
        """Get the handle to the child window with the class name "Button" and caption Cancel (top level)"""
        self.cancel_button_handle = win32gui.FindWindowEx(self.window_handle, None, "Button", "Cancel")
        return True

    def __set_save_as_file_name_handle(self):
        element_path = [["DUIViewWndClassName", None], ["DirectUIHWND", None], ["FloatNotifySink", None],
                        ["ComboBox", None], ["Edit", None]]
        # noinspection SpellCheckingInspection
        hwnds = _get_child_windows_by_class_and_caption_path(self.window_handle, element_path)
        for hwnd in hwnds:  # It looks like there is only 1.
            if not _get_text_from_dialog_box(hwnd) is None:
                self.file_name_handle = hwnd
                return True

    def __set_save_as_type_handle(self):
        element_path = [["DUIViewWndClassName", None], ["DirectUIHWND", None],
                        ["FloatNotifySink", None], ["ComboBox", None]]
        # noinspection SpellCheckingInspection
        hwnds = _get_child_windows_by_class_and_caption_path(self.window_handle, element_path)
        for hwnd in hwnds:  # Look for child, which only exits in 'File name' field.
            if len(_get_child_windows_by_class_and_caption_path(hwnd, [["Edit", None]])) == 0:
                # if "Adobe Acrobat Document (*.pdf)" in __get_text_from_save_as_dialog_box(hwnd) is None:
                self.type_handle = hwnd
                return True

    def __init__(self):
        """The constructor for the class."""
        logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
        self.window_handle = _wait_for_window_open("#32770", "Save As")
        self.__set_save_as_window_handles()

    """__save_as_file_name_handle"""
    @property  # Get
    def file_name_handle(self):
        return self.__save_as_file_name_handle

    @file_name_handle.setter  # Set
    def file_name_handle(self, value):
        self.__save_as_file_name_handle = value

    """__find_save_as_type_handle"""
    @property  # Get
    def type_handle(self):
        return self.__save_as_type_handle

    @type_handle.setter  # Set
    def type_handle(self, value):
        self.__save_as_type_handle = value

    """__find_save_button_handle"""
    @property  # Get
    def save_button_handle(self):
        return self.__save_as_save_button_handle

    @save_button_handle.setter  # Set
    def save_button_handle(self, value):
        self.__save_as_save_button_handle = value

    """__find_cancel_button_handle"""
    @property  # Get
    def cancel_button_handle(self):
        return self.__save_as_cancel_button_handle

    @cancel_button_handle.setter  # Set
    def cancel_button_handle(self, value):
        self.__save_as_cancel_button_handle = value

    """___Global accessors___"""
    """_common_dlg_window_handle"""
    @property  # Get
    def window_handle(self):
        global _common_dlg_window_handle
        return _common_dlg_window_handle

    @window_handle.setter  # Set
    def window_handle(self, value):
        global _common_dlg_window_handle
        _common_dlg_window_handle = value

    """__common_dlg_file_path"""
    @property  # Get
    def file_path(self):
        global _common_dlg_path
        return _common_dlg_path

    @file_path.setter  # Set
    def file_path(self, value):
        global _common_dlg_path
        _common_dlg_path = value


# copy_dir(filepath + "\\" + patient_folder_template, patient_full_path)
#
# Window Focus
# win32gui.FlashWindow(hWnd, 0)
# win32gui.SetFocus(hWnd)
# visible = win32gui.IsWindowVisible(hWnd)
# enabled = win32gui.IsWindowEnabled(hWnd)
# win32gui.SetForegroundWindow(hWnd)
# win32gui.SetActiveWindow(hwnd_parent)
#
# Send Keys
# win32gui.SendKeys("{HOME}")
#
# path_text = "C:\\"
# pyperclip.copy(path_text)
# # test = pyperclip.paste()
#
# # Delete the current selection.
# #win32gui.SendKeys("{DEL}")
# # Paste the clipboard contents.
# win32gui.SendKeys(path_text)
# # win32gui.SendKeys("{V}")
# # Tab four times.
# win32gui.SendKeys("{TAB}", 3)
# win32gui.SendKeys("{ENTER}")
#
# def simulate_mouse_click(hWnd):
#     """Simulates a mouse click on the window with the specified handle."""
# noinspection SpellCheckingInspection
#     # Send the WM_L_BUTTON_DOWN message to the window.
#     win32gui.SendMessage(hWnd, win32con.WM_L_BUTTON_DOWN, 0, 0)
# noinspection SpellCheckingInspection
#     # Send the WM_L_BUTTON_UP message to the window.
#     win32gui.SendMessage(hWnd, win32con.WM_L_BUTTON_UP, 0, 0)
#
# # Get the current working directory
# saved_working_directory = os.get_cwd()
# # Set a new working directory
# os.chdir(os.path.dir_name(pdf_filepath))

# webdriver locations
# width = driver.get_window_size().get("width")
# height = driver.get_window_size().get("height")
# x = driver.get_window_position().get('x')
# y = driver.get_window_position().get('y')
