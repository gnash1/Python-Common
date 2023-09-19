"""
Methods to interact with PDFViewer iframe web elements
https://stackoverflow.com/questions/56986848/how-to-download-embedded-pdf-from-webpage-using-selenium
"""

import inspect
import pyautogui
from com_on_dlg_man import SaveAsCommonDlg
import win32gui
from time import sleep
from PIL import ImageGrab  # Install via 'Pillow'
from PIL.Image import Image
from selenium.webdriver.common.action_chains import ActionChains
# from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.support import expected_conditions as ec
# from selenium.webdriver.common.by import By
import logging


def _error_message(e, details):
    calling_function_name = inspect.getframeinfo(details).function
    logging.error("The exception type is:", type(e), f"An error occurred in function: "
                                                     f"{calling_function_name}", f"Error details: {e}")


class ClassPDFView:
    __pdf_view_status = ["Unknown"]  # mutable
    __pdf_view_element = None  # selenium element
    __pdf_view_save_full_path = None
    __xy_hit_point = [0, 0]
    __hwndParent = 0
    __pdf_loaded = False
    __pdf_view_is_initialized = None

    def __refresh_pdf_view_status(self):  # , x_pos = 1593, y_pos = 375) -> str:
        """ Tests image in PDFViewer for color to determine status. Returns (str): "Unknown", "Empty", "Loading",
        "Loaded"" """
        try:
            image = self.__get_pixel_color_imagegrab()
            # win32api.SetCursorPos((save_dx, save_dy))
            self.pdf_view_status = "Unknown"
            if image == (50, 54, 57):
                self.pdf_view_status = "Loading"
            elif image == (169, 169, 169):
                self.pdf_view_status = "Empty"
            elif image == (241, 241, 241) or image == (66, 70, 73):
                self.pdf_view_status = "Loaded"
            return self.pdf_view_status
        except Exception as e:
            _error_message(e, inspect.currentframe())

    def __update_pdf_view_hotspot(self):
        """ finds the pdfView hotspot test pixel """
        # Get the top-left and bottom-right coordinates of the window.
        browser_rect = win32gui.GetWindowRect(self.hwnd_parent)
        # pyautogui.moveTo(rect[2] - rect[0], rect[3] - rect[1])  # This is bottom right corner of browser window.
        element_x = self.pdf_view_element.location["x"]  # X-coordinate
        element_y = self.pdf_view_element.location["y"]  # Y-coordinate
        element_width = self.pdf_view_element.size["width"]  # Width
        self.hit_point = ([browser_rect[0] + (element_x + element_width) - 108, browser_rect[1] + element_y + 159])
        # __y_hit_point = element_y + 169  # 375 When: "Chrome is being controlled by automated test
        # pyautogui.moveTo(self.hit_point[0], self.hit_point[1])  # Test

    def __get_pixel_color_imagegrab(self) -> Image:
        # Capture a region around the pixel (1x1)
        pixel_region = (self.hit_point[0], self.hit_point[1], self.hit_point[0] + 1,
                        self.hit_point[1] + 1)

        # Capture the pixel
        pixel_image = ImageGrab.grab(bbox=pixel_region)

        # Get the color of the pixel
        pixel_color = pixel_image.getpixel((0, 0))
        return pixel_color

    # def __get_pixel_color(self, hwnd, x, y):
    #     """Get the color of a pixel in an inactive window."""
    #     pixel_color = win32gui.GetPixel(hwnd, x, y)
    #
    #     # Check for CLR_INVALID.
    #     if pixel_color == win32con.CLR_INVALID:
    #         print("The pixel is outside the current clipping region.")
    #     else:
    #         # Print the red, green, and blue components of the pixel color.
    #         print("Red:", pixel_color >> 16 & 0xFF)
    #         print("Green:", pixel_color >> 8 & 0xFF)
    #         print("Blue:", pixel_color & 0xFF)

    def save_pdf(self, full_path, make_directory=True, file_overwrite=True):
        try:
            # Set the process as the foreground window
            # Could also save the window Z-order and reset here.
            # win32gui.SetForegroundWindow(hwnd)
            self.pdf_view_save_full_path = full_path

            self.__update_pdf_view_hotspot()  # Browser window can move.

            # Wait for unload of previous pdf after first load
            if self.__pdf_view_is_initialized:
                while self.__refresh_pdf_view_status() == "Loaded":
                    sleep(0.1)

            while self.__refresh_pdf_view_status() != "Loaded":
                sleep(0.1)
            self.pdf_view_is_initialized = True

            self.__click_pdfview_download_button()

            obj = SaveAsCommonDlg()
            obj.save_as_window_interact(self.pdf_view_save_full_path, make_directory, file_overwrite)

            # os.chdir(saved_working_directory)
        except Exception as e:
            _error_message(e, inspect.currentframe())

    def __click_pdfview_download_button_javascript(self):
        """Tested as a way to remove the mouse move"""
        """Not working"""
        self.__update_pdf_view_hotspot()
        try:  # Click the Download button with JavaScript
            # # driver = webdriver.Chrome()
            # # element = _get_element(driver, By.CSS_SELECTOR, "div[class='pdf-container expanded-width']")
            # # element = _get_element(driver, By.CSS_SELECTOR, "iframe[class='pdfView']")
            # driver = self.pdf_view_element.parent
            # # contentWindow = self.pdf_view_element.contentWindow
            #
            # x_coord = 1591
            # y_coord = 379
            # driver.execute_script("document.createEvent('MouseEvent').initMouseEvent('click', true, true, window, 1, 0, 0, 1573, 237, false, false, false, false, 1, null);")
            # # script = '''
            # # var event = document.createEvent('MouseEvent');
            # # document.createEvent('MouseEvent').initMouseEvent('click', true, true, window, 1, 0, 0, 1573, 237, false, false, false, false, 1, null);
            # # return event;
            # # '''
            #
            # # driver.execute_script(f"document.elementFromPoint({x_coord}, {y_coord}).click();")

            return
        except Exception as e:
            _error_message(e, inspect.currentframe())

    def __click_pdfview_download_button_actionchains(self):
        """Tested as a way to remove the mouse move"""
        """Works the first time through, however not the second - not sure why?"""
        self.__update_pdf_view_hotspot()
        try:  # Click the Download button
            driver = self.pdf_view_element.parent
            # element = _get_element(driver, By.CSS_SELECTOR, "div[class='pdf-container expanded-width']")
            # element = _get_element(driver, By.CSS_SELECTOR, "iframe[class='pdfView']")

            # win32gui.SetActiveWindow(self.hwnd_parent)
            # driver.switch_to.frame(self.pdf_view_element)
            actions = ActionChains(driver)
            actions.move_to_element_with_offset(self.pdf_view_element, 400, -454)
            actions.click()
            actions.perform()  # Works!
            # https://www.selenium.dev/selenium/docs/api/py/webdriver/selenium.webdriver.common.action_chains.html
            # driver.switch_to.default_content()

            return
        except Exception as e:
            _error_message(e, inspect.currentframe())

    def __click_pdfview_download_button_ctypes_send_input(self):
        # """Tested as a way to remove the mouse move"""
        # """Not working"""
        # import ctypes
        # import win32con
        # # Define constants for mouse events
        # # win32con.MOUSEEVENTF_LEFTDOWN
        # # MOUSEEVENTF_LEFTDOWN = 0x0002
        # # win32con.MOUSEEVENTF_LEFTUP
        # # MOUSEEVENTF_LEFTUP = 0x0004
        #
        # # Define the INPUT structure
        # class INPUT(ctypes.Structure):
        #     _fields_ = [
        #         ("type", ctypes.c_ulong),
        #         ("mi", ctypes.c_ulong * 3)
        #     ]
        #
        # # Define the input event
        # input_event = INPUT()
        # input_event.type = 0x0000  # INPUT_MOUSE
        # input_event.mi[0] = 1591  # dx
        # input_event.mi[1] = 379  # dy
        # input_event.mi[2] = win32con.MOUSEEVENTF_LEFTDOWN  # dwFlags
        #
        # # Send the mouse click event
        # ctypes.windll.user32.SendInput(1, ctypes.pointer(input_event), ctypes.sizeof(INPUT))
        #
        # # Release the mouse click event (optional)
        # input_event.mi[2] = win32con.MOUSEEVENTF_LEFTUP
        # ctypes.windll.user32.SendInput(1, ctypes.pointer(input_event), ctypes.sizeof(INPUT))
        return

    def __click_pdfview_download_button(self):
        self.__update_pdf_view_hotspot()
        try:  # Click the Download button
            saved_x, saved_y = pyautogui.position()
            # win32gui.SetActiveWindow(hwnd)
            pyautogui.click(self.hit_point[0], self.hit_point[1])
            pyautogui.moveTo(saved_x, saved_y)
            return
        except Exception as e:
            _error_message(e, inspect.currentframe())

    def __init__(self, hwnd_parent, pdf_iframe):
        self.hwnd_parent = hwnd_parent
        self.pdf_view_element = pdf_iframe
        # self.pdf_view_save_full_path = pdf_full_path
        return

    # def __del__(self):
    #     """The destructor for the class. Not Usually required for Python classes."""
    #     # if self.file_handle:
    #     #     self.file_handle.close()

    @property  # Get
    def hwnd_parent(self):
        return self.__hwndParent

    @hwnd_parent.setter  # Set
    def hwnd_parent(self, value):
        self.__hwndParent = value

    @property  # Get
    def pdf_view_status(self):
        return self.__pdf_view_status[0]

    @pdf_view_status.setter  # Set
    def pdf_view_status(self, value):
        self.__pdf_view_status[0] = value

    @property  # Get
    def pdf_view_element(self):
        return self.__pdf_view_element

    @pdf_view_element.setter  # Set
    def pdf_view_element(self, value):
        # element = _get_element(driver, By.CSS_SELECTOR, "iframe[class='pdfView']")
        self.__pdf_view_element = value

    @property  # Get
    def pdf_view_save_full_path(self):
        return self.__pdf_view_save_full_path

    @pdf_view_save_full_path.setter  # Set
    def pdf_view_save_full_path(self, value):
        self.__pdf_view_save_full_path = value

    @property  # Get
    def hit_point(self):
        return self.__xy_hit_point

    @hit_point.setter  # Set
    def hit_point(self, value):
        self.__xy_hit_point = value

    @property  # Get
    def pdf_view_is_initialized(self):
        return self.__pdf_view_is_initialized

    @pdf_view_is_initialized.setter  # Set
    def pdf_view_is_initialized(self, value):
        self.__pdf_view_is_initialized = value

    # __all__ = [save_pdf]


"""
# Example usage:
# if __name__ == "__main__":
#     file_url = "https://example.com/yourfile.pdf"  # Replace with the actual file URL
# noinspection SpellCheckingInspection
#     download_path = os.path.join(os.get_cwd(), "downloads")
#
#     if not os.path.exists(download_path):
#         os.makedirs(download_path)
#
#     downloader = FileDownloader(driver=None)  # You can pass a WebDriver instance here
#
#     downloaded_file = downloader.download_file(file_url, download_path)
#     print(f"Downloaded file: {downloaded_file}")
"""
