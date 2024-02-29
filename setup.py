
import sys
from cx_Freeze import setup, Executable, IncludeFiles, Options

# Define the base (optional on non-Windows)
base = None
if sys.platform == "win32":
    base = "Win32GUI"

# Correct file paths and include necessary modules/data
includefiles = [
    (r"Images\app.ico", r"."),  # Include app icon in working directory
    (r"Images\bill.ico", r"."),   # Include bill icon in working directory
    (r"bills", r"."),             # Include entire "bills" folder
    (r"Kd_Images", r"."),         # Include entire "Kd_Images" folder
    (r"SK_Daten.xlsx", r"."),   # Include SK_Daten.xlsx
    (r"SK_Point.xlsx", r"."),    # Include SK_Point.xlsx
]

# Executable configuration (single executable with both icons)
executable = Executable(
    script="RegBillSys.py",
    base=base,
    icon=r"Images\app.ico",  # Set primary icon for executable
    includes=["tkinter", "matplotlib"],  # Include required Python modules (if applicable)
    shortcutName="RegBillSys",  # Name of the shortcut
    shortcutDir="desktop",  # Create a shortcut on the desktop
    shortcutFolder="RegBillSys",  # Shortcut folder name (if applicable)
)

# Setup options
setup(
    name="RegBillSys",
    version="0.1",
    description="RegBillSys application",
    author="ABC",
    options={"build_exe": {"include_files": includefiles}},
    executables=[executable],
)

#############################################################
# from cx_Freeze import setup,Executable,sys 
# includefiles=['Images\icon.ico','Images\app.ico']
# excludes=[]
# packages=[]
# base=None
# if sys.platform=="win32":
#     base="Win32GUI"
    
# shortcut_table=[
#     ("DesktopShortcut",
#      "DesktopFolder",
#      "RegBillSys",
#      "TARGETDIR",
#      "[TARGETDIR]\RegBillSys.exe",
#      None,
#      None,
#      None,
#      None,
#      None,
#      None,
#      "TARGETDIR",
#      )
# ]
# msi_data={"Shortcut":shortcut_table}

# bdist_msi_options={'data':msi_data}
# setup(
#     version="0.1",
#     description="RegBillSys",
#     author="ABC",
#     name="RegBillSys",
#     options={'build_exe':{'include_files':includefiles},'bdist_msi':bdist_msi_options,},
#     executables=[
#         Executable(
#             scrift="RegBillSys.PY",
#             base=base,
#             icon='Images/icon.ico',
#             icon='Images/app.ico',
#         )
#     ]
# )
# Pip install cx_Freeze
# Python setup.Py bdist_msi

# pyinstaller Exemple.py --onefile   --- lệnh này sẽ xuất hiện cửa sổ CMD khi chạy chương trình
# pyinstaller Exemple.py --onefile --windowed
# pyinstaller Exemple.py --onefile --noconsole   (Khi sử dụng flag --noconsole, bạn sẽ không thể sử dụng các chức năng nhập/xuất dữ liệu từ console.)
# pyinstaller Exemple.py --onefile --icon icon.ico  (Sử dụng flag --icon và --splash có thể giúp chương trình của bạn trông chuyên nghiệp hơn.)
# pyinstaller Exemple.py --onefile --splash splash.png
# cxfreeze Exemple.py -o Exemple.exe    (Cx_Freeze là một thư viện đóng gói khác cho Python. Nó có thể tạo ra các chương trình thực thi Windows mà không cần cửa sổ console.)

######################### lệnh đóng gói one file với pyinstaller -----

#  pyinstaller --onefile --windowed --icon=Images\app.ico --add-data "Images\app.ico;Images" --add-data "Images\bill.ico;Images" --add-data "bills;bills" --add-data "Kd_Images;Kd_Images" --add-data "SK_Daten.xlsx;." --add-data "SK_Point.xlsx;." RegBillSys.py

##################################

Có một số cách để bạn có thể chạy chương trình Python đóng gói bằng PyInstaller mà không xuất hiện cửa sổ đen cmd:

1. Sử dụng '--windowed' flag:

Thêm flag --windowed vào lệnh PyInstaller của bạn. Flag này sẽ tạo ra một cửa sổ console Windows ẩn khi chương trình khởi động.

Python
pyinstaller Exemple.py --onefile --windowed
Hãy thận trọng khi sử dụng các đoạn mã.
2. Sử dụng '--noconsole' flag:

Thêm flag --noconsole vào lệnh PyInstaller của bạn. Flag này sẽ ngăn chặn việc tạo ra cửa sổ console hoàn toàn.

Python
pyinstaller Exemple.py --onefile --noconsole
Hãy thận trọng khi sử dụng các đoạn mã.
3. Sử dụng '--icon' flag:

Thêm flag --icon cùng với đường dẫn đến file icon vào lệnh PyInstaller của bạn. Flag này sẽ tạo ra một biểu tượng cho chương trình của bạn, 
giúp bạn dễ dàng khởi động chương trình mà không cần lo lắng về cửa sổ console.

Python
pyinstaller Exemple.py --onefile --icon icon.ico
Hãy thận trọng khi sử dụng các đoạn mã.
4. Sử dụng '--splash' flag:

Thêm flag --splash cùng với đường dẫn đến file ảnh splash vào lệnh PyInstaller của bạn. 
Flag này sẽ hiển thị ảnh splash trong vài giây khi chương trình khởi động, che đi cửa sổ console.

Python
pyinstaller Exemple.py --onefile --splash splash.png
Hãy thận trọng khi sử dụng các đoạn mã.
5. Sử dụng cx_Freeze:

Cx_Freeze là một thư viện đóng gói khác cho Python. Nó có thể tạo ra các chương trình thực thi Windows mà không cần cửa sổ console.

Python
cxfreeze Exemple.py -o Exemple.exe
Hãy thận trọng khi sử dụng các đoạn mã.
Lưu ý:

Khi sử dụng flag --noconsole, bạn sẽ không thể sử dụng các chức năng nhập/xuất dữ liệu từ console.
Sử dụng flag --icon và --splash có thể giúp chương trình của bạn trông chuyên nghiệp hơn.
