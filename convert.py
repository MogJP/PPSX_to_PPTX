import sys
import os
import platform # OS判別のために追加 (エラーメッセージ用)

# --- UNOモジュールのインポート ---
# このスクリプトはLibreOfficeのUNOライブラリを使用します。
# 'uno' モジュールがインポートできる環境 (LibreOfficeにバンドルされたPythonなど) で実行してください。
try:
    import uno
    import unohelper
except ImportError:
    print("Error: Could not import 'uno' or 'unohelper' module.")
    print("This script must be run with the Python interpreter bundled with LibreOffice,")
    print("or in a Python environment correctly configured to find the LibreOffice UNO SDK.")
    print("Example command (path to LibreOffice python depends on your installation):")
    current_platform = platform.system()
    if current_platform == "Windows":
         print("  'C:\\Program Files\\LibreOffice\\program\\python.exe' your_script_name.py <input_ppsx_path> <output_pptx_path> <password>")
    elif current_platform == "Linux":
         print("  /usr/lib/libreoffice/program/python your_script_name.py <input_ppsx_path> <output_pptx_path> <password>")
    elif current_platform == "Darwin": # macOS
         print("  /Applications/LibreOffice.app/Contents/Resources/python your_script_name.py <input_ppsx_path> <output_pptx_path> <password>")
    else:
         print("  (Please find the path to the python executable within your LibreOffice installation directory)")
    sys.exit(1) # UNOモジュールがない場合はスクリプトを終了

# --- UNO連携関数 ---

def get_uno_context():
    """
    Get the UNO component context from the current LibreOffice Python environment.
    This function assumes the script is run by the LibreOffice bundled Python.
    """
    print("Attempting to get UNO component context.")
    print("Calling uno.getComponentContext()...")
    try:
        local_context = uno.getComponentContext()
        print("uno.getComponentContext() returned.")
        print("UNO component context obtained successfully.")
        return local_context
    except Exception as e:
        print("Exception occurred during UNO context acquisition.")
        print(f"Error getting UNO context: {e}")
        print("Please ensure this script is executed by the LibreOffice bundled Python interpreter.")
        return None


def convert_ppsx_to_pptx(input_path, output_path, password):
    """
    Opens a password-protected PPSX file in LibreOffice Impress,
    enters the password, and saves it as a PPTX file.
    """
    print("Starting convert_ppsx_to_pptx function.")
    print("Attempting to get UNO context.")
    ctx = get_uno_context()
    print(f"UNO context obtained: {ctx}")

    if ctx is None:
        print("Could not obtain UNO context. Exiting conversion.")
        return

    print("Attempting to get ServiceManager and Desktop.")
    print("Calling ctx.getServiceManager()...")
    smgr = ctx.getServiceManager()
    print("ctx.getServiceManager() returned.")
    print("Calling smgr.createInstanceWithContext('com.sun.star.frame.Desktop', ctx)...")
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    print("smgr.createInstanceWithContext() returned.") # 追加
    print("ServiceManager and Desktop obtained.")

    # ファイルパスを UNO URL 形式に変換
    input_url = unohelper.systemPathToFileUrl(input_path)
    output_url = unohelper.systemPathToFileUrl(output_path)

    print(f"Input system path: {input_path}")
    print(f"Input UNO URL: {input_url}")
    print(f"Output system path: {output_path}")
    print(f"Output UNO URL: {output_url}")


    # ファイルを開くためのプロパティを設定
    load_properties = (
        uno.createUnoStruct("com.sun.star.beans.PropertyValue"),
        uno.createUnoStruct("com.sun.star.beans.PropertyValue")
    )
    load_properties[0].Name = "Password"
    load_properties[0].Value = password
    load_properties[1].Name = "ReadOnly"
    load_properties[1].Value = False

    # ロードプロパティの内容を表示 (パスワードは表示しない)
    print("Load properties:")
    for prop in load_properties:
        if prop.Name != "Password":
            print(f"  {prop.Name}: {prop.Value}")
        else:
            print(f"  {prop.Name}: ***")


    print(f"Attempting to load document from URL: {input_url}")
    try:
        # desktop.loadComponentFromURL() でドキュメントを開きます。
        # 第1引数: ファイルURL
        # 第2引数: フレーム名 ("" に変更)
        # 第3引数: フラグ (0でデフォルト)
        # 第4引数: プロパティ (パスワードなど)
        # LibreOffice Pythonで実行している場合、LibreOfficeのGUIが表示される可能性があります。
        # ヘッドレスモードで実行したい場合は、soffice --headless ... -env:python=/path/to/this/script.py のように
        # LibreOffice自体をヘッドレスで起動し、その環境でこのスクリプトを実行する方法があります。
        document = desktop.loadComponentFromURL(input_url, "", 0, load_properties) # Changed "_blank" to ""
        print("Document loaded successfully.")
    except Exception as e:
         print(f"Error loading document: {e}")
         document = None # エラー発生時はdocumentをNoneにする

    if document is None:
        print(f"Failed to open the document or incorrect password for: {input_url}")
        return

    print("Document is valid. Checking service type.")
    if not document.supportsService("com.sun.star.presentation.PresentationDocument"):
        print("Opened document is not a presentation document.")
        try:
            document.close(True)
        except Exception as e:
            print(f"Error closing non-presentation document: {e}")
        return
    print("Document is a presentation document.")

    # 保存するためのプロパティを設定
    save_properties = (
        uno.createUnoStruct("com.sun.star.beans.PropertyValue"),
    )
    save_properties[0].Name = "FilterName"
    save_properties[0].Value = "MS PowerPoint 2007 XML"

    print(f"Attempting to save document to: {output_url} as PPTX")
    try:
        document.storeToURL(output_url, save_properties)
        print("Document saved successfully.")
    except Exception as e:
        print(f"Error saving document: {e}")

    try:
        document.close(True)
        print("Document closed.")
    except Exception as e:
        print(f"Error closing document: {e}")


# --- スクリプトの実行部分 ---
if __name__ == "__main__":
    if len(sys.argv) != 4:
        print("Usage: Please run this script with the LibreOffice bundled Python interpreter.")
        print("Example: 'C:\\Program Files\\LibreOffice\\program\\python.exe' your_script_name.py <input_ppsx_path> <output_pptx_path> <password>")
        sys.exit(1)

    input_file = sys.argv[1]
    output_file = sys.argv[2]
    file_password = sys.argv[3]

    if not os.path.exists(input_file):
        print(f"Error: Input file not found at {input_file}")
        sys.exit(1)

    convert_ppsx_to_pptx(input_file, output_file, file_password)

    print("Conversion process finished.")
    print("Note: The LibreOffice application may remain running if not started in headless mode.")

