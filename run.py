import subprocess
import sys
import os
import platform # OS判別のために追加

# --- 設定 ---
# 実行したいLibreOfficeにバンドルされているPythonインタープリタへのパスを設定してください。
# ユーザーから指定されたパスをデフォルトとします。
LIBREOFFICE_PYTHON_PATH = r"C:\Program Files\LibreOffice\program\python.exe"

# 実行したい変換スクリプト (前回のコード) のパスを設定してください。
# 例: このスクリプトと同じディレクトリにある場合
SCRIPT_TO_RUN_PATH = os.path.join(os.path.dirname(__file__), "convert.py")

# --- スクリプトの実行 ---

def run_script_with_libreoffice_python(script_path, args):
    """
    指定されたLibreOffice Pythonインタープリタを使用して、別のスクリプトを実行します。

    Args:
        script_path (str): 実行したいPythonスクリプトのフルパス。
        args (list): 実行したいスクリプトに渡すコマンドライン引数のリスト (スクリプト名を除く)。
    """
    if not os.path.exists(LIBREOFFICE_PYTHON_PATH):
        print(f"Error: LibreOffice Python interpreter not found at '{LIBREOFFICE_PYTHON_PATH}'.")
        print("Please set the correct path in the script.")
        # OSに応じた一般的なパスのヒントを表示
        current_platform = platform.system()
        print("Common paths:")
        if current_platform == "Windows":
             print("  'C:\\Program Files\\LibreOffice\\program\\python.exe'")
             print("  'C:\\Program Files (x86)\\LibreOffice\\program\\python.exe'")
        elif current_platform == "Linux":
             print("  /usr/lib/libreoffice/program/python")
             print("  /opt/libreoffice*/program/python")
        elif current_platform == "Darwin": # macOS
             print("  /Applications/LibreOffice.app/Contents/Resources/python")
        return

    if not os.path.exists(script_path):
        print(f"Error: Script to run not found at '{script_path}'.")
        print("Please set the correct path in the script.")
        return

    # 実行コマンドを構築
    # 最初の要素は実行するインタープリタ、次が実行するスクリプト、その後に引数
    command = [LIBREOFFICE_PYTHON_PATH, script_path] + args

    print(f"Running command: {' '.join(command)}")

    try:
        # subprocess.run() を使用してコマンドを実行
        # capture_output=True で標準出力と標準エラー出力をキャプチャ
        # text=True で出力をテキストとして扱う (Python 3.7+)
        # check=True で、コマンドがゼロ以外の終了コードを返した場合にCalledProcessErrorを発生させる
        result = subprocess.run(command, capture_output=True, text=True, check=True)

        print("\n--- Script Output ---")
        print(result.stdout)

        if result.stderr:
            print("\n--- Script Errors (if any) ---")
            print(result.stderr)

        print(f"\nScript finished with return code: {result.returncode}")

    except FileNotFoundError:
        print(f"Error: The command or script was not found.")
    except subprocess.CalledProcessError as e:
        print(f"\n--- Script Execution Failed ---")
        print(f"Command: {e.cmd}")
        print(f"Return Code: {e.returncode}")
        print(f"Stdout: {e.stdout}")
        print(f"Stderr: {e.stderr}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")


# --- メインの実行部分 ---
if __name__ == "__main__":
    # このスクリプト自身のコマンドライン引数を取得
    # 引数は このスクリプト名, 入力ファイルパス, 出力ファイルパス, パスワード の順を想定
    if len(sys.argv) < 4:
        print("Usage: python run_script_with_lo_python.py <input_ppsx_path> <output_pptx_path> <password>")
        sys.exit(1)

    # 実行したいスクリプトに渡す引数を取得 (このスクリプト名を除く)
    # sys.argv[1:] は最初の引数から最後までのリスト
    script_args = sys.argv[1:]

    # LibreOffice Pythonで実行するスクリプトのパスと引数を渡して実行関数を呼び出す
    run_script_with_libreoffice_python(SCRIPT_TO_RUN_PATH, script_args)