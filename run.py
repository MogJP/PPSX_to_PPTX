import subprocess
import sys
import os
import platform # OS判別のために追加
import threading # リアルタイム出力を扱うために追加
import queue # リアルタイム出力を扱うために追加
import time # timeモジュールを追加

# --- 設定 ---
# 実行したいLibreOfficeにバンドルされているPythonインタープリタへのパスを設定してください。
# ユーザーから指定されたパスをデフォルトとします。
LIBREOFFICE_PYTHON_PATH = r"C:\Program Files\LibreOffice\program\python.exe"

# 実行したい変換スクリプト (前回のコード) のパスを設定してください。
# 例: このスクリプトと同じディレクトリにある場合
SCRIPT_TO_RUN_PATH = os.path.join(os.path.dirname(__file__), "convert.py") # convert.py に合わせて修正

# --- スクリプトの実行 ---

def enqueue_output(out, queue):
    """
    指定されたストリーム (stdout/stderr) から行を読み込み、キューに入れるヘルパー関数。
    スレッドで使用されます。
    """
    for line in iter(out.readline, b''): # バイト列として読み込み
        queue.put(line)
    out.close()

def run_script_with_libreoffice_python_realtime(script_path, args):
    """
    指定されたLibreOffice Pythonインタープリタを使用して、別のスクリプトを実行し、
    出力をリアルタイムで表示します。

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
    command = [LIBREOFFICE_PYTHON_PATH, script_path] + args

    print(f"Running command: {' '.join(command)}")

    try:
        # subprocess.Popen() を使用してプロセスを起動
        # stdout=subprocess.PIPE, stderr=subprocess.PIPE で出力をパイプにリダイレクト
        # bufsize=1 で行バッファリングを有効にする (リアルタイムに近い出力のため)
        # universal_newlines=False または text=False (Python 3.7+) でバイト列として扱う
        # (リアルタイム読み込みではバイト列で扱い、デコードは読み込み側で行うのが安全)
        process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, bufsize=1)

        # リアルタイム出力のためのキューとスレッドを作成
        q_stdout = queue.Queue()
        q_stderr = queue.Queue()

        # stdoutとstderrを読み込むスレッドを開始
        t_stdout = threading.Thread(target=enqueue_output, args=(process.stdout, q_stdout))
        t_stderr = threading.Thread(target=enqueue_output, args=(process.stderr, q_stderr))
        t_stdout.daemon = True # メインスレッド終了時に一緒に終了
        t_stderr.daemon = True
        t_stdout.start()
        t_stderr.start()

        print("\n--- Script Output (Realtime) ---")

        # プロセスが実行中の間、キューから出力を読み込んで表示
        while process.poll() is None or not q_stdout.empty() or not q_stderr.empty():
            try:
                # キューからデータを非ブロックで取得
                line_stdout = q_stdout.get_nowait()
                print(line_stdout.decode(sys.stdout.encoding or 'utf-8').strip()) # デコードして表示
            except queue.Empty:
                pass # キューが空の場合は何もしない

            try:
                line_stderr = q_stderr.get_nowait()
                # 標準エラー出力は区別して表示しても良い
                print(f"ERR: {line_stderr.decode(sys.stderr.encoding or 'utf-8').strip()}") # デコードして表示
            except queue.Empty:
                pass

            time.sleep(0.01) # 短時間待機してCPU使用率を抑える

        # プロセスが終了し、すべての出力がキューから読み込まれたことを確認
        t_stdout.join()
        t_stderr.join()

        # プロセスの終了コードを取得
        returncode = process.wait()

        print(f"\n--- Script finished with return code: {returncode} ---")

        if returncode != 0:
             print(f"Script execution failed with return code: {returncode}")
             # 必要に応じて、エラーの詳細をさらに表示
             # 例: process.communicate() を使って残りの出力を取得 (ただし、上記で全て読み込んでいるはず)

    except FileNotFoundError:
        print(f"Error: The command or script was not found.")
    except Exception as e:
        # 修正点: 例外オブジェクト e を直接表示
        print(f"An unexpected error occurred during subprocess execution: {e}")
        # 念のため、元のトレースバックも表示
        import traceback
        traceback.print_exc()


# --- メインの実行部分 ---
if __name__ == "__main__":
    # コマンドライン引数を確認
    # 引数は このスクリプト名, 入力ファイルパス, 出力ファイルパス, パスワード の順を想定
    if len(sys.argv) < 4:
        print("Usage: python run_convert.py <input_ppsx_path> <output_pptx_path> <password>")
        sys.exit(1)

    # 実行したいスクリプトに渡す引数を取得 (このスクリプト名を除く)
    script_args = sys.argv[1:]

    # LibreOffice Pythonで実行するスクリプトのパスと引数を渡して実行関数を呼び出す
    run_script_with_libreoffice_python_realtime(SCRIPT_TO_RUN_PATH, script_args)

