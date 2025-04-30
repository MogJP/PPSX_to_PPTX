import uno
import unohelper
import sys
import os
import platform # OS判別のために追加
import time # 待機のために追加
import traceback # 例外情報の表示のために追加

# --- UNO例外クラスのインポート ---
from com.sun.star.uno import Exception as UnoException
from com.sun.star.connection import NoConnectException
from com.sun.star.io import IOException
from com.sun.star.lang import DisposedException, IllegalArgumentException
from com.sun.star.script import CannotConvertException
from com.sun.star.uno import RuntimeException

# --- UNO列挙型（Enum）のインポート ---
from com.sun.star.document.UpdateDocMode import NO_UPDATE, QUIET_UPDATE # NO_UPDATEをインポート

# --- 設定 ---
# 接続するOfficeインスタンスのUNO URLを設定します。
# LibreOffice/Apache OpenOffice共通
UNO_HOST = "localhost"
UNO_PORT = "2002"
UNO_URL = f"uno:socket,host={UNO_HOST},port={UNO_PORT};urp;StarOffice.ComponentContext"

# 接続リトライ設定
MAX_RETRY_ATTEMPTS = 15 # 最大リトライ回数
RETRY_DELAY_SECONDS = 2 # リトライ間の待機時間 (秒)

# --- UNO連携ヘルパー関数 ---

def UnoProps(**args):
    """
    UNOプロパティのタプルを作成するヘルパー関数 (unoconv.pyより)
    """
    props = []
    for key in args:
        prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        prop.Name = key
        prop.Value = args[key]
        props.append(prop)
    return tuple(props) # タプルを返す

# --- UNO連携クラス ---

class OfficeConverter:
    """
    Officeインスタンスへの接続とドキュメント変換処理を管理するクラス。
    unoconv.pyのConvertorクラスのパターンを参考に。
    """
    def __init__(self, uno_url):
        self.uno_url = uno_url
        self.context = None
        self.svcmgr = None
        self.desktop = None
        self.cwd = None

        # UNOコンテキストの取得を試みる (リトライ含む)
        self.context = self._get_uno_context_with_retry(self.uno_url)

        if self.context is None:
            raise ConnectionError("Failed to obtain UNO context after multiple retries.")

        # ServiceManagerとDesktopの取得
        print("Attempting to get ServiceManager and Desktop.")
        try:
            self.svcmgr = self.context.getServiceManager()
            print("ctx.getServiceManager() returned.")
            self.desktop = self.svcmgr.createInstanceWithContext("com.sun.star.frame.Desktop", self.context)
            print("smgr.createInstanceWithContext('com.sun.star.frame.Desktop', ctx) returned.")
            print("ServiceManager and Desktop obtained.")
        except Exception as e:
            raise RuntimeError(f"Error obtaining ServiceManager or Desktop: {e}") from e

        # 現在の作業ディレクトリをUNO URL形式で取得
        self.cwd = unohelper.systemPathToFileUrl(os.getcwd())
        print(f"Current working directory UNO URL: {self.cwd}")


    def _get_uno_context_with_retry(self, uno_url):
        """
        指定されたUNO URLで実行中のOfficeインスタンスに接続し、UNOコンポーネントコンテキストを取得します。
        接続できるまでリトライを行います。
        """
        print(f"Attempting to connect to UNO context at {uno_url}")

        for attempt in range(MAX_RETRY_ATTEMPTS):
            try:
                # PyUNOランタイムからローカルコンテキストを取得
                localContext = uno.getComponentContext()
                # UnoUrlResolverを作成
                resolver = localContext.getServiceManager().createInstanceWithContext(
                    "com.sun.star.bridge.UnoUrlResolver", localContext)

                # 実行中のOfficeインスタンスに接続
                print(f"Attempt {attempt + 1}/{MAX_RETRY_ATTEMPTS}: Resolving UNO URL...")
                ctx = resolver.resolve(uno_url)

                # 接続成功
                print("UNO URL resolved. Context obtained.")
                return ctx # 接続成功したらコンテキストを返して関数を終了

            except Exception as e:
                # 接続エラーが発生した場合の処理
                print(f"Attempt {attempt + 1}/{MAX_RETRY_ATTEMPTS}: Connection failed - {e}")
                print(f"Exception Type: {type(e)}") # 例外の型を表示

                if attempt < MAX_RETRY_ATTEMPTS - 1:
                    print(f"Retrying in {RETRY_DELAY_SECONDS} seconds...")
                    time.sleep(RETRY_DELAY_SECONDS)
                    print("Finished sleeping, attempting next retry.")
                else:
                    print("Max retry attempts reached.")
                    print(f"Final Error connecting to UNO context at {uno_url}: {e}")
                    return None # 最大リトライ回数に達したらNoneを返して関数を終了

        # この行は通常到達しないはずですが、念のためNoneを返します
        return None


    def convert(self, input_path, output_path, password=None): # パスワード引数をオプションに変更
        """
        指定された入力ファイル (PPSXまたはPPTX) を読み込み、PPTX形式で保存します。
        unoconv.pyのconvertメソッドのパターンを参考に。

        Args:
            input_path (str): 入力ファイルのフルパス (PPSXまたはPPTX)。
            output_path (str): 出力するPPTXファイルのフルパス。
            password (str, optional): 入力ファイルのパスワード。デフォルトはNone。
        """
        print(f"Starting conversion process for: {input_path}")

        document = None # ドキュメント変数を初期化
        phase = "loading" # 処理フェーズを示す

        try:
            # ファイルパスを UNO URL 形式に変換
            input_url = unohelper.systemPathToFileUrl(input_path)
            output_url = unohelper.systemPathToFileUrl(output_path)

            print(f"Input system path: {input_path}")
            print(f"Input UNO URL: {input_url}")
            print(f"Output system path: {output_path}")
            print(f"Output UNO URL: {output_url}")

            # ファイルを開くためのプロパティを設定 (unoconv.pyを参考にHidden=True, ReadOnly=Trueを追加)
            load_properties = UnoProps(
                Hidden=True,       # GUIに表示しない
                ReadOnly=True,     # 読み取り専用で開く (変換目的なら通常これで十分)
                UpdateDocMode=NO_UPDATE # リンク等の自動更新を無効化
            )

            # パスワードが指定されている場合のみPasswordプロパティを追加
            if password: # パスワードがNoneでないかチェック
                 load_properties += UnoProps(Password=password)

            # ロードプロパティの内容を表示 (パスワードは表示しない)
            print("Load properties:")
            for prop in load_properties:
                if prop.Name != "Password":
                    print(f"  {prop.Name}: {prop.Value}")
                else:
                    print(f"  {prop.Name}: ***")

            print(f"Attempting to load document from URL: {input_url}")
            # desktop.loadComponentFromURL() でドキュメントを開きます。
            # 第1引数: ファイルURL (UNO URL形式)
            # 第2引数: フレーム名 ("_blank" に変更 - 正常動作を確認)
            # 第3引数: フラグ (0でデフォルト)
            # 第4引数: プロパティ (パスワードなど)
            document = self.desktop.loadComponentFromURL(input_url, "_blank", 0, load_properties) # ここを修正

            print("desktop.loadComponentFromURL() called.") # 呼び出しが完了したことを示すログ
            print(f"Document object after loadComponentFromURL: {document}") # documentの値を表示

            if document is None:
                 # unoconv.pyではNoneが返る場合は例外を発生させている
                 # ここで正しいUNO例外クラスを使用
                 raise UnoException(f"Failed to load document from URL: {input_url}. loadComponentFromURL returned None.", self.context)

            print("Document loaded successfully.") # documentがNoneでない場合のみ表示される

            # ドキュメントが PresentationDocument であることを確認
            phase = "checking_doctype"
            if not document.supportsService("com.sun.star.presentation.PresentationDocument"):
                raise TypeError("Opened document is not a presentation document.")
            print("Document is a presentation document.")

            # 保存するためのプロパティを設定
            phase = "saving"
            # 'FilterName' で保存形式を指定する
            # PPTX 形式のフィルター名 ('MS PowerPoint 2007 XML') はApache OpenOffice/LibreOffice共通
            # Overwrite=True プロパティも追加
            save_properties = UnoProps(
                FilterName="MS PowerPoint 2007 XML", # PPTX 形式のフィルター名
                Overwrite=True # 既存ファイルを上書き
            )

            print(f"Attempting to save document to: {output_url} as PPTX using storeAsURL()...") # ログメッセージを修正
            # --- 修正点: storeToURL() を storeAsURL() に変更 ---
            # UnoPropsがタプルを返すため、tuple()で再度囲む必要はない
            document.storeAsURL(output_url, save_properties) # storeAsURLを使用
            # --- 修正点ここまで ---
            print("Document saved successfully.")

            phase = "disposing"
            # ドキュメントを閉じる (disposeはリソース解放、closeはドキュメント自体を閉じる)
            # unoconv.pyではdisposeしてからcloseしている
            try:
                if hasattr(document, 'dispose'):
                    document.dispose()
                    print("Document disposed.")
            except Exception as e:
                 print(f"Warning: Error during document dispose: {e}")

            try:
                if hasattr(document, 'close'):
                    document.close(True) # Trueで変更を破棄して閉じる (保存済みなので問題ない)
                    print("Document closed.")
            except Exception as e:
                 print(f"Warning: Error during document close: {e}")


        except ConnectionError as e:
             print(f"Fatal Connection Error: {e}")
             # 接続エラーはget_uno_contextで処理されるため、ここでは通常発生しないが念のため
             sys.exit(1) # 致命的なエラーとして終了

        except UnoException as e:
            # UNO例外の捕捉 (unoconv.pyのパターンを参考に)
            print(f"UNO Exception during {phase} phase:")
            if hasattr(e, 'Message'):
                 print(f"  Message: {e.Message}")
            if hasattr(e, 'ErrCode'):
                 print(f"  ErrCode: {e.ErrCode}")
            print(f"  Exception Type: {type(e)}")
            # traceback.print_exc() # 必要であれば詳細なトレースバックを表示 (冗長になる可能性あり)
            sys.exit(1) # エラーとして終了

        except Exception as e:
            # その他の予期しない例外
            print(f"An unexpected error occurred during {phase} phase: {e}")
            print(f"Exception Type: {type(e)}")
            # traceback.print_exc() # 必要であれば詳細なトレースバックを表示 (冗長になる可能性あり)
            sys.exit(1) # エラーとして終了

        finally:
            # 後処理 (ここでは特に何もしないが、リソース解放などが必要な場合に備えて)
            pass


# --- スクリプトの実行部分 ---
if __name__ == "__main__":
    # コマンドライン引数を確認
    # 引数は スクリプト名, 入力ファイルパス, 出力ファイルパス (+ パスワード) の順
    if len(sys.argv) not in [3, 4]: # 引数の数が3または4であることを許可
        print("Usage: python convert.py <input_file_path> <output_pptx_path> [password]") # Usageメッセージを更新
        print("Please ensure Office (LibreOffice or Apache OpenOffice) is running in UNO listening mode (e.g., soffice --accept=socket,host=localhost,port=2002;urp; --headless)")
        sys.exit(1)

    input_file = sys.argv[1]
    output_file = sys.argv[2]
    # 引数の数が4つの場合のみパスワードを取得
    file_password = sys.argv[3] if len(sys.argv) == 4 else None # 引数が4つの場合のみパスワードを取得

    if not os.path.exists(input_file):
        print(f"Error: Input file not found at {input_file}")
        sys.exit(1)

    # OfficeConverterクラスのインスタンスを作成し、変換を実行
    try:
        converter = OfficeConverter(UNO_URL)
        # パスワードをconvertメソッドに渡す (Noneの場合もある)
        converter.convert(input_file, output_file, file_password)
        print("Conversion process finished successfully.")
        sys.exit(0) # 成功終了
    except ConnectionError as e:
        print(f"Failed to connect to Office: {e}")
        print("Please ensure Office is running in UNO listening mode.")
        sys.exit(1) # 接続エラーで終了
    except Exception as e:
        print(f"An error occurred during the conversion process: {e}")
        # 上記のconvertメソッド内のtry-exceptで捕捉されなかった例外
        sys.exit(1) # その他のエラーで終了


# このスクリプトは実行中のOfficeインスタンスに接続しているため、
# スクリプトが終了してもOfficeプロセス自体は終了しません。
# Officeを終了するには手動で行うか、run.cmdなどの外部スクリプトで行う必要があります。
