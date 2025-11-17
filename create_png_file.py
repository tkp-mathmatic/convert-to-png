import os
import io
import datetime
import shutil

import cv2
import numpy as np
from PIL import Image
from pdf2image import convert_from_path

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload


# ================================
# Google Drive API 関連の関数
# ================================

def get_drive_service():
    """
    サービスアカウントの credentials.json を使って
    Google Drive API のクライアントを作成する。
    """
    # GitHub Actions 側で、GOOGLE_CREDENTIALS を credentials.json に書き出しておく前提
    credentials_file = "credentials.json"

    creds = Credentials.from_service_account_file(
        credentials_file,
        scopes=["https://www.googleapis.com/auth/drive"]
    )
    service = build("drive", "v3", credentials=creds)
    return service


def list_pdfs_in_folder(service, folder_id):
    """
    指定したフォルダID配下の PDF ファイル一覧を取得する。
    """
    query = (
        f"'{folder_id}' in parents "
        "and mimeType='application/pdf' "
        "and trashed = false"
    )
    results = service.files().list(
        q=query,
        fields="files(id, name)"
    ).execute()
    return results.get("files", [])


def download_pdf(service, file_id, dest_path):
    """
    Drive 上の PDF をローカルにダウンロードする。
    """
    request = service.files().get_media(fileId=file_id)
    with io.FileIO(dest_path, "wb") as fh:
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            status, done = downloader.next_chunk()
            # 進捗を見たい場合は以下をコメントアウト解除
            # print(f"Download {int(status.progress() * 100)}%.")


def upload_png(service, local_path, new_name, parent_folder_id):
    """
    ローカルの PNG ファイルを Drive の指定フォルダにアップロードする。
    """
    file_metadata = {
        "name": new_name,
        "parents": [parent_folder_id]
    }
    media = MediaFileUpload(local_path, mimetype="image/png")
    file = service.files().create(
        body=file_metadata,
        media_body=media,
        fields="id"
    ).execute()
    return file["id"]


# ================================
# PDF → 縦長PNG 変換クラス
# ================================

class QPngCreator:
    """
    Qpdfをpngに変換するためのクラス
    （複数ページのPDFを、縦に長い1枚のPNGにする）
    """

    def __init__(self, resize_flg, output_path, v_width, h_width, logpath=None):
        """
        :param resize_flg: 余白カットをするかどうか（True/False）
        :param output_path: PNG を出力するローカルフォルダ
        :param v_width: 縦長のときの出力横幅(px)
        :param h_width: 横長のときの出力横幅(px)
        :param logpath: ログファイルのパス。指定がなければ output_path/log.csv
        """
        self.BASE_PATH = output_path  # ここでは output_path をベースとして使う
        self.TEMP_PATH = os.path.join(self.BASE_PATH, "temp")
        self.resize_flg = resize_flg

        # 出力フォルダを作成
        self.output_path = output_path
        if not os.path.isdir(self.output_path):
            os.makedirs(self.output_path, exist_ok=True)

        # ログファイルの設定
        if logpath:
            self.logpath = logpath
        else:
            self.logpath = os.path.join(self.output_path, "log.csv")

        # png仕様による最大サイズ（縦の最大ピクセル数）
        self.SIZE_MAX = 65535

        # dpi設定
        self.dpi = 350

        # 出力横幅の設定（縦長用・横長用）
        self.V_WIDTH = v_width
        self.H_WIDTH = h_width
        self.output_width = self.V_WIDTH  # 初期値は縦長想定

        # とりあえず作成する高さ（後で切り詰める）
        self.TEMP_HEIGHT = 100000
        self.pdf_path = None
        self.png_list = []

    def _write_log(self, target, message):
        """ログをCSV形式で1行追記する関数"""
        now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        line = f"{now_str},{target},{message}\n"
        with open(self.logpath, mode="a", encoding="utf-8") as f:
            f.write(line)

    def _cv2_read(self, filename, flags=cv2.IMREAD_COLOR, dtype=np.uint8):
        """
        日本語ファイル名にも対応した画像読み込み関数
        （cv2.imreadの代わりに使う）
        """
        n = np.fromfile(filename, dtype)
        img = cv2.imdecode(n, flags)
        return img

    def _create_png(self):
        """
        PDF（self.pdf_path）をページごとのPNGに分解して
        一時フォルダ（TEMP_PATH）に保存する。
        """
        # 一時フォルダを作り直す
        if os.path.isdir(self.TEMP_PATH):
            shutil.rmtree(self.TEMP_PATH)
        os.makedirs(self.TEMP_PATH, exist_ok=True)

        # PDF → 各ページの画像（Pillow Image のリスト）
        pages = convert_from_path(self.pdf_path, dpi=self.dpi, fmt="png", thread_count=4)

        self.png_list = []
        for i, page in enumerate(pages):
            filepath = os.path.join(self.TEMP_PATH, f"_{i+1:02d}.png")
            self.png_list.append(filepath)
            page.save(filepath, "PNG")

            # 余白をカットするモードの場合
            if self.resize_flg:
                page = Image.open(filepath)
                width, height = page.size

                # ここは元コードのクロップ範囲を踏襲
                left = 250
                top = 350
                right = width - 250
                bottom = height - 350

                page_crop = page.crop((left, top, right, bottom))
                page_crop.save(filepath)

        self.png_list = sorted(self.png_list)
        return self.png_list

    def _set_output_width(self, png_path):
        """
        最初のページのサイズから「縦長 or 横長」を判定して
        出力時の横幅（output_width）を決める。
        """
        h, w, _ = self._cv2_read(png_path).shape
        if h >= w:
            # 縦長
            self.output_width = self.V_WIDTH
        else:
            # 横長
            self.output_width = self.H_WIDTH

    def _create_board(self):
        """
        ベースとなる白紙（縦長キャンバス）を作る
        """
        self.base_img = np.zeros((self.TEMP_HEIGHT, self.output_width, 3), np.uint8)
        self.base_img.fill(255)  # 真っ白に塗る
        return self.base_img

    def _paste_image(self, png_path):
        """
        1ページ分の画像を、キャンバス（base_img）に貼り付ける。
        """
        page_img = self._cv2_read(png_path)

        # 横幅が output_width になるようにスケーリング（縦横比は維持）
        scaling_rate = self.output_width / page_img.shape[1]
        new_size = (self.output_width, round(page_img.shape[0] * scaling_rate))
        page_img = cv2.resize(page_img, new_size, interpolation=cv2.INTER_AREA)

        # TEMP_HEIGHT を超えるようなら False を返して終了
        if self.pasted_line + new_size[1] > self.TEMP_HEIGHT:
            return False

        # base_img に貼り付け
        self.base_img[self.pasted_line:self.pasted_line + new_size[1], :, :] = page_img
        self.pasted_line += new_size[1]
        return True

    def _save_png(self):
        """
        縦長PNGをファイルとして保存する。
        """
        if not os.path.isdir(self.output_path):
            os.makedirs(self.output_path, exist_ok=True)

        filepath = os.path.join(self.output_path, self.save_name + ".png")

        # OpenCV(BGR) → Pillow(RGB) に変換して保存
        img_rgb = cv2.cvtColor(self.base_img, cv2.COLOR_BGR2RGB)
        img_pil = Image.fromarray(img_rgb)
        img_pil.save(filepath, dpi=(self.dpi, self.dpi))

    def resize_image(self):
        """
        縦方向のサイズが PNG の上限を超える場合、
        上限まで縮小する。
        """
        if self.base_img.shape[0] < self.SIZE_MAX:
            return self.base_img

        scaling_rate = self.SIZE_MAX / self.base_img.shape[0]
        new_size = (round(self.base_img.shape[1] * scaling_rate), self.SIZE_MAX)
        self.base_img = cv2.resize(self.base_img, new_size, interpolation=cv2.INTER_AREA)
        return self.base_img

    def execute(self, pdf_path=None, png_list=None, save_name=None):
        """
        メイン実行関数：
        - PDFパスを渡された場合：PDF → ページごとPNG → 縦長PNG
        - PNGリストを渡された場合：PNGの束 → 縦長PNG
        """
        if pdf_path:
            # PDFから変換するモード
            self.save_name = os.path.basename(pdf_path)[:-4]  # 拡張子 .pdf を除く
            self.pdf_path = pdf_path
            self.png_list = self._create_png()
            target_name = self.save_name
        elif png_list and save_name:
            # 既にPNGがあるモード
            self.save_name = save_name
            self.pdf_path = None
            self.png_list = png_list
            target_name = save_name
        else:
            return False

        # キャンバス作成のための初期化
        self.pasted_line = 10
        self._set_output_width(self.png_list[0])
        self._create_board()

        over_flg = False

        for page_num, png_path in enumerate(self.png_list):
            if self.pasted_line >= self.SIZE_MAX and not over_flg:
                print(
                    f"WARNING on {os.path.basename(png_path)}: "
                    f"{page_num + 1}ページ目でサイズオーバーしたため、リサイズされます。"
                )
                over_flg = True

            result = self._paste_image(png_path)
            if not result:
                print(
                    f"ERROR on {os.path.basename(png_path)}: "
                    "処理できるサイズをオーバーしたため、このファイルのpng化を中止します。"
                )
                self._write_log(target_name, "処理可能サイズ超過")
                break

        # 実際に使った部分だけ切り出し
        self.base_img = self.base_img[:self.pasted_line, :]

        if over_flg:
            self.resize_image()
            self._write_log(target_name, "リサイズ処理済")
        else:
            self._write_log(target_name, "正常")

        # PNG保存
        self._save_png()

        # 一時フォルダ削除（PDFモードのときのみ）
        if self.pdf_path and os.path.isdir(self.TEMP_PATH):
            shutil.rmtree(self.TEMP_PATH)

        return True


# ================================
# メイン処理
# ================================

def main():
    """
    - 環境変数からフォルダIDや設定値を受け取る
    - DriveからPDFをダウンロード
    - QPngCreatorでPNG化
    - PNGをDriveにアップロード
    を一括で行う。
    """

    # ----- 環境変数から値を取得（GAS → GitHub Actions から渡される想定） -----
    input_folder_id = os.environ["INPUT_FOLDER_ID"]    # PDFが入っているDriveフォルダID
    output_folder_id = os.environ["OUTPUT_FOLDER_ID"]  # PNGを保存したいDriveフォルダID

    # V_WIDTH / H_WIDTH / RESIZE_FLG は、渡されなければデフォルト値を使う
    v_width = int(os.environ.get("V_WIDTH", "640"))
    h_width = int(os.environ.get("H_WIDTH", "1000"))
    resize_flg_str = os.environ.get("RESIZE_FLG", "false").lower()
    resize_flg = resize_flg_str == "true"

    # ローカルの作業用フォルダ
    work_dir = "./work"
    pdf_dir = os.path.join(work_dir, "pdf")
    png_dir = os.path.join(work_dir, "png")

    os.makedirs(pdf_dir, exist_ok=True)
    os.makedirs(png_dir, exist_ok=True)

    # ログファイルのパス
    log_path = os.path.join(work_dir, "log.csv")

    # Drive API クライアント作成
    service = get_drive_service()

    # PDF ファイル一覧を取得
    pdf_files = list_pdfs_in_folder(service, input_folder_id)
    print(f"Found {len(pdf_files)} pdf files in folder: {input_folder_id}")

    # PNG 変換クラスを初期化（出力先は png_dir）
    qpc = QPngCreator(
        resize_flg=resize_flg,
        output_path=png_dir,
        v_width=v_width,
        h_width=h_width,
        logpath=log_path
    )

    for file in pdf_files:
        pdf_id = file["id"]
        pdf_name = file["name"]  # 例: "A001.pdf"
        base_name, _ = os.path.splitext(pdf_name)

        local_pdf_path = os.path.join(pdf_dir, pdf_name)
        local_png_path = os.path.join(png_dir, f"{base_name}.png")

        print(f"Processing: {pdf_name}")

        # 1) Drive から PDF をローカルにダウンロード
        download_pdf(service, pdf_id, local_pdf_path)

        # 2) PDF → 縦長PNGに変換
        success = qpc.execute(pdf_path=local_pdf_path)
        if not success:
            print(f"Failed to convert: {pdf_name}")
            continue

        # 3) 生成された PNG を Drive にアップロード
        upload_png(service, local_png_path, f"{base_name}.png", output_folder_id)

        print(f"Uploaded PNG for: {pdf_name}")

    print("All done.")


if __name__ == "__main__":
    main()
