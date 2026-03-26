## この関数の説明

## ローカルで PNG 化する方法

### 1. 依存関係を入れる

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

`pdf2image` を使う場合は `poppler` も必要です。

macOS:

```bash
brew install poppler
```

### 2. フォルダを作る

プロジェクト直下に `input` フォルダを作り、変換したい PDF を入れます。

```bash
mkdir -p input output
```

### 3. 実行する

```bash
python create_png_file.py
```

### 4. 出力先

- 入力: `input`
- 出力: `output`
- ログ: `output/log.csv`

`INPUT_FOLDER_ID` と `OUTPUT_FOLDER_ID` を指定しない場合は、自動でローカルモードになります。

必要なら環境変数で変更できます。

```bash
LOCAL_MODE=true \
LOCAL_INPUT_DIR=./input \
LOCAL_OUTPUT_DIR=./output \
V_WIDTH=640 \
H_WIDTH=1000 \
RESIZE_FLG=false \
python create_png_file.py
```
