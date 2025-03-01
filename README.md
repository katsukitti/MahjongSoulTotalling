## MahjongSoulTotalling
# Discordに雀魂の対局結果画像をアップすると自動的に以下の処理を行うPythonプログラムです。
# 1. Discordの対局結果画像を保存
# 2. 対局結果（1着～4着のプレイヤー名、素点、順位点）を抽出
# 3. 対局結果を集計用Excelに追記
# 4. 集計用Excelから通算成績を抽出
# 5. Discordに対局結果と通算成績を投稿
#
## How to
# 1. コードエディタなどで動作に必要なPythonライブラリを調べ、インストール
# 2. Tesseract-OCRをインストール（日本語の訓練データを含む）
# 3. DiscordのBOTを作成し、投稿したいDiscordサーバの管理権限を付与
# 4. .envファイルの各種パラメータを設定
# 5. python check_discord.py を実行
#
## Creation record
# 以下のnoteをご覧ください
# https://note.com/katsukitti/n/na644c00696e4
#
## License
# This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
