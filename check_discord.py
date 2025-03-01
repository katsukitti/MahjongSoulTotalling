from discord.ext import commands
import discord
import subprocess
import os
from dotenv import load_dotenv
import aiohttp
from my_logger import set_logger, getLogger

load_dotenv()
set_logger()
logger = getLogger(__name__)

# 結果抽出プログラム
RESULT_PY = os.getenv('RESULT_PY')
INPUT_IMAGE = os.getenv('INPUT_IMAGE')

# Discordの権限設定
intents = discord.Intents.default()
intents.members = True # メンバー管理の権限
intents.message_content = True # メッセージの内容を取得する権限

bot = commands.Bot(
    command_prefix="$", # $コマンド名　でコマンドを実行できるようになる
    case_insensitive=True, # 大文字小文字を区別しない
    intents=intents # 権限を設定
)

# Discordにログイン
@bot.event
async def on_ready():
    print(f"Logged in as {bot.user.name}")
    logger.info(f"Logged in as {bot.user.name}")

# 投稿を確認
@bot.event
async def on_message(message: discord.Message):

    # BOTのメッセージは無視
    if message.author.bot:
        return
    
    # 投稿されたメッセージが特定のチャンネルIDの場合
    target_channel_id = os.getenv('TARGET_CHANNEL_ID')
    if message.channel.id == int(target_channel_id):

        logger.info(f"投稿メッセージ: {message}")

        # ファイルが添付されている場合
        if message.attachments:

            # 添付ファイルごとに処理
            for attachment in message.attachments:

                logger.info(f"添付ファイル: {attachment}")

                # 添付ファイルをダウンロード
                async with aiohttp.ClientSession() as session: 
                    async with session.get(attachment.url) as resp:
                        if resp.status == 200:
                            with open(INPUT_IMAGE, "wb") as f:
                                f.write(await resp.read())

                            try:
                                # 結果抽出プログラムを実行
                                command = ["python", RESULT_PY] 
                                proc = subprocess.run(command, capture_output=True, text=True)
                                print(f"標準出力: {proc.stdout}")
                                logger.info(f"標準出力: {proc.stdout}")
                                await message.reply(proc.stdout) # 標準出力のみDiscordに投稿
                            except subprocess.CalledProcessError as e:
                                print(f"実行エラー: {e.stderr}")
                                logger.error(f"実行エラー: {e.stderr}")
                            except Exception as ex:
                                print(f"予期せぬエラー: {str(ex)}")
                                logger.error(f"予期せぬエラー: {str(ex)}")

                            #os.remove(INPUT_IMAGE) # 一時ファイルを削除

    # on_message をオーバーライドすると、コマンドが実行されなくなる対策
    await bot.process_commands(message)

# BOTの実行
TOKEN = os.getenv('DISCORD_BOT_TOKEN')
bot.run(TOKEN)