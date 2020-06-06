# coding=utf-8
import tweepy
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, timedelta
import os

# 以下4つ「xxxxx」を、先ほど控えた値で書き換える。
CONSUMER_KEY = os.environ['CONSUMER_KEY']
CONSUMER_SECRET = os.environ['CONSUMER_SECRET']
ACCESS_TOKEN = os.environ['ACCESS_TOKEN']
ACCESS_TOKEN_SECRET = os.environ['ACCESS_TOKEN_SECRET']

auth = tweepy.OAuthHandler(CONSUMER_KEY, CONSUMER_SECRET)
auth.set_access_token(ACCESS_TOKEN, ACCESS_TOKEN_SECRET)
api = tweepy.API(auth)

# ↓探したいワードを指定する
SEARCH_WORD = 'Progate'
# ↓Twitterで検索する回数を指定する。回数が増えすぎるとTwitter APIのアクセス上限を超えてしまう可能性もあるので注意。
SEARCH_COUNT = 3

# ↓Spread Sheetへアクセスするための認証情報
PROJECT_ID = os.environ['PROJECT_ID']
PRIVATE_KEY_ID = os.environ['PRIVATE_KEY_ID']
PRIVATE_KEY = os.environ['PRIVATE_KEY']
CLIENT_EMAIL = os.environ['CLIENT_EMAIL']
CLIENT_ID = os.environ['CLIENT_ID']
CLIENT_X509_CERT_URL = os.environ['CLIENT_X509_CERT_URL']
# ↓JSONファイルを作成する時のテンプレートファイル名
TEMPLATE_FILE_NAME = 'spread_sheet_credential_template.txt'

CREDENTIAL_FILE_NAME = 'spread_sheet_credential.json'
SCOPE_URL = 'https://spreadsheets.google.com/feeds'
# ↓アクセスしたいSpread SheetのID
GID = os.environ['GID']
# ↓アクセスしたいSpread Sheetのシート名
SHEET_NAME = 'シート4'
TWEET_ID_COLUMN_NUM = 1
TEXT_COLUMN_NUM = 3
TIME_DIFFERENCE = 9
USER_ENTERED_OPTION = 'USER_ENTERED'
RAW_OPTION = 'RAW'

def update_tweets(gid, sheet_name, search_word, search_count):
    # Spread Sheetにアクセスする
    sheet = access_to_sheet(gid)
    worksheet = sheet.worksheet(sheet_name)

    # すでにsheetに書かれているtweet_idのリストを取得する
    tweet_ids = worksheet.col_values(TWEET_ID_COLUMN_NUM)
    # すでにsheetに値が書かれている場合も考慮して、空白が始まる最初の行番号を特定する。
    blank_start_num = len(tweet_ids) + 1
    # 一行目は列の名前が書いてあって、値ではないのでリストから除いておく。
    tweet_ids.pop(0)
    texts = worksheet.col_values(TEXT_COLUMN_NUM)
    texts.pop(0)

    max_id = None
    collected_tweets = []
    collected_texts = []
    lower_search_name = search_word.lower()
    # search_countの回数だけ、検索APIを叩く
    for i in range(0, search_count):
        print('#########################')
        print('「{}」での{}回目のツイート検索開始'.format(search_word, i + 1))
        # 2度目以降の検索では一つ前の検索以降のつぶやきを検索していく
        if max_id is None:
            search_results = api.search(q=search_word, count=100, lang='ja', result_type='mixed')
        else:
            search_results = api.search(q=search_word, count=100, lang='ja', max_id=max_id, result_type='mixed')

        for status in search_results:
            # 文章が重複しているツイートが増えるとノイズになるので、以下のような不要な重複ツイートは除く。
            # ・引用文章なしのリツイート
            # ・取得済み文章に（前方30文字が）同じつぶやきが含まれている
            # ・アカウント名/ユーザー名に検索語が含まれているためにヒットしまっている
            user_info = status.user
            text = status.text
            pre_text = text[:30]
            lower_screen_name = user_info.screen_name.lower()
            lower_name = user_info.name.lower()
            if 'RT @' in text or pre_text in [text[:30] for text in collected_texts]\
                    or (lower_search_name in lower_screen_name or lower_search_name in lower_name):
                continue
            # 重複したデータを弾くため、重複チェックで利用するデータをリストに詰めておく。
            collected_texts.append(text)
            collected_tweets.append(status)

        # 一番古いツイートよりも古いツイートを探すため、一番古いツイートID取得。検索結果が0件の場合は検索処理のループ終了
        if len(search_results) > 1:
            max_id = search_results[-1].id
        else:
            break

    # SpreadSheetに1列ずつ新規ツイートのデータを一括書き込みしていく
    end_cell_num = blank_start_num + len(collected_tweets) - 1
    if blank_start_num < end_cell_num:
        added_count = register_tweets(worksheet, blank_start_num, end_cell_num, collected_tweets)
        print('シートに{}個のツイート情報を新規追加しました。'.format(added_count))
    else:
        print('シートへの追加対象ツイートは0個だったぽい')

    # SpreadSheetにアクセスするための認証ファイルを削除する
    os.remove(CREDENTIAL_FILE_NAME)

def access_to_sheet(gid):
    # 書き込み用ファイル
    credential_file = open(CREDENTIAL_FILE_NAME, 'w')

    # テンプレートファイルを読み込み、書き込みファイルに書き込み
    template_file = open(TEMPLATE_FILE_NAME)
    template_file_lines = template_file.readlines()
    line_num = 0
    for line in template_file_lines:
        line_num += 1
        if line_num == 3:
            line = line.format(PROJECT_ID)
        elif line_num == 4:
            line = line.format(PRIVATE_KEY_ID)
        elif line_num == 5:
            line = line.format(PRIVATE_KEY)
        elif line_num == 6:
            line = line.format(CLIENT_EMAIL)
        elif line_num == 7:
            line = line.format(CLIENT_ID)
        elif line_num == 11:
            line = line.format(CLIENT_X509_CERT_URL)
        credential_file.writelines(line)
    template_file.close()
    credential_file.close()

    credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIAL_FILE_NAME, SCOPE_URL)
    client = gspread.authorize(credentials)
    return client.open_by_key(gid)

def register_tweets(sheet, start_row, end_row, tweets):
    id_list = []
    url_dict = {}
    text_dict = {}
    tweet_datetime_dict = {}
    tweet_fav_dict = {}
    retweet_dict = {}
    account_name_dict = {}
    user_name_dict = {}
    follower_count_dict = {}
    follow_count_dict = {}
    row_num = start_row
    now = slushed_current_datetime()
    added_count = 0
    for tweet in tweets:
        id_str = tweet.id_str
        id_list.append(id_str)
        # データが存在しない場合にエラー発生する可能性があるので、例外処理してそのデータは無視する。
        try:
            t_user = tweet.user
            account_name = t_user.screen_name
            account_name_dict[id_str] = account_name
            url_dict[id_str] = 'https://twitter.com/' + account_name + '/status/' + id_str
            text_dict[id_str] = tweet.text
            tweet_datetime_dict[id_str] = slushed_datetime(add_hour_from(tweet.created_at, TIME_DIFFERENCE))
            tweet_fav_dict[id_str] = tweet.favorite_count
            retweet_dict[id_str] = tweet.retweet_count
            user_name_dict[id_str] = t_user.name
            follower_count_dict[id_str] = t_user.followers_count
            follow_count_dict[id_str] = t_user.friends_count
            added_count += 1
        except Exception as e:
            print("id_str:{}, 例外args:{}".format(id_str, e.args))
        row_num += 1

    update_cells_with_list(sheet, 'A' + str(start_row), 'A' + str(end_row), id_list, value_input_option=RAW_OPTION)
    update_cells(sheet, 'B' + str(start_row), 'B' + str(end_row), id_list, url_dict, value_input_option=RAW_OPTION)
    update_cells(sheet, 'C' + str(start_row), 'C' + str(end_row), id_list, text_dict, value_input_option=RAW_OPTION)
    update_cells(sheet, 'D' + str(start_row), 'D' + str(end_row), id_list, tweet_datetime_dict, value_input_option=USER_ENTERED_OPTION)
    update_cells(sheet, 'E' + str(start_row), 'E' + str(end_row), id_list, tweet_fav_dict, value_input_option=RAW_OPTION)
    update_cells(sheet, 'F' + str(start_row), 'F' + str(end_row), id_list, retweet_dict, value_input_option=RAW_OPTION)
    update_cells(sheet, 'G' + str(start_row), 'G' + str(end_row), id_list, account_name_dict, value_input_option=RAW_OPTION)
    update_cells(sheet, 'H' + str(start_row), 'H' + str(end_row), id_list, user_name_dict, value_input_option=RAW_OPTION)
    update_cells(sheet, 'I' + str(start_row), 'I' + str(end_row), id_list, follower_count_dict, value_input_option=RAW_OPTION)
    update_cells(sheet, 'J' + str(start_row), 'J' + str(end_row), id_list, follow_count_dict, value_input_option=RAW_OPTION)
    update_cells_by_value(sheet, 'K' + str(start_row), 'K' + str(end_row), now, value_input_option=USER_ENTERED_OPTION)

    return added_count

def update_cells_with_list(sheet, from_cell, to_cell, list, value_input_option):
    """
    リストの値をシートに書き込んでいく。
    :param sheet: スプレッドシートの特定のワークシート
    :param from_cell: 書き込みを始めるセル
    :param to_cell:  書き込みを終了するセル
    :param list: 値のリスト
    :param value_input_option: RAW:文字列(ex.「=1+1」と入力すると「=1+1」になる)。USER_ENTERED:関数や数値など(ex.「=1+1」と入力すると「2」になる)
    :return: 更新したセル数
    """
    cell_list = sheet.range('{}:{}'.format(from_cell, to_cell))
    count_num = -1
    for cell in cell_list:
        count_num += 1
        try:
            val = list[count_num]
        except Exception as e:
            continue
        if val is None:
            continue
        cell.value = val
    print('{}から{}まで書き込むよ'.format(from_cell, to_cell))
    sheet.update_cells(cell_list, value_input_option=value_input_option)

def update_cells(sheet, from_cell, to_cell, id_list, dict, value_input_option):
    """
    {id: value}の辞書の値をシートに書き込んでいく。id_listの値を持っているdictのvalueをsheetのidの行に書き込んでいく。
    開始セルの位置=id_listの最初のid, 終了セルの位置=id_listの最後のid で対応している必要あり。
    :param sheet: スプレッドシートの特定のワークシート
    :param from_cell: 書き込みを始めるセル
    :param to_cell:  書き込みを終了するセル
    :param id_list: IDのリスト
    :param dict: {id: value}の辞書
    :param value_input_option: RAW:文字列(ex.「=1+1」と入力すると「=1+1」になる)。USER_ENTERED:関数や数値など(ex.「=1+1」と入力すると「2」になる)
    :return: 更新したセル数
    """
    cell_list = sheet.range('{}:{}'.format(from_cell, to_cell))
    count_num = -1
    updated_num = 0
    for cell in cell_list:
        count_num += 1
        try:
            val = dict[id_list[count_num]]
        except Exception as e:
            continue
        if val is None:
            continue
        cell.value = val
        updated_num += 1
    print('{}から{}まで書き込むよ'.format(from_cell, to_cell))
    sheet.update_cells(cell_list, value_input_option=value_input_option)
    return updated_num

def update_cells_by_value(sheet, from_cell, to_cell, value, value_input_option):
    """
    指定した行から指定した行まで、一括で一つのvalueに更新する。
    :param sheet: Spread Sheet
    :param from_cell: 開始セル
    :param to_cell: 終了セル
    :param value: 入力値
    :param value_input_option: RAW:文字列(ex.「=1+1」と入力すると「=1+1」になる)。USER_ENTERED:関数や数値など(ex.「=1+1」と入力すると「2」になる)
    """
    cell_list = sheet.range('{}:{}'.format(from_cell, to_cell))
    for cell in cell_list:
        cell.value = value
    print('{}から{}まで書き込むよ'.format(from_cell, to_cell))
    sheet.update_cells(cell_list, value_input_option=value_input_option)

def slushed_datetime(datetime):
    """
    return japanese style formatted datetime (ex. 2018/12/01 12:10:00)
    """
    return datetime.strftime('%Y/%m/%d %H:%M:%S')

def slushed_current_datetime():
    """
    return japanese style formatted current datetime (ex. 2018/12/01 12:10:00)
    """
    return datetime.now().strftime('%Y/%m/%d %H:%M:%S')

def add_hour_from(dt, hour):
    """
    :param dt: 指定日時(datetime)
    :param hour: 時間（int）
    :return: 今から指定時間後の日付
    """
    return dt + timedelta(hours=hour)

# ツイート検索してSpread Sheetに書き込む処理を実行
update_tweets(GID, SHEET_NAME, SEARCH_WORD, SEARCH_COUNT)
