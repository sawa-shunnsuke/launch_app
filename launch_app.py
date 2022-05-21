"""
追記：学校のpython自由課題のコードなのでおかしいところ等多いと思います。



gui側で実行したいアプリの配列の要素数を指定
バックエンド側で要素数が参照できる絶対パスのリストを作成、起動


呼び出す必要のある関数/関数の説明

launch_call_inter(絶対パスが入ったファイル名(○○.txt), gui?から渡されるデータ([0,1,2...]))
    txtファイルから絶対パスを取得し、guiから渡されるデータを参照し起動させる
    絶対パスが使用しているPC上に存在しない場合の処理はしてあるが、GUI処理で不都合があれば変更してもらって大丈夫です

write_path(絶対パスが入ったファイル名(○○.txt), 絶対パスの通称？と絶対パスが入った2次元リスト)
    2次元リストを「絶対パスの通称？,絶対パス」でtxtファイルに書き込む
    1つずつ追加する処理は任せてあるので一括で書き込むタイプです

read_path(絶対パスが入ったファイル名(○○.txt))
    「絶対パスの通称？,絶対パス」の形式で記入されているtxtファイルから2次元配列で取得する
    何かと便利に使えると思いますが、txtファイルが無い場合のエラー処理はしていないので注意

get_path()
    1度だけファイルダイアログ？を使用して選択したファイルから絶対パスを取得する
    何も選択しなかった場合文字列で「nofile」が返されるので、ループさせたい場合はそれで判断するといいです
"""

import subprocess as su
import webbrowser as we
import configparser as co
import os
import win32com.client
import sys
from tkinter import filedialog
import tkinter as tk
from tkinter import messagebox


# 引数のリストに入った絶対パスをもとにアプリを実行する関数
# 管理者権限が必要なアプリの実行はできないことがある？
# 現在はフォルダ・.exeファイル・一部のファイル（テスト不足、xslxファイルは開けた）・インターネットショートカットが対応している
# 絶対パス上にファイル又はドルだが存在していなかった場合のエラー処理はしていないので別途作った関数を仕様するように
# 一応private化？（関数名の前に__を追加するとprivateにできるらしい？）してあるので大丈夫だと思うが上記の理由があるため呼び出さないように
# print文はデバッグ用なのでコメント化してあるが消してもOK
def __launch_path(input_path_list):
    # print(input_path_list)
    # 引数のリスト分だけループさせる
    for i_p_l_num in range(len(input_path_list)):
        # print(repr(input_path_list[i_p_l_num]) + "を起動")
        # 絶対パスを元に選択したアプリ等の起動を開始
        # パスにそもそも拡張子が無い場合もフォルダとして認識
        if not "." in input_path_list[i_p_l_num]:
            # print("拡張子を認識しなかったのでフォルダとして起動")
            su.Popen(["explorer", input_path_list[i_p_l_num]], shell=True)
            continue
        try:
            # パスが.exeかどうか（.exe専用のがないとmicrosoftのアプリがきどうできない？)
            if ".exe" in input_path_list[i_p_l_num].lower():
                # print(".exeを認識したのでexeファイルとして起動")
                # 一部のmicrosoftアプリのショートカットは通常のexeでは開けないので先に処理をする

                # teamsの場合
                if r"Teams\Update.exe" and r"Teams/Update.exe" in input_path_list[i_p_l_num]:
                    # print("teamsと判断")
                    su.Popen(input_path_list[i_p_l_num] + ' --processStart "Teams.exe"', shell=True)
                    continue
                else:
                    # teams以外の場合
                    # Excel
                    if "xlicons.exe" in input_path_list[i_p_l_num]:
                        # print("excelと判断")
                        launch_app_key = win32com.client.Dispatch('Excel.Application')
                        launch_app_key.Visible = True
                        launch_app_key.Workbooks.Add()
                        continue
                    # Word
                    if "wordicon.exe" in input_path_list[i_p_l_num]:
                        # print("Wordと判断")
                        launch_app_key = win32com.client.Dispatch('Word.Application')
                        launch_app_key.Visible = True
                        launch_app_key.Documents.Add()
                        continue
                    # PowerPoint
                    if "pptico.exe" in input_path_list[i_p_l_num]:
                        # print("PowerPointと判断")
                        launch_app_key = win32com.client.Dispatch('PowerPoint.Application')
                        launch_app_key.Visible = True
                        launch_app_key.Presentations.Add()
                        continue
                    # microsoftアプリのショートカット以外（対応していないアプリもここに入る)
                    su.Popen(input_path_list[i_p_l_num], shell=True)
                    continue

            # インターネットショートカットの場合
            if ".url" in input_path_list[i_p_l_num].lower():
                # print(".urlを認識したのでwebブラウザとして起動")
                # ini形式でファイルを開く。interpolation=Noneの指定は本文参照。
                url_file = co.ConfigParser(interpolation=None)
                url_file.read(input_path_list[i_p_l_num])
                # URLを読んでデフォルトのブラウザーで表示する。
                url = url_file['InternetShortcut']['URL']
                we.open(url)
                continue

            # どれも当てはまらなかったら適当なファイルとして起動させる
            # print("適当なファイルとして起動")
            su.Popen(['start', input_path_list[i_p_l_num]], shell=True)
        except:
            print("ERROR!!!:ファイルが実行できません")


# リストに入っている絶対パス上にファイル又はフォルダが存在するかの確認
# 返却値が０だと全てのパス上にファイル又はフォルダが存在、0以外の整数だと存在していないとする
# 返却地は１０進数で返されるが、ファイルがあるかどうかは２進数に変換すると分かるようになる
# ２進数で０はファイルが存在、１はファイルが無いと判断されるようにする
# 返却値の２進数はリストの０番地が2進数の１とする（012の順番だとファイルの存在がttfなら4、tffなら6、fffなら7になる）
# pythonのint型の最大値はほぼないらしい？から２進数が多すぎてエラーになることは多分無いと思う？
# 尚2進数にした理由はなんかやってみたかったからである！
def __is_existence_path(input_path_list):
    # 返却値用の変数
    re_num = 0
    # 返却地に入れるための２進数変数
    binary_num = 1
    # リスト分だけループ
    for i_p_l_num in range(len(input_path_list)):
        # 絶対パス上にファイル又はフォルダが存在しない場合その位置に２進数の１を入れる（１０進数）
        if not os.path.exists(input_path_list[i_p_l_num]):
            re_num += binary_num
        # 2進数用の変数を１つ勧める
        binary_num *= 2
    return re_num


# is_existence_pathの返却値を使用し、ファイルが存在しなかったパスのリストの絶対パスを表示し、
# 存在しない絶対パスをリストから除外してlaunch_pathで起動させるか起動自体をさせないかユーザ側に決めさせる関数
# 尚この関数は仕様には無い関数なので必要ない場合launch_call_inter関数の最終行の処理を消せば使われなくなる
def __no_exist_exclusion_path(input_binary_n, input_path_list, input_name_list):
    # ２進数変換後用のリスト
    bin_exist_path_list = []
    # 除外する絶対パス用のリスト
    exist_path_name = []
    # 除外する絶対パスの名前のリスト（メッセージボックスに出力します)
    exist_name = []
    # 引数の１０進数の値を２進数に変換してリストに格納する
    while input_binary_n != 0:
        bin_exist_path_list.append(input_binary_n % 2)
        input_binary_n = input_binary_n // 2
    # print(bin_exist_path_list)
    # ２進数のリストをループで回す
    for bin_num in range(len(bin_exist_path_list)):
        # １（存在しない絶対パスの要素の場所）
        if bin_exist_path_list[bin_num] == 1:
            # 除外用の絶対パスリストに絶対パスを追加する
            exist_path_name.append(input_path_list[bin_num])
            exist_name.append(input_name_list[bin_num])
    # ユーザへの確認
    # print(*exist_path_name, sep='\n')
    # flg = input("以外のアプリケーションを起動しますか？(yes=y,no=anykey)>")
    ret = messagebox.askyesno('実行エラー', '\n'.join(exist_name) + '\nが絶対パス上にありません\nこれ以外のアプリケーションを起動しますか？')
    if ret:
        # 除外用のリスト分だけループする
        for e_p_n_num in range(len(exist_path_name)):
            # ユーザが存在しない絶対パス以外のパスのリストを起動すると選択した場合、リストから名前検索で要素を削除
            # remove関数は１つの要素しか削除しないが、除外用リストも同じ名前でも複数格納できるので問題なし
            input_path_list.remove(exist_path_name[e_p_n_num])
        # 存在しない絶対パスを削除したので起動する
        __launch_path(input_path_list)


# txtファイルから２次元リストに入力
# 改行文字でリストを分ける
def read_path(input_file):
    # 返却地用の変数
    read_data = []
    try:
        # ファイルを開く
        f = open(input_file, 'r', encoding='utf-8')
    except Exception:
        print("ファイルが開けませんでした：", str(input_file))
        # ファイル名が間違っている等でファイルが開けなかった場合異常終了としてシステムを強制終了させる
        sys.exit(1)
    # 要素ごとに処理（名前 絶対パス)
    for line in f:
        # 前後空白削除
        line = line.strip()
        # 末尾の\nの削除
        line = line.replace('\n', '')
        # 分割文字の指定
        line = line.split(",")
        # 名前と絶対パスの２個の要素をリストに追加
        read_data.append(line)
    f.close()
    return read_data


# txtファイルを読み込みgui？から受け取った値(0,1,2,3...)からパスを起動する？
# guiからの値はファイル起動のパターン(0,1,2...)のみらしいのでこちらでも絶対パスを読み込む必要がある
# read_pathで絶対パスのリストを読み込みfor文で回す
# for文の要素用の変数と受け取った値を照合し、合った場合はその要素に入っている絶対パスをリストに格納
# リストに格納した絶対パスの集まりをis_existence_pathで存在するか確認し、launch_pathで一括起動させる
def launch_call_inter(input_file, input_path_pattern):
    # 引数であるパスのパターン用の変数
    i_p_p_num = 0
    # 実行用のリスト
    launch_list = []
    # 実行時に絶対パス上にファイル等が無かった時用のリスト
    launch_list_name = []
    # txtファイルからの読み込み
    path_list = read_path(input_file)
    # print(path_list)
    # for文で読み込んだリストを回す
    for p_l_num in range(len(path_list)):
        # 読み込んだリストの行番号と引数が合致するか

        if p_l_num == input_path_pattern[i_p_p_num]:
            # 合致したら引数のリスト用変数に+１する
            i_p_p_num += 1
            # 絶対パスを実行用のリストに追加する
            # print(launch_list)
            launch_list.append(path_list[p_l_num][1])
            launch_list_name.append(path_list[p_l_num][0])
            # 引数のリストが一番最後まで行ったらこのfor文を抜ける
            if i_p_p_num >= len(input_path_pattern):
                break
    # 実行前に絶対パスが存在するかの確認
    if __is_existence_path(launch_list) == 0:
        # 全ての絶対パスが存在したら一括起動させる
        __launch_path(launch_list)
    else:
        bin_num = __is_existence_path(launch_list)
        # print("ERROR!:リストに存在しない絶対パスがあります")
        # リストに格納されている絶対パスのどれが存在していないかを表示
        # print(bin_num)
        # print("存在しない絶対パスは2進数で1として出力されます")
        # print(bin(bin_num))
        __no_exist_exclusion_path(bin_num, launch_list, launch_list_name)


# ファイル名には空白を入れることができるためカンマ区切りで出力する必要がある
# また、引数の2次元配列は2個になる
def write_path(input_file, in_path_pattern):
    # print(in_path_pattern)
    f = open(input_file, 'w')
    for in_path in in_path_pattern:
        f.write(in_path[0] + "," + in_path[1] + "\n")
    f.close()


# 一度だけファイルの絶対パスを探しに行く
# ファイルが選択されなかった場合、nullが返される
def get_path():
    # 一度だけファイルの絶対パスを取得する
    # ファイルパス検索時の最初の位置
    dir = r'C:\Users\user\Desktop'
    # C:\\
    # ファイルパスの拡張子の指定（これがないとたぶん.exeが表示できない。）
    filetype = [("全てのファイル", "*")]
    # 一度に大量に探しに行く必要があるため無限ループで回す。
    # ファイルパスを取得しに行く
    fld = tk.filedialog.askopenfilename(filetypes=filetype, initialdir=dir)
    if not fld:
        # print("ファイルが選択されませんでした")
        return "nofile"
    else:
        # print(fld)
        return fld


# /-------------ここから動作確認（gui側の処理等で使う場面）-----------------/

print("gui側の処理等を抜いた単体テストとなります。")


# テスト用のリスト
file = 'filepath.txt'

# testlistはtxtファイルへの出力用、launch_testlistはtxtファイルからの入力と実行用の変数
# testlistの要素数5と6の位置にあるものは存在しない絶対パスのチェック用です
testlist = [["???", r'C:\aaaa\aaaaa.aaa']]

s = input("初回起動ですか？yesの場合はy、noの場合はy以外を入力してください")
if s == 'y':
    # txtファイルへの書き込み(txtファイルが無い場合は自動で追加)
    write_path(file, testlist)
# 要素ごとに改行させて表示
print("現在のリストは\n")
print(testlist, sep='\n')
s = input("リストに絶対パスを追加する場合はy、実行させない場合はy以外を入力してください>")
if s == 'y':
    print("絶対パスを連続で取得します。終わる場合はキャンセルを押してください")
    path = get_path()
    while path != "nofile":
        name = input(path+"の名前を決めてください>")
        testlist.append([name, path])
        path = get_path()
    # txtファイルへの書き込み(txtファイルが無い場合は自動で追加)
    write_path(file, testlist)

# txtファイルから読み込み
launch_testlist = read_path('filepath.txt')
# 要素ごとに改行させて表示
print(*launch_testlist, sep='\n')

# リスト番後の入力
# print(len(launch_testlist))
launch_data = []
print("表示されたアプリケーションを実行させる場合はy、実行させない場合はy以外を入力してください>")
# dataには[0,1,2,3...]のように格納される
# 仮にプリセットを実装するのであれば入力値は[0,1,2,3...]としないと動かないので注意
for i in range(len(testlist)):
    s = input(testlist[i][0] + ">")
    if s == 'y':
        launch_data.append(i)

print(launch_data)
print(file)
# 一括起動
launch_call_inter(file, launch_data)
