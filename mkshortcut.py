#%% 
# initialize
import comtypes.client
import os

SAVE_DIR_PARENT = r"C:\Users\aaa\AppData\Roaming\Microsoft\Windows\Start Menu\Programs"
SAVE_DIR = os.path.join(SAVE_DIR_PARENT, "FreeSofts")
if not os.path.exists(SAVE_DIR):
	os.mkdir(SAVE_DIR)
	print("Target folder was created")
else :
	print("Target folder already exists")
#%%
# Define making shortcut function
def mkshortcut(target_file):
	basename = os.path.basename(target_file)
	filename = os.path.splitext(basename)[0]
	save_path=os.path.join(SAVE_DIR_PARENT, filename+ ".lnk")
	#WSHを生成
	wsh=comtypes.client.CreateObject("wScript.Shell",dynamic=True)
	#ショートカットの作成先を指定して、ショートカットファイルを開く。作成先のファイルが存在しない場合は、自動作成される。
	short=wsh.CreateShortcut(save_path)
	#以下、ショートカットにリンク先やコメントといった情報を指定する。
	#リンク先を指定
	short.TargetPath=target_file
	#コメントを指定する
	short.Description="テストショートカット"
	#引数を指定したい場合は、下記のコメントを解除して、引数を指定する。
	#short.arguments="/param1"
	#アイコンを指定したい場合は、下記のコメントを解除してアイコンのパスを指定する。
	#short.IconLocation="C:\\test\\test.ico"
	#作業ディレクトリを指定したい場合は、下記のコメントを解除してディレクトリのパスを指定する。
	#short.workingDirectory="c:\\test\\"
	#ショートカットファイルを作成する

	short.Save()

#%%
# executing
f = open('lists.txt', 'r', encoding="utf-8")
while 1:
	line = f.readline()
	if not line:
		break
	line = line.replace('\n','')
	if not line:
		continue
	print(line)
	mkshortcut(line)
# %%
