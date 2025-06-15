import ctypes
from ctypes import wintypes
import win32gui
import win32con
import speech_recognition as sr
import threading

user32 = ctypes.windll.user32
gdi32 = ctypes.windll.gdi32


def get_workerw():
    progman = win32gui.FindWindow("Progman", None)
    result = ctypes.c_ulong()
    user32.SendMessageTimeoutW(progman, 0x052C, 0, 0, 0, 1000, ctypes.byref(result))

    def enum_windows_proc(hwnd, lParam):
        p = win32gui.FindWindowEx(hwnd, 0, "SHELLDLL_DefView", None)
        if p != 0:
            lParam[0] = win32gui.FindWindowEx(0, hwnd, "WorkerW", None)
            return False
        return True

    workerw = [0]
    win32gui.EnumWindows(enum_windows_proc, workerw)
    return workerw[0]

# 壁紙UIにテキスト描画関数
user32.FillRect.argtypes = [wintypes.HDC, ctypes.POINTER(wintypes.RECT), wintypes.HBRUSH]
user32.FillRect.restype = wintypes.INT

def draw_text_on_workerw(hwnd, text):
    hdc = user32.GetDC(hwnd)
    rect = wintypes.RECT()
    user32.GetClientRect(hwnd, ctypes.byref(rect))

    brush = gdi32.CreateSolidBrush(0x002233)  # ダークブルー
    user32.FillRect(hdc, ctypes.byref(rect), brush)  

    gdi32.SetTextColor(hdc, 0xFFFFFF)  # 白色
    gdi32.SetBkMode(hdc, win32con.TRANSPARENT)
    gdi32.TextOutW(hdc, 50, 50, text, len(text))

    user32.ReleaseDC(hwnd, hdc)

# 音声認識→壁紙UI更新の処理
def recognize_and_update_ui(hwnd):
    r = sr.Recognizer()
    mic = sr.Microphone()

    with mic as source:
        print("音声入力開始：話してください")
        r.adjust_for_ambient_noise(source)
        audio = r.listen(source)

    try:
        command = r.recognize_google(audio, language='ja-JP')
        print(f"認識結果: {command}")
        draw_text_on_workerw(hwnd, f"認識: {command}")

        # ここにコマンド解析＋アプリ起動などを追加可能

    except sr.UnknownValueError:
        print("音声が認識できませんでした。")
        draw_text_on_workerw(hwnd, "音声が認識できませんでした。")
    except sr.RequestError as e:
        print(f"音声認識サービスに接続できません: {e}")
        draw_text_on_workerw(hwnd, "認識サービスに接続できません。")

def main():
    hwnd = get_workerw()
    if not hwnd:
        print("WorkerW ウィンドウが見つかりませんでした。")
        return

    # 最初にUIをクリア
    draw_text_on_workerw(hwnd, "準備完了。音声を入力してください。")

    # 音声認識を別スレッドで実行（UIフリーズ防止）
    threading.Thread(target=recognize_and_update_ui, args=(hwnd,), daemon=True).start()

    # メインループで継続的にUI描画を保つ
    import time
    while True:
        # ここで必要に応じてUIの更新を行う
        time.sleep(1)

if __name__ == "__main__":
    main()

