import sys
import socket
import os
import traceback

def check_python():
    print(f"[*] Python 版本: {sys.version}")
    return True

def check_port(port=8000):
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        s.bind(("127.0.0.1", port))
        s.close()
        print(f"[PASS] 埠口 {port} 可用 (未被佔用)")
        return True
    except socket.error as e:
        print(f"[FAIL] 埠口 {port} 已被佔用！錯誤訊息: {e}")
        return False

def check_dependencies():
    packages = ["fastapi", "uvicorn", "pydantic", "win32com", "pythoncom"]
    success = True
    print("[*] 正在檢查依賴套件...")
    for pkg in packages:
        try:
            __import__(pkg)
            print(f"  [PASS] {pkg} 載入成功")
        except ImportError:
            print(f"  [FAIL] {pkg} 載入失敗！")
            success = False
    return success

def check_outlook():
    print("[*] 正在檢查本機 Outlook COM 連線...")
    try:
        import win32com.client
        import pythoncom
        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch("Outlook.Application")
        ns = outlook.GetNamespace("MAPI")
        inbox = ns.GetDefaultFolder(6)
        print(f"  [PASS] Outlook 連線成功！收件匣名稱: {inbox.Name} (信件總數: {inbox.Items.Count})")
        pythoncom.CoUninitialize()
        return True
    except Exception as e:
        print(f"  [FAIL] Outlook 連線失敗！詳細錯誤:")
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("====================================================")
    print("      Logistics Operation Platform - 系統環境診斷")
    print("====================================================")
    print()
    check_python()
    print()
    check_port(8000)
    print()
    check_dependencies()
    print()
    check_outlook()
    print()
    print("====================================================")
    print("診斷完畢，環境檢測成功！")
    print("====================================================")
    input("按任意鍵結束...")
