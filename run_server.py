"""
Runner script for Exam Student Matcher & Discrepancy Finder.
Starts FastAPI backend server on 0.0.0.0:8000 (accessible on LAN)
and opens the default web browser.
"""
import os
import sys
import time
import socket
import webbrowser
import threading
import uvicorn

def get_lan_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"

def open_browser_delayed(url):
    time.sleep(1.2)
    webbrowser.open(url)

if __name__ == "__main__":
    lan_ip = get_lan_ip()
    port = 8000
    
    print("=" * 65)
    print("  EXAM STUDENT MATCHER & DISCREPANCY FINDER")
    print("=" * 65)
    print(f"[*] Local access:       http://localhost:{port}")
    print(f"[*] Team LAN access:    http://{lan_ip}:{port}")
    print("=" * 65)
    print("[*] Share the Team LAN link with your colleagues on the same WiFi/Network.")
    print("[*] Press Ctrl+C in this window to stop the server.")
    print("=" * 65)
    
    # Launch browser in background thread
    threading.Thread(target=open_browser_delayed, args=(f"http://localhost:{port}",), daemon=True).start()
    
    # Run uvicorn server on all interfaces
    uvicorn.run("main:app", host="0.0.0.0", port=port, reload=False)
