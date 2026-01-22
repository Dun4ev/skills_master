import sys
import os
import json
from datetime import datetime, timedelta

def get_emails_mac(days=1):
    """
    Extracts emails from Outlook for Mac using AppleScript. 
    Writes results directly to a file to handle large content and special characters.
    """
    import subprocess
    import tempfile
    import os
    
    output_temp = tempfile.NamedTemporaryFile(suffix='.txt', delete=False).name
    
    # In Outlook Mac, message 1 is typically the NEWEST
    script = f'''
    set output_path to "{output_temp}"
    set fp to open for access POSIX file output_path with write permission
    set eof of fp to 0
    
    tell application "Microsoft Outlook"
        try
            set accList to every exchange account
            repeat with acc in accList
                set allFs to mail folders of acc
                repeat with f in allFs
                    set fn to (name of f) as string
                    if (fn contains "Входящ") or (fn contains "Отправл") or (fn contains "Банат") or (fn contains "Inbox") or (fn contains "Sent") then
                        try
                            set msgCount to count messages of f
                            set take to 100 -- Large scan since we'll filter by date
                            if msgCount < take then set take to msgCount
                            
                            -- Loop from 1 to take (newest to oldest)
                            repeat with i from 1 to take
                                try
                                    set msg to message i of f
                                    
                                    set subj to ""
                                    try
                                        set subj to (subject of msg) as string
                                    end try
                                    
                                    set sndr to ""
                                    try
                                        set sndr to (name of sender of msg) as string
                                    on error
                                        try
                                            set sndr to (address of sender of msg) as string
                                        end try
                                    end try
                                    
                                    set recv to ""
                                    try
                                        set recv to (time received of msg) as string
                                    end try
                                    
                                    set cont to ""
                                    try
                                        set cont to (plain text content of msg) as string
                                        if length of cont > 500 then set cont to (text 1 thru 500 of cont)
                                    end try
                                    
                                    set lineBody to fn & "|SEP|" & subj & "|SEP|" & sndr & "|SEP|" & recv & "|SEP|" & cont & "|MSG_END|" & return
                                    write lineBody to fp as «class utf8»
                                on error e
                                end try
                            end repeat
                        on error
                        end try
                    end if
                end repeat
            end repeat
        on error e
            write "ERROR: " & e to fp as «class utf8»
        end try
    end tell
    close access fp
    '''
    
    with tempfile.NamedTemporaryFile(mode='w', suffix='.applescript', delete=False) as tf:
        tf.write(script)
        temp_name = tf.name

    try:
        process = subprocess.Popen(['osascript', temp_name], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        process.communicate()
        os.unlink(temp_name)
        
        if not os.path.exists(output_temp):
            return {"error": "Output file not found"}
            
        with open(output_temp, 'r', encoding='utf-8', errors='ignore') as f:
            data = f.read()
        os.unlink(output_temp)
        
        results = []
        messages = data.strip().split("|MSG_END|")
        for m in messages:
            m = m.strip()
            if not m: continue
            
            parts = m.split("|SEP|")
            if len(parts) >= 5:
                results.append({
                    "folder": parts[0].strip(),
                    "subject": parts[1].strip(),
                    "sender": parts[2].strip(),
                    "received": parts[3].strip(),
                    "content": parts[4].strip()
                })
        return results
    except Exception as e:
        if os.path.exists(output_temp): os.unlink(output_temp)
        return {"error": str(e)}

def get_emails_windows(days=1):
    # (Remains same)
    try:
        import win32com.client
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        
        # 6 = olFolderInbox, 5 = olFolderSentMail
        folders = [6, 5]
        folder_names = {6: "Inbox", 5: "Sent Items"}
        results = []
        target_date = datetime.now() - timedelta(days=days)
        
        for folder_id in folders:
            folder = outlook.GetDefaultFolder(folder_id)
            messages = folder.Items
            messages.Sort("[ReceivedTime]", True)
            
            for msg in messages:
                try:
                    if msg.ReceivedTime.timestamp() < target_date.timestamp():
                        break
                    results.append({
                        "folder": folder_names[folder_id],
                        "subject": msg.Subject,
                        "sender": msg.SenderName if folder_id == 6 else "Me",
                        "received": str(msg.ReceivedTime),
                        "content": msg.Body[:1000]
                    })
                except:
                    continue
        return results
    except ImportError:
        return {"error": "pywin32 not installed. Run: pip install pywin32"}
    except Exception as e:
        return {"error": str(e)}

def main():
    print(f"Detecting OS: {sys.platform}")
    days = 1
    
    if sys.platform == "darwin":
        print(f"Using Mac Outlook automation (scanning latest messages)...")
        emails = get_emails_mac(days)
    elif sys.platform == "win32":
        print(f"Using Windows Outlook automation (scanning last {days} days)...")
        emails = get_emails_windows(days)
    else:
        print(f"Platform {sys.platform} not supported for direct Outlook access.")
        return

    output_file = os.path.join(os.path.dirname(__file__), "daily_emails.json")
    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(emails, f, ensure_ascii=False, indent=2)
    
    print(f"Done! Results saved to {output_file}")

if __name__ == "__main__":
    main()
