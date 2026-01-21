import sys
import os
import json
from datetime import datetime, timedelta

def get_emails_mac(days=1):
    """
    Extracts emails from Outlook for Mac (Inbox and Sent Items) using AppleScript.
    """
    import subprocess
    
    # AppleScript to fetch messages from Inbox and Sent Items
    script = f'''
    set emailData to {{}}
    set startDate to (current date) - ({days} * days)
    
    tell application "Microsoft Outlook"
        -- Folders to scan
        set targetFolders to {{inbox, sent items folder}}
        
        repeat with mailFolder in targetFolders
            set folderName to name of mailFolder
            set allMessages to (every message of mailFolder whose time received is greater than startDate)
            
            repeat with msg in allMessages
                set end of emailData to {{folder:folderName, subject:subject of msg, sender:name of sender of msg, received:time received of msg as string, content:plain text content of msg}}
            end repeat
        end repeat
    end tell
    return emailData
    '''
    try:
        process = subprocess.Popen(['osascript', '-e', script], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        out, err = process.communicate()
        if err:
            return {"error": err}
        return out
    except Exception as e:
        return {"error": str(e)}

def get_emails_windows(days=1):
    """
    Extracts emails from Outlook for Windows (Inbox and Sent Items) using pywin32.
    """
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
                # ComObject can sometimes error on ReceivedTime if item is still being synced
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
        print("Using Mac Outlook automation...")
        # Note: In a real environment, parsing osascript output needs more care
        emails = get_emails_mac(days)
    elif sys.platform == "win32":
        print("Using Windows Outlook automation...")
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
