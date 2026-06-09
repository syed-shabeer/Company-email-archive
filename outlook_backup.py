import os
import json
import sys
import subprocess
from datetime import datetime
from pathlib import Path
from tkinter import Tk, Toplevel, Label, Button, StringVar, OptionMenu, Frame, messagebox

# ── Dependency check with clear message ────────────────────────────────────
try:
    import win32com.client
except ImportError:
    print("[!] ERROR: pywin32 is not installed.")
    print("    This script must be run via the packaged .exe or run_backup.bat.")
    input("\nPress Enter to exit...")
    sys.exit(1)


class OutlookBackup:
    def __init__(self):
        self.root = Path(r"C:\Users\T0415IL\OneDrive - Stellantis\EmailBackup")
        self.emails_dir = self.root / "Emails"
        self.index_path = self.root / "email_index.json"
        self.error_log_path = self.root / "backup_errors.log"

        # Ensure folders exist
        self.emails_dir.mkdir(parents=True, exist_ok=True)

        self.index_data = self._load_index()
        self._error_count = 0

    # ── Index helpers ───────────────────────────────────────────────────────

    def _load_index(self):
        if self.index_path.exists():
            try:
                with open(self.index_path, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except (json.JSONDecodeError, OSError) as e:
                self._log_error(f"Could not load index file: {e}. Starting fresh.")
        return {"total_emails": 0, "emails": []}

    def save_index(self):
        try:
            with open(self.index_path, 'w', encoding='utf-8') as f:
                json.dump(self.index_data, f, indent=2, ensure_ascii=False)
        except OSError as e:
            self._log_error(f"Failed to save index: {e}")

    # ── Error logging ───────────────────────────────────────────────────────

    def _log_error(self, message: str):
        """Append an error to the log file instead of silently swallowing it."""
        self._error_count += 1
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        entry = f"[{timestamp}] {message}\n"
        try:
            with open(self.error_log_path, 'a', encoding='utf-8') as f:
                f.write(entry)
        except OSError:
            pass  # If we can't log, at least don't crash
        print(f"  [!] {message}")

    # ── Subject blocklist ───────────────────────────────────────────────────

    SKIP_SUBJECTS = ["The Hub", "Digest"]

    # ── Sender blocklist ────────────────────────────────────────────────────
    # Case-insensitive substrings (checked on lowercased sender)
    SKIP_SENDER_SUBSTRINGS_CI = ["noreply", "donotreply", "hrnotifications"]
    # Case-sensitive substrings (checked on original sender string)
    SKIP_SENDER_SUBSTRINGS_CS = ["ICT"]

    def _should_skip(self, subject: str, sender: str = "") -> bool:
        if any(phrase in subject for phrase in self.SKIP_SUBJECTS):
            return True
        sender_lower = sender.lower()
        if any(s in sender_lower for s in self.SKIP_SENDER_SUBSTRINGS_CI):
            return True
        if any(s in sender for s in self.SKIP_SENDER_SUBSTRINGS_CS):
            return True
        return False

    # ── Core backup ─────────────────────────────────────────────────────────

    def _ask_schedule_time(self):
        """Show a popup to ask the user what time daily backup should run.
        Only runs once — skipped if scheduler_done.txt already exists."""
        flag = self.root / "scheduler_done.txt"
        if flag.exists():
            return  # Already scheduled — don't ask again

        root = Tk()
        root.withdraw()  # Hide root window

        popup = Toplevel(root)
        popup.title("Backup Complete — Set Daily Schedule")
        popup.resizable(False, False)

        # Center on screen
        popup.update_idletasks()
        w, h = 400, 240
        x = (popup.winfo_screenwidth() - w) // 2
        y = (popup.winfo_screenheight() - h) // 2
        popup.geometry(f"{w}x{h}+{x}+{y}")
        popup.configure(bg="#f0f4fa")

        # Force popup to front
        popup.lift()
        popup.attributes("-topmost", True)
        popup.focus_force()
        popup.grab_set()

        Label(popup, text="✅  Backup Complete!",
              font=("Segoe UI", 14, "bold"),
              bg="#f0f4fa", fg="#003A8C").pack(pady=(20, 2))
        Label(popup, text="Your emails have been backed up successfully.",
              font=("Segoe UI", 9), bg="#f0f4fa", fg="#64748b").pack(pady=(0, 4))
        Label(popup, text="What time should the daily backup run?",
              font=("Segoe UI", 10, "bold"), bg="#f0f4fa", fg="#334155").pack(pady=(0, 12))

        row = Frame(popup, bg="#f0f4fa")
        row.pack()

        # Hour
        hour_var = StringVar(value="08")
        hours = [f"{h:02d}" for h in range(1, 13)]
        OptionMenu(row, hour_var, *hours).pack(side="left", padx=4)

        Label(row, text=":", font=("Segoe UI", 13, "bold"), bg="#f0f4fa").pack(side="left")

        # Minute
        min_var = StringVar(value="00")
        minutes = [f"{m:02d}" for m in range(0, 60, 5)]
        OptionMenu(row, min_var, *minutes).pack(side="left", padx=4)

        # AM/PM
        ampm_var = StringVar(value="AM")
        OptionMenu(row, ampm_var, "AM", "PM").pack(side="left", padx=4)

        scheduled = [False]

        def confirm():
            time_str = f"{hour_var.get()}:{min_var.get()} {ampm_var.get()}"
            scheduled[0] = True
            popup.destroy()
            root.destroy()
            self._create_scheduled_task(time_str)
            self._show_final_instructions()

        def skip():
            popup.destroy()
            root.destroy()
            self._show_final_instructions()

        btn_frame = Frame(popup, bg="#f0f4fa")
        btn_frame.pack(pady=16)
        Button(btn_frame, text="Set Schedule", command=confirm,
               bg="#003A8C", fg="white", font=("Segoe UI", 10, "bold"),
               relief="flat", padx=14, pady=6, cursor="hand2").pack(side="left", padx=8)
        Button(btn_frame, text="Skip", command=skip,
               bg="#e2e8f0", fg="#334155", font=("Segoe UI", 10),
               relief="flat", padx=14, pady=6, cursor="hand2").pack(side="left", padx=8)

        root.mainloop()

    def _show_final_instructions(self):
        """Open a new CMD window with red-text instructions to launch the archive."""
        bat_path = self.root / "Stellantis_Email_Archive.bat"
        script = (
            "color 04 & "  # red text on black background
            "cls & "
            "echo. & "
            "echo ======================================================= & "
            "echo   SETUP COMPLETE - Ready to Use! & "
            "echo ======================================================= & "
            "echo. & "
            "echo  Your emails are backed up and ready to view. & "
            "echo. & "
            "echo  TO VIEW YOUR EMAIL ARCHIVE: & "
            "echo  ----------------------------------------- & "
            "echo  Double-click: Stellantis_Email_Archive.bat & "
            f"echo  Location: {self.root} & "
            "echo. & "
            "echo  TIP: Right-click Stellantis_Email_Archive.bat & "
            "echo       and choose 'Send to Desktop (shortcut)' & "
            "echo       for quick access every day. & "
            "echo. & "
            "echo  Once the website opens, click the '? Help' & "
            "echo  button in the top right to learn how to use it. & "
            "echo. & "
            "echo ======================================================= & "
            "echo. & "
            "pause"
        )
        subprocess.Popen(f'start "Setup Complete" cmd /c "{script}"', shell=True)

    def _create_scheduled_task(self, time_str):
        """Call SetBackupSchedule.bat with the chosen time as an argument."""
        bat_path = self.root / "SetBackupSchedule.bat"
        if not bat_path.exists():
            print(f"[!] SetBackupSchedule.bat not found at {bat_path}")
            return
        try:
            subprocess.Popen(
                f'cmd /c ""{bat_path}" "{time_str}""',
                shell=True,
                cwd=str(self.root)
            )
            # Write flag so this popup never appears again
            flag = self.root / "scheduler_done.txt"
            flag.write_text(f"Scheduled at {time_str}\n", encoding="utf-8")
        except Exception as e:
            print(f"[!] Could not launch scheduler: {e}")

    # All Outlook default folder IDs — used to get their EntryIDs at runtime
    # 3=Deleted Items, 4=Outbox, 5=Sent Items, 6=Inbox, 9=Calendar, 10=Contacts,
    # 11=Journal, 12=Notes, 13=Tasks, 16=Drafts, 17=Archive, 18=Junk Email, 23=Junk Email (alt)
    SYSTEM_FOLDER_IDS = {3, 4, 5, 6, 9, 10, 11, 12, 13, 16, 17, 18, 23}

    # Name-based blocklist for system/cache folders that have no fixed default ID
    SYSTEM_NAME_BLOCKLIST = {
        'search folders', 'rss feeds', 'rss subscriptions',
        'sync issues', 'conflicts', 'local failures', 'server failures',
        'conversation history', 'conversation action settings',
        'social activity notifications', 'quick step settings',
        'externalcontacts', 'peoplecentricconversation buddies',
        'organizational contacts', 'gal contacts', 'companies',
        'recipient cache', 'yammer root', 'team chat',
        'outbound', 'inbound', 'feeds', 'files',
        'archive',  # built-in archive, not user-created
        'stellantis records',  # always excluded
    }

    def _get_system_entry_ids(self, outlook):
        """Get EntryIDs of all true system default folders so we can exclude them by ID."""
        ids = set()
        for fid in self.SYSTEM_FOLDER_IDS:
            try:
                ids.add(outlook.GetDefaultFolder(fid).EntryID)
            except Exception:
                pass
        # Also add EntryIDs of known hidden folders by iterating root
        try:
            root = outlook.GetDefaultFolder(6).Parent
            for folder in root.Folders:
                name = folder.Name.strip().lower()
                if name in self.SYSTEM_NAME_BLOCKLIST or (name.startswith('{') and name.endswith('}')):
                    try:
                        ids.add(folder.EntryID)
                    except Exception:
                        pass
        except Exception:
            pass
        return ids

    def _is_real_mail_folder(self, folder, system_entry_ids):
        """Return True only if this is a real user-created mail folder."""
        name = folder.Name.strip()
        # Block GUIDs like {A9E2BC46-...}
        if name.startswith('{') and name.endswith('}'):
            return False
        # Block by EntryID — catches system folders regardless of name
        try:
            if folder.EntryID in system_entry_ids:
                return False
        except Exception:
            pass
        # Block known junk names — only as fallback
        if name.lower() in self.SYSTEM_NAME_BLOCKLIST:
            return False
        return True

    def _get_all_outlook_folders(self, outlook):
        """Return structured list of selectable custom folders (outside Inbox/Sent).
        Inbox/Sent subfolders are auto-scanned and NOT shown here.
        Each entry: (display_label, folder_key, folder_object, indent_level)
        folder_key uses Parent/Child format for website filter display."""
        system_entry_ids = self._get_system_entry_ids(outlook)
        result = []

        # Only show top-level custom folders and their subfolders (outside Inbox/Sent)
        try:
            root_folder = outlook.GetDefaultFolder(6).Parent
            for folder in root_folder.Folders:
                if not self._is_real_mail_folder(folder, system_entry_ids):
                    continue
                parent_name = folder.Name
                result.append((parent_name, parent_name, folder, 0))
                try:
                    for sub in folder.Folders:
                        if self._is_real_mail_folder(sub, system_entry_ids):
                            # Key is Parent/Sub so website shows it correctly
                            sub_key = f"{parent_name}/{sub.Name}"
                            result.append((f"  └─ {sub.Name}", sub_key, sub, 1))
                            try:
                                for subsub in sub.Folders:
                                    if self._is_real_mail_folder(subsub, system_entry_ids):
                                        subsub_key = f"{parent_name}/{sub.Name}/{subsub.Name}"
                                        result.append((f"    └─ {subsub.Name}", subsub_key, subsub, 2))
                            except Exception:
                                pass
                except Exception:
                    pass
        except Exception as e:
            self._log_error(f"Error listing Outlook folders: {e}")

        return result

    def _select_folders_from_list(self, all_folders, prompt_title="Select folders to include"):
        """Show numbered list of folders and let user pick by number(s).
        Selecting a parent folder auto-includes all its subfolders."""
        print()
        print("-" * 55)
        print(f"  {prompt_title}")
        print("-" * 55)
        # all_folders: (display_label, flat_name, folder_object, indent_level)
        num = 1
        index_map = {}  # displayed number -> position in all_folders
        for pos, (label, name, obj, level) in enumerate(all_folders):
            if obj is None:  # header row, not selectable
                print(f"\n  {label}")
            else:
                print(f"  [{num:2}] {label}")
                index_map[num] = pos
                num += 1
        print()
        print("  Enter number(s) separated by commas to select.")
        print("  Example: 1,3  or just  2  for a single folder.")
        print("  Selecting a parent folder includes all its subfolders.")
        print("  Type A to include ALL folders — Press Enter to skip all.")
        print("-" * 55)

        max_num = num - 1
        while True:
            raw = input("  Your selection: ").strip()
            if not raw:
                return []

            # A = select everything
            if raw.upper() == 'A':
                selected = []
                selected_names = set()
                for _, name, obj, _ in all_folders:
                    if obj is not None and name not in selected_names:
                        selected.append((name, obj))
                        selected_names.add(name)
                print(f"  ✓ All {len(selected)} folders selected.")
                return selected

            try:
                indices = [int(x.strip()) for x in raw.split(",")]
                invalid = [i for i in indices if i < 1 or i > max_num]
                if invalid:
                    print(f"  [!] Invalid numbers: {invalid}. Try again.")
                    continue

                selected = []
                selected_names = set()
                for i in indices:
                    pos = index_map[i]
                    label, name, obj, level = all_folders[pos]
                    if name not in selected_names:
                        selected.append((name, obj))
                        selected_names.add(name)
                    # If parent (level 0), auto-add ALL its subfolders
                    if level == 0:
                        j = pos + 1
                        while j < len(all_folders) and all_folders[j][3] in (1, 2):
                            _, skey, sobj, slevel = all_folders[j]
                            if sobj is not None and skey not in selected_names:
                                selected.append((skey, sobj))
                                selected_names.add(skey)
                            j += 1
                    # If level 1, also auto-add its children (level 2)
                    elif level == 1:
                        j = pos + 1
                        while j < len(all_folders) and all_folders[j][3] == 2:
                            _, skey, sobj, _ = all_folders[j]
                            if sobj is not None and skey not in selected_names:
                                selected.append((skey, sobj))
                                selected_names.add(skey)
                            j += 1
                return selected
            except ValueError:
                print("  [!] Please enter numbers only, separated by commas.")

    def _ask_custom_folders(self, outlook):
        """Show all Outlook folders and let user pick.
        Saves choices to custom_folders.json and reuses on future runs.
        Auto-detects newly added folders on each run and asks about them."""
        config_path = self.root / "custom_folders.json"
        all_folders = self._get_all_outlook_folders(outlook)
        all_folder_names = [name for _, name, _, _ in all_folders]

        # ── Load saved config if exists ───────────────────────────────────
        if config_path.exists():
            try:
                with open(config_path, 'r', encoding='utf-8') as f:
                    saved = json.load(f)
                saved_names    = saved.get("folders", [])
                seen_names     = saved.get("seen_folders", [])

                # Detect folders that are new since last run (4-tuple)
                new_folders = [(label, name, obj, level) for label, name, obj, level in all_folders
                               if name not in seen_names]

                if new_folders:
                    print()
                    print("-" * 55)
                    print("  New Outlook folders detected since last backup:")
                    for label, name, _, _ in new_folders:
                        print(f"    • {name}")
                    print("-" * 55)
                    selected_new = self._select_folders_from_list(
                        new_folders,
                        prompt_title="Select any new folders to include in backup"
                    )
                    for name, _ in selected_new:
                        if name not in saved_names:
                            saved_names.append(name)

                    # Update seen list and save
                    saved["folders"]      = saved_names
                    saved["seen_folders"] = all_folder_names
                    with open(config_path, 'w', encoding='utf-8') as f:
                        json.dump(saved, f, indent=2)

                if saved_names:
                    print()
                    print("-" * 55)
                    print(f"  Backing up custom folders: {', '.join(saved_names)}")
                    print("-" * 55)
                else:
                    print()
                    print("-" * 55)
                    print("  No custom folders configured.")
                    print("-" * 55)

                # Resolve saved names to folder objects (4-tuple now)
                custom_folders = []
                for name in saved_names:
                    match = next((obj for _, n, obj, _ in all_folders if n == name), None)
                    if match:
                        custom_folders.append((name, match))
                    else:
                        print(f"  [!] Saved folder '{name}' not found in Outlook — skipping.")
                return custom_folders

            except Exception as e:
                self._log_error(f"Could not read custom_folders.json: {e}")

        # ── First time: ask Y/N before showing folder list ──────────────
        print()
        print("-" * 55)
        print("  This backup will scan your Inbox and Sent Items.")
        print("  Do you want to include any other folders too?")
        print("-" * 55)

        while True:
            ans = input("  Include custom folders? (Y/N): ").strip().upper()
            if ans == 'N':
                selected = []
                break
            elif ans == 'Y':
                if not all_folders:
                    print("  No additional folders found in Outlook.")
                    selected = []
                else:
                    selected = self._select_folders_from_list(
                        all_folders,
                        prompt_title="Select folders to include (or press Enter to skip)"
                    )
                break
            else:
                print("  [!] Please enter Y or N.")

        folder_names = [name for name, _ in selected]

        if folder_names:
            print(f"  ✓ Selected: {', '.join(folder_names)}")
        else:
            print("  ✓ No custom folders selected.")

        # Save choices + full seen list for future new-folder detection
        try:
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump({
                    "folders": folder_names,
                    "seen_folders": all_folder_names
                }, f, indent=2)
            print("  ✓ Preferences saved.")
        except Exception as e:
            self._log_error(f"Could not save custom_folders.json: {e}")

        print("-" * 55)
        return selected

    def _find_outlook_folder(self, outlook, folder_name):
        """Search for a folder by name under the default store root."""
        try:
            root_folder = outlook.GetDefaultFolder(6).Parent
            for folder in root_folder.Folders:
                if folder.Name.lower() == folder_name.lower():
                    return folder
                try:
                    for subfolder in folder.Folders:
                        if subfolder.Name.lower() == folder_name.lower():
                            return subfolder
                except Exception:
                    continue
        except Exception as e:
            self._log_error(f"Error searching for folder '{folder_name}': {e}")
        return None

    def backup_emails(self):
        print("=" * 55)
        print("   STELLANTIS EMAIL BACKUP")
        print("=" * 55)

        # Connect to Outlook
        try:
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        except Exception as e:
            print(f"\n[!] Could not connect to Outlook: {e}")
            print("    Make sure Outlook is installed and your profile is set up.")
            input("\nPress Enter to exit...")
            sys.exit(1)

        # Ask user about custom folders before starting
        custom_folders = self._ask_custom_folders(outlook)

        # Build scan list: (folder_name, folder_object_or_id)
        # Default folders use IDs; custom folders use objects directly
        default_folders = [(6, "Inbox"), (5, "Sent Items")]

        indexed_ids = {e['email_id'] for e in self.index_data['emails']}
        new_count = 0
        skip_count = 0

        SAVE_EVERY = 50  # Save progress to disk every N emails

        try:
            # ── Scan default folders ──────────────────────────────────────
            for f_id, f_name in default_folders:
                print(f"\n[*] Scanning {f_name}...")

                try:
                    outlook_folder = outlook.GetDefaultFolder(f_id)
                    items = outlook_folder.Items
                    items.Sort("[ReceivedTime]", True)  # Newest first
                except Exception as e:
                    self._log_error(f"Could not access folder '{f_name}': {e}")
                    continue

                for item in items:
                    try:
                        # Only process standard mail items (Class 43)
                        if item.Class != 43:
                            continue

                        eid = item.EntryID
                        if eid in indexed_ids:
                            continue

                        subject = item.Subject or "[No Subject]"
                        sender = item.SenderName or ""

                        if self._should_skip(subject, sender):
                            skip_count += 1
                            continue

                        # ── Save JSON (for search.html viewer) ──────────────
                        email_json = {
                            "subject": subject,
                            "from": item.SenderName,
                            "received_date": item.ReceivedTime.isoformat(),
                            "body": item.Body,
                            "folder": f_name,
                        }
                        json_path = self.emails_dir / f"{eid}.json"
                        with open(json_path, 'w', encoding='utf-8') as f:
                            json.dump(email_json, f, indent=2, ensure_ascii=False)

                        # ── Save MSG (for opening in Outlook) ────────────────
                        msg_path = self.emails_dir / f"{eid}.msg"
                        item.SaveAs(str(msg_path))

                        # ── Add to index ──────────────────────────────────────
                        self.index_data['emails'].append({
                            "email_id": eid,
                            "subject": subject,
                            "from": item.SenderName,
                            "received_date": item.ReceivedTime.isoformat(),
                            "folder": f_name,
                        })
                        indexed_ids.add(eid)
                        new_count += 1

                        display_subject = subject[:50] + "..." if len(subject) > 50 else subject
                        print(f"  [+] ({new_count}) {display_subject}")

                        # ── Save progress every SAVE_EVERY emails ─────────────
                        if new_count % SAVE_EVERY == 0:
                            self.index_data['total_emails'] = len(self.index_data['emails'])
                            self.save_index()
                            print(f"\n  [✓] Progress saved — {new_count} emails backed up so far\n")

                    except Exception as e:
                        subject_hint = getattr(item, 'Subject', 'unknown')
                        self._log_error(f"Failed to save email '{subject_hint}': {e}")
                        continue

            # ── Scan subfolders of Inbox and Sent Items automatically ────
            for f_id, f_name in default_folders:
                try:
                    default_folder = outlook.GetDefaultFolder(f_id)
                    system_entry_ids = self._get_system_entry_ids(outlook)
                    for sub in default_folder.Folders:
                        if not self._is_real_mail_folder(sub, system_entry_ids):
                            continue
                        if sub.Name.lower() == 'stellantis records':
                            continue
                        sub_full_name = f"{f_name}/{sub.Name}"
                        print(f"\n[*] Scanning {sub_full_name}...")
                        try:
                            items = sub.Items
                            items.Sort("[ReceivedTime]", True)
                        except Exception as e:
                            self._log_error(f"Could not access subfolder '{sub_full_name}': {e}")
                            continue
                        for item in items:
                            try:
                                if item.Class != 43:
                                    continue
                                eid = item.EntryID
                                subject = item.Subject or "[No Subject]"
                                sender = item.SenderName or ""
                                if self._should_skip(subject, sender):
                                    skip_count += 1
                                    continue
                                if eid in indexed_ids:
                                    for entry in self.index_data['emails']:
                                        if entry['email_id'] == eid and entry['folder'] != sub_full_name:
                                            entry['folder'] = sub_full_name
                                            json_path = self.emails_dir / f"{eid}.json"
                                            if json_path.exists():
                                                try:
                                                    with open(json_path, 'r', encoding='utf-8') as f:
                                                        ej = json.load(f)
                                                    ej['folder'] = sub_full_name
                                                    with open(json_path, 'w', encoding='utf-8') as f:
                                                        json.dump(ej, f, indent=2, ensure_ascii=False)
                                                except Exception:
                                                    pass
                                    continue
                                email_json = {
                                    "subject": subject, "from": item.SenderName,
                                    "received_date": item.ReceivedTime.isoformat(),
                                    "body": item.Body, "folder": sub_full_name,
                                }
                                json_path = self.emails_dir / f"{eid}.json"
                                with open(json_path, 'w', encoding='utf-8') as f:
                                    json.dump(email_json, f, indent=2, ensure_ascii=False)
                                item.SaveAs(str(self.emails_dir / f"{eid}.msg"))
                                self.index_data['emails'].append({
                                    "email_id": eid, "subject": subject,
                                    "from": item.SenderName,
                                    "received_date": item.ReceivedTime.isoformat(),
                                    "folder": sub_full_name,
                                })
                                indexed_ids.add(eid)
                                new_count += 1
                                display_subject = subject[:50] + "..." if len(subject) > 50 else subject
                                print(f"  [+] ({new_count}) {display_subject}")
                                if new_count % SAVE_EVERY == 0:
                                    self.index_data['total_emails'] = len(self.index_data['emails'])
                                    self.save_index()
                                    print(f"\n  [✓] Progress saved — {new_count} emails backed up so far\n")
                            except Exception as e:
                                self._log_error(f"Failed to save email: {e}")
                                continue
                except Exception as e:
                    self._log_error(f"Error scanning subfolders of {f_name}: {e}")

            # ── Scan custom folders ──────────────────────────────────────
            for f_name, outlook_folder in custom_folders:
                if f_name.lower() == 'stellantis records' or f_name.lower().endswith('/stellantis records'):
                    print(f"\n[*] Skipping excluded folder: {f_name}")
                    continue
                print(f"\n[*] Scanning custom folder: {f_name}...")

                try:
                    items = outlook_folder.Items
                    items.Sort("[ReceivedTime]", True)
                except Exception as e:
                    self._log_error(f"Could not access custom folder '{f_name}': {e}")
                    continue

                for item in items:
                    try:
                        if item.Class != 43:
                            continue

                        eid = item.EntryID
                        subject = item.Subject or "[No Subject]"
                        sender = item.SenderName or ""

                        if self._should_skip(subject, sender):
                            skip_count += 1
                            continue

                        # If already indexed under a different folder (e.g. Inbox),
                        # update the folder name to the correct custom folder
                        if eid in indexed_ids:
                            for entry in self.index_data['emails']:
                                if entry['email_id'] == eid and entry['folder'] != f_name:
                                    entry['folder'] = f_name
                                    # Also update the JSON file on disk
                                    json_path = self.emails_dir / f"{eid}.json"
                                    if json_path.exists():
                                        try:
                                            with open(json_path, 'r', encoding='utf-8') as f:
                                                email_json = json.load(f)
                                            email_json['folder'] = f_name
                                            with open(json_path, 'w', encoding='utf-8') as f:
                                                json.dump(email_json, f, indent=2, ensure_ascii=False)
                                            print(f"  [~] Updated folder for: {subject[:50]}")
                                        except Exception as e:
                                            self._log_error(f"Failed to update folder in JSON for '{subject}': {e}")
                            continue

                        email_json = {
                            "subject": subject,
                            "from": item.SenderName,
                            "received_date": item.ReceivedTime.isoformat(),
                            "body": item.Body,
                            "folder": f_name,
                        }
                        json_path = self.emails_dir / f"{eid}.json"
                        with open(json_path, 'w', encoding='utf-8') as f:
                            json.dump(email_json, f, indent=2, ensure_ascii=False)

                        msg_path = self.emails_dir / f"{eid}.msg"
                        item.SaveAs(str(msg_path))

                        self.index_data['emails'].append({
                            "email_id": eid,
                            "subject": subject,
                            "from": item.SenderName,
                            "received_date": item.ReceivedTime.isoformat(),
                            "folder": f_name,
                        })
                        indexed_ids.add(eid)
                        new_count += 1

                        display_subject = subject[:50] + "..." if len(subject) > 50 else subject
                        print(f"  [+] ({new_count}) {display_subject}")

                        if new_count % SAVE_EVERY == 0:
                            self.index_data['total_emails'] = len(self.index_data['emails'])
                            self.save_index()
                            print(f"\n  [✓] Progress saved — {new_count} emails backed up so far\n")

                    except Exception as e:
                        subject_hint = getattr(item, 'Subject', 'unknown')
                        self._log_error(f"Failed to save email '{subject_hint}': {e}")
                        continue

        except KeyboardInterrupt:
            # User closed the window or pressed Ctrl+C — save whatever we have
            print("\n\n  [!] Interrupted — saving progress...")

        # ── Final save (always runs) ──────────────────────────────────────────
        if new_count > 0:
            self.index_data['total_emails'] = len(self.index_data['emails'])
            self.save_index()

        print("\n" + "=" * 55)
        print(f"  Done!")
        print(f"  New emails saved  : {new_count}")
        print(f"  Skipped (digest)  : {skip_count}")
        print(f"  Total in archive  : {self.index_data['total_emails']}")
        if self._error_count:
            print(f"  Errors logged     : {self._error_count}  (see backup_errors.log)")
        print("=" * 55)
        print("\n  TIP: Run again anytime — it only picks up new emails.")

        self._ask_schedule_time()


if __name__ == "__main__":
    bot = OutlookBackup()
    bot.backup_emails()
