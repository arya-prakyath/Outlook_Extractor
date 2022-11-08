import win32com.client

print('Loading...\n')
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
folders = inbox.folders
all_folders = []

for folder in folders:
    all_folders.append(folder)
    try:
        sub_folders = folder.folders
        for sub_folder in sub_folders:
            all_folders.append(f'\t--> {sub_folder}')

            try:
                sub_sub_folders = sub_folder.folders
                for sub_sub_folder in sub_sub_folders:
                    all_folders.append(f'\t\t--> {sub_sub_folder}')
            except Exception:
                pass

    except Exception:
        pass

print('---Folders---')
for folder_name in all_folders:
    print(folder_name)

input("\nPress 'Enter' to exit")
