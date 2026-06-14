# Redline with Word

A macOS Finder Quick Action that compares two Microsoft Word documents and creates a redline file next to the revised document.

The first selected document is treated as the original/before version. The second selected document is treated as the revised/after version. The output is named:

```text
<revised-file-name>.redline.docx
```

The workflow ignores formatting-only changes and focuses on substantive text changes.

## Requirements

- macOS
- Microsoft Word for Mac
- Terminal access for installation

## Install

Download or clone this folder, then run:

```bash
chmod +x install.sh
./install.sh
```

This installs the Quick Action here:

```text
~/Library/Services/Redline with Word.workflow
```

## First Run Permissions

The first time you run the Quick Action, macOS may ask for permission for Automator to control Microsoft Word. Click **Allow**.

If your files are in Google Drive, iCloud Drive, Dropbox, or a shared/network folder, you may also need to give Automator Full Disk Access:

```text
System Settings > Privacy & Security > Full Disk Access > Automator
```

## How To Use

1. In Finder, select exactly two Word documents.
2. Select the original/before version first.
3. Select the revised/after version second.
4. Right-click the selected files.
5. Choose **Quick Actions > Redline with Word**.
6. Word will come to the front while the comparison runs.
7. A `.redline.docx` file will appear next to the revised document.

## Sharing With Someone Else

To share this with another Mac user, send them this folder/repo with these files:

- `applescript-v2`
- `build_workflow.py`
- `clean_redline.py`
- `install.sh`
- `README.md`

They should run:

```bash
chmod +x install.sh
./install.sh
```

You can also zip and share the installed workflow bundle directly:

```text
~/Library/Services/Redline with Word.workflow
```

If sharing the workflow bundle directly, the recipient should place it in their own:

```text
~/Library/Services/
```

Sharing the full folder plus `install.sh` is usually more reliable because it rebuilds the workflow cleanly on the recipient's machine.

## Notes

- Word comes to the front while the workflow runs. This is intentional because background Word automation can stall.
- The selected file order matters.
- Existing tracked changes in the input documents are accepted in temporary working copies before comparison. The original files on disk are not modified.
- The workflow saves the comparison result through a temporary Desktop file first, then moves it next to the revised document. This avoids some Word sandbox/write-access problems with cloud folders.
- After Word creates the comparison, the workflow removes obvious no-op redlines such as whitespace-only changes and identical delete/reinsert pairs caused by DOCX structure differences.
