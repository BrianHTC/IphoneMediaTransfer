# iPhone-to-Windows Media Transfer Tool 

A PowerShell tool for transfering photos and videos from an iPhone through a USB cable. The tool supports date filtering, automatic date-based and file type sorting, configurable conflict handling methods, and result verification.

![image](https://github.com/BrianHTC/IphoneMediaTransfer/blob/main/Interface.png)

## Features

- Windows Forms graphical interface
- Copies media from an iPhone connected through USB, using MTP protocol. 
- Scans iPhone storage recursively
- Filters imports by date range
- Fast scanning of relevant `YYYYMM` source folders when a date filter is enabled
- Supports most file types. 
- Filters by common photo and video file types
- Supports custom file types
- Organizes output into:
  - `YYYY`
  - `YYYY\MM`
  - `YYYY\MM\DD`
- Handles duplicate filenames using:
  - Skip
  - Overwrite
  - Rename with `_1`, `_2`, and so on
- iPhone puts content that its system consider "identical" into a special folder named in YYYYMM_b format, and give them the same name, causing file conflicts when transfering medias the traditional method. This tool keeps the files in YYYYMM_b or so on folders with a suffix [number] so that all the files are preserved well. 
- Optional size and video length verification, to make sure the media were transfered correctly.
- Abort button that stops the copying process and cleans up temporary files generated.
- Displays destination disk-space information

## Source Folder Filename Rules

Some iPhone exports may contain multiple source folders for the same month, such as:

```text
202207
202207_a
202207_b
202207_c
202207_d
```

The script applies these filename rules before applying the selected conflict option:

```text
202207\IMG_5161.JPG   -> IMG_5161.JPG
202207_a\IMG_5161.JPG -> IMG_5161.JPG
202207_b\IMG_5161.JPG -> IMG_5161[1].JPG
202207_c\IMG_5161.JPG -> IMG_5161[2].JPG
202207_d\IMG_5161.JPG -> IMG_5161[3].JPG
```

The processing order is:

1. Determine the source-folder suffix.
2. Build the proposed destination filename.
3. Apply Skip, Overwrite, or Rename conflict handling.
4. Copy and verify the file.

## Requirements

- Windows 10 or Windows 11 (Compatibility with other versions unchecked)
- Windows PowerShell 5.1 or above (Compatibility with older versions unchecked)
### -iTunes service **MUST** be installed first. The iPhone must be visible under **This PC** folder in File Explorer
- An iPhone connected by a functioning USB cable that transfers data
- The iPhone must be unlocked after connecting to the PC and select "Trust this device"

## How to use

1. Download the PowerShell script to Desktop:

   ```text
   iPhoneMediaTransfer.ps1
   ```

2. Connect and unlock the iPhone, select trust and enter your password.

3. Confirm that the iPhone appears under **This PC** in file explorer.

4. Run the script with Windows PowerShell:

   ```powershell
   powershell -NoProfile -ExecutionPolicy Bypass -File "C:\Users\user\Desktop\iPhoneMediaTransfer.ps1"
   ```
   OR
   
   Download the command line launcher
   
   ```text
   Run iPhone Media Transfer.bat
   ```
   And double click to run. 
   
## Default Destination

The default destination of output is the desktop. You can select your own folder via the **Browse** button or type in the file path.

## Output Data Structure

```text
-YYYY
--MM
---DD
```

## Date Filtering and Fast Scan

When the date-range filter is enabled, the script first uses source-folder names to avoid scanning unrelated months, wasting time.

For example, when importing files from `2000/01/01-2000/01/01`, the script can enter folders such as:

```text
200001
200001_a
200001_b
200001_c
```

It can skip unrelated month folders such as:

```text
202211_a
202306_a
202411_b
```

After scanning the relevant month folders, the script applies the exact file-level date filter.

A typical log includes:

```text
Date range: 2022/07/17 -> 2022/07/17
Scanning folders...
Fast scan: source month folders outside the selected range will be skipped.
Found 125 candidate files.
After date filter: 23 files.
```

## File Conflict Options

You can choose to skip, over write, or rename it with _1 suffix when a file conflict happens before the transfering process starts. It prevents the annoying pop-up window from showing up and stopping the process. 

### Skip

If the proposed destination filename already exists, the source file is not copied.

Example:

```text
Skip (exists): source=202207_b\IMG_5161.JPG | destination=2022\07\IMG_5161[1].JPG
```

### Overwrite

The existing proposed destination file is replaced.

Use this option carefully because the existing file is removed before the replacement copy begins.

### Rename with `_1`, `_2`, and so on

When a proposed destination filename already exists, the script finds the next available filename.

Example:

```text
IMG_5161.JPG
IMG_5161_1.JPG
IMG_5161_2.JPG
```

## Conditional Temporary Staging

The script uses temporary staging only when necessary:

1. The source folder ends in `_b`, `_c`, `_d`, or a later letter, because the copied file must receive a `[number]` suffix.
2. Rename mode is selected and the proposed destination filename already exists.

Normal files from folders without a suffix, or from `_a`, are copied directly without a staging folder when possible.

The temporary staging location is:

```text
%LOCALAPPDATA%\iPhoneMediaCopy\Staging
```

This normally resolves to:

```text
C:\Users\<WindowsUserName>\AppData\Local\iPhoneMediaCopy\Staging
```

The script attempts to delete:

```text
%LOCALAPPDATA%\iPhoneMediaCopy\Staging
```

and the empty parent folder:

```text
%LOCALAPPDATA%\iPhoneMediaCopy
```

Cleanup runs:

- After each staged copy
- When **Abort** is pressed
- After a copy error
- When the complete operation exits

## File Integrity Verification

When verification is enabled, the script compares the copied local file size with the size reported by the MTP source item.

For supported video and audio formats, the script also attempts to read the local media duration.

Example log output:

```text
Copied: IMG_5161.JPG -> 2022\07\IMG_5161.JPG [VERIFY OK size=4281932]
```

MTP metadata may occasionally be missing or unavailable. A verification failure does not always mean the file content is corrupt, so manually inspect files reported as failed.

## Log Details

The log identifies the original source folder, original filename, proposed destination name, and source-folder suffix number.

Example:

```text
Source: 202207_b\IMG_5161.JPG | Proposed: IMG_5161[1].JPG | FolderSuffix=1
```

This makes it easier to diagnose incorrect source-folder detection or conflict handling.

## Abort Button

Selecting **Abort** requests cancellation as soon as the current MTP operation allows it.

Because Windows Shell MTP copying can be asynchronous, cancellation may not be instantaneous. The script performs best-effort cleanup immediately and retries cleanup when the active copy operation returns.

For direct copies, the script also attempts to remove an incomplete destination file after an abort.

## Troubleshooting

These are possible solutions, not one-by-one steps or procedure. Try them as you please. 

### iPhone not found

- Unlock the iPhone.
- Tap **Trust** if prompted.
- Reconnect the USB cable.
- Confirm the iPhone is visible under **This PC** in file explorer.
- Try another USB port or cable.
- Close other applications that may be accessing the iPhone.

### No files match the date range

- Confirm that the files expose a readable Date Taken or Date Modified value through Windows MTP.
- Try disabling the date filter to confirm that files are visible.
- Review the log for the number of files without a readable media date.
- Confirm the source folders use the expected `YYYYMM_a` naming convention.

### `_b` files do not receive `[1]`

The log should show:

```text
Source: 202207_b\IMG_5161.JPG | Proposed: IMG_5161[1].JPG | FolderSuffix=1
```

If `FolderSuffix=0`, confirm that the source folder name actually ends with `_b` and does not contain trailing spaces or other characters.

### Scanning takes too long

- Make sure that the iPhone is connected and unlocked. 
- Enable the date-range filter so fast month-folder filtering can skip unrelated folders.
- Use the narrowest practical date range.
- Keep the iPhone unlocked during scanning.
- Avoid running other applications that are reading the iPhone simultaneously.

### Temporary folder remains after abort

Close the script and manually remove the following folder if no import is still running:

```text
%LOCALAPPDATA%\iPhoneMediaCopy
```

Do not delete the folder while the script or Windows MTP copy operation is still active.

## Limitations

- Windows Shell MTP operations are asynchronous and can pause while the iPhone prepares large files.
- Abort requests may take time to take effect.
### - iCloud-optimized media may need to be downloaded by the iPhone before transfer. This is usually why you can't find the photo or video that you see on your phone. 
- MTP may not expose reliable capture dates or file sizes for every item.
- Video duration metadata may be unavailable until Windows finishes indexing the copied file.
- The script depends on the source-folder naming convention for fast month filtering and `[number]` suffix generation.

## Safety Recommendations

- Test with a small date range and a separate destination folder first.
- Use **Skip** for the safest initial import.
- Review verification failures before deleting media from the iPhone.
- Keep a separate backup of important media.
- Do not disconnect the iPhone during an active copy.

## Contributing and Licensing

I do not intend to maintain this project often. Feel free to pull request or make a version of your own base on this project (but cite me please). 
Bug reports, security reports and improvements are welcome. When opening an issue, include:

- Windows version
- PowerShell version
- iTunes service version
- iPhone and iOS version
- Source-folder name
- Selected conflict mode
- Relevant log lines
- Whether the date-range filter was enabled

Do not include personal photos, private filenames, device identifiers, or other sensitive information.

## Disclaimer and Disclosure

This project is created under the assitance of Microsoft Copilot AI and Grok. 

This project is provided without warranty. Verify copied files before removing originals from the iPhone. The project is not affiliated with or endorsed by Apple or Microsoft.

Oh, and fuck Apple. I had to spend time making this all because of your shitty design. 
