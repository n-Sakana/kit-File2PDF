# File2PDF

File2PDF is a Windows tool that combines PDF, PowerPoint, Word and Excel files into one PDF for duplex printing. It adds a blank page after each document with an odd page count, including the last, so each document starts on the front of a sheet. Each document gets a bookmark.

## Requirements

- Windows PowerShell 5.1.
- The corresponding Microsoft Office desktop app to convert PowerPoint, Word or Excel files. PDF files do not require Office.

No installer or administrator rights are required. Save [File2PDF.cmd](File2PDF.cmd) to a writable folder.

## Usage

Double-click `File2PDF.cmd` and choose a folder, or drag a folder onto the file. It includes subfolders and lists supported files by relative path, ignoring case.

Use the Up and Down keys to move, Enter to select, and Esc to go back.

- **Merge in listed order** includes every listed file.
- **Choose merge order** lets you select files in the order to merge them. Press Enter again to remove a selection. Choose **Continue** when ready; unselected files are excluded.
- **Settings** controls A4 page sizing, Excel scaling, error handling and what happens after a merge. Changes are saved to `File2PDF.ini` beside the script.

Review the file order and output path, then choose **Merge**. If the output exists, the confirmation menu offers **Overwrite and merge**.

The PDF is saved in the selected folder as `<folder>結合版<yyyyMMdd>.pdf`. The suffix `結合版` means "merged version" and is kept so previous outputs in that folder can be excluded from later merges.

## Supported files and limitations

Supported extensions: `.pdf`, `.ppt`, `.pptx`, `.doc`, `.docx`, `.xls`, `.xlsx`, `.xlsm`.

Source files are opened read-only and are not saved. Encrypted PDFs are not supported. By default, an error stops the merge; enable **Skip files with errors** to continue with the remaining files.

A4 sizing and fitting Excel sheets to one page wide are enabled by default. A4 sizing preserves orientation and proportions, but removes annotations, including links, from resized pages. The output creates document bookmarks and does not preserve the source PDFs' tag structure.
