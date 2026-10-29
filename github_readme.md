# Certificate Batch Generator (Offline Version) - User & Maintenance Manual

**Applicable File:** `certificate-generator-offline.html`  
**Update Date:** October 29, 2026

---

## 1. What is This?

The Certificate Batch Generator is a single HTML file (approximately 3.2 MB). Once opened in a browser, you can load a certificate background template, adjust text positions and styles on the preview, upload an Excel or CSV name list, generate a PDF certificate for each individual in bulk, and download them packaged as a ZIP file.

* **No installation, no internet connection, and no server required.**
* **Name lists and certificates are processed entirely within your browser and are never uploaded anywhere.**
* **Not included:** Email delivery, cloud storage, multi-user collaboration, or account permissions.

---

## 2. Quick Start

1. Double-click to open the file using **Chrome** or **Edge**.
2. **Template:** Use the preset forest template directly, or upload your own background under "Template Background".
3. **Name List:** Upload an Excel or CSV file under "2. Name List". The system will automatically map the columns.
4. **Text Fields:** Confirm the corresponding columns for each text field under "3. Text Fields", and drag the text to the correct position in the right-side preview.
5. Click **"Generate & Download ZIP"**, and extract the folder to find individual certificates inside the `pdf` folder.

> *Tip: It is recommended to test with 2 to 3 rows of sample data to verify the layout before processing the complete list.*

---

## 3. Detailed Instructions

### 3.1 Template Background

| Format | Description |
| :--- | :--- |
| **PNG, JPG, WebP** | Use directly. |
| **PDF** | Reads only the **1st page**, converting it into an image with a long edge of approximately 2400 pixels; multi-page PDFs will prompt a notice that only the 1st page is used. |
| **PowerPoint, Keynote** | Cannot be uploaded directly. Please export to PDF first in PowerPoint (`File -> Export -> PDF`) or save as PNG; in Google Slides, use `Download -> PDF`. |

**Key Points:**
* The background must be clean and free of placeholder text such as pre-existing names. The system cannot read text embedded in the image nor erase it; text positioning is entirely determined by your dragging.
* Both landscape and portrait orientations are supported. The template card will display the background dimensions and orientation.
* When changing templates, texts will retain their previous relative positions, so you will usually need to re-drag and adjust font sizes—especially for landscape templates.
* Background resolution determines the final output quality; images that are too small will appear blurry when scaled up.
* Uploaded backgrounds are not saved; refreshing the page reverts to the preset template. Click "Restore Default Template" to switch back at any time.

### 3.2 Preparing the Name List

* **Formats:** `.xlsx`, `.xls`, `.csv`. Only the **first worksheet** is read.
* **Header Row:** The system automatically identifies the topmost row containing at least two non-empty cells as the header row; any large titles or blank rows above it are ignored.
* **Blank Headers:** Columns with blank headers but existing data are displayed as "Column C", etc.; entirely blank columns are hidden; duplicate headers are automatically numbered, e.g., "Name (2)".
* **Name Column:** A name column is mandatory. Rows without a name are skipped, and the PDF file name is derived from the name.
* **Dates:** Excel date formats are automatically converted to Malay format, e.g., `28 Oktober 2026`; text-format dates are displayed as-is; CSV dates like `2026-10-28` are also converted.
* **CSV Encoding:** Automatically recognizes UTF-8 and GBK. Big5 may cause garbled text; please save as "CSV UTF-8" in Excel or use `.xlsx` directly.
* **Headers:** Supports Chinese, Malay, and English. See the Appendix for automatic mapping keywords.

### 3.3 Text Field Settings

Click the small tabs at the top (`NAMA`, `NAMA SEKOLAH`, etc.) to select the text you want to edit.

| Setting | Function |
| :--- | :--- |
| **Name** | For identification purposes only; will not be printed on the certificate. |
| **Data Source Column** | Selects the name list header. The dropdown will display the first row of data as an example. Selecting "None / Fixed Text" uses the fixed text below. |
| **Fixed Text** | Fixed text to be printed, or fallback text when a column has no data. |
| **Font Size** | Relative unit from 5 to 60; WYSIWYG preview. |
| **Max Width** | Range from 40 to 400; text exceeding this width will automatically shrink or wrap. |
| **Color, Alignment, Font** | Color, center/left/right alignment, and 6 available fonts. |
| **Bold, Uppercase, Display** | Unchecking "Display" temporarily hides the text without losing settings. |
| **Delete Field** | Permanently deletes this text field (at least one must remain). |
| **+ Add Text** | Adds custom text mapped to any list column, such as class or certificate number. |

* **Auto-adaptation:** Excessively long text shrinks first (to a minimum of ~65% of its original size); if it still doesn't fit and contains spaces, it splits evenly into two lines. Chinese names do not contain spaces and will not wrap; they will continue to shrink.

### 3.4 Preview Operations

* Drag text with your mouse or finger to reposition it; click on text to adjust its settings on the left panel.
* Clicking a text item displays a green border in the preview, allowing you to fine-tune its position using the arrow keys (1 pixel per press, or 10 pixels while holding `Shift`).
* Pressing `Delete` or `Backspace` triggers a confirmation prompt to delete the selected text field.
* Pressing `Backspace` or `Delete` inside a text input box on the left will only delete the characters inside, not the entire text field.
* The preview image is smaller than the actual output; the actual output has a higher resolution.

### 3.5 Generation and Output

* Click **"Generate & Download ZIP"**. The page will show progress, and automatically download `certificates.zip` upon completion.
* The ZIP file contains a `pdf` folder with an individual PDF for each person. Checking "Also include combined PDF" adds an `all_certificates.pdf` for convenient batch printing.
* File names are derived from names; illegal characters are removed, and duplicate names automatically append `_2`, `_3`, etc.
* **PDF Page Width:** A4 size (Portrait: 595 pt, Landscape: 842 pt); height scales according to the background ratio. The default template is slightly longer than A4—please select "Fit to page" when printing.
* **Output Image Width:** Approximately 1600 to 2400 pixels at 92% JPEG quality.
* In the output PDF, text is part of the image and cannot be selected or copied.

---

## 4. Important Notes

1. Please back up your name list independently; the webpage does not save name lists.
2. Before formal generation, use a few records to check the layout—especially for the longest names and school titles.
3. Fonts depend on the user's computer. The default serif font sequentially checks for `Liberation Serif`, `Times New Roman`, and `Times`; appearance may vary slightly across computers. For Chinese names, select "Chinese Sans-Serif".
4. Layout settings are automatically saved in the browser. Switching browsers, switching computers, or clearing browser data will erase them. There is currently no "Reset Layout" button or layout export feature.
5. Since name lists contain student personal data, ensure proper consent has been obtained and use them solely for issuing certificates.
6. The embedded Excel reading library (`SheetJS 0.18.5`) is an older version with known security vulnerabilities. The primary risk involves opening untrusted files. Only process name lists you trust.
7. The default template contains Tzu Chi and PaGamO logos and designs; please verify authorization before sharing publicly.
8. For physical printing, the original background resolution must be sufficient. The default template is only 391×564 pixels, which is suitable only for screens and standard office printing.

---

## 5. Troubleshooting FAQ

| Symptom | Solution |
| :--- | :--- |
| **Opens to a bunch of code** | Open with Chrome or Edge, not a text editor; if received on a mobile device, download it to a computer first. |
| **Upload list shows read failure** | File may be corrupted or password-protected; save as `.xlsx` and try again. |
| **Dropdown doesn't show desired column** | Check if the header row is positioned at the top and contains at least two non-empty cells. |
| **Prompt asking to map name column** | Make sure the data source for the `NAMA` text field is mapped to the name column in your list. |
| **Clicking generate does nothing** | Upload a name list first. |
| **Text is too small or mispositioned** | Landscape templates require manual adjustment of position and font size. |
| **Chinese characters show as squares** | Select "Chinese Sans-Serif" as the font. |
| **Overlapping text on the certificate** | The background image originally contained text; please remove it in your original design software. |
| **PDF template reading failed** | The PDF may be encrypted or corrupted; export it as a PNG instead. |
| **Garbled CSV content** | Save as "CSV UTF-8" or use `.xlsx`. |

---

## 6. How It Is Built

This tool is a pure frontend webpage with no backend, no database, and no build tools—packaged entirely within a single HTML file.

* **Browser (Local HTML):**
  * **Interface & Logic:** Native HTML / CSS / JavaScript
  * **SheetJS 0.18.5:** Reads Excel and CSV files
  * **pdf.js 3.11.174:** Reads PDF templates
  * **Canvas:** Renders background and text, outputs to JPEG
  * **jsPDF 2.5.1:** Generates one PDF page per person
  * **JSZip 3.10.1:** Packages files into a ZIP download

**Data Flow:** Upload the name list $\rightarrow$ parse headers and data $\rightarrow$ for each row, draw the background on the Canvas, draw text according to field settings, convert to JPEG, embed into a PDF, and bundle everything into a ZIP.

**Key Design Elements:**
* **Positioning:** Saved as percentages (relative to background width and height). Font size and maximum width use a "400-pixel wide background" as a baseline unit, maintaining general alignment across different background sizes.
* **Offline Capability:** All four libraries are directly embedded within the file. The `pdf.js` worker is base64-encoded and loaded as a Blob at runtime, eliminating the need for external network sources.
* **Persistence:** Browser `localStorage` only saves text field layout configurations, not background images or name lists.
* **Asynchronous Generation:** Yields execution time back to the page during batch generation so the progress bar updates smoothly without freezing.
* **File Size:** Around 3.2 MB, with the embedded libraries accounting for the majority of the size.

---

## 7. Maintenance & Modification

Routine use requires no maintenance: no accounts, no services, and no expiring licenses. It is recommended to keep a backup copy of the file and copy it directly when switching computers.

To modify the code, open the HTML file in a text editor and search for the following keywords:

| Desired Change | Search Keyword |
| :--- | :--- |
| **Default text positions and sizes** | `defaults` and `mk(` |
| **Available fonts** | `FONTS` |
| **Auto-mapping header keywords** | `CANDS` |
| **Month names and date formats** | `MONTHS`, `fmtDate` |
| **Default text color** | `const C =` |
| **Default template image** | `DEFAULT_TPL` (Base64-encoded PNG string; replace directly) |
| **Output image quality** | `outW` and `image/jpeg` |

> *Always keep a backup of the original file before making modifications. Test previews and generation results using sample data after editing. Upgrading embedded libraries requires re-downloading and replacing corresponding code blocks; this is recommended for web developers.*

---

## 8. Distribution

* Simply share this single HTML file with others; the default template is already embedded, and no internet is required.
* Gmail or Outlook may occasionally block `.html` attachments; it is recommended to compress the file into a ZIP or share it via cloud storage.
* Advise recipients to use a computer; mobile browsers typically cannot properly run local HTML files.
* Layout adjustments do not travel with the file; recipients will need to configure layouts themselves.
* If you wish to distribute a version without the default template, you must create a blank version with the template removed, as embedded images can be extracted from the source code.

---

## 9. Test Log & Known Limitations

**Tested (Chromium 141, fully offline):**
* All four libraries loaded successfully with zero network requests.
* Excel files with top banners and Chinese titles; GBK and UTF-8 encoded CSVs.
* Chinese names in certificates and file names; generating ZIP archives.
* Landscape PDF templates generating landscape A4 PDFs; uploading `.pptx` files triggers an export prompt.
* Arrow key fine-tuning, `Delete` confirmation, and input box `Backspace` isolation.
* Flawless page execution with zero browser errors.

**Untested:** Safari, Firefox, physical Windows & Mac environments, mobile devices, large lists with hundreds of users, and actual physical print outputs.

**Known Limitations:**
* Cannot directly read PowerPoint files.
* Text in output PDFs is rasterized as images and cannot be selected.
* Background text cannot be automatically erased.
* Only reads the first worksheet of a name list and the 1st page of a PDF template.
* Lacks layout export, import, and reset functions.

---

## 10. Appendix

### Default Text Fields

| Name | Default Style |
| :--- | :--- |
| **NAMA** | Bold, Uppercase, Centered |
| **NAMA SEKOLAH** | Uppercase, Centered |
| **TARIKH** | Centered, automatic date conversion to Malay format |
| **EMAIL** | Small font, located to the right of the signature; uncheck "Display" if unused |

### Available Fonts
Serif (`Times`), Sans-Serif (`Arial`), `Georgia`, Monospace (`Courier`), Script (`Script`), and Chinese Sans-Serif (`中文 黑体`).

### Auto-Mapping Header Keywords
* **Name:** `nama`, `name`, `nama penuh`, `姓名`, `名字`, `学生姓名`, `參加者`, etc.
* **School:** `nama sekolah`, `sekolah`, `school`, `学校`, `學校`
* **Date:** `tarikh`, `date`, `日期`
* **Email:** `email`, `e-mail`, `邮箱`, `邮件`, `电邮`, `电子邮件`, `信箱`, etc. (Simplified & Traditional Chinese)

Headers containing these terms will be automatically matched. Any unmatched columns can be manually selected from the dropdown menu.