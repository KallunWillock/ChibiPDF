
![vbaXray logo](logo_chibipdf_small.png)
# ChibiPDF

![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)
![Platform: Windows](https://img.shields.io/badge/Platform-Windows-blue.svg)
![VBA](https://img.shields.io/badge/VBA-7.0+-purple.svg)
![Dependencies](https://img.shields.io/badge/Dependencies-None-brightgreen.svg)

A dependency-free PDF rendering and OCR library for VBA. No Adobe Acrobat. No Tesseract. No external DLLs.

ChibiEx allows VBA applications to:

* Extract text from searchable and scanned PDFs using OCR
* Render PDF pages to PNG image files, `StdPicture` objects, memory, or streams
* OCR image files, HBITMAPs, `StdPicture`s, and in-memory streams

Just a single class module and the Windows APIs already present on Windows 10 and later.

---

## Features

### PDF Text Extraction

Extract text from individual pages or entire documents.

```vb
Dim pdf As New ChibiEx

If pdf.LoadFile("C:\Docs\Report.pdf") Then
    ' Extract text from a single page
    Debug.Print pdf.ExtractText(1)
    
    ' Extract text from the entire document
    Debug.Print pdf.ExtractAllText()
End If
```
### OCR for Scanned PDFs
ChibiEx automatically handles image-based PDFs by rendering pages and passing them directly to the Windows OCR engine entirely in memory.

Because ExtractText handles the rendering and OCR pipeline under the hood, extracting text from a scanned PDF requires zero extra code:
```vb
' This works exactly the same for scanned PDFs!
Dim text As String
text = pdf.ExtractText(1) 
```
No temporary files are created.

### PDF Rendering
Render PDF pages to PNG files, byte arrays, StdPictures, or streams.

```vb
' Render all pages
pdf.ToFile "C:\Output\"

' Render a specific page
pdf.ToFile "C:\Output\", 5

' Render a range of pages
pdf.ToFile "C:\Output\", 2, 10

' Render specific pages
pdf.ToFile "C:\Output\", Array(1, 5, 12)

' Render to a StdPicture (for UserForms or Image controls)
Dim pic As StdPicture
Set pic = pdf.ToPicture(1)
```

### Standalone OCR
The OCR engine can be used independently of PDF functionality.
```vb
Dim ocr As New ChibiEx

If ocr.RecogniseFile("C:\OCR\Scan.png") Then
    Debug.Print ocr.ResultText
End If
```

## Why ChibiEx?

Most VBA OCR solutions require third-party software (Adobe Acrobat, Tesseract, GhostScript, etc.) that must be installed on the host machine. From personal experience, this is ordinarily a non-starter in locked-down corporate environments.

ChibiEx uses native Windows WinRT APIs and ships as a single VBA class module.

## Requirements

| Requirement      | Details                                                        |
| ---------------- | -------------------------------------------------------------- |
| Operating System | Windows 10 (Build 10240+) or later                             |
| VBA Host         | Excel, Word, Access, Outlook, PowerPoint, VB6-compatible hosts |
| Architecture     | 32-bit and 64-bit                                              |
| Dependencies     | None                                                           |

---

## Roadmap

ChibiEx is part of the (intended) wider ChibiPDF project.

Planned components include:

| Module      | Purpose                              |
| ----------- | ------------------------------------ |
| ChibiEx     | PDF extraction, rendering and OCR    |
| ChibiScribe | PDF generation and document creation |

---

## Credits

ChibiPDF was created by Kallun Willock.

Special thanks to Frank Schüler (activevb.de) for pioneering 32-bit WinRT interop techniques that helped make this project possible.

---

## License

MIT License
