PrivateTools DOCX → PDF Fix 6

This version replaces the manual block-splitting workaround with Paged.js for normal browser pagination. It still uses docx-preview to render the Word document and html2canvas/jsPDF for PDF capture.

Files to replace in the existing PrivateTools repository:
- docx-to-pdf.html
- office-tools.js

The page also loads Paged.js 0.4.3 from unpkg. The DOCX file remains in the browser; the file itself is not sent to a PrivateTools backend.

Known limitations:
- This is not Microsoft Word's native layout engine.
- Complex Word features can still render differently.
- The PDF is image-based in this implementation, so PDF text is not independently selectable.
