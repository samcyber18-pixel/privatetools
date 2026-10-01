DOCX to PDF pagination fix 5

Fixes the pagination pass so source page geometry remains attached while the renderer measures content.
Flowing block-level content is grouped into fixed-size page containers before PDF capture.

Replace docx-to-pdf.html and office-tools.js in the PrivateTools repository.
