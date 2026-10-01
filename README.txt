PrivateTools DOCX -> PDF fix 3

Replace the existing docx-to-pdf.html and office-tools.js in the repository.

This version bypasses docx-preview's renderAsync DOM helper and uses parseAsync + renderDocument so the application controls both output containers directly. It also keeps Word's last-rendered-page-break markers enabled for better pagination.
