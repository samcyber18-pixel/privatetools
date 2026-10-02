PrivateTools — PDF file-selection fix 2

Root cause fixed: the previous validation inferred the INPUT file type from the OUTPUT tool name. For example, "pdf-to-docx" incorrectly matched "docx" and rejected PDFs. This version stores sourceExts explicitly per tool and uses those for validation and error messages.

Replace the existing office-tools.js in the GitHub repository, commit to main, and wait for GitHub Pages deployment.
