# PrivateTools v2

A small static toolbox for browser-based image and PDF utilities.

## Included tools

1. Image Compressor
2. Image Converter & Resizer
3. PDF Merger
4. PDF Splitter

## Privacy model

The four core file operations are designed to process selected files in the browser without an application backend receiving the file contents. The current PDF tools load `pdf-lib` from a pinned jsDelivr URL, so the PDF engine requires an internet connection to initialize.

## Hosting

The project is static HTML/CSS/JavaScript and can be deployed to GitHub Pages or another static host.

## Local testing

From this directory, run:

```bash
python -m http.server 8000
```

Then open `http://localhost:8000/`.

## Notes

- PNG browser encoding is lossless; the quality slider does not make PNG output lossy.
- Very large files can be limited by device/browser memory.
- Unusual or encrypted PDFs may not be supported by the PDF library.
