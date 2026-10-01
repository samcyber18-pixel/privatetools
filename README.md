# PrivateTools v1

Four browser tools:

1. Compress images
2. Convert & resize images
3. Merge PDFs
4. Split PDFs

## Current behavior

Image files are decoded and processed in the browser. The PDF operations are performed in the browser with `pdf-lib` loaded from a pinned jsDelivr URL.

The selected file contents are not sent to an application backend by this code. The PDF engine itself is loaded from the CDN as JavaScript, so an internet connection is currently required for the PDF tools to initialize.

For the strongest privacy and offline story before public launch, vendor/bundle the pinned PDF library locally instead of using the CDN script in `index.html`.

## Local testing

Run a local HTTP server from this directory:

```bash
python -m http.server 8000
```

Open `http://localhost:8000`.

## Deployment

This is a static site and does not need a database or application server for these four tools. It can be deployed to a static host.

## Notes

- Browser PNG encoding is lossless; the quality slider does not make PNG output lossy. The compressor offers WebP/JPG for stronger reduction.
- Very large files are limited by device/browser memory.
- Some encrypted or unusual PDFs may not be supported by the PDF library. The UI reports failures instead of silently producing a bad result.
