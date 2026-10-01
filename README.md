# PrivateTools v3

A small static toolbox for browser-based image and PDF utilities, with a community roadmap layer.

## Included tools

1. Image Compressor
2. Image Converter & Resizer
3. PDF Merger
4. PDF Splitter

## Community

The site now has a community page that points users to GitHub Discussions for tool requests, discussion and polls during the zero-budget phase.

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

## Next planned steps

- Enable GitHub Discussions on the repository.
- Add a Giscus-backed comment experience once Discussions is enabled.
- Vendor `pdf-lib` locally so the PDF engine is no longer fetched from a CDN.
- Add analytics/Search Console after the product pages are stable.
- Use actual request and traffic data to choose the next tools.
