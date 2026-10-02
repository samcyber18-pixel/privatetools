PrivateTools — PDF file selection fix

Replace the existing office-tools.js in the GitHub repository with the supplied file.
This keeps the current Office conversion implementations and makes file selection more robust:
- responds to both change and input events
- accepts PDF by extension or MIME type
- lets clicking the drop zone open the file picker
- shows the selected filename and size
- handles browser FileList timing more defensively
