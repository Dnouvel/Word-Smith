# Word Synonym Finder Add-in

## Overview
An MS Word add-in that displays synonyms and related words in a side panel. Users can search for synonyms by typing a word or selecting text in their Word document.

## Project Structure
```
word-synonym-addin/
├── server.js           # Express server to host the add-in
├── manifest.xml        # Office Add-in manifest configuration
├── package.json        # Node.js dependencies
└── src/
    └── taskpane/
        ├── taskpane.html   # Task pane UI
        ├── taskpane.css    # Styling
        └── taskpane.js     # Office.js logic & API calls
```

## Features
- Search for synonyms by typing a word
- Get selected word from Word document
- Three categories of word results:
  - **Synonyms**: Words with the same meaning
  - **Related Words**: Conceptually related words
  - **Similar Sounding**: Words that sound alike
- Click any word to insert it into the document

## API
Uses the free Datamuse API (https://www.datamuse.com/api/) - no API key required.

## How to Use with MS Word

### Sideloading for Testing
1. Run the development server (starts automatically)
2. Get the public URL where the add-in is hosted
3. Update `manifest.xml` with your actual URLs (replace `localhost:5000`)
4. In Word:
   - Go to **Insert** > **My Add-ins** > **Upload My Add-in**
   - Upload the `manifest.xml` file
5. The add-in will appear in the Home tab

### For Production
- Deploy to HTTPS hosting (required for Office Add-ins)
- Update manifest.xml with production URLs
- Submit to Microsoft AppSource or deploy via SharePoint/Admin Center

## Development
- Server runs on port 5000
- Live reload not included - restart server after changes
- Test UI directly by visiting the server URL in browser

## Recent Changes
- Initial creation (January 2026)
