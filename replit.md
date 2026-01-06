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
        ├── taskpane.js     # Office.js logic & API calls
        └── functions.html  # Required for Office add-in loading
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

### Step 1: Get Your Public URL
After publishing/deploying, note your public HTTPS URL.

### Step 2: Update the Manifest
Edit `manifest.xml` and replace the Replit URL with your deployed URL:
- Update the `<SourceLocation>` URL
- Update the `<SupportUrl>` URL
- Update the `<AppDomain>` entries

### Step 3: Sideload the Add-in
**For Word Desktop:**
1. Open Word
2. Go to **Insert** > **My Add-ins** > **Upload My Add-in**
3. Browse and select your `manifest.xml` file
4. Click **Upload**

**For Word Online:**
1. Open Word Online (office.com)
2. Go to **Insert** > **Office Add-ins**
3. Click **Upload My Add-in** in the top-right
4. Upload your `manifest.xml` file

### Step 4: Open the Side Panel
Once sideloaded, the add-in opens automatically as a side panel.

## Using the Add-in
1. **Type a word** in the search box and click "Find"
2. **Or select text** in your Word document and click "Get Selected Word"
3. Browse the results in three categories: Synonyms, Related Words, Similar Sounding
4. **Click any word** to replace your selected text in the document

## Development Notes
- Server runs on port 5000
- Office.js warning in browser is normal (appears when not running inside Word)
- The add-in requires HTTPS hosting - Replit handles this automatically

## Recent Changes
- January 2026: Initial creation with simplified manifest for easier sideloading
