# Word Smith - MS Word Add-in

## Overview
An MS Word add-in that displays synonyms, describing words, and related words in a compact side panel. Auto-fetches as you type.

## Project Structure
```
├── index.html          # Redirect to taskpane
├── taskpane.html       # Main task pane UI
├── taskpane.css        # Styling (compact design)
├── taskpane.js         # Office.js logic & API calls
├── functions.html      # Required for Office add-in
├── manifest.xml        # Office Add-in manifest
└── README.md           # GitHub instructions
```

## Features
- Auto-fetch synonyms as you type
- Describing Words (adjectives for nouns, like describingwords.io)
- Synonyms, Related Words, Similar Sounding
- Compact layout with color-coded categories
- Click any word to replace in document

## API
Uses the free Datamuse API (https://www.datamuse.com/api/) - no API key required.

## GitHub Pages Deployment

1. Push this code to GitHub
2. Enable GitHub Pages in repository Settings → Pages
3. Update manifest.xml with your GitHub Pages URL
4. Sideload manifest.xml into Word

## Recent Changes
- January 2026: Restructured for GitHub Pages hosting
