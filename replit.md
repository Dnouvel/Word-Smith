# Word Synonym Finder Add-in

## Overview
An MS Word add-in that displays synonyms and related words in a side panel. This is a static web application designed to be hosted on GitHub Pages.

## Project Structure
```
├── index.html          # Redirect to taskpane
├── taskpane.html       # Main task pane UI
├── taskpane.css        # Styling
├── taskpane.js         # Office.js logic & API calls
├── functions.html      # Required for Office add-in
├── manifest.xml        # Office Add-in manifest (update with your GitHub URL)
└── README.md           # GitHub instructions
```

## Features
- Search for synonyms by typing a word
- Get selected word from Word document
- Three categories of results: Synonyms, Related Words, Similar Sounding
- Click any word to insert it into the document

## API
Uses the free Datamuse API (https://www.datamuse.com/api/) - no API key required.

## GitHub Pages Deployment

1. Push this code to GitHub
2. Enable GitHub Pages in repository Settings → Pages
3. Update manifest.xml with your GitHub Pages URL
4. Sideload manifest.xml into Word

## Recent Changes
- January 2026: Restructured for GitHub Pages hosting
