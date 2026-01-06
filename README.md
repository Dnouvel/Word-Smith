# Word Smith - MS Word Add-in

An MS Word add-in that displays synonyms, describing words, and related words in a compact side panel as you type.

## Features

- **Auto-fetch** - Synonyms appear automatically as you type
- **Describing Words** - Adjectives commonly used with your word (like describingwords.io)
- **Synonyms** - Words with the same meaning
- **Related Words** - Conceptually related words
- **Similar Sounding** - Words that sound alike
- **One-click insert** - Click any word to replace it in your document
- **Compact design** - Small fonts to see all options at once

## Setup Instructions

### Step 1: Enable GitHub Pages

1. Go to your repository on GitHub
2. Click **Settings** → **Pages**
3. Under "Source", select **Deploy from a branch**
4. Choose **main** branch and **/ (root)** folder
5. Click **Save**
6. Wait a few minutes for GitHub Pages to deploy

Your site will be available at: `https://dnouvel.github.io/Word-Smith/`

### Step 2: Manifest Already Configured

The `manifest.xml` is already configured for this repository:
- Source URL: `https://dnouvel.github.io/Word-Smith/taskpane.html`

### Step 3: Sideload the Add-in in Word

**For Word Desktop:**
1. Open Word
2. Go to **Insert** → **My Add-ins** → **Upload My Add-in**
3. Upload your updated `manifest.xml` file

**For Word Online:**
1. Open Word Online at office.com
2. Go to **Insert** → **Office Add-ins**
3. Click **Upload My Add-in**
4. Upload your updated `manifest.xml` file

## Using the Add-in

1. Type a word in the search box and click **Find**
2. Or select text in your document and click **Get Selected Word**
3. Browse Synonyms, Related Words, and Similar Sounding results
4. Click any word to insert it into your document

## API

Uses the free [Datamuse API](https://www.datamuse.com/api/) - no API key required.

## License

MIT
