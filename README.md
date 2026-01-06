# Word Synonym Finder Add-in

An MS Word add-in that displays synonyms and related words in a side panel.

## Features

- Search for synonyms by typing a word
- Get selected word from your Word document
- Three categories of results:
  - **Synonyms**: Words with the same meaning
  - **Related Words**: Conceptually related words
  - **Similar Sounding**: Words that sound alike
- Click any word to insert it into your document

## Setup Instructions

### Step 1: Enable GitHub Pages

1. Go to your repository on GitHub
2. Click **Settings** → **Pages**
3. Under "Source", select **Deploy from a branch**
4. Choose **main** branch and **/ (root)** folder
5. Click **Save**
6. Wait a few minutes for GitHub Pages to deploy

Your site will be available at: `https://YOUR_USERNAME.github.io/YOUR_REPO/`

### Step 2: Update the Manifest

Edit `manifest.xml` and replace `YOUR_USERNAME` and `YOUR_REPO` with your actual values:

- `YOUR_USERNAME` → your GitHub username
- `YOUR_REPO` → your repository name

For example, if your repo is `github.com/johnsmith/word-synonym-finder`:
- Replace `YOUR_USERNAME` with `johnsmith`
- Replace `YOUR_REPO` with `word-synonym-finder`

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
