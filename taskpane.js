let isOfficeInitialized = false;
let autoFetchEnabled = true;
let lastDetectedWord = '';
let isSearching = false;

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        isOfficeInitialized = true;
        console.log('Office.js initialized for Word');
        startAutoDetection();
    }
    
    document.getElementById('search-btn').addEventListener('click', handleSearch);
    document.getElementById('word-input').addEventListener('keypress', (e) => {
        if (e.key === 'Enter') {
            handleSearch();
        }
    });
    document.getElementById('get-selection-btn').addEventListener('click', getSelectedWord);
    document.getElementById('auto-fetch-toggle').addEventListener('change', (e) => {
        autoFetchEnabled = e.target.checked;
        updateAutoFetchStatus();
    });
    
    updateAutoFetchStatus();
});

function updateAutoFetchStatus() {
    const statusEl = document.getElementById('auto-status');
    if (autoFetchEnabled) {
        statusEl.textContent = 'Auto-fetch: ON - Type in your document, synonyms appear automatically!';
        statusEl.className = 'auto-status active';
    } else {
        statusEl.textContent = 'Auto-fetch: OFF - Use manual search above';
        statusEl.className = 'auto-status inactive';
    }
}

function startAutoDetection() {
    setInterval(detectCurrentWord, 300);
}

async function detectCurrentWord() {
    if (!isOfficeInitialized || !autoFetchEnabled || isSearching) return;
    
    try {
        await Word.run(async (context) => {
            const selection = context.document.getSelection();
            selection.load('text');
            
            const range = selection.getRange('Start');
            const paragraph = range.paragraphs.getFirst();
            paragraph.load('text');
            
            await context.sync();
            
            const paragraphText = paragraph.text;
            const selectionText = selection.text;
            
            if (selectionText && selectionText.trim().length > 0) {
                return;
            }
            
            const cursorPosition = findCursorPosition(paragraphText, selectionText);
            const wordBeforeCursor = getWordBeforeCursor(paragraphText, cursorPosition);
            
            if (wordBeforeCursor && 
                wordBeforeCursor.length > 1 && 
                wordBeforeCursor !== lastDetectedWord &&
                /^[a-zA-Z]+$/.test(wordBeforeCursor)) {
                
                lastDetectedWord = wordBeforeCursor;
                document.getElementById('word-input').value = wordBeforeCursor;
                await autoSearch(wordBeforeCursor);
            }
        });
    } catch (error) {
        console.log('Auto-detection cycle:', error.message);
    }
}

function findCursorPosition(paragraphText, selectionText) {
    return paragraphText.length;
}

function getWordBeforeCursor(text, cursorPos) {
    const textUpToCursor = text.substring(0, cursorPos).trim();
    
    const lastSpaceIndex = textUpToCursor.lastIndexOf(' ');
    const lastNewlineIndex = textUpToCursor.lastIndexOf('\n');
    const lastSeparator = Math.max(lastSpaceIndex, lastNewlineIndex);
    
    let word;
    if (lastSeparator === -1) {
        word = textUpToCursor;
    } else {
        word = textUpToCursor.substring(lastSeparator + 1);
    }
    
    return word.replace(/[^a-zA-Z]/g, '').toLowerCase();
}

async function autoSearch(word) {
    if (isSearching) return;
    isSearching = true;
    
    showLoading(true);
    hideError();
    hideResults();
    
    try {
        const results = await fetchSynonyms(word);
        displayResults(word, results);
    } catch (error) {
        console.error('Auto-search error:', error);
    } finally {
        showLoading(false);
        isSearching = false;
    }
}

async function getSelectedWord() {
    if (!isOfficeInitialized) {
        document.getElementById('word-input').value = 'happy';
        handleSearch();
        return;
    }

    try {
        await Word.run(async (context) => {
            const selection = context.document.getSelection();
            selection.load('text');
            await context.sync();
            
            const selectedText = selection.text.trim();
            if (selectedText) {
                const word = selectedText.split(/\s+/)[0].replace(/[^a-zA-Z]/g, '');
                if (word) {
                    document.getElementById('word-input').value = word;
                    lastDetectedWord = word.toLowerCase();
                    handleSearch();
                }
            }
        });
    } catch (error) {
        showError('Could not get selected text. Please select a word in your document.');
        console.error('Error getting selection:', error);
    }
}

async function handleSearch() {
    const input = document.getElementById('word-input');
    const word = input.value.trim().toLowerCase().replace(/[^a-zA-Z]/g, '');
    
    if (!word) {
        showError('Please enter a word to search');
        return;
    }
    
    isSearching = true;
    showLoading(true);
    hideError();
    hideResults();
    
    try {
        const results = await fetchSynonyms(word);
        displayResults(word, results);
    } catch (error) {
        showError('Failed to fetch synonyms. Please try again.');
        console.error('Search error:', error);
    } finally {
        showLoading(false);
        isSearching = false;
    }
}

async function fetchSynonyms(word) {
    const [synonymsRes, relatedRes, similarRes] = await Promise.all([
        fetch(`https://api.datamuse.com/words?rel_syn=${encodeURIComponent(word)}&max=20`),
        fetch(`https://api.datamuse.com/words?ml=${encodeURIComponent(word)}&max=15`),
        fetch(`https://api.datamuse.com/words?sl=${encodeURIComponent(word)}&max=10`)
    ]);
    
    const [synonyms, related, similar] = await Promise.all([
        synonymsRes.json(),
        relatedRes.json(),
        similarRes.json()
    ]);
    
    return {
        synonyms: synonyms.map(w => w.word),
        related: related.map(w => w.word).filter(w => !synonyms.some(s => s.word === w)),
        similar: similar.map(w => w.word).filter(w => w !== word)
    };
}

function displayResults(word, results) {
    document.getElementById('searched-word').textContent = word;
    document.getElementById('current-word').classList.remove('hidden');
    
    const hasResults = results.synonyms.length > 0 || results.related.length > 0 || results.similar.length > 0;
    
    if (!hasResults) {
        document.getElementById('no-results').classList.remove('hidden');
        return;
    }
    
    if (results.synonyms.length > 0) {
        const synonymsList = document.getElementById('synonyms-list');
        synonymsList.innerHTML = results.synonyms.map(w => 
            `<span class="word-chip" data-word="${w}">${w}</span>`
        ).join('');
        document.getElementById('synonyms-section').classList.remove('hidden');
    }
    
    if (results.related.length > 0) {
        const relatedList = document.getElementById('related-list');
        relatedList.innerHTML = results.related.map(w => 
            `<span class="word-chip" data-word="${w}">${w}</span>`
        ).join('');
        document.getElementById('related-section').classList.remove('hidden');
    }
    
    if (results.similar.length > 0) {
        const similarList = document.getElementById('similar-list');
        similarList.innerHTML = results.similar.map(w => 
            `<span class="word-chip" data-word="${w}">${w}</span>`
        ).join('');
        document.getElementById('similar-section').classList.remove('hidden');
    }
    
    document.querySelectorAll('.word-chip').forEach(chip => {
        chip.addEventListener('click', () => insertWord(chip.dataset.word));
    });
}

async function insertWord(word) {
    if (!isOfficeInitialized) {
        alert(`Word to insert: "${word}"\n\nNote: This feature works when running inside MS Word.`);
        return;
    }
    
    try {
        await Word.run(async (context) => {
            const selection = context.document.getSelection();
            const range = selection.getRange('Start');
            const paragraph = range.paragraphs.getFirst();
            paragraph.load('text');
            await context.sync();
            
            const paragraphText = paragraph.text;
            const wordToReplace = lastDetectedWord;
            
            if (wordToReplace && paragraphText.toLowerCase().includes(wordToReplace)) {
                const searchResults = context.document.body.search(wordToReplace, { matchCase: false, matchWholeWord: true });
                searchResults.load('items');
                await context.sync();
                
                if (searchResults.items.length > 0) {
                    const lastMatch = searchResults.items[searchResults.items.length - 1];
                    lastMatch.insertText(word, Word.InsertLocation.replace);
                    await context.sync();
                    lastDetectedWord = word.toLowerCase();
                }
            } else {
                selection.insertText(word, Word.InsertLocation.replace);
                await context.sync();
            }
        });
    } catch (error) {
        showError('Failed to insert word. Please try again.');
        console.error('Insert error:', error);
    }
}

function showLoading(show) {
    document.getElementById('loading').classList.toggle('hidden', !show);
}

function showError(message) {
    const errorEl = document.getElementById('error-message');
    errorEl.textContent = message;
    errorEl.classList.remove('hidden');
}

function hideError() {
    document.getElementById('error-message').classList.add('hidden');
}

function hideResults() {
    document.getElementById('current-word').classList.add('hidden');
    document.getElementById('synonyms-section').classList.add('hidden');
    document.getElementById('related-section').classList.add('hidden');
    document.getElementById('similar-section').classList.add('hidden');
    document.getElementById('no-results').classList.add('hidden');
    document.getElementById('synonyms-list').innerHTML = '';
    document.getElementById('related-list').innerHTML = '';
    document.getElementById('similar-list').innerHTML = '';
}
