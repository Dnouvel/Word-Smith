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
    statusEl.textContent = autoFetchEnabled ? 'Auto' : 'Manual';
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
            
            const wordBeforeCursor = getWordBeforeCursor(paragraphText, paragraphText.length);
            
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

function getWordBeforeCursor(text, cursorPos) {
    const textUpToCursor = text.substring(0, cursorPos).trim();
    const lastSpaceIndex = textUpToCursor.lastIndexOf(' ');
    const lastNewlineIndex = textUpToCursor.lastIndexOf('\n');
    const lastSeparator = Math.max(lastSpaceIndex, lastNewlineIndex);
    
    let word = lastSeparator === -1 ? textUpToCursor : textUpToCursor.substring(lastSeparator + 1);
    return word.replace(/[^a-zA-Z]/g, '').toLowerCase();
}

async function autoSearch(word) {
    if (isSearching) return;
    isSearching = true;
    
    showLoading(true);
    hideError();
    hideResults();
    
    try {
        const results = await fetchAllWordData(word);
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
        document.getElementById('word-input').value = 'ocean';
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
        showError('Could not get selected text.');
        console.error('Error getting selection:', error);
    }
}

async function handleSearch() {
    const input = document.getElementById('word-input');
    const word = input.value.trim().toLowerCase().replace(/[^a-zA-Z]/g, '');
    
    if (!word) {
        showError('Please enter a word');
        return;
    }
    
    isSearching = true;
    showLoading(true);
    hideError();
    hideResults();
    
    try {
        const results = await fetchAllWordData(word);
        displayResults(word, results);
    } catch (error) {
        showError('Failed to fetch. Try again.');
        console.error('Search error:', error);
    } finally {
        showLoading(false);
        isSearching = false;
    }
}

async function fetchAllWordData(word) {
    const [describingRes, synonymsRes, relatedRes, similarRes] = await Promise.all([
        fetch(`https://api.datamuse.com/words?rel_jjb=${encodeURIComponent(word)}&max=25`),
        fetch(`https://api.datamuse.com/words?rel_syn=${encodeURIComponent(word)}&max=20`),
        fetch(`https://api.datamuse.com/words?ml=${encodeURIComponent(word)}&max=15`),
        fetch(`https://api.datamuse.com/words?sl=${encodeURIComponent(word)}&max=10`)
    ]);
    
    const [describing, synonyms, related, similar] = await Promise.all([
        describingRes.json(),
        synonymsRes.json(),
        relatedRes.json(),
        similarRes.json()
    ]);
    
    const synonymWords = synonyms.map(w => w.word);
    return {
        describing: describing.map(w => w.word),
        synonyms: synonymWords,
        related: related.map(w => w.word).filter(w => !synonymWords.includes(w)),
        similar: similar.map(w => w.word).filter(w => w !== word)
    };
}

function displayResults(word, results) {
    document.getElementById('searched-word').textContent = word;
    document.getElementById('current-word').classList.remove('hidden');
    
    const hasResults = results.describing.length > 0 || results.synonyms.length > 0 || 
                       results.related.length > 0 || results.similar.length > 0;
    
    if (!hasResults) {
        document.getElementById('no-results').classList.remove('hidden');
        return;
    }
    
    if (results.describing.length > 0) {
        const list = document.getElementById('describing-list');
        list.innerHTML = results.describing.map(w => 
            `<span class="word-chip describing" data-word="${w}">${w}</span>`
        ).join('');
        document.getElementById('describing-section').classList.remove('hidden');
    }
    
    if (results.synonyms.length > 0) {
        const list = document.getElementById('synonyms-list');
        list.innerHTML = results.synonyms.map(w => 
            `<span class="word-chip synonym" data-word="${w}">${w}</span>`
        ).join('');
        document.getElementById('synonyms-section').classList.remove('hidden');
    }
    
    if (results.related.length > 0) {
        const list = document.getElementById('related-list');
        list.innerHTML = results.related.map(w => 
            `<span class="word-chip related" data-word="${w}">${w}</span>`
        ).join('');
        document.getElementById('related-section').classList.remove('hidden');
    }
    
    if (results.similar.length > 0) {
        const list = document.getElementById('similar-list');
        list.innerHTML = results.similar.map(w => 
            `<span class="word-chip similar" data-word="${w}">${w}</span>`
        ).join('');
        document.getElementById('similar-section').classList.remove('hidden');
    }
    
    document.querySelectorAll('.word-chip').forEach(chip => {
        chip.addEventListener('click', () => insertWord(chip.dataset.word));
    });
}

async function insertWord(word) {
    if (!isOfficeInitialized) {
        alert(`Word: "${word}"\n\nWorks inside MS Word.`);
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
        showError('Failed to insert word.');
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
    document.getElementById('describing-section').classList.add('hidden');
    document.getElementById('synonyms-section').classList.add('hidden');
    document.getElementById('related-section').classList.add('hidden');
    document.getElementById('similar-section').classList.add('hidden');
    document.getElementById('no-results').classList.add('hidden');
    document.getElementById('describing-list').innerHTML = '';
    document.getElementById('synonyms-list').innerHTML = '';
    document.getElementById('related-list').innerHTML = '';
    document.getElementById('similar-list').innerHTML = '';
}
