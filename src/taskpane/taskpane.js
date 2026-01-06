let isOfficeInitialized = false;

// Initialize Office.js
Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        isOfficeInitialized = true;
        console.log('Office.js initialized for Word');
    }
    
    // Set up event listeners
    document.getElementById('search-btn').addEventListener('click', handleSearch);
    document.getElementById('word-input').addEventListener('keypress', (e) => {
        if (e.key === 'Enter') {
            handleSearch();
        }
    });
    document.getElementById('get-selection-btn').addEventListener('click', getSelectedWord);
    
    // Auto-detect selection changes
    if (isOfficeInitialized) {
        setInterval(checkSelection, 1500);
    }
});

// Get selected word from Word document
async function getSelectedWord() {
    if (!isOfficeInitialized) {
        // For testing outside Word, use a sample word
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
                // Get first word if multiple words selected
                const word = selectedText.split(/\s+/)[0].replace(/[^a-zA-Z]/g, '');
                if (word) {
                    document.getElementById('word-input').value = word;
                    handleSearch();
                }
            }
        });
    } catch (error) {
        showError('Could not get selected text. Please select a word in your document.');
        console.error('Error getting selection:', error);
    }
}

// Check selection periodically for auto-update
let lastCheckedWord = '';
async function checkSelection() {
    if (!isOfficeInitialized) return;
    
    try {
        await Word.run(async (context) => {
            const selection = context.document.getSelection();
            selection.load('text');
            await context.sync();
            
            const selectedText = selection.text.trim();
            if (selectedText && selectedText.length < 30) {
                const word = selectedText.split(/\s+/)[0].replace(/[^a-zA-Z]/g, '').toLowerCase();
                if (word && word !== lastCheckedWord && word.length > 1) {
                    lastCheckedWord = word;
                    document.getElementById('word-input').value = word;
                    // Optionally auto-search
                    // handleSearch();
                }
            }
        });
    } catch (error) {
        // Silently fail for background checks
    }
}

// Main search function
async function handleSearch() {
    const input = document.getElementById('word-input');
    const word = input.value.trim().toLowerCase().replace(/[^a-zA-Z]/g, '');
    
    if (!word) {
        showError('Please enter a word to search');
        return;
    }
    
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
    }
}

// Fetch synonyms from Datamuse API (free, no API key needed)
async function fetchSynonyms(word) {
    // Fetch multiple types of related words in parallel
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

// Display results
function displayResults(word, results) {
    document.getElementById('searched-word').textContent = word;
    document.getElementById('current-word').classList.remove('hidden');
    
    const hasResults = results.synonyms.length > 0 || results.related.length > 0 || results.similar.length > 0;
    
    if (!hasResults) {
        document.getElementById('no-results').classList.remove('hidden');
        return;
    }
    
    // Display synonyms
    if (results.synonyms.length > 0) {
        const synonymsList = document.getElementById('synonyms-list');
        synonymsList.innerHTML = results.synonyms.map(w => 
            `<span class="word-chip" data-word="${w}">${w}</span>`
        ).join('');
        document.getElementById('synonyms-section').classList.remove('hidden');
    }
    
    // Display related words
    if (results.related.length > 0) {
        const relatedList = document.getElementById('related-list');
        relatedList.innerHTML = results.related.map(w => 
            `<span class="word-chip" data-word="${w}">${w}</span>`
        ).join('');
        document.getElementById('related-section').classList.remove('hidden');
    }
    
    // Display similar sounding words
    if (results.similar.length > 0) {
        const similarList = document.getElementById('similar-list');
        similarList.innerHTML = results.similar.map(w => 
            `<span class="word-chip" data-word="${w}">${w}</span>`
        ).join('');
        document.getElementById('similar-section').classList.remove('hidden');
    }
    
    // Add click handlers to word chips
    document.querySelectorAll('.word-chip').forEach(chip => {
        chip.addEventListener('click', () => insertWord(chip.dataset.word));
    });
}

// Insert word into Word document
async function insertWord(word) {
    if (!isOfficeInitialized) {
        alert(`Word to insert: "${word}"\n\nNote: This feature works when running inside MS Word.`);
        return;
    }
    
    try {
        await Word.run(async (context) => {
            const selection = context.document.getSelection();
            selection.insertText(word, Word.InsertLocation.replace);
            await context.sync();
        });
    } catch (error) {
        showError('Failed to insert word. Please try again.');
        console.error('Insert error:', error);
    }
}

// UI helper functions
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
