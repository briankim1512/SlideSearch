const state = {
    file: null,
    searchQuery: {
        text: '',
        title: '',
        time_range: ['', '']
    },
    searchTimeout: null,
    searchPaneVisible: false,
    resultsRange: [0, 20],
    results: [],
    selectedSlides: [],
};

const dom = {
    // Panes
    landingPane: document.getElementById('landing-pane'),
    searchPane: document.getElementById('search-pane'),
    zoomPane: document.getElementById('zoom-pane'),
    
    // Inputs
    fileInput: document.getElementById('file-input'),
    uploadZone: document.getElementById('upload-zone'),
    backButton: document.getElementById('back-button'),
    stitchButton: document.getElementById('stitch-button'),
    searchInput: document.getElementById('search-input'),
    searchQuery: document.getElementById('search-query'),
    advancedSearchButton: document.getElementById('advanced-search-button'),
    advancedSearchTitle: document.getElementById('advanced-search-title'),
    advancedSearchDateRangeStart: document.getElementById('advanced-search-date-range-start'),
    advancedSearchDateRangeEnd: document.getElementById('advanced-search-date-range-end'),

    // Views
    results: document.getElementById('results'),
    loading: document.getElementById('loading'),
    loadingText: document.getElementById('loading-text'),
    loadingSpinner: document.getElementById('loading-spinner'),
    loadingBar: document.getElementById('loading-bar'),
    loadingProgress: document.getElementById('loading-progress'),
    zoomedImage: document.getElementById('zoomed-image'),
    advancedSearchOptions: document.getElementById('advanced-search-options'),
};

function showPane(pane) {
    if (pane === 'search') {
        dom.landingPane.style.display = 'none';
        dom.searchPane.style.display = 'flex';
    } else if (pane === 'landing') {
        dom.landingPane.style.display = 'block';
        dom.searchPane.style.display = 'none';
    }
}

async function uploadPresentations() {
    const paths = await pywebview.api.pick_files();
    dom.uploadZone.style.display = 'none';
    dom.loading.style.display = 'block';

    if (!paths.length) {
        dom.uploadZone.style.display = 'block';
        dom.loading.style.display = 'none';
        return;
    }

    const ingestResults = await pywebview.api.ingest_files(paths);
    console.log(ingestResults)
    
    dom.uploadZone.style.display = 'block';
    dom.loading.style.display = 'none';

    dom.searchInput.focus();
}

function updateUploadStatus(text) {
    dom.loadingText.textContent = text;
}

function resetUploadStatus() {
    dom.loadingText.textContent = '업로드 시작 중...';
    dom.loadingProgress.style.width = '0%';
    dom.loadingSpinner.style.display = 'block';
    dom.loadingBar.style.display = 'none';
}

function updateUploadProgress(percentage) {
    dom.loadingSpinner.style.display = 'none';
    dom.loadingBar.style.display = 'block';
    dom.loadingProgress.style.width = `${percentage}%`;
}

function updateResults(results, prev_results=[]) {
    const new_hashes = new Set(results.map(result => result['hash']));
    const resultsBody = document.getElementById('results-body');
    resultsBody.innerHTML = '';
    console.log(results);

    if (prev_results.length > 0) {
        // Add prev_results to top of results if checked to new results
        prev_results.forEach(result => {
            if (state.selectedSlides.includes(result['hash']) && !new_hashes.has(result['hash'])) {
                results.unshift(result);
            }
        });
    }

    results.forEach(result => {
        const row = document.createElement('tr');
        const selectCell = document.createElement('td');
        const checkbox = document.createElement('input');
        
        checkbox.type = 'checkbox';
        checkbox.checked = state.selectedSlides.includes(result['hash']);
        checkbox.addEventListener('change', () => {
            selectSlide(result['hash']);
        });
        selectCell.appendChild(checkbox);
        row.appendChild(selectCell);

        for (const key in result) {
            const cell = document.createElement('td');
            if (key === 'image_path') {
                const img = document.createElement('img');
                img.src = result[key];
                img.alt = '미리보기 이미지';
                img.style.maxWidth = '80px';
                img.style.maxHeight = '60px';
                img.addEventListener('click', () => {
                    zoomOnPhoto(result[key]);
                });
                cell.appendChild(img);
            } else if (key === 'hash') {
                continue;
            } else if (key === 'text') {
                console.log('Adding text div for:', result[key]);
                const textDiv = document.createElement('pre');
                textDiv.className = 'result-text-pre';
                textDiv.textContent = result[key];
                cell.appendChild(textDiv);
            } else {
                cell.textContent = result[key];
            }
            row.appendChild(cell);
        }
        resultsBody.appendChild(row);
    });

    document.getElementById('result-count').textContent = results.length;
}

function zoomOnPhoto(path) {
    dom.zoomedImage.src = path;
    dom.zoomPane.style.display = 'flex';
    dom.zoomPane.addEventListener('click', () => {
        dom.zoomPane.style.display = 'none';
    });
}

function selectSlide (hash) {
    const index = state.selectedSlides.indexOf(hash);
    if (index === -1) {
        state.selectedSlides.push(hash);
    } else {
        state.selectedSlides.splice(index, 1);
    }
    document.getElementById('stitch-button').disabled = state.selectedSlides.length === 0;
    console.log('Selected slides:', state.selectedSlides);
}

function searchSlides () {
    clearTimeout(state.searchTimeout);
    state.searchQuery = {
        text: dom.searchQuery.value,
        title: dom.advancedSearchTitle.value,
        time_range: [
            dom.advancedSearchDateRangeStart.value,
            dom.advancedSearchDateRangeEnd.value
        ]
    }

    if (
        state.searchQuery.text.trim() === '' &&
        state.searchQuery.title.trim() === '' &&
        state.searchQuery.time_range[0] === '' &&
        state.searchQuery.time_range[1] === ''
    ) {
        if (state.selectedSlides.length > 0) {
            const prev_results = state.results;
            state.results = state.results.filter(r => state.selectedSlides.includes(r['hash']));
            updateResults(state.results, prev_results);
            dom.results.scrollTop = 0;
        } else {
            state.results = [];
            updateResults([]);
        }

        dom.stitchButton.disabled = state.selectedSlides.length === 0;
        return;
    }

    state.searchTimeout = setTimeout(async () => {
        const prev_results = state.results;
        const results = await pywebview.api.search_slides(state.searchQuery);
        state.results = results;
        updateResults(results, prev_results);
    }, 500);
}

function checkDateInputs() {
    if (dom.advancedSearchDateRangeStart.value && dom.advancedSearchDateRangeEnd.value) {
        if (dom.advancedSearchDateRangeStart.value > dom.advancedSearchDateRangeEnd.value) {
            dom.advancedSearchDateRangeEnd.value = dom.advancedSearchDateRangeStart.value;
        }
        searchSlides();
    }
}

function resetSearch() {
    dom.searchQuery.value = '';
    dom.advancedSearchTitle.value = '';
    dom.advancedSearchDateRangeStart.value = '';
    dom.advancedSearchDateRangeEnd.value = '';
    dom.searchInput.value = '';
    dom.searchInput.blur();
    
    state.searchQuery = {
        text: '',
        title: '',
        time_range: ['', '']
    };
    state.results = [];
    state.selectedSlides = [];
    dom.stitchButton.disabled = true;
    dom.advancedSearchOptions.style.display = 'none';
    updateResults([]);
    showPane('landing');
}

dom.uploadZone.addEventListener('click', uploadPresentations);

dom.searchInput.addEventListener('click', (event) => {
    showPane('search');
    dom.searchQuery.focus();
});

dom.backButton.addEventListener('click', resetSearch);

dom.advancedSearchButton.addEventListener('click', () => {
    if (dom.advancedSearchOptions.style.display === 'none') {
        dom.advancedSearchOptions.style.display = 'flex';
    } else {
        dom.advancedSearchDateRangeStart.value = '';
        dom.advancedSearchDateRangeEnd.value = '';
        dom.advancedSearchTitle.value = '';
        state.searchQuery.title = '';
        state.searchQuery.time_range = ['','']
        dom.advancedSearchOptions.style.display = 'none';
        searchSlides();
    }
});

dom.stitchButton.addEventListener('click', async () => {
    if (state.selectedSlides.length === 0) {
        alert('장표를 선택해 주세요')
    };
    
    dom.stitchButton.disabled = true;
    dom.stitchButton.textContent = '처리 중...';
    
    const message = await pywebview.api.stitch_slides(state.selectedSlides);
    alert(message);
    
    const checkboxes = document.querySelectorAll('#results-body input[type="checkbox"]');
    checkboxes.forEach(checkbox => {
        checkbox.checked = false;
    });
    state.selectedSlides = [];
    
    dom.results.scrollTop = 0;
    dom.stitchButton.textContent = '선택한 장표로 새 PPT 만들기';
    dom.stitchButton.disabled = false;
});

dom.searchQuery.addEventListener('input', searchSlides);
dom.advancedSearchTitle.addEventListener('input', searchSlides);
dom.advancedSearchDateRangeStart.addEventListener('change', checkDateInputs);
dom.advancedSearchDateRangeEnd.addEventListener('change', checkDateInputs);
