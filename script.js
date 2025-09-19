document.addEventListener('DOMContentLoaded', () => {
    const queryInput = document.getElementById('query-input');
    const loadingSpinner = document.getElementById('loading-spinner');
    const errorMessage = document.getElementById('error-message');
    const resultsContainer = document.getElementById('results-container');
    const assumptionsTableContainer = document.getElementById('assumptions-table-container');
    const forecastTableContainer = document.getElementById('forecast-table-container');
    
    queryInput.addEventListener('keypress', (e) => {
        if (e.key === 'Enter') {
            e.preventDefault();
            handleSearch();
        }
    });

    async function handleSearch() {
        const query = queryInput.value.trim();
        if (query === '') {
            showError('Please enter a query.');
            return;
        }

        showLoading(true);
        hideElements();

        try {
            const response = await fetch('http://localhost:8000/forecast_from_prompt', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({ prompt: query })
            });

            if (!response.ok) {
                const errorData = await response.json();
                throw new Error(errorData.detail || 'Network response was not ok');
            }

            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.style.display = 'none';
            a.href = url;
            a.download = 'forecast.xlsx';
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
            
            hideElements();
            showLoading(false);

        } catch (error) {
            console.error('API Error:', error);
            showError('Failed to fetch data. Please check the backend service. Error: ' + error.message);
            showLoading(false);
        }
    }

    function showLoading(isLoading) {
        if (isLoading) {
            loadingSpinner.classList.remove('hidden');
        } else {
            loadingSpinner.classList.add('hidden');
        }
    }

    function showError(message) {
        errorMessage.textContent = message;
        errorMessage.classList.remove('hidden');
    }

    function hideElements() {
        errorMessage.classList.add('hidden');
        if (resultsContainer) {
            resultsContainer.classList.add('hidden');
        }
    }
});
