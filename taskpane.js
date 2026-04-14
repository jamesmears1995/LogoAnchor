/* global Office, PowerPoint */

// Your Brandfetch API Key
const BRANDFETCH_API_KEY = '2n6DXvTWzbmovzyQCMDWuYBYKxRKgRKbZE2oAE3MXmbco0ST4pKMWPM4vToVVFLkGOs5XqVvnLxFH2EivInzpg';

Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        console.log("LogoAnchor Engine Active.");
    }
});

async function searchBrands() {
    const query = document.getElementById('search-input').value;
    const resultsContainer = document.getElementById('results');
    
    if (!query) return;
    resultsContainer.innerHTML = '<p class="text-xs text-slate-400 animate-pulse">Searching...</p>';

    try {
        const response = await fetch(`https://api.brandfetch.io/v2/search/${query}`, {
            method: 'GET',
            headers: { 'Authorization': `Bearer ${BRANDFETCH_API_KEY}` }
        });
        const data = await response.json();
        renderResults(data);
    } catch (error) {
        resultsContainer.innerHTML = '<p class="text-xs text-red-500">Search error. Check connection.</p>';
    }
}

function renderResults(brands) {
    const container = document.getElementById('results');
    container.innerHTML = ''; 

    brands.forEach(brand => {
        const div = document.createElement('div');
        div.className = "p-3 border border-slate-100 rounded-lg hover:bg-blue-50 cursor-pointer transition flex flex-col items-center justify-center bg-white shadow-sm";
        // Using the icon or logo provided by the search
        const imgUrl = brand.icon || brand.logo || `https://img.logo.dev/${brand.domain}?token=638a1f2e`;
        
        div.onclick = () => insertLogoToSlide(imgUrl);
        
        div.innerHTML = `
            <img src="${imgUrl}" class="h-10 w-10 object-contain mb-2">
            <span class="text-[10px] font-bold text-slate-600 truncate w-full text-center">${brand.name}</span>
        `;
        container.appendChild(div);
    });
}

async function insertLogoToSlide(url) {
    try {
        const response = await fetch(url);
        const blob = await response.blob();
        const reader = new FileReader();

        reader.onloadend = () => {
            const base64Data = reader.result.split(',')[1];
            PowerPoint.run(async (context) => {
                const sheet = context.presentation.getSelectedSlides().getItemAt(0);
                sheet.insertImageFromBase64(base64Data, { height: 100, width: 100 });
                await context.sync();
            });
        };
        reader.readAsDataURL(blob);
    } catch (error) {
        console.error("Insertion error: ", error);
    }
}